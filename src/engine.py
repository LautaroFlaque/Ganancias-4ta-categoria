# -*- coding: utf-8 -*-
"""
Motor de control y estimacion de retenciones del Impuesto a las Ganancias
4ta categoria (empleados en relacion de dependencia) - Regimen RG (AFIP/ARCA) 4003
con la escala del art. 94 LIG segun Ley 27.743.

Replica la logica de la hoja "Planilla Retenciones" de los liquidadores oficiales
de Ganancias 4ta categoria: calculo acumulado mes a mes, criterio de lo percibido,
deducciones del art. 30 proporcionadas a los meses transcurridos, tope del 35% de
la remuneracion bruta del mes, y "valvula de escape" por retenciones informadas
manualmente.

NO contiene datos de contribuyentes. Los parametros del periodo fiscal se cargan
desde 'params.json' (ver extraer_parametros.py).
"""

import json
import os

# ---------------------------------------------------------------------------
# Parametros del periodo fiscal (se cargan desde params.json)
# ---------------------------------------------------------------------------
_P = json.load(open(os.path.join(os.path.dirname(__file__), "params.json"), encoding="utf-8"))

# Deducciones personales art. 30 - valores MENSUALES del 1er semestre
DP_SEM1 = _P["ded_personales_sem1"]
# Deducciones personales art. 30 - valores ANUALES del 2do semestre
# (desde julio se suma 1/12 del valor anual por mes, arrancando del acumulado de junio)
DP_SEM2_ANUAL = _P["ded_personales_sem2_anual"]

# Topes mensuales a la base imponible de aportes de seguridad social (ANSES), mes por mes
TOPE_SS = _P["topes_ss"]

# Escalas mensuales ACUMULADAS del art. 94 - {ABREV_MES: [[desde, hasta(None=inf), fijo, alicuota, excedente], ...]}
SCALES = {}
for m, rows in _P["scales"].items():
    SCALES[m] = []
    for desde, hasta, fijo, alic, exc in rows:
        hasta = None if isinstance(hasta, str) else float(hasta)
        SCALES[m].append((float(desde), hasta, float(fijo), float(alic), float(exc)))

APORTE_JUB = 0.11   # jubilacion
APORTE_OS = 0.03    # obra social
APORTE_PAMI = 0.03  # Ley 19.032 / INSSJP
APORTE_TOTAL = APORTE_JUB + APORTE_OS + APORTE_PAMI

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
ABREV = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]


def escala(mes_nombre, base):
    """Aplica la escala mensual acumulada del art. 94 correspondiente al mes de pago.
    Devuelve el impuesto determinado acumulado."""
    if base <= 0:
        return 0.0
    tramos = SCALES[ABREV[MESES.index(mes_nombre)]]
    fijo, alic, exc = tramos[-1][2], tramos[-1][3], tramos[-1][4]
    for desde, hasta, f, a, e in tramos:
        if base > desde and (hasta is None or base <= hasta):
            fijo, alic, exc = f, a, e
            break
    return fijo + (base - exc) * alic


def dp_acumulada(concepto, m_pago):
    """Deduccion personal del art. 30 acumulada al mes de pago m_pago (1..12).
    1er semestre: valor mensual x meses. 2do semestre: acumulado de junio + (m-6) x (anual/12)."""
    men = DP_SEM1[concepto]
    if m_pago <= 6:
        return men * m_pago
    return men * 6 + (m_pago - 6) * (DP_SEM2_ANUAL[concepto] / 12.0)


def liquidar(datos, modalidad_mes_siguiente=True, usar_aportes_reales=True,
             fix_bug_carryover=False, sac_prorrateo=True, tope_ss=None):
    """
    Calcula, recibo por recibo, la retencion/(devolucion) de Ganancias 4ta categoria.

    datos: dict con listas alineadas (indice 0 = primer recibo del periodo). Claves:
        remun            : remuneracion bruta del recibo (habitual + no habitual gravado)
        sac              : SAC efectivamente percibido en el recibo (0 si se usa prorrateo)
        no_computable    : conceptos que restan de la base (horas extra exentas, etc.), en positivo
        ap_jub, ap_os, ap_pami : aportes reales del recibo (se usan si usar_aportes_reales=True)
        sindical         : cuota sindical del recibo
        otros_desc       : otros descuentos computables del recibo
        alq_beneficio    : importe del beneficio por alquiler 40% del recibo (no acumulado)
        ded_grales_otras : otras deducciones generales del recibo (no acumulado, ya con tope)
        cajas_compl      : aportes a cajas complementarias / SGR / profesionales del recibo (sin tope)
        conyuge, hijos, hijo_incap : cantidad de cargas de familia por recibo (0/1/n)
        ret_informada    : retencion real informada para el recibo (None = usar la teorica)
      flags:
        aplica_conyuge, aplica_hijos, aplica_hijo_incap : bool

    modalidad_mes_siguiente : True  -> el recibo del mes M se paga en M+1 (el indice i mapea a mes de pago i+2)
                              False -> el indice i ya es el mes de pago (i -> mes de pago i+1)
    usar_aportes_reales     : True  -> 17% s/ tope, salvo que el aporte real informado sea menor
                              False -> 17% s/ min(remun, tope)
    sac_prorrateo           : True  -> prorratea 1/12 de la remuneracion como SAC (Ap. C Anexo II RG 4003)
    fix_bug_carryover       : True  -> no descarta el arrastre de retenciones cuando la GNSI del mes da 0
    tope_ss                 : lista de topes ANSES alineada a los datos (default: TOPE_SS del params.json)

    Devuelve una lista de dicts (uno por recibo) con todos los subtotales.
    """
    n = len([r for r in datos["remun"] if r is not None])
    TOPE = tope_ss if tope_ss is not None else TOPE_SS
    out = []
    prev_ret_acum = 0.0
    prev_ret_real = 0.0
    acum_remun = acum_ap = acum_alq = acum_otras = acum_cajas = 0.0

    def g(clave, i, default=0.0):
        v = datos.get(clave, [default] * 12)[i]
        return default if v is None else float(v)

    for i in range(n):
        mpago = (i + 2) if modalidad_mes_siguiente else (i + 1)
        if mpago > 12:
            mpago = 12
        mes_pago_nombre = MESES[mpago - 1]

        remun = g("remun", i)
        sac = g("sac", i)
        nocomp = g("no_computable", i)

        sac_prorr = ((remun / 12.0 if sac == 0 else 0.0) if sac_prorrateo else 0.0)
        rem_mes = remun + sac + sac_prorr - nocomp
        acum_remun += rem_mes

        # --- aportes de seguridad social + obra social
        ap_tope = TOPE[i] * APORTE_TOTAL
        ap_real = g("ap_jub", i) + g("ap_os", i) + g("ap_pami", i)
        if usar_aportes_reales:
            ap_mes = ap_real if (0 < ap_real < ap_tope) else ap_tope
        else:
            ap_mes = min(remun, TOPE[i]) * APORTE_TOTAL
        ap_mes += g("sindical", i) + g("otros_desc", i)
        ap_mes += (ap_mes / 12.0 if (sac == 0 and sac_prorrateo) else 0.0)
        acum_ap += ap_mes

        # --- deducciones generales
        acum_alq += g("alq_beneficio", i)
        alq_comp = min(acum_alq, dp_acumulada("MNI", mpago))   # tope = MNI acumulado
        acum_otras += g("ded_grales_otras", i)
        acum_cajas += g("cajas_compl", i)
        ded_grales_acum = alq_comp + acum_otras + acum_cajas

        gcia_neta_acum = acum_remun - acum_ap - ded_grales_acum

        # --- deducciones personales acumuladas (a la posicion mes de pago)
        dp = dp_acumulada("MNI", mpago)
        if datos.get("aplica_conyuge"):
            dp += dp_acumulada("CONYUGE", mpago) * g("conyuge", i)
        if datos.get("aplica_hijos"):
            dp += dp_acumulada("HIJO", mpago) * g("hijos", i)
        if datos.get("aplica_hijo_incap"):
            dp += dp_acumulada("HIJO_INCAP", mpago) * g("hijo_incap", i)
        dp_esp = dp_acumulada("DED_ESP", mpago)
        doceava = (dp + dp_esp) / 12.0                          # "doceava parte" adicional
        dp_total = dp + dp_esp + doceava

        gnsi = max(0.0, gcia_neta_acum - dp_total)
        imp_det = escala(mes_pago_nombre, gnsi)

        if fix_bug_carryover:
            ret_antes = prev_ret_acum + prev_ret_real
        else:
            ret_antes = 0.0 if gnsi == 0 else (prev_ret_acum + prev_ret_real)

        if i == 0 and modalidad_mes_siguiente:
            a_retener = max(0.0, imp_det)     # el primer recibo no resta retenciones previas
        else:
            a_retener = imp_det - ret_antes   # puede ser negativo (devolucion)

        tope35 = (remun + sac) * 0.35
        if a_retener == 0:
            ret_dev = 0.0
        elif a_retener < tope35:
            ret_dev = a_retener
        else:
            ret_dev = tope35

        ri = datos.get("ret_informada", [None] * 12)[i]
        ret_real = ret_dev if (ri is None or ri == 0) else float(ri)

        out.append(dict(
            recibo=i + 1, mes_pago=mes_pago_nombre,
            remun=remun, sac=sac, sac_prorr=sac_prorr,
            rem_acum=acum_remun, ap_mes=ap_mes, ap_acum=acum_ap,
            alq_comp=alq_comp, ded_grales_acum=ded_grales_acum,
            gcia_neta_acum=gcia_neta_acum, dp_total=dp_total, gnsi=gnsi,
            imp_det=imp_det, ret_antes=ret_antes, a_retener=a_retener,
            tope35=tope35, ret_dev=ret_dev, ret_real=ret_real,
        ))
        prev_ret_acum = ret_antes
        prev_ret_real = ret_real

    return out


def fmt(x):
    """Formato de moneda estilo es-AR."""
    return f"{x:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")


def tabla(res, cols=("recibo", "mes_pago", "rem_acum", "ap_acum", "ded_grales_acum",
                     "gcia_neta_acum", "dp_total", "gnsi", "imp_det", "ret_antes",
                     "a_retener", "tope35", "ret_dev")):
    hdr = {"recibo": "Recibo", "mes_pago": "Mes pago", "rem_acum": "Rem.Bruta Acum",
           "ap_acum": "Aportes Acum", "ded_grales_acum": "Ded.Grales Acum",
           "gcia_neta_acum": "Gcia Neta Acum", "dp_total": "Ded.Pers Acum", "gnsi": "GNSI",
           "imp_det": "Imp.Determ", "ret_antes": "Ret.previas", "a_retener": "A retener/(dev)",
           "tope35": "Tope 35%", "ret_dev": "Ret/(Dev) mes"}
    w = {c: max(len(hdr[c]), 16) for c in cols}
    print(" | ".join(hdr[c].rjust(w[c]) for c in cols))
    print("-+-".join("-" * w[c] for c in cols))
    for r in res:
        print(" | ".join((str(r[c]) if isinstance(r[c], str) else fmt(r[c])).rjust(w[c]) for c in cols))
