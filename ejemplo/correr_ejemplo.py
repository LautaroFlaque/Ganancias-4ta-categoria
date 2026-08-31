# -*- coding: utf-8 -*-
"""
Ejemplo minimo, con datos SINTETICOS (no corresponden a ningun contribuyente).

Corre el motor sobre 'datos_ejemplo.csv' e imprime la tabla de liquidacion,
y ademas genera 'ejemplo_salida.xlsx'.

    python correr_ejemplo.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from engine import liquidar, tabla, fmt          # noqa: E402
from armar_planilla import leer_csv, armar        # noqa: E402

AQUI = os.path.dirname(__file__)
CSV = os.path.join(AQUI, "datos_ejemplo.csv")
XLSX = os.path.join(AQUI, "ejemplo_salida.xlsx")

datos, n = leer_csv(CSV)
res = liquidar(datos, modalidad_mes_siguiente=False, sac_prorrateo=False)

print(f"Ejemplo: {n} recibos, remuneracion ~6.000.000/mes, alquiler 40% 300.000/mes, SAC en jun y dic.\n")
tabla(res)
print(f"\nRetencion total del periodo (teorica): {fmt(sum(r['ret_dev'] for r in res))}")

armar(CSV, XLSX, mes_siguiente=False, prorrateo_sac=False)
