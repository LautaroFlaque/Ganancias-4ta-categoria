# Ganancias 4ª categoría — control y estimación de retenciones

Herramienta para **reconstruir, controlar y proyectar** la retención del Impuesto a las
Ganancias sobre sueldos (4ª categoría, empleados en relación de dependencia) bajo el
**Régimen de Retención RG (AFIP/ARCA) 4003**, con la escala del **art. 94 LIG** según
**Ley 27.743**.

El motor replica la lógica de la hoja *"Planilla Retenciones"* de los liquidadores
oficiales de Ganancias 4ª categoría y permite:

- reconstruir mes a mes cuánto **debió** retenerse (retención teórica),
- compararlo contra lo que el empleador **efectivamente** retuvo (retención real),
- detectar **sub-retención / sobre-retención** y su causa (típicamente deducciones mal
  declaradas en el SiRADIG / F.572),
- estimar el **saldo de la DDJJ / liquidación anual**,
- correr **escenarios** (agregar/quitar deducciones, cambiar la modalidad de pago, etc.).


---

## Contenido

```
src/
  engine.py             Motor de cálculo (RG 4003 / art. 94 Ley 27.743). Sin dependencias más allá de la stdlib.
  params.json           Parámetros del período fiscal: escala mensual acumulada, deducciones art. 30, topes ANSES.
  extraer_parametros.py Extrae params.json desde un liquidador oficial (.xlsm) y verifica contra los PDFs de ARCA.
  armar_planilla.py     Arma una planilla de control (.xlsx) desde un CSV de datos por recibo.
ejemplo/
  datos_ejemplo.csv     Datos SINTÉTICOS (no corresponden a ningún contribuyente).
  correr_ejemplo.py     Corre el ejemplo e imprime la liquidación + genera un .xlsx.
normativa/
  Deducciones-personales-art-30-*.pdf   Publicaciones oficiales de ARCA (deducciones art. 30, 2º sem. y anual 2026).
  Tabla-Art-94-LIG-*.pdf                Publicaciones oficiales de ARCA (escala art. 94, 2º sem. y anual 2026).
```

---

## Empezar (paso a paso)

### Requisitos

- **Python 3.10 o superior** (`python --version`). En Windows, instalá desde
  [python.org](https://www.python.org/downloads/) marcando *"Add Python to PATH"*.
- Git (sólo si vas a clonar en vez de bajar el ZIP).

### 1. Obtener el código

```bash
git clone https://github.com/LautaroFlaque/Ganancias-4ta-categoria.git
cd Ganancias-4ta-categoria
```

*(o desde GitHub: botón verde **Code → Download ZIP** y descomprimir.)*

### 2. Instalar las dependencias

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

pip install -r requirements.txt
```

*(El `venv` es opcional pero recomendado. Sin él, `pip install -r requirements.txt` a secas también funciona.)*

### 3. Probar que anda (ejemplo con datos sintéticos)

```bash
cd ejemplo
python correr_ejemplo.py
```

Debería imprimir una tabla de liquidación de 12 recibos y generar `ejemplo/ejemplo_salida.xlsx`.
Los datos del ejemplo son inventados (`datos_ejemplo.csv`): sueldo 6.000.000/mes, alquiler 40 %
de 300.000/mes, SAC en junio y diciembre.

### 4. Usarlo con un caso propio

1. Creá una carpeta para tus datos (git la ignora):
   ```bash
   mkdir privado
   ```
2. Copiá `ejemplo/datos_ejemplo.csv` a `privado/mi_caso.csv` y completá **una fila por recibo**
   (delimitador `;`, números en formato `1.234.567,89` o `1234567.89`). Columnas:

   | columna | qué va | vacío = |
   |---|---|---|
   | `remun` | remuneración bruta del recibo (sin SAC) | 0 |
   | `sac` | SAC efectivamente cobrado en ese recibo | 0 |
   | `no_computable` | conceptos exentos / no computables (en positivo) | 0 |
   | `aporte_real` | aporte jub+PAMI+OS real del recibo | usa 17 % s/ tope |
   | `sindical` / `otros_desc` | otros descuentos computables | 0 |
   | `alq_beneficio` | beneficio alquiler 40 % de ese recibo | 0 |
   | `ded_grales_otras` | otras deducciones generales del SiRADIG (ya con tope) | 0 |
   | `cajas_compl` | incremento de aportes a SGR / cajas de ese recibo | 0 |
   | `conyuge` / `hijos` / `hijo_incap` | cantidad de cargas de familia | 0 |
   | `ret_informada` | retención real de ARCA de ese recibo | usa la teórica |

3. Generá la planilla:
   ```bash
   cd ../src
   python armar_planilla.py  ../privado/mi_caso.csv  ../privado/control.xlsx  --mes-en-curso
   ```
   Opciones: `--mes-siguiente` (el recibo del mes M se paga en M+1), `--con-prorrateo-sac`,
   `--aportes-tope`, `--conyuge` / `--hijos` / `--hijo-incap` (habilitan esas cargas).

4. Abrí `privado/control.xlsx`. Hoja **Resumen**: recibo por recibo, retención **teórica** vs
   **real** (columna `ret_informada`) y la diferencia. La suma de esa diferencia es el saldo
   estimado de la liquidación anual / DDJJ.

### 5. Para otro período fiscal

`params.json` trae los valores de **2026**. Para otro año necesitás el liquidador oficial de
Ganancias 4ª de ese período y:

```bash
cd src
python extraer_parametros.py  "ruta/al/Liquidador_Ganancias_4ta_<año>.xlsm"
```

Eso reescribe `params.json` y lo verifica contra los PDFs de `normativa/` (reemplazá también
esos PDFs por los del período nuevo).

---

## Uso del motor desde Python

```python
import sys; sys.path.insert(0, "src")
from engine import liquidar, tabla

datos = dict(
    remun=[6_000_000]*12,          # remuneración bruta de cada recibo
    alq_beneficio=[300_000]*12,    # beneficio alquiler 40% (Art. 85 inc. h)
    ret_informada=[None]*12,       # retención real de cada recibo (None = usar la teórica)
    # ... resto de campos: ver el docstring de liquidar()
)
res = liquidar(datos, modalidad_mes_siguiente=False, sac_prorrateo=False)
tabla(res)                          # imprime la tabla
print(res[7]["imp_det"])            # impuesto determinado acumulado al 8º recibo
```

Cada elemento de `res` es un `dict` con todos los subtotales de ese recibo.

---

## Cómo funciona el régimen RG 4003 (resumen)

1. **Quién retiene.** El empleador es agente de retención. Con varios empleadores, retiene
   el que paga la mayor remuneración.
2. **Criterio.** Rentas de 4ª categoría → **percibido** (art. 18 LIG). El período fiscal
   abarca todo lo *puesto a disposición* entre el 1/1 y el 31/12. El sueldo de diciembre
   cobrado en enero del año siguiente pertenece a ese año siguiente; el aguinaldo cobrado
   en diciembre pertenece al año en curso.
3. **Base mensual (acumulada).** Cada recibo, sobre la remuneración bruta acumulada del
   año se restan:
   - aportes obligatorios: jubilación 11 % + INSSJP/Ley 19.032 3 % + obra social 3 %, con
     tope sobre la base imponible máxima de ANSES;
   - **deducciones generales del SiRADIG (F.572 web)**: alquiler 40 % (tope = ganancia no
     imponible), cuota médico-asistencial, primas de seguro de vida, servicio doméstico,
     donaciones (5 %), intereses de créditos hipotecarios, aportes a SGR / cajas, cuota
     sindical, etc. — cada una con su tope;
   - **deducciones personales del art. 30** (ganancia no imponible, cargas de familia,
     deducción especial), **proporcionadas a los meses transcurridos** (N/12 del valor
     anual), actualizadas por IPC en enero y julio;
   - una **doceava parte adicional** de la deducción especial (mecanismo del SAC).
4. **Impuesto y retención.**
   ```
   Ganancia neta sujeta a impuesto acumulada  →  escala art. 94 (5 %–35 %)  →  impuesto determinado acumulado
   Retención del mes = impuesto determinado acumulado  −  retenciones ya practicadas en el año
   ```
   El resultado se **topea al 35 % de la remuneración bruta del mes** (lo que excede se
   traslada al mes siguiente). Si da negativo, el empleador **devuelve** en ese recibo.
5. **Liquidación anual (F.1359).** Hasta el 30/4 del año siguiente el empleador recalcula
   el año con valores definitivos y retiene/devuelve la diferencia.
6. **DDJJ del empleado.** Si queda saldo que el empleador no alcanzó a retener, o el
   contribuyente tiene otras rentas o está observado, presenta su propia DDJJ y paga el
   saldo, con intereses resarcitorios desde el vencimiento.

**Punto clave:** la retención mensual sólo es tan buena como los datos del SiRADIG. Si el
SiRADIG está inflado con deducciones sin respaldo, la retención sale baja y el ajuste
explota en la liquidación anual o en la DDJJ.

---

## Metodología del motor (`engine.py`)

El motor reproduce, columna por columna, la hoja *"Planilla Retenciones"* de un liquidador
oficial de Ganancias 4ª categoría:

| Concepto | Implementación |
|---|---|
| Remuneración bruta computable | `remun + SAC (+ prorrateo 1/12 opcional) − conceptos no computables` |
| Aportes | 17 % sobre el **tope ANSES** del mes; si el aporte real informado es menor, se usa el real. También `min(remun, tope)·17 %` como alternativa. |
| Alquiler 40 % | acumulado, topeado a la ganancia no imponible acumulada del mes de pago |
| Deducciones personales art. 30 | 1er semestre: valor mensual × meses. 2º semestre: acumulado de junio + (m−6) × (valor anual / 12) |
| "Doceava parte" | `(MNI + cargas + deducción especial acumuladas) / 12` |
| Escala art. 94 | la del **mes de pago**, mensual acumulada (12 tablas distintas por IPC) |
| Retención del mes | `impuesto determinado acumulado − retenciones previas`, topeado al 35 % de la remuneración bruta del recibo |
| Retenciones informadas | override manual por recibo (la "válvula de escape" de la fila 192 del liquidador) |
| Modalidad de pago | `mes en curso` (índice = mes de pago) o `mes siguiente` (el recibo del mes M se paga en M+1) |

### Validación

- El motor se contrastó **al centavo** contra el primer mes de un liquidador oficial
  (ganancia neta acumulada, deducciones personales, ganancia neta sujeta a impuesto).
- Los parámetros de `params.json` se verificaron contra los **PDFs oficiales de ARCA**
  (`normativa/`): los acumulados mensuales de deducciones del art. 30 y la escala del
  art. 94 coinciden con lo publicado para el período jul-dic 2026 y para la liquidación
  anual 2026.

### Particularidades replicadas del liquidador oficial

- El **primer recibo del período** no resta retenciones previas (columna D del liquidador).
- Con `modalidad_mes_siguiente`, cada recibo se liquida a la posición del **mes de pago**
  (deducciones y escala de M+1), lo que adelanta 1/12 de deducciones respecto de los
  meses de ingreso acumulados — es la contrapartida del criterio percibido.
- Hay una opción (`fix_bug_carryover`) para no descartar el arrastre de retenciones
  cuando la ganancia neta sujeta a impuesto de un mes da 0 (comportamiento que algunos
  liquidadores traen y que "pierde" la pista de lo retenido).

---

## Flujo de trabajo típico para un caso real (fuera de git)

1. **Reunir fuentes** (van a una carpeta privada, nunca al repo):
   - histórico *"Mis Aportes"* de ARCA → remuneración bruta y aportes devengados mes a mes;
   - *"Mis Retenciones"* (IMP_RET) de ARCA → retenciones/devoluciones reales por recibo;
   - recibos de sueldo → separar básico / SAC / bonos, confirmar fechas de pago;
   - presentaciones de **SiRADIG (F.572)** → qué deducciones declaró el empleado y desde cuándo;
   - liquidador oficial de Ganancias 4ª del período.
2. **Extraer parámetros** del liquidador → `python extraer_parametros.py <liquidador.xlsm>`.
3. **Armar el CSV** del caso (un recibo por fila, criterio percibido: el primer recibo del
   período es el sueldo de diciembre anterior cobrado en enero).
4. **Correr** `armar_planilla.py` → planilla con *Resumen* (teórico vs real), *Datos* y *Parámetros*.
5. **Analizar**: la diferencia acumulada entre impuesto determinado y retención real es el
   saldo estimado de la DDJJ; contrastar contra las deducciones del SiRADIG para explicar
   cada devolución/retención.

---

