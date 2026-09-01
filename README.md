# Ganancias 4ª categoría — control y estimación de retenciones

Herramienta para **reconstruir, controlar y proyectar** la retención del Impuesto a las
Ganancias sobre sueldos (4ª categoría, empleados en relación de dependencia) bajo el
**Régimen de Retención RG (AFIP/ARCA) 4003**, con la escala del **art. 94 LIG** según
**Ley 27.743**.

El motor replica la lógica de un papel de trabajo de Ganancias 4ª categoría y permite:

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

`params.json` trae los valores de **2026**. Para otro año se necesita el papel de trabajo de
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

### Validación

- Los parámetros de `params.json` se verificaron contra los **PDFs oficiales de ARCA**
  (`normativa/`): los acumulados mensuales de deducciones del art. 30 y la escala del
  art. 94 coinciden con lo publicado para el período jul-dic 2026 y para la liquidación
  anual 2026.

---

## Datos de entrada de un caso real

Todos van a una carpeta privada, **nunca al repo**. Cada fuente alimenta una parte del CSV:

| Fuente | Origen | Qué aporta | Columnas del CSV |
|---|---|---|---|
| Histórico **"Mis Aportes"** | Portal ARCA (Clave Fiscal) | Remuneración bruta y aportes (jubilación / obra social) mes a mes | `remun`, `aporte_real` |
| **"Mis Retenciones"** (SICORE) | Portal ARCA | Retenciones y devoluciones reales por recibo, con fecha | `ret_informada` |
| **Recibos de sueldo** | Cliente / empleador | Separar básico, SAC y bonos; fecha de pago (define el primer recibo del período) | `remun`, `sac`, `no_computable` |
| Presentaciones de **SiRADIG (F.572)** | Portal ARCA | Deducciones declaradas y desde cuándo: alquiler, cargas de familia, aportes a SGR, cuota médica, etc. | `alq_beneficio`, `cajas_compl`, `ded_grales_otras`, `conyuge` / `hijos` / `hijo_incap` |
| Liquidador oficial de Ganancias 4ª | Publicación anual | Sólo si hay que regenerar `params.json` para un período nuevo | — |

**Mínimo indispensable:** Mis Aportes + Mis Retenciones + SiRADIG cubren la comparación
teórico vs. real. Los recibos afinan SAC, bonos y fechas.

### Pasos

1. Reunir las fuentes en la carpeta privada.
2. `python extraer_parametros.py <liquidador.xlsm>` si cambió el período fiscal.
3. Armar el CSV (un recibo por fila; criterio percibido: el primer recibo del período es el
   sueldo de diciembre anterior cobrado en enero).
4. `python armar_planilla.py <caso.csv> <control.xlsx>` → planilla *Resumen* / *Datos* / *Parámetros*.
5. Analizar: la diferencia acumulada entre impuesto determinado y retención real es el saldo
   estimado de la DDJJ; se contrasta contra el SiRADIG para explicar cada retención o devolución.

---

## Roadmap

Hoy los tres archivos de ARCA (Mis Aportes, Mis Retenciones, SiRADIG) se descargan **a mano**
desde el portal. El objetivo es **automatizar esa descarga**, dejando como entrada manual sólo
lo que aporta el cliente (recibos de sueldo y respaldo de deducciones).

- ARCA **no expone API pública** para los servicios a nivel contribuyente (Mis Aportes, Mis
  Retenciones, SiRADIG). Sí hay web services oficiales para padrón / constancia de inscripción
  y facturación electrónica (autenticación WSAA con certificado), que no cubren estos datos.
- Vía prevista: **scraper autenticado** con Clave Fiscal (Playwright / Selenium) que reproduzca
  la descarga de los mismos `.xls` / PDF, contemplando el 2FA y los términos de uso del portal.
  El contador accede con la relación delegada del cliente ("Administrador de Relaciones").
- Paso siguiente: **parsers** que normalicen esas descargas directo al formato del CSV, para
  que armar un caso sea "pegar las credenciales → planilla".

