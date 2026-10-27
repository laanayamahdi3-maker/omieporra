# OMIEporra ⚡

Cada mes cada alumno apuesta el **precio medio (€/MWh) del mercado diario de OMIE (España)** del mes siguiente.
Este repositorio descarga solo los datos de OMIE todos los días, calcula la media del mes y hace el ranking.

## ¿Qué hace cada archivo?

| Archivo | Para qué sirve |
|---|---|
| `apuestas.csv` | Las apuestas del mes (lo cambias tú cada mes). |
| `omieporra.py` | El programa que descarga OMIE y calcula todo. |
| `requirements.txt` | Librerías de Python que necesita el programa. |
| `.github/workflows/omieporra.yml` | La "tarea automática" de GitHub Actions: se ejecuta cada día a las 14:00 UTC (15:00 en invierno / 16:00 en verano, hora de España). |
| `datos/` | Ficheros originales de OMIE descargados (uno por día). |
| `omieporra_AAAA-MM.xlsx` / `.png` | Los resultados de cada mes. |

## Cada mes: actualizar `apuestas.csv`

Hazlo **antes del día 1** del mes de la apuesta.

1. Exporta las respuestas de Microsoft Forms a Excel y guárdalas como **CSV (delimitado por punto y coma)**.
2. Deja solo dos columnas, con esta primera fila exacta:

   ```
   nombre;prediccion
   Ana;75,5
   Luis;82
   ```
   (Puedes usar coma decimal, como en Excel en español.)
3. En GitHub, abre `apuestas.csv` → icono del lápiz ✏️ (o *Add file → Upload files* y sube tu archivo con el mismo nombre) → **Commit changes**.

> Los resultados del mes anterior no se pierden: su `omieporra_AAAA-MM.xlsx` y `.png` se quedan guardados en el repositorio.

## Ver los resultados

- En la página principal del repositorio abre `omieporra_AAAA-MM.png` (se ve directamente) o descarga `omieporra_AAAA-MM.xlsx`
  (hoja **Ranking** con la gráfica, y hoja **Datos OMIE** con el precio de cada día).
- Puesto 1 = la apuesta más cercana a la media acumulada del mes. La clasificación es "provisional" hasta que termina el mes.

## Ejecutarlo a mano

Pestaña **Actions** → **OMIEporra** → botón **Run workflow**.
Si dejas el campo vacío calcula el mes actual; si escribes `2026-10` calcula ese mes.

## Si sale en rojo ❌

La ejecución falla a propósito (y GitHub te avisa por correo) cuando:
- OMIE no responde o bloquea la descarga,
- falta el precio de un día que ya debería estar publicado (no hay datos nuevos),
- `apuestas.csv` tiene un formato incorrecto (p. ej. una apuesta que no es un número).

Pulsa sobre la ejecución en rojo → paso con la ❌ → al final del texto verás un mensaje `ERROR: ...` explicando el motivo.

## Configuración inicial (una sola vez)

*Settings → Actions → General → Workflow permissions →* **Read and write permissions** → *Save*.
