# OMIEporra ⚡

Cada mes cada alumno apuesta el **precio medio (€/MWh) del mercado diario de OMIE (España)** del mes siguiente.
Este repositorio descarga solo los datos de OMIE todos los días, calcula la media del mes y hace el ranking.

## ¿Qué hace cada archivo?

| Archivo | Para qué sirve |
|---|---|
| `apuestas.csv` | Reserva: apuestas a mano, solo si un mes no tiene apuestas en el formulario. |
| `omieporra.py` | El programa que descarga OMIE y calcula todo. |
| `requirements.txt` | Librerías de Python que necesita el programa. |
| `.github/workflows/omieporra.yml` | La "tarea automática" de GitHub Actions: se ejecuta cada día a las 14:00 UTC (15:00 en invierno / 16:00 en verano, hora de España). |
| `datos/` | Ficheros originales de OMIE descargados (uno por día). |
| `omieporra_AAAA-MM.xlsx` / `.png` | Los resultados de cada mes. |
| `resultados/AAAA-MM.json` | Resumen de cada mes que usa la web (historial). |
| `plantilla_web.html` | El diseño de la página web. |

## Cómo apuesta cada compañero (Google Forms)

Cada uno rellena el formulario de Google con su **nombre**, el **mes** por el que apuesta y su **predicción (€/MWh)**.
El programa lee las respuestas solo, cada día. Reglas:
- Una apuesta para un mes **solo cuenta si se envía antes del día 1** de ese mes.
- Si alguien apuesta varias veces para el mismo mes, **cuenta la última**.
- En la web se ven ya las apuestas para el mes siguiente (nombre y cifra), comparadas con la media del mes actual y la de los últimos 7 días.

### Crear el formulario (una sola vez)
1. En [Google Forms](https://forms.google.com) crea un formulario con 3 preguntas obligatorias:
   - **Nombre** (respuesta corta)
   - **Mes** (desplegable con opciones como `Noviembre 2026`, `Diciembre 2026`...)
   - **Predicción (€/MWh)** (respuesta corta; en *⋮ → Validación de respuestas* elige *Número*)
2. Pestaña **Respuestas → Vincular con Hojas de cálculo** → *Crear una hoja de cálculo nueva*.
3. En la hoja: **Archivo → Compartir → Publicar en la Web** → elige la hoja de respuestas y el formato **Valores separados por comas (.csv)** → **Publicar** → copia el enlace.
4. Pon ese enlace en `.github/workflows/omieporra.yml`, en la línea `APUESTAS_URL: ""` (entre las comillas).
5. Pon el enlace para **rellenar** el formulario (botón *Enviar → 🔗* del formulario) en la línea `FORMULARIO_URL: ""`. Así la web muestra el formulario dentro de la página (botón **Apostar ahora**) mientras las apuestas están abiertas.

Cada mes solo tienes que añadir al desplegable "Mes" el mes siguiente.

> `apuestas.csv` sigue funcionando como reserva: se usa si no hay enlace del formulario o si el formulario no tiene apuestas para ese mes.

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

## Página web para la clase 🌐

Cada ejecución publica el ranking y la gráfica en GitHub Pages:
**https://laanayamahdi3-maker.github.io/omieporra/**

Para que funcione (una sola vez):
1. *Settings → General →* abajo del todo, *Danger Zone → Change visibility →* **Public** (GitHub Pages gratis solo funciona en repositorios públicos).
2. *Settings → Pages → Build and deployment → Source:* **GitHub Actions**.
3. *Actions → OMIEporra → Run workflow* para publicarla la primera vez.
