"""OMIEporra: compara las apuestas de la clase con el precio medio del mercado
diario de OMIE (España) del mes.

Uso:
    python omieporra.py            # mes actual
    python omieporra.py 2026-10    # un mes concreto

Genera:
    omieporra_AAAA-MM.xlsx  (hojas "Ranking" y "Datos OMIE")
    omieporra_AAAA-MM.png   (gráfica)
    web/index.html          (página para GitHub Pages)

Si OMIE no responde, si un fichero tiene un formato raro o si faltan datos que
ya deberían estar publicados, el script termina con error (código 1) y explica
el motivo, para que GitHub Actions marque la ejecución en rojo.
"""

import calendar
import html
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")  # sin ventana: necesario en GitHub Actions
import matplotlib.pyplot as plt
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

URL_OMIE = (
    "https://www.omie.es/es/file-download?parents%5B0%5D=marginalpdbc"
    "&filename=marginalpdbc_{fecha}.1"
)
CARPETA_DATOS = Path("datos")
FICHERO_APUESTAS = Path("apuestas.csv")
ZONA = ZoneInfo("Europe/Madrid")
INTENTOS = 4


class ErrorOmie(Exception):
    """Error que se muestra al usuario tal cual y para el programa."""


# ---------------------------------------------------------------- descarga


def leer_fichero_omie(texto, dia):
    """Convierte el texto de un marginalpdbc en una lista de precios de España.

    Formato esperado:
        MARGINALPDBC;
        2026;10;26;1;63.42;63.42;
        ...
        *
    Columnas: año;mes;día;periodo;precio_PT;precio_ES
    Acepta cualquier número de periodos (24 horarios, 96 cuartohorarios,
    92/100 en los días de cambio de hora...).
    """
    lineas = [l.strip() for l in texto.strip().splitlines() if l.strip()]
    if not lineas or not lineas[0].upper().startswith("MARGINALPDBC"):
        raise ErrorOmie(
            f"El fichero de OMIE del {dia} no empieza por 'MARGINALPDBC' "
            f"(empieza por: {lineas[0][:60]!r} si hay algo). "
            "Puede que OMIE haya cambiado el formato o devuelto una página de error."
        )
    precios = []
    for linea in lineas[1:]:
        if linea.startswith("*"):
            break
        campos = linea.rstrip(";").split(";")
        if len(campos) < 6:
            raise ErrorOmie(f"Línea con menos de 6 columnas en el fichero del {dia}: {linea!r}")
        anio, mes, d = int(campos[0]), int(campos[1]), int(campos[2])
        if date(anio, mes, d) != dia:
            raise ErrorOmie(f"El fichero del {dia} contiene datos de otra fecha: {linea!r}")
        precios.append(float(campos[5].replace(",", ".")))
    if not precios:
        raise ErrorOmie(f"El fichero de OMIE del {dia} no tiene ningún periodo.")
    return precios


def descargar_dia(dia, sesion):
    """Devuelve el texto del fichero del día, o None si OMIE aún no lo publica.

    Usa la carpeta datos/ como caché. Lanza ErrorOmie si OMIE no responde.
    """
    nombre = f"marginalpdbc_{dia:%Y%m%d}.1"
    cache = CARPETA_DATOS / nombre
    if cache.exists():
        return cache.read_text(encoding="latin-1")

    url = URL_OMIE.format(fecha=f"{dia:%Y%m%d}")
    ultimo_error = None
    for intento in range(1, INTENTOS + 1):
        try:
            r = sesion.get(url, timeout=30)
        except requests.RequestException as e:
            ultimo_error = f"sin respuesta ({e.__class__.__name__}: {e})"
        else:
            if r.status_code == 404:
                return None
            if r.status_code == 200:
                texto = r.content.decode("latin-1")
                if not texto.strip():
                    return None  # OMIE a veces devuelve vacío si no existe
                if not texto.lstrip().upper().startswith("MARGINALPDBC"):
                    # Página HTML en vez del fichero: lo tratamos como "no publicado"
                    if "<html" in texto[:500].lower():
                        return None
                leer_fichero_omie(texto, dia)  # valida antes de guardar
                CARPETA_DATOS.mkdir(exist_ok=True)
                cache.write_text(texto, encoding="latin-1")
                print(f"  descargado {nombre}")
                return texto
            ultimo_error = f"HTTP {r.status_code}"
        if intento < INTENTOS:
            espera = 2**intento
            print(f"  {nombre}: {ultimo_error}; reintento en {espera} s")
            time.sleep(espera)
    raise ErrorOmie(f"OMIE no responde al pedir {nombre}: {ultimo_error}\nURL: {url}")


def datos_del_mes(anio, mes, hoy):
    """Tabla diaria (fecha, periodos, precio medio) del mes, hasta donde haya datos."""
    ultimo_dia_mes = date(anio, mes, calendar.monthrange(anio, mes)[1])
    # OMIE publica sobre las 13:00 el precio del día siguiente.
    hasta = min(ultimo_dia_mes, hoy + timedelta(days=1))
    if hasta < date(anio, mes, 1):
        raise ErrorOmie(f"El mes {anio}-{mes:02d} todavía no ha empezado.")

    filas = []
    sesion = requests.Session()
    sesion.headers["User-Agent"] = "omieporra (GitHub Actions, uso educativo)"
    dia = date(anio, mes, 1)
    while dia <= hasta:
        texto = descargar_dia(dia, sesion)
        if texto is None:
            # Solo es normal que falte el día de mañana.
            if dia <= hoy:
                raise ErrorOmie(
                    f"No hay datos de OMIE para el {dia:%d/%m/%Y}, y ya deberían estar "
                    "publicados. Revisa si la web de OMIE funciona o si ha cambiado la URL:\n"
                    + URL_OMIE.format(fecha=f"{dia:%Y%m%d}")
                )
            print(f"  {dia}: aún no publicado (normal si es mañana)")
        else:
            precios = leer_fichero_omie(texto, dia)
            filas.append(
                {
                    "Fecha": dia,
                    "Periodos": len(precios),
                    "Precio medio diario (€/MWh)": round(sum(precios) / len(precios), 2),
                    "Mínimo (€/MWh)": min(precios),
                    "Máximo (€/MWh)": max(precios),
                }
            )
        dia += timedelta(days=1)

    tabla = pd.DataFrame(filas)
    tabla["Media acumulada (€/MWh)"] = (
        tabla["Precio medio diario (€/MWh)"].expanding().mean().round(2)
    )
    return tabla


# ---------------------------------------------------------------- apuestas


def leer_apuestas():
    if not FICHERO_APUESTAS.exists():
        raise ErrorOmie("No encuentro apuestas.csv en la carpeta del repositorio.")
    for codificacion in ("utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(FICHERO_APUESTAS, sep=";", dtype=str, encoding=codificacion)
            break
        except UnicodeDecodeError:
            continue
    df.columns = [c.strip().lower() for c in df.columns]
    if not {"nombre", "prediccion"} <= set(df.columns):
        raise ErrorOmie(
            "apuestas.csv debe tener las columnas 'nombre;prediccion'. "
            f"Columnas encontradas: {list(df.columns)}"
        )
    df = df[["nombre", "prediccion"]].dropna(how="all")
    df["nombre"] = df["nombre"].str.strip()
    limpio = (
        df["prediccion"].str.replace("€", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["prediccion"] = pd.to_numeric(limpio, errors="coerce")
    malas = df[df["prediccion"].isna()]
    if not malas.empty:
        raise ErrorOmie(
            "Hay apuestas que no son un número en apuestas.csv: "
            + ", ".join(malas["nombre"].fillna("(sin nombre)"))
        )
    if df.empty:
        raise ErrorOmie("apuestas.csv no tiene ninguna apuesta.")
    return df


def ranking(apuestas, media):
    r = apuestas.rename(columns={"nombre": "Nombre", "prediccion": "Apuesta (€/MWh)"})
    r["Media OMIE (€/MWh)"] = round(media, 2)
    r["Diferencia (€/MWh)"] = (r["Apuesta (€/MWh)"] - media).round(2)
    r["Error absoluto"] = r["Diferencia (€/MWh)"].abs()
    r = r.sort_values("Error absoluto").reset_index(drop=True)
    r.insert(0, "Puesto", r["Error absoluto"].rank(method="min").astype(int))
    return r.drop(columns="Error absoluto")


# ---------------------------------------------------------------- salidas


def grafica(tabla, apuestas, etiqueta_mes, fichero):
    fig, ax = plt.subplots(figsize=(11, 6))
    fechas = pd.to_datetime(tabla["Fecha"])
    ax.plot(fechas, tabla["Precio medio diario (€/MWh)"], marker="o", ms=4,
            color="#9aa5b1", label="Precio medio diario")
    ax.plot(fechas, tabla["Media acumulada (€/MWh)"], lw=3, color="#d1495b",
            label="Media acumulada del mes")

    anio, mes = map(int, etiqueta_mes.split("-"))
    inicio = pd.Timestamp(anio, mes, 1)
    fin = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1])
    for _, fila in apuestas.iterrows():
        ax.axhline(fila["prediccion"], color="#2e86ab", lw=0.8, ls="--", alpha=0.7)
        ax.annotate(f"{fila['nombre']} ({fila['prediccion']:g})",
                    xy=(fin, fila["prediccion"]), xytext=(4, 0),
                    textcoords="offset points", va="center", fontsize=8, color="#2e86ab")

    ax.set_xlim(inicio - pd.Timedelta(days=0.5), fin + pd.Timedelta(days=0.5))
    ax.set_title(f"OMIEporra {etiqueta_mes} · media acumulada: "
                 f"{tabla['Media acumulada (€/MWh)'].iloc[-1]:.2f} €/MWh")
    ax.set_ylabel("€/MWh")
    ax.grid(alpha=0.3)
    ax.legend(loc="upper left")
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(fichero, dpi=120, bbox_inches="tight")
    plt.close(fig)


def excel(tabla_ranking, tabla_datos, imagen, fichero):
    with pd.ExcelWriter(fichero, engine="openpyxl") as w:
        tabla_ranking.to_excel(w, sheet_name="Ranking", index=False)
        tabla_datos.to_excel(w, sheet_name="Datos OMIE", index=False)
    libro = load_workbook(fichero)
    for hoja in libro.worksheets:
        for col in hoja.columns:
            ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            hoja.column_dimensions[col[0].column_letter].width = ancho + 2
    img = XLImage(str(imagen))
    img.width, img.height = img.width * 0.6, img.height * 0.6
    libro["Ranking"].add_image(img, "H2")
    libro.save(fichero)


def pagina_web(tabla_ranking, tabla_datos, etiqueta, carpeta=Path("web")):
    """Crea web/index.html (para GitHub Pages) con el ranking y la gráfica del mes,
    y copia ahí las gráficas y Excels de todos los meses."""
    carpeta.mkdir(exist_ok=True)
    meses = sorted({p.stem.split("_")[1] for p in Path(".").glob("omieporra_*.png")}, reverse=True)
    for mes in meses:
        for ext in ("png", "xlsx"):
            origen = Path(f"omieporra_{mes}.{ext}")
            if origen.exists():
                (carpeta / origen.name).write_bytes(origen.read_bytes())

    def num(x):
        return f"{x:.2f}".replace(".", ",")

    filas = "\n".join(
        f"<tr><td>{f['Puesto']}</td><td>{html.escape(str(f['Nombre']))}</td>"
        f"<td>{num(f['Apuesta (€/MWh)'])}</td><td>{num(f['Diferencia (€/MWh)'])}</td></tr>"
        for _, f in tabla_ranking.iterrows()
    )
    otros = "\n".join(
        f'<li>{m}: <a href="omieporra_{m}.png">gráfica</a> · <a href="omieporra_{m}.xlsx">Excel</a></li>'
        for m in meses
    )
    media = tabla_datos["Media acumulada (€/MWh)"].iloc[-1]
    ultimo = tabla_datos["Fecha"].iloc[-1]
    actualizado = datetime.now(ZONA).strftime("%d/%m/%Y %H:%M")
    (carpeta / "index.html").write_text(f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OMIEporra {etiqueta}</title>
<style>
body {{ font-family: system-ui, sans-serif; max-width: 960px; margin: 0 auto; padding: 16px; color: #1f2933; background: #fff; }}
h1 {{ margin-bottom: 0; }} .sub {{ color: #616e7c; margin-top: 4px; }}
.media {{ font-size: 1.4em; margin: 16px 0; }} .media b {{ color: #d1495b; }}
table {{ border-collapse: collapse; width: 100%; margin: 16px 0; }}
th, td {{ padding: 6px 10px; border-bottom: 1px solid #e4e7eb; text-align: right; }}
th:nth-child(2), td:nth-child(2) {{ text-align: left; }}
tbody tr:first-child {{ background: #fff4d6; font-weight: bold; }}
img {{ max-width: 100%; height: auto; }}
</style></head><body>
<h1>⚡ OMIEporra {etiqueta}</h1>
<p class="sub">Datos hasta el {ultimo:%d/%m/%Y} · actualizado el {actualizado} (hora de Madrid)</p>
<p class="media">Media acumulada del mes: <b>{num(media)} €/MWh</b></p>
<table><thead><tr><th>Puesto</th><th>Nombre</th><th>Apuesta (€/MWh)</th><th>Diferencia (€/MWh)</th></tr></thead>
<tbody>
{filas}
</tbody></table>
<img src="omieporra_{etiqueta}.png" alt="Gráfica OMIEporra {etiqueta}">
<h2>Todos los meses</h2>
<ul>
{otros}
</ul>
<p class="sub">Fuente: OMIE, mercado diario, precio marginal de España.</p>
</body></html>
""", encoding="utf-8")


# ---------------------------------------------------------------- principal


def main():
    ahora = datetime.now(ZONA)
    if len(sys.argv) > 1 and sys.argv[1].strip():
        try:
            anio, mes = map(int, sys.argv[1].strip().split("-"))
            date(anio, mes, 1)
        except ValueError:
            raise ErrorOmie(f"Mes no válido: {sys.argv[1]!r}. Usa el formato AAAA-MM, p. ej. 2026-10")
    else:
        anio, mes = ahora.year, ahora.month
    etiqueta = f"{anio}-{mes:02d}"
    print(f"OMIEporra {etiqueta} (hoy {ahora:%Y-%m-%d %H:%M} hora de Madrid)")

    apuestas = leer_apuestas()
    tabla = datos_del_mes(anio, mes, ahora.date())
    if tabla.empty:
        raise ErrorOmie(f"Todavía no hay ningún dato de OMIE para {etiqueta}.")
    media = tabla["Media acumulada (€/MWh)"].iloc[-1]
    tabla_ranking = ranking(apuestas, media)

    png = Path(f"omieporra_{etiqueta}.png")
    xlsx = Path(f"omieporra_{etiqueta}.xlsx")
    grafica(tabla, apuestas, etiqueta, png)
    excel(tabla_ranking, tabla, png, xlsx)
    pagina_web(tabla_ranking, tabla, etiqueta)

    print(f"\nDías con datos: {len(tabla)} (último: {tabla['Fecha'].iloc[-1]})")
    print(f"Media acumulada: {media:.2f} €/MWh\n")
    print(tabla_ranking.to_string(index=False))
    print(f"\nGenerados: {xlsx} y {png}")


if __name__ == "__main__":
    try:
        main()
    except ErrorOmie as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        sys.exit(1)
