"""Cálculos extra de la OMIEporra para la web: curva del día por cuartos de hora,
España frente a Portugal, el mismo mes de otros años, rivales automáticos y
clasificación general.

Todo es opcional: omieporra.py llama a estas funciones dentro de try/except, así
que si algo falla aquí la actualización diaria sigue funcionando igual.
"""

from datetime import date
from pathlib import Path

CARPETA_DATOS = Path("datos")
DIAS_TENDENCIA = 7  # «últimos 7 días» del mes anterior


# ---------------------------------------------------------------- precios por cuarto de hora


def leer_es_pt(dia):
    """Precios de España y Portugal de cada periodo de un día, desde la caché de datos/."""
    fichero = CARPETA_DATOS / f"marginalpdbc_{dia:%Y%m%d}.1"
    if not fichero.exists():
        return None
    es, pt = [], []
    for linea in fichero.read_text(encoding="latin-1").splitlines()[1:]:
        linea = linea.strip()
        if not linea or linea.startswith("*"):
            continue
        campos = linea.rstrip(";").split(";")
        if len(campos) < 6:
            continue
        pt.append(float(campos[4].replace(",", ".")))
        es.append(float(campos[5].replace(",", ".")))
    return (es, pt) if es else None


def a_cuartos(periodos):
    """Convierte los periodos de un día en 96 cuartos de hora (00:00 a 23:45).

    - 96 periodos: día normal de 15 minutos.
    - 100: día del cambio de hora de octubre (la hora de 02:00 se repite): se quita la repetición.
    - 92: día del cambio de marzo (no existe la hora de 02:00): esos cuartos quedan vacíos (None).
    - 24, 25 o 23: días antiguos con precios por hora; cada hora vale para sus 4 cuartos.
    """
    n = len(periodos)
    if n == 96:
        return list(periodos)
    if n == 100:
        return list(periodos[:12]) + list(periodos[16:])
    if n == 92:
        return list(periodos[:8]) + [None] * 4 + list(periodos[8:])
    if n in (23, 24, 25):
        horas = list(periodos)
        if n == 25:
            horas = horas[:3] + horas[4:]
        if n == 23:
            horas = horas[:2] + [None] + horas[2:]
        return [h for h in horas for _ in range(4)]
    return None  # formato desconocido: mejor no inventar


def detalle_de_dias(fechas):
    """Para cada día: sus 96 cuartos de hora (España), la media de Portugal y cuántos
    periodos tuvieron un precio distinto en los dos países."""
    detalle = {}
    for dia in fechas:
        datos = leer_es_pt(dia)
        if not datos:
            continue
        es, pt = datos
        cuartos = a_cuartos(es)
        detalle[dia.isoformat()] = {
            "cuartos": [None if c is None else round(c, 2) for c in cuartos] if cuartos else None,
            "pt": round(sum(pt) / len(pt), 2),
            "distintos": sum(1 for a, b in zip(es, pt) if abs(a - b) > 0.005),
            "periodos_totales": len(es),
        }
    return detalle


# ---------------------------------------------------------------- otros años y rivales


def mes_anterior(anio, mes):
    return (anio - 1, 12) if mes == 1 else (anio, mes - 1)


def otros_anios(anio, mes, ahora, datos_del_mes, cuantos=2):
    """El mismo mes de los años anteriores (precio medio de cada día y media del mes).
    Si un año no se puede descargar, se salta con un aviso."""
    resultado = []
    for atras in range(1, cuantos + 1):
        try:
            tabla = datos_del_mes(anio - atras, mes, ahora)
        except Exception as e:  # noqa: BLE001 - nunca debe parar la actualización diaria
            print(f"  AVISO: no se ha podido obtener {mes:02d}/{anio - atras}: {e}")
            continue
        if tabla.empty:
            continue
        resultado.append({
            "anio": anio - atras,
            "media": round(float(tabla["Media acumulada (€/MWh)"].iloc[-1]), 2),
            "dias": [
                {"dia": f["Fecha"].day, "precio": float(f["Precio medio diario (€/MWh)"])}
                for _, f in tabla.iterrows()
            ],
        })
    return resultado


def rivales(anio, mes, ahora, datos_del_mes, anios_anteriores):
    """Predicciones automáticas con lo que se sabía antes de empezar el mes."""
    lista = []
    a, m = mes_anterior(anio, mes)
    try:
        anterior = datos_del_mes(a, m, ahora)
        precios = anterior["Precio medio diario (€/MWh)"]
        lista.append({
            "nombre": "🤖 Persistencia",
            "apuesta": round(float(precios.mean()), 2),
            "explicacion": f"La media del mes anterior ({m:02d}/{a}).",
        })
        lista.append({
            "nombre": "🤖 Últimos 7 días",
            "apuesta": round(float(precios.tail(DIAS_TENDENCIA).mean()), 2),
            "explicacion": f"La media de los últimos {DIAS_TENDENCIA} días del mes anterior.",
        })
    except Exception as e:  # noqa: BLE001
        print(f"  AVISO: sin rivales del mes anterior: {e}")
    hace_un_anio = next((x for x in anios_anteriores if x["anio"] == anio - 1), None)
    if hace_un_anio:
        lista.append({
            "nombre": "🤖 Año pasado",
            "apuesta": hace_un_anio["media"],
            "explicacion": f"La media del mismo mes del año anterior ({mes:02d}/{anio - 1}).",
        })
    return lista


# ---------------------------------------------------------------- clasificación general


def clave_nombre(nombre):
    return " ".join(str(nombre).lower().split())


def clasificacion_general(resultados):
    """Suma de todos los meses TERMINADOS: meses jugados, victorias, error medio,
    sesgo (si suele pasarse o quedarse corto) y mejor puesto de cada uno."""
    jugadores = {}
    for r in resultados:
        if not r.get("final"):
            continue
        for f in r.get("ranking", []):
            j = jugadores.setdefault(clave_nombre(f["nombre"]), {
                "nombre": f["nombre"], "meses": 0, "victorias": 0, "errores": [], "difs": [], "mejor": None,
            })
            j["meses"] += 1
            j["victorias"] += 1 if f["puesto"] == 1 else 0
            j["errores"].append(abs(f["diferencia"]))
            j["difs"].append(f["diferencia"])
            j["mejor"] = f["puesto"] if j["mejor"] is None else min(j["mejor"], f["puesto"])
    tabla = [
        {
            "nombre": j["nombre"],
            "meses": j["meses"],
            "victorias": j["victorias"],
            "error_medio": round(sum(j["errores"]) / len(j["errores"]), 2),
            "sesgo": round(sum(j["difs"]) / len(j["difs"]), 2),
            "mejor_puesto": j["mejor"],
        }
        for j in jugadores.values()
    ]
    # Primero quien más meses ha ganado; a igualdad, quien menos se equivoca de media
    return sorted(tabla, key=lambda x: (-x["victorias"], x["error_medio"]))


def fechas_de(tabla):
    return [f if isinstance(f, date) else date.fromisoformat(str(f)) for f in tabla["Fecha"]]
