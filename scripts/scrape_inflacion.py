"""
Scraper mensual del IPC de Bolivia.

Fuentes:
  - INE (Instituto Nacional de Estadistica) - nube.ine.gob.bo
  - CEPALSTAT (CEPAL) - api-cepalstat.cepal.org

Salidas:
  data/ipc_general.json        - indice nacional + variaciones
  data/ipc_divisiones.json     - indice por division COICOP
  data/ipc_ciudades.json       - indice + variaciones por ciudad/conurbacion
  data/ipc_alimentos.json      - alimentos vs no alimentos (nacional)
  data/ipc_ciudades_alimentos.json - alimentos vs no alimentos por ciudad
  data/ipc_productos.json      - productos con mayor/menor variacion + ponderaciones
  data/ipc_productos_hist.json - series historicas de top productos
  data/ipc_ciudades_top.json   - top 5 subidas/bajadas por ciudad (para mapa)
  data/ipc_pesos.json          - ponderacion de cada division (para la contribucion a la inflacion)
  data/ipc_nucleo.json         - inflacion nucleo: IPC sin alimentos ni energia (ponderado por producto)
  data/ipc_transables.json     - IPC transables vs no transables (CEPALSTAT)
  data/ipc_regional.json       - IPC comparativo regional (CEPALSTAT)
  data/metadata.json           - fecha actualizacion, fuente, cobertura
"""

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from io import BytesIO

import requests
import openpyxl

# ── Configuracion ────────────────────────────────────────────────────────────

DATA_DIR = Path(__file__).parent.parent / "data"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

# Paginas del INE de donde se resuelven dinamicamente los enlaces de descarga.
# Los enlaces de nube.ine.gob.bo (Nextcloud) CAMBIAN con cada publicacion mensual,
# por eso se leen de la pagina en cada corrida (ver resolver_fuentes). Los enlaces
# hardcodeados abajo son solo el FALLBACK si el parseo de la pagina falla; conviene
# mantenerlos al dia con la ultima publicacion conocida.
PAGINA_NACIONAL = "https://www.ine.gob.bo/index.php/nacional/"
PAGINA_CIUDADES = "https://www.ine.gob.bo/index.php/ciudades-y-conurbaciones/"

FUENTES_NACIONAL = {
    "general":       "https://nube.ine.gob.bo/index.php/s/P2HkvtlKILPhbvB/download",
    "divisiones":    "https://nube.ine.gob.bo/index.php/s/xiffVcALyTuppvB/download",
    "productos":     "https://nube.ine.gob.bo/index.php/s/lkqCU9CqhvvqtJK/download",
    "ponderaciones": "https://nube.ine.gob.bo/index.php/s/dPhl5fLuXu3n9Zp/download",
    "alimentos":     "https://nube.ine.gob.bo/index.php/s/mKbUGrcfcJaJJEC/download",
    "no_alimentos":  "https://nube.ine.gob.bo/index.php/s/CetvQmQxPiYHAbe/download",
}

FUENTES_CIUDADES = {
    "precios_promedio":      "https://nube.ine.gob.bo/index.php/s/2n5L4vHgP6tnuVG/download",
    "ciudades_variaciones":  "https://nube.ine.gob.bo/index.php/s/S1QROlXt3Kyq3x7/download",
    "ciudades_divisiones":   "https://nube.ine.gob.bo/index.php/s/uDVNHK8ZEjez25L/download",
    "ciudades_alimentos":    "https://nube.ine.gob.bo/index.php/s/l9co01Sa9hx4wyH/download",
    "ciudades_no_alimentos": "https://nube.ine.gob.bo/index.php/s/El65Tj7p83X8l6Z/download",
    "ciudades_productos":    "https://nube.ine.gob.bo/index.php/s/XXPHDFsJcb8F2V4/download",
    "ciudades_ponderaciones":"https://nube.ine.gob.bo/index.php/s/LTCPw7GQL7NiXrG/download",
    "ponderacion_ciudades":  "https://nube.ine.gob.bo/index.php/s/4GhQZKvpEbDQh7c/download",
}

MESES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}

CIUDAD_DEPTO = {
    "BOLIVIA": "Nacional",
    "SUCRE": "Chuquisaca",
    "CONURBACION LA PAZ": "La Paz",
    "REGION METROPOLITANA KANATA": "Cochabamba",
    "ORURO": "Oruro",
    "POTOSI": "Potosi",
    "TARIJA": "Tarija",
    "CONURBACION SANTA CRUZ": "Santa Cruz",
    "TRINIDAD": "Beni",
    "COBIJA": "Pando",
}


def _sin_acentos(t: str) -> str:
    t = t.lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        t = t.replace(a, b)
    return t


def _extraer_links(html: str) -> list[tuple[str, str]]:
    """Devuelve [(texto_normalizado, url), ...] de los anchors a nube.ine.gob.bo."""
    import re
    links = []
    patron = r'<a[^>]+href="([^"]*nube\.ine\.gob\.bo[^"]*)"[^>]*>(.*?)</a>'
    for m in re.finditer(patron, html, re.I | re.S):
        url = m.group(1)
        texto = re.sub(r"<[^>]+>", " ", m.group(2))
        texto = _sin_acentos(re.sub(r"\s+", " ", texto).strip())
        links.append((texto, url))
    return links


def _match(texto: str, key: str) -> bool:
    """Reglas para mapear el texto de un enlace del INE a una clave de fuente."""
    t = texto
    reglas = {
        # Nacional
        "no_alimentos":  "no aliment" in t and "ciudad" not in t,
        "alimentos":     "aliment" in t and "no aliment" not in t and "ciudad" not in t,
        "ponderaciones": "ponderac" in t and "producto" in t and "ciudad" not in t,
        "productos":     "indice a nivel producto" in t and "ciudad" not in t,
        "divisiones":    "division" in t and "ciudad" not in t,
        "general":       "indice general" in t and "division" not in t and "ciudad" not in t,
        # Ciudades
        "precios_promedio":       "precios promedio" in t,
        "ciudades_no_alimentos":  "no aliment" in t and "ciudad" in t,
        "ciudades_alimentos":     "aliment" in t and "no aliment" not in t and "ciudad" in t,
        "ciudades_divisiones":    "division" in t and "ciudad" in t,
        "ponderacion_ciudades":   "ponderac" in t and "ciudad" in t,
        "ciudades_ponderaciones": "ponderac" in t and "producto" in t and "ciudad" not in t,
        "ciudades_productos":     "indices a nivel producto" in t,
        "ciudades_variaciones":   "por ciudad" in t and "division" not in t and "aliment" not in t,
    }
    return reglas.get(key, False)


def resolver_fuentes() -> tuple[dict, dict]:
    """
    Lee las paginas del INE y resuelve los enlaces de descarga actuales por el
    texto de cada enlace. Cae al enlace hardcodeado si no encuentra alguna clave.
    Devuelve (fuentes_nacional, fuentes_ciudades).
    """
    def resolver(pagina: str, fallback: dict) -> dict:
        try:
            html = requests.get(pagina, headers=HEADERS, timeout=60, verify=False).text
            links = _extraer_links(html)
        except Exception as e:
            print(f"  [WARN] No se pudo leer {pagina}: {e} -> usando fallback", file=sys.stderr)
            return dict(fallback)

        resuelto = {}
        for texto, url in links:
            for key in fallback:
                if key in resuelto:
                    continue
                if _match(texto, key):
                    resuelto[key] = url
                    break

        for key, url in fallback.items():
            if key not in resuelto:
                print(f"  [WARN] '{key}' no encontrado en la pagina -> fallback", file=sys.stderr)
                resuelto[key] = url
        return resuelto

    nacional = resolver(PAGINA_NACIONAL, FUENTES_NACIONAL)
    ciudades = resolver(PAGINA_CIUDADES, FUENTES_CIUDADES)
    return nacional, ciudades


def descargar(url: str) -> bytes:
    r = requests.get(url, headers=HEADERS, timeout=60, verify=False)
    r.raise_for_status()
    return r.content


def abrir_excel(contenido: bytes) -> openpyxl.Workbook:
    return openpyxl.load_workbook(BytesIO(contenido), data_only=True)


def safe_float(val) -> float | None:
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def normalizar_ciudad(nombre: str) -> str:
    """Quita notas (1), (2), (3) y normaliza."""
    import re
    nombre = re.sub(r'\s*\(\d+\)\s*$', '', nombre).strip()
    return nombre


def depto_de_ciudad(ciudad: str) -> str:
    norm = ciudad.upper().replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
    for key, val in CIUDAD_DEPTO.items():
        if key in norm:
            return val
    return ciudad


# ── Estructura tipo A: meses en filas, anos en columnas ─────────────────────
# Usado por: general, alimentos, no_alimentos
# R5: MES | 2018 | 2019 | 2020 | ...
# R7: Enero | val | val | val | ...

def leer_tipo_A(ws, fila_anios=5, fila_datos_inicio=7, col_nombre=1, col_datos_inicio=2):
    """
    Lee una hoja con meses en filas y anos en columnas.
    Retorna lista de {fecha: 'YYYY-MM', valor: float}.
    """
    # Leer anos del header
    anios = []
    for col in range(col_datos_inicio, ws.max_column + 1):
        val = ws.cell(row=fila_anios, column=col).value
        if val is not None:
            try:
                anios.append((col, int(val)))
            except (ValueError, TypeError):
                pass

    # Los Excel de alimentos y no alimentos traen en UNA hoja cuatro bloques de enero a diciembre: índice,
    # variación mensual, acumulada y a 12 meses, en ese orden. Sólo vale el primero (el índice): sin este
    # control la serie salía con cuatro filas por mes y las páginas tenían que adivinar cuál era cuál.
    serie, vistas, bloques = [], set(), 1
    for fila in range(fila_datos_inicio, ws.max_row + 1):
        mes_raw = ws.cell(row=fila, column=col_nombre).value
        if mes_raw is None:
            continue
        mes_str = str(mes_raw).strip().upper()
        mes_num = MESES.get(mes_str)
        if mes_num is None:
            continue

        for col_anio, anio in anios:
            val = safe_float(ws.cell(row=fila, column=col_anio).value)
            if val is not None:
                fecha = f"{anio}-{mes_num:02d}"
                if fecha in vistas:
                    bloques += mes_num == 1 and col_anio == anios[0][0]
                    continue
                vistas.add(fecha)
                serie.append({"fecha": fecha, "valor": round(val, 6)})

    if bloques > 1:
        print(f"  [i] hoja '{ws.title}': {bloques} bloques de meses; se toma el primero (índice)")
    serie.sort(key=lambda x: x["fecha"])
    return serie


def leer_tipo_A_multifila(ws, fila_anios=5, fila_datos_inicio=7, col_nombre=1, col_datos_inicio=2):
    """
    Lee una hoja tipo A con multiples filas de datos (ciudades, divisiones).
    Retorna {nombre: [{fecha, valor}, ...]}.
    """
    anios = []
    for col in range(col_datos_inicio, ws.max_column + 1):
        val = ws.cell(row=fila_anios, column=col).value
        if val is not None:
            try:
                anios.append((col, int(val)))
            except (ValueError, TypeError):
                pass

    # Leer meses (fila_datos_inicio en adelante)
    # Pero primero necesitamos identificar los nombres de las filas
    resultado = {}
    for fila in range(fila_datos_inicio, ws.max_row + 1):
        nombre_raw = ws.cell(row=fila, column=col_nombre).value
        if nombre_raw is None:
            continue
        nombre = str(nombre_raw).strip()
        mes_num = MESES.get(nombre.upper())
        if mes_num is not None:
            continue
        if not nombre or nombre.upper() in ("MES", ""):
            continue

        serie = []
        for col_anio, anio in anios:
            val = safe_float(ws.cell(row=fila, column=col_anio).value)
            if val is not None:
                serie.append({"fecha": f"{anio}", "valor": round(val, 6)})

        if serie:
            resultado[nombre] = serie

    return resultado


# ── Estructura tipo B: ciudades en filas, meses en columnas ─────────────────
# Usado por: ciudades_variaciones
# R5: CIUDADES CAPITALES | 2018 | ... (merged)
# R6:                    | ENERO | FEBRERO | ...
# R8: BOLIVIA | val | val | ...

def leer_tipo_B(ws, fila_anios=5, fila_meses=6, fila_datos_inicio=8, col_nombre=1, col_datos_inicio=2):
    """
    Lee una hoja con ciudades en filas y meses en columnas.
    Retorna {ciudad: [{fecha, valor}, ...]}.
    """
    # Leer anos (se propagan a la derecha)
    anio_actual = None
    col_anios = {}
    for col in range(col_datos_inicio, ws.max_column + 1):
        val = ws.cell(row=fila_anios, column=col).value
        if val is not None:
            try:
                anio_actual = int(val)
            except (ValueError, TypeError):
                pass
        if anio_actual is not None:
            col_anios[col] = anio_actual

    # Leer meses
    col_fechas = {}
    for col in range(col_datos_inicio, ws.max_column + 1):
        val_mes = ws.cell(row=fila_meses, column=col).value
        if val_mes is None:
            continue
        mes_str = str(val_mes).strip().upper()
        mes_num = MESES.get(mes_str)
        if mes_num is None:
            continue
        anio = col_anios.get(col)
        if anio is not None:
            col_fechas[col] = f"{anio}-{mes_num:02d}"

    # Leer datos
    resultado = {}
    for fila in range(fila_datos_inicio, ws.max_row + 1):
        nombre_raw = ws.cell(row=fila, column=col_nombre).value
        if nombre_raw is None:
            continue
        nombre = normalizar_ciudad(str(nombre_raw).strip())
        if not nombre:
            continue

        serie = []
        for col, fecha in col_fechas.items():
            val = safe_float(ws.cell(row=fila, column=col).value)
            if val is not None:
                serie.append({"fecha": fecha, "valor": round(val, 6)})

        if serie:
            serie.sort(key=lambda x: x["fecha"])
            resultado[nombre] = serie

    return resultado


# ── Estructura tipo C: divisiones con codigo + descripcion ──────────────────
# R5: DIVISION | DESCRIPCION | 2018 | ... (merged)
# R6:          |             | ENERO | FEBRERO | ...
# R8: 0        | INDICE GENERAL | val | val | ...

def leer_tipo_C(ws, fila_anios=5, fila_meses=6, fila_datos_inicio=8,
                col_codigo=1, col_desc=2, col_datos_inicio=3):
    """
    Lee hoja con codigo + descripcion + datos mensuales.
    Retorna {descripcion: [{fecha, valor}, ...]}.
    """
    anio_actual = None
    col_anios = {}
    for col in range(col_datos_inicio, ws.max_column + 1):
        val = ws.cell(row=fila_anios, column=col).value
        if val is not None:
            try:
                anio_actual = int(val)
            except (ValueError, TypeError):
                pass
        if anio_actual is not None:
            col_anios[col] = anio_actual

    col_fechas = {}
    for col in range(col_datos_inicio, ws.max_column + 1):
        val_mes = ws.cell(row=fila_meses, column=col).value
        if val_mes is None:
            continue
        mes_str = str(val_mes).strip().upper()
        mes_num = MESES.get(mes_str)
        if mes_num is None:
            continue
        anio = col_anios.get(col)
        if anio is not None:
            col_fechas[col] = f"{anio}-{mes_num:02d}"

    resultado = {}
    for fila in range(fila_datos_inicio, ws.max_row + 1):
        desc_raw = ws.cell(row=fila, column=col_desc).value
        codigo_raw = ws.cell(row=fila, column=col_codigo).value
        if desc_raw is None:
            continue
        desc = str(desc_raw).strip()
        codigo = str(codigo_raw).strip() if codigo_raw is not None else ""
        if not desc:
            continue

        key = f"{codigo}. {desc}" if codigo and codigo != "0" else desc

        serie = []
        for col, fecha in col_fechas.items():
            val = safe_float(ws.cell(row=fila, column=col).value)
            if val is not None:
                serie.append({"fecha": fecha, "valor": round(val, 6)})

        if serie:
            serie.sort(key=lambda x: x["fecha"])
            resultado[key] = serie

    return resultado


# ── Procesadores por archivo ─────────────────────────────────────────────────

def procesar_general(contenido: bytes) -> dict:
    wb = abrir_excel(contenido)
    resultado = {}

    for sheet_name in wb.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue

        ws = wb[sheet_name]

        if "INDICE" in upper or "ÍNDICE" in upper:
            if "VAR" not in upper:
                resultado["indice"] = leer_tipo_A(ws)
        if "VAR" in upper and "MENSUAL" in upper and "ACUMULADA" not in upper and "12" not in upper:
            resultado["var_mensual"] = leer_tipo_A(ws)
        elif "ACUMULADA" in upper:
            resultado["var_acumulada"] = leer_tipo_A(ws)
        elif "12" in upper:
            resultado["var_interanual"] = leer_tipo_A(ws)

    return resultado


def procesar_divisiones(contenido: bytes) -> dict:
    wb = abrir_excel(contenido)
    resultado = {}

    for sheet_name in wb.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue
        if "INDICE" in upper or "ÍNDICE" in upper:
            if "VAR" not in upper:
                ws = wb[sheet_name]
                resultado = leer_tipo_C(ws)
                break

    return resultado


def procesar_ciudades(contenido: bytes) -> dict:
    wb = abrir_excel(contenido)
    resultado = {}

    for sheet_name in wb.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue

        ws = wb[sheet_name]

        if "INDICE" in upper or "ÍNDICE" in upper:
            if "VAR" not in upper:
                tipo = "indice"
            else:
                continue
        elif "MENSUAL" in upper and "ACUMULADA" not in upper and "12" not in upper:
            tipo = "var_mensual"
        elif "ACUMULADA" in upper:
            tipo = "var_acumulada"
        elif "12" in upper:
            tipo = "var_interanual"
        else:
            continue

        datos = leer_tipo_B(ws)
        for ciudad, serie in datos.items():
            if ciudad not in resultado:
                resultado[ciudad] = {"departamento": depto_de_ciudad(ciudad)}
            resultado[ciudad][tipo] = serie

    return resultado


def procesar_alimentos_simple(contenido: bytes) -> list:
    """Procesa un Excel de alimentos o no_alimentos (estructura tipo A)."""
    wb = abrir_excel(contenido)
    for sheet_name in wb.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue
        ws = wb[sheet_name]
        serie = leer_tipo_A(ws)
        if serie:
            return serie
    return []


def procesar_productos(contenido_prod: bytes, contenido_pond: bytes) -> dict:
    wb_pond = abrir_excel(contenido_pond)

    # Leer ponderaciones (estructura tipo C: codigo | descripcion | ponderacion)
    ponderaciones = {}
    for sheet_name in wb_pond.sheetnames:
        ws = wb_pond[sheet_name]
        if "INICIO" in sheet_name.upper():
            continue
        for fila in range(1, ws.max_row + 1):
            desc = ws.cell(row=fila, column=2).value
            pond = safe_float(ws.cell(row=fila, column=3).value)
            if pond is None:
                pond = safe_float(ws.cell(row=fila, column=2).value)
                desc = ws.cell(row=fila, column=1).value
            if desc and pond is not None:
                ponderaciones[str(desc).strip()] = pond

    # Leer indices de productos (estructura tipo C: codigo | descripcion | datos)
    wb_prod = abrir_excel(contenido_prod)
    productos = []

    for sheet_name in wb_prod.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue

        ws = wb_prod[sheet_name]
        datos = leer_tipo_C(ws)

        for nombre, serie in datos.items():
            if len(serie) < 13:
                continue
            desc = nombre.split(". ", 1)[-1] if ". " in nombre else nombre
            # el índice general no es un producto (antes se contaba: «398»). Sólo ESE rótulo: «Consulta médica
            # general» sí es un producto (un filtro por «GENERAL» lo dejaba afuera).
            if desc.strip().upper().lstrip("Í").startswith(("NDICE GENERAL", "INDICE GENERAL")):
                continue
            ultimo = serie[-1]["valor"]
            hace_12, prev_mes = hace_meses(serie, 12), hace_meses(serie, 1)
            if not hace_12 or not prev_mes:
                continue
            var_12 = ((ultimo / hace_12) - 1) * 100
            var_mes = ((ultimo / prev_mes) - 1) * 100
            pond = ponderaciones.get(desc.strip(), None)
            productos.append({
                "producto": desc.strip(),
                "var_interanual": round(var_12, 4),
                "var_mensual": round(var_mes, 4),
                "fecha": serie[-1]["fecha"],
                "ponderacion": pond,
                "_i": (ultimo, hace_12, prev_mes),
            })
        break

    # APORTE de cada producto a la variación del IPC (índice de Laspeyres de base fija: Σ peso·índice / Σ peso
    # reproduce el general del INE exacto). Aporte = peso × (I_t − I_antes) / Σ peso × I_antes, en pp: los 397
    # suman la variación del general. Un aporte dice dónde se registró la suba, no qué la causó.
    con_peso = [x for x in productos if x["ponderacion"]]
    den_i = sum(x["ponderacion"] * x["_i"][1] for x in con_peso)
    den_m = sum(x["ponderacion"] * x["_i"][2] for x in con_peso)
    for x in productos:
        w = x["ponderacion"]
        x["aporte_interanual"] = round(w * (x["_i"][0] - x["_i"][1]) / den_i * 100, 5) if w and den_i else None
        x["aporte_mensual"] = round(w * (x["_i"][0] - x["_i"][2]) / den_m * 100, 5) if w and den_m else None
    for x in productos:
        x.pop("_i", None)
    suma_i = sum(x["aporte_interanual"] or 0 for x in productos)
    print(f"  [OK] aportes por producto: {len(con_peso)} con ponderación, suman {suma_i:.3f} pp (interanual)")

    fecha = productos[0]["fecha"] if productos else None

    def _top(metric: str) -> dict:
        ordenado = sorted(productos, key=lambda x: x.get(metric) or 0, reverse=True)
        return {
            "top_subidas": ordenado[:15],
            "top_bajadas": list(reversed(ordenado[-15:])),
            "total_productos": len(ordenado),
        }

    def _aporte(metric: str) -> dict:
        con = [x for x in productos if x.get(metric) is not None]
        ordenado = sorted(con, key=lambda x: x[metric], reverse=True)
        return {"suman": ordenado[:10], "restan": list(reversed(ordenado[-10:])), "total_pp": round(sum(x[metric] for x in con), 4)}

    return {
        "fecha": fecha,
        "interanual": _top("var_interanual"),
        "mensual": _top("var_mensual"),
        # lo que más MOVIÓ el índice (pp), no lo que más subió (%): un producto que pesa mucho y sube poco
        # puede aportar más que uno que se dispara y casi no pesa
        "aporte": {"interanual": _aporte("aporte_interanual"), "mensual": _aporte("aporte_mensual")},
    }


def procesar_ciudades_alimentos(contenido_ali: bytes, contenido_no_ali: bytes) -> dict:
    """Alimentos y no alimentos por ciudad. Cada ciudad es una hoja separada (tipo A)."""
    resultado = {}

    for label, contenido in [("alimentos", contenido_ali), ("no_alimentos", contenido_no_ali)]:
        wb = abrir_excel(contenido)
        for sheet_name in wb.sheetnames:
            upper = sheet_name.upper()
            if "INICIO" in upper:
                continue

            # Extraer nombre de ciudad del nombre de hoja (ej: "2 - SUCRE")
            parts = sheet_name.split(" - ", 1)
            ciudad = parts[-1].strip() if len(parts) > 1 else sheet_name.strip()
            ciudad = normalizar_ciudad(ciudad)

            ws = wb[sheet_name]
            serie = leer_tipo_A(ws)
            if serie:
                if ciudad not in resultado:
                    resultado[ciudad] = {"departamento": depto_de_ciudad(ciudad)}
                resultado[ciudad][label] = serie

    return resultado


def procesar_productos_historico(contenido_prod: bytes, extra: list | None = None) -> dict:
    """Series historicas de los top 10 productos que mas subieron y bajaron, y de los `extra` (los de mayor aporte)."""
    wb = abrir_excel(contenido_prod)

    all_prods = {}
    for sheet_name in wb.sheetnames:
        if "INICIO" in sheet_name.upper():
            continue
        ws = wb[sheet_name]
        datos = leer_tipo_C(ws)
        for nombre, serie in datos.items():
            desc = nombre.split(". ", 1)[-1] if ". " in nombre else nombre
            all_prods[desc.strip()] = serie
        break

    # Calcular var interanual y seleccionar top/bottom 10
    ranked_i = []
    ranked_m = []
    for nombre, serie in all_prods.items():
        if len(serie) < 13:
            continue
        ultimo = serie[-1]["valor"]
        hace_12, prev_mes = hace_meses(serie, 12), hace_meses(serie, 1)
        if not hace_12 or not prev_mes:
            continue
        var_12 = ((ultimo / hace_12) - 1) * 100
        var_mes = ((ultimo / prev_mes) - 1) * 100
        ranked_i.append((nombre, var_12))
        ranked_m.append((nombre, var_mes))

    ranked_i.sort(key=lambda x: x[1], reverse=True)
    ranked_m.sort(key=lambda x: x[1], reverse=True)
    # Union de top/bottom 10 por interanual y por mensual (los modos del toggle).
    top_names = (
        [r[0] for r in ranked_i[:10]] + [r[0] for r in ranked_i[-10:]] +
        [r[0] for r in ranked_m[:10]] + [r[0] for r in ranked_m[-10:]]
    )

    top_names += list(extra or [])
    resultado = {}
    for nombre in top_names:
        if nombre in all_prods and nombre not in resultado:
            resultado[nombre] = all_prods[nombre]

    return resultado


def procesar_ciudades_productos_top(contenido: bytes) -> dict:
    """Top 5 subidas y bajadas por ciudad (para tooltips del mapa)."""
    wb = abrir_excel(contenido)
    resultado = {}

    for sheet_name in wb.sheetnames:
        upper = sheet_name.upper()
        if "INICIO" in upper:
            continue

        ciudad = sheet_name.strip()
        ciudad = normalizar_ciudad(ciudad)

        ws = wb[sheet_name]
        datos = leer_tipo_C(ws)

        prods = []
        for nombre, serie in datos.items():
            if len(serie) < 13:
                continue
            desc = nombre.split(". ", 1)[-1] if ". " in nombre else nombre
            if "GENERAL" in desc.upper() or "INDICE" in desc.upper():
                continue
            ultimo = serie[-1]["valor"]
            hace_12, prev_mes = hace_meses(serie, 12), hace_meses(serie, 1)
            if not hace_12 or not prev_mes:
                continue
            var_12 = ((ultimo / hace_12) - 1) * 100
            var_mes = ((ultimo / prev_mes) - 1) * 100
            prods.append({"p": desc.strip(), "i": round(var_12, 2), "m": round(var_mes, 2)})

        def _topbot(metric: str) -> dict:
            ordenado = sorted(prods, key=lambda x: x[metric], reverse=True)
            return {
                "subidas": [{"p": x["p"], "v": x[metric]} for x in ordenado[:5]],
                "bajadas": [{"p": x["p"], "v": x[metric]} for x in reversed(ordenado[-5:])],
            }

        resultado[ciudad] = {
            "departamento": depto_de_ciudad(ciudad),
            "fecha": max((sr[-1]["fecha"] for sr in datos.values() if sr), default=None),   # el mes de estos rankings
            "interanual": _topbot("i"),
            "mensual": _topbot("m"),
        }

    return resultado


# Definición ESTÁNDAR del núcleo (decisión de Carlos, 2026-10-05): IPC sin alimentos ni energía.
# - Alimentos: división 01 (alimentos y bebidas no alcohólicas).
# - Energía, por la clasificación COICOP (no por lista a mano): 04.5 electricidad, gas y otros
#   combustibles del hogar + 07.2.2 combustibles y lubricantes para vehículos. Con la canasta 2016
#   son 6 productos y 4,29 % del IPC: electricidad, gas por red, GLP, garrafa, gasolina y GNV.
# ⛔ Antes «energía» era la división 07 ENTERA (pasajes de minibús, micro y taxi, vehículos, llantas…:
#   12 %), y el núcleo excluía todo el transporte: en sep-2026 daba 3,5 % contra 5,7 % del estándar,
#   porque dejaba fuera justo el traslado del precio del combustible a los pasajes.
def hace_meses(serie: list, meses: int):
    """Valor de la serie `meses` antes de su último dato, buscado por FECHA (no por posición: si a un producto le
    falta un mes, serie[-13] compararía contra el mes equivocado sin avisar). None si ese mes no está."""
    u = serie[-1]["fecha"]
    t = int(u[:4]) * 12 + int(u[5:7]) - 1 - meses
    clave = f"{t // 12}-{t % 12 + 1:02d}"
    return next((x["valor"] for x in reversed(serie) if x["fecha"] == clave), None)


ENERGIA_PREFIJOS = ("045", "0722")
DEFINICIONES = {
    "nucleo": "IPC sin alimentos ni energía (ponderado por producto con las ponderaciones del INE)",
    "alimentos": "Alimentos y bebidas no alcohólicas (división 01 de la COICOP)",
    "energia": "Electricidad, gas y combustibles (COICOP 04.5 y 07.2.2: electricidad, gas por red, GLP, gasolina, GNV)",
}


def clasificar_producto(codigo: str) -> str:
    if codigo[:2] == "01":
        return "alimentos"
    if codigo.startswith(ENERGIA_PREFIJOS):
        return "energia"
    return "nucleo"


def pesos_divisiones(contenido_pond: bytes, divisiones: dict) -> dict:
    """
    Ponderación de cada división = suma de las ponderaciones de sus productos (cuadro del INE, base 2016).
    Con un índice de Laspeyres de base fija, el índice general es Σ peso·índice / 100: se VERIFICA mes a mes
    contra el índice general del INE antes de publicar. Con los pesos, la contribución de cada división a la
    inflación sale exacta (suman la variación del índice general).
    """
    import re as _re
    wb = abrir_excel(contenido_pond)
    por_div = {}
    for sheet_name in wb.sheetnames:
        if "INICIO" in sheet_name.upper():
            continue
        ws = wb[sheet_name]
        for fila in range(1, ws.max_row + 1):
            code = ws.cell(row=fila, column=1).value
            desc = ws.cell(row=fila, column=2).value
            pond = safe_float(ws.cell(row=fila, column=3).value)
            if code and desc and pond is not None and len(str(code).strip()) >= 4:
                k = str(code).strip()[:2]
                por_div[k] = por_div.get(k, 0.0) + pond
    nombres = {}
    for nombre in divisiones:
        m = _re.match(r"(\d+)\.", nombre)
        if m:
            nombres[m.group(1).zfill(2)] = nombre
    if set(nombres) != set(por_div):
        raise SystemExit(f"✗ pesos: las divisiones no coinciden ({sorted(nombres)} vs {sorted(por_div)})")
    general = next((v for k, v in divisiones.items() if "GENERAL" in k.upper()), None)
    if not general:
        raise SystemExit("✗ pesos: falta el índice general en ipc_divisiones")
    idx = {k: {x["fecha"]: x["valor"] for x in divisiones[n]} for k, n in nombres.items()}
    peor = 0.0
    for punto in general:
        f = punto["fecha"]
        if all(f in idx[k] for k in idx):
            calc = sum(por_div[k] * idx[k][f] for k in idx) / 100
            peor = max(peor, abs(calc - punto["valor"]))
    if peor > 0.05:
        raise SystemExit(f"✗ pesos: Σ peso·índice no reproduce el índice general (desvío máx {peor:.3f})")
    print(f"  [OK] pesos de división: suman {sum(por_div.values()):.2f}; Σ peso·índice reproduce el general (desvío máx {peor:.4f})")
    return {
        "base": "2016",
        "fuente": "INE — ponderaciones de la canasta del IPC por producto, sumadas por división",
        "nota": "Contribución de una división a la variación del IPC = peso × (índice hoy − índice antes) / Σ peso × índice antes.",
        "divisiones": {nombres[k]: round(por_div[k], 4) for k in sorted(por_div)},
    }


def pesos_ciudades(contenido: bytes, ciudades: dict) -> dict:
    """
    Ponderación OFICIAL de cada ciudad en el índice nacional (cuadro 1.1 del INE, base 2016). Se verifica que
    Σ peso·índice de ciudad / 100 reproduzca el índice de BOLIVIA mes a mes antes de publicarla.
    """
    import unicodedata as _u
    norm = lambda t: " ".join(_u.normalize("NFD", str(t)).encode("ascii", "ignore").decode().lower().replace("(", " (").split(" (")[0].split())
    wb = abrir_excel(contenido)
    ws = wb[wb.sheetnames[0]]
    oficiales = {}
    for fila in range(1, ws.max_row + 1):
        code, desc, pond = ws.cell(row=fila, column=1).value, ws.cell(row=fila, column=2).value, safe_float(ws.cell(row=fila, column=3).value)
        if code is not None and str(code).strip().isdigit() and desc and pond is not None:
            oficiales[norm(desc)] = pond
    nombres = {norm(c): c for c in ciudades if c.upper() != "BOLIVIA"}
    if set(nombres) != set(oficiales):
        raise SystemExit(f"✗ pesos de ciudad: no coinciden {sorted(set(nombres) ^ set(oficiales))}")
    nac = {x["fecha"]: x["valor"] for x in ciudades["BOLIVIA"]["indice"]}
    idx = {k: {x["fecha"]: x["valor"] for x in ciudades[c]["indice"]} for k, c in nombres.items()}
    peor = 0.0
    for f, v in nac.items():
        if all(f in idx[k] for k in idx):
            peor = max(peor, abs(sum(oficiales[k] * idx[k][f] for k in idx) / 100 - v))
    if peor > 0.05:
        raise SystemExit(f"✗ pesos de ciudad: Σ peso·índice no reproduce Bolivia (desvío máx {peor:.3f})")
    print(f"  [OK] pesos de ciudad (oficiales): Σ peso·índice reproduce Bolivia (desvío máx {peor:.4f})")
    return {nombres[k]: round(oficiales[k], 4) for k in sorted(nombres)}


def calcular_descomposicion(contenido_prod: bytes, contenido_pond: bytes) -> dict:
    """
    Calcula indices ponderados para: general, nucleo, alimentos, energia.
    Usa datos a nivel de producto con ponderaciones reescaladas por categoria.
    """
    wb_pond = abrir_excel(contenido_pond)
    ponderaciones = {}
    for sheet_name in wb_pond.sheetnames:
        if "INICIO" in sheet_name.upper():
            continue
        ws = wb_pond[sheet_name]
        for fila in range(1, ws.max_row + 1):
            code = ws.cell(row=fila, column=1).value
            desc = ws.cell(row=fila, column=2).value
            pond = safe_float(ws.cell(row=fila, column=3).value)
            if code and desc and pond is not None:
                code_str = str(code).strip()
                if len(code_str) >= 4:
                    ponderaciones[code_str] = {
                        "desc": str(desc).strip(),
                        "pond": pond,
                        "cat": clasificar_producto(code_str),
                    }

    wb_prod = abrir_excel(contenido_prod)
    series_prod = {}
    for sheet_name in wb_prod.sheetnames:
        if "INICIO" in sheet_name.upper():
            continue
        ws = wb_prod[sheet_name]
        datos = leer_tipo_C(ws)
        for nombre, serie in datos.items():
            codigo = nombre.split(".")[0].strip()
            desc = nombre.split(". ", 1)[-1].strip() if ". " in nombre else nombre.strip()
            if codigo in ponderaciones:
                series_prod[codigo] = serie
            else:
                for pc, pi in ponderaciones.items():
                    if pi["desc"] == desc:
                        series_prod[pc] = serie
                        break
        break

    cats = {"general": {}, "nucleo": {}, "alimentos": {}, "energia": {}}
    for codigo, info in ponderaciones.items():
        if codigo not in series_prod:
            continue
        cat = info["cat"]
        cats["general"][codigo] = info["pond"]
        cats[cat][codigo] = info["pond"]

    resultado = {}
    for cat_name, prods_pond in cats.items():
        if not prods_pond:
            continue
        total_pond = sum(prods_pond.values())
        if total_pond == 0:
            continue
        pesos = {c: p / total_pond for c, p in prods_pond.items()}

        por_fecha = {}
        for codigo, peso in pesos.items():
            serie = series_prod.get(codigo, [])
            for punto in serie:
                f = punto["fecha"]
                if f not in por_fecha:
                    por_fecha[f] = 0.0
                por_fecha[f] += punto["valor"] * peso

        indice = [{"fecha": f, "valor": round(v, 6)}
                  for f, v in sorted(por_fecha.items())]
        resultado[cat_name] = indice

    meta = {}
    for cat_name, prods_pond in cats.items():
        total = sum(prods_pond.values())
        meta[cat_name] = {
            "productos": len(prods_pond),
            "ponderacion_total": round(total, 2),
        }
    resultado["meta"] = meta

    return resultado




# ── CEPALSTAT ────────────────────────────────────────────────────────────────

CEPALSTAT_BASE = "https://api-cepalstat.cepal.org/cepalstat/api/v1"

CEPALSTAT_MES_MAP = {
    516: 1, 517: 2, 518: 3, 519: 4, 825: 5, 821: 6,
    822: 7, 823: 8, 824: 9, 826: 10, 827: 11, 828: 12,
}

CEPALSTAT_PAISES = {
    "BOL": "Bolivia", "ARG": "Argentina", "BRA": "Brasil",
    "CHL": "Chile", "COL": "Colombia", "ECU": "Ecuador",
    "PRY": "Paraguay", "PER": "Perú", "URY": "Uruguay",
}


def cepalstat_año(dim_id: int) -> int:
    return 1980 + (dim_id - 29150)


def cepalstat_fetch(indicator_id: int) -> list[dict]:
    url = f"{CEPALSTAT_BASE}/indicator/{indicator_id}/data?format=json&in=1&lang=es"
    r = requests.get(url, headers=HEADERS, timeout=120)
    r.raise_for_status()
    body = r.json().get("body", {})
    return body.get("data", [])


def cepalstat_to_series(raw: list[dict], filter_iso3: set | None = None) -> dict[str, list[dict]]:
    result = {}
    for rec in raw:
        iso3 = rec.get("iso3", "")
        if filter_iso3 and iso3 not in filter_iso3:
            continue
        val = safe_float(rec.get("value"))
        if val is None:
            continue
        dim_year = rec.get("dim_29117")
        dim_month = rec.get("dim_515")
        if dim_year is None or dim_month is None:
            continue
        year = cepalstat_año(dim_year)
        month = CEPALSTAT_MES_MAP.get(dim_month)
        if month is None or year < 2010:
            continue
        country = CEPALSTAT_PAISES.get(iso3, iso3)
        if country not in result:
            result[country] = []
        result[country].append({"fecha": f"{year}-{month:02d}", "valor": round(val, 4)})
    for series in result.values():
        series.sort(key=lambda x: x["fecha"])
    return result


def procesar_cepalstat_transables() -> dict:
    print("  > CEPALSTAT: IPC transables (762)...")
    raw_t = cepalstat_fetch(762)
    print("  > CEPALSTAT: IPC no transables (763)...")
    raw_nt = cepalstat_fetch(763)
    t = cepalstat_to_series(raw_t, {"BOL"})
    nt = cepalstat_to_series(raw_nt, {"BOL"})
    return {
        "transables": t.get("Bolivia", []),
        "no_transables": nt.get("Bolivia", []),
    }


def procesar_cepalstat_regional() -> dict:
    print("  > CEPALSTAT: IPC general regional (365)...")
    raw = cepalstat_fetch(365)
    return cepalstat_to_series(raw, set(CEPALSTAT_PAISES.keys()))


# ── Principal ────────────────────────────────────────────────────────────────

def main() -> None:
    import warnings
    warnings.filterwarnings("ignore", message="Unverified HTTPS")

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print("Resolviendo enlaces actuales desde la pagina del INE...")
    fuentes_nacional, fuentes_ciudades = resolver_fuentes()

    print("Descargando Excel del INE...")

    archivos = {}
    todas_fuentes = {**fuentes_nacional, **fuentes_ciudades}

    for nombre, url in todas_fuentes.items():
        try:
            print(f"  > {nombre}...")
            archivos[nombre] = descargar(url)
        except Exception as e:
            print(f"  [ERROR] {nombre}: {e}", file=sys.stderr, flush=True)

    # ── Procesar nacional ────────────────────────────────────────────────
    print("\nProcesando datos nacionales...")

    general = None   # lo usa también el control de la descomposición
    divisiones = None
    productos = None
    if "general" in archivos:
        general = procesar_general(archivos["general"])
        out_path = DATA_DIR / "ipc_general.json"
        out_path.write_text(json.dumps(general, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        n = len(general.get("indice", []))
        print(f"  [OK] ipc_general.json - {n} meses")

    if "divisiones" in archivos:
        divisiones = procesar_divisiones(archivos["divisiones"])
        out_path = DATA_DIR / "ipc_divisiones.json"
        out_path.write_text(json.dumps(divisiones, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_divisiones.json - {len(divisiones)} divisiones")

        # El núcleo sale SÓLO de la descomposición ponderada (más abajo). El viejo respaldo de promedio simple de
        # divisiones tenía otra definición y, si la descomposición fallaba, quedaba publicado en silencio.

    if divisiones and "ponderaciones" in archivos:
        pesos = pesos_divisiones(archivos["ponderaciones"], divisiones)
        (DATA_DIR / "ipc_pesos.json").write_text(json.dumps(pesos, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print("  [OK] ipc_pesos.json")

    if "alimentos" in archivos and "no_alimentos" in archivos:
        ali = procesar_alimentos_simple(archivos["alimentos"])
        no_ali = procesar_alimentos_simple(archivos["no_alimentos"])
        alimentos = {"alimentos": ali, "no_alimentos": no_ali}
        out_path = DATA_DIR / "ipc_alimentos.json"
        out_path.write_text(json.dumps(alimentos, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_alimentos.json - ali: {len(ali)}, no_ali: {len(no_ali)}")

    if "productos" in archivos and "ponderaciones" in archivos:
        productos = procesar_productos(archivos["productos"], archivos["ponderaciones"])
        out_path = DATA_DIR / "ipc_productos.json"
        out_path.write_text(json.dumps(productos, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_productos.json - {productos['interanual']['total_productos']} productos (interanual + mensual)")

        descomp = calcular_descomposicion(archivos["productos"], archivos["ponderaciones"])
        descomp["meta"]["definiciones"] = DEFINICIONES
        # Control: el «general» reconstruido con los productos tiene que reproducir la interanual oficial del INE.
        # Si en algún mes de los últimos 24 se aparta más de 0,3 pp, la descomposición está mal leída: no se publica.
        oficial = {x["fecha"]: x["valor"] for x in (general or {}).get("var_interanual", [])}
        rec = {x["fecha"]: x["valor"] for x in descomp.get("general", [])}
        desvios = []
        for f in sorted(oficial)[-24:]:
            a = f"{int(f[:4]) - 1}{f[4:]}"
            if f in rec and a in rec and rec[a]:
                d = (rec[f] / rec[a] - 1) * 100 - oficial[f]
                if abs(d) > 0.3:
                    desvios.append(f"{f}: {d:+.2f} pp")
        if desvios:
            raise SystemExit("✗ la descomposición no reproduce la interanual del INE: " + ", ".join(desvios[:6]))
        print(f"  [OK] la descomposición reproduce la interanual del INE (últimos {min(24, len(oficial))} meses, ±0,3 pp)")
        out_path = DATA_DIR / "ipc_descomposicion.json"
        out_path.write_text(json.dumps(descomp, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        meta = descomp.get("meta", {})
        for cat, info in meta.items():
            if cat == "definiciones":
                continue
            print(f"    {cat}: {info['productos']} prods, pond={info['ponderacion_total']}%")
        print(f"  [OK] ipc_descomposicion.json")

        if "nucleo" in descomp and descomp["nucleo"]:
            nucleo_path = DATA_DIR / "ipc_nucleo.json"
            nucleo_path.write_text(
                json.dumps(descomp["nucleo"], ensure_ascii=False, separators=(",", ":")),
                encoding="utf-8",
            )
            print(f"  [OK] ipc_nucleo.json (ponderado) - {len(descomp['nucleo'])} meses")

    if "productos" in archivos:
        extra = []
        if isinstance(productos, dict):
            for modo in ("interanual", "mensual"):
                ap = productos.get("aporte", {}).get(modo, {})
                extra += [x["producto"] for x in ap.get("suman", []) + ap.get("restan", [])]
        prod_hist = procesar_productos_historico(archivos["productos"], extra)
        out_path = DATA_DIR / "ipc_productos_hist.json"
        out_path.write_text(json.dumps(prod_hist, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_productos_hist.json - {len(prod_hist)} productos")

    # ── Procesar ciudades ────────────────────────────────────────────────
    print("\nProcesando datos por ciudad...")

    if "ciudades_variaciones" in archivos:
        ciudades = procesar_ciudades(archivos["ciudades_variaciones"])
        out_path = DATA_DIR / "ipc_ciudades.json"
        out_path.write_text(json.dumps(ciudades, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_ciudades.json - {len(ciudades)} ciudades")
        if "ponderacion_ciudades" in archivos:
            ruta_pesos = DATA_DIR / "ipc_pesos.json"
            pesos = json.loads(ruta_pesos.read_text(encoding="utf-8")) if ruta_pesos.exists() else {}
            pesos["ciudades"] = pesos_ciudades(archivos["ponderacion_ciudades"], ciudades)
            pesos["fuente_ciudades"] = "INE — ponderaciones del IPC por ciudad capital y conurbación (cuadro 1.1, base 2016)"
            ruta_pesos.write_text(json.dumps(pesos, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    if "ciudades_alimentos" in archivos and "ciudades_no_alimentos" in archivos:
        ciudades_ali = procesar_ciudades_alimentos(
            archivos["ciudades_alimentos"], archivos["ciudades_no_alimentos"]
        )
        out_path = DATA_DIR / "ipc_ciudades_alimentos.json"
        out_path.write_text(json.dumps(ciudades_ali, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_ciudades_alimentos.json - {len(ciudades_ali)} ciudades")

    if "ciudades_productos" in archivos:
        ciudades_top = procesar_ciudades_productos_top(archivos["ciudades_productos"])
        out_path = DATA_DIR / "ipc_ciudades_top.json"
        out_path.write_text(json.dumps(ciudades_top, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_ciudades_top.json - {len(ciudades_top)} ciudades")

    # ── CEPALSTAT ────────────────────────────────────────────────────────
    print("\nDescargando datos CEPALSTAT...")
    try:
        transables = procesar_cepalstat_transables()
        out_path = DATA_DIR / "ipc_transables.json"
        out_path.write_text(json.dumps(transables, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_transables.json - T: {len(transables['transables'])}, NT: {len(transables['no_transables'])}")
    except Exception as e:
        print(f"  [ERROR] transables: {e}", file=sys.stderr, flush=True)

    try:
        regional = procesar_cepalstat_regional()
        out_path = DATA_DIR / "ipc_regional.json"
        out_path.write_text(json.dumps(regional, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_regional.json - {len(regional)} países")
    except Exception as e:
        print(f"  [ERROR] regional: {e}", file=sys.stderr, flush=True)

    # ── Metadata ─────────────────────────────────────────────────────────
    metadata = {
        "actualizado": datetime.now(timezone.utc).isoformat(),
        "fuente": "INE Bolivia - Indice de Precios al Consumidor (Base 2016)",
        "url_nacional": "https://www.ine.gob.bo/index.php/nacional/",
        "url_ciudades": "https://www.ine.gob.bo/index.php/ciudades-y-conurbaciones/",
        "cobertura": "9 ciudades capitales y conurbaciones",
        "frecuencia": "Mensual",
        "base": 2016,
    }
    meta_path = DATA_DIR / "metadata.json"
    meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n[OK] metadata.json")
    print("Proceso completado!")


def main_cepalstat_only() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    print("Descargando solo datos CEPALSTAT...")
    try:
        transables = procesar_cepalstat_transables()
        out_path = DATA_DIR / "ipc_transables.json"
        out_path.write_text(json.dumps(transables, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_transables.json - T: {len(transables['transables'])}, NT: {len(transables['no_transables'])}")
    except Exception as e:
        print(f"  [ERROR] transables: {e}", file=sys.stderr, flush=True)
    try:
        regional = procesar_cepalstat_regional()
        out_path = DATA_DIR / "ipc_regional.json"
        out_path.write_text(json.dumps(regional, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        print(f"  [OK] ipc_regional.json - {len(regional)} países")
    except Exception as e:
        print(f"  [ERROR] regional: {e}", file=sys.stderr, flush=True)
    print("CEPALSTAT completado!")


if __name__ == "__main__":
    if "--cepalstat-only" in sys.argv:
        main_cepalstat_only()
    else:
        main()
