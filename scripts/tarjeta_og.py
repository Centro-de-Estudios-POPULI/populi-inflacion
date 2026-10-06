"""
Tarjeta para compartir (og.png, 1200×630): la «carátula» que muestran WhatsApp, Facebook, X o LinkedIn al pegar
el enlace de la página. Lleva el último dato, así que se regenera con cada actualización del IPC (paso del robot,
después del scraper).

    python scripts/tarjeta_og.py

Ningún número se escribe acá: todo sale de data/ipc_general.json y data/ipc_descomposicion.json.
Además pone el mes del dato en la URL de la imagen dentro de index.html (og.png?v=AAAA-MM): las redes guardan la
tarjeta de cada URL, y con el mes en la URL la tarjeta nueva se ve apenas cambia el dato.
Paleta, firma y tipografía de la marca (populi-marca; firma como la tarjeta del Banco de Gráficos).
"""
import json
import re
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding="utf-8")
RAIZ = Path(__file__).resolve().parent.parent
FUENTES = Path(__file__).resolve().parent / "fuentes"
DATA = RAIZ / "data"
SALIDA = RAIZ / "og.png"
INDEX = RAIZ / "index.html"

# paleta (copia de populi-marca/paleta.py)
ROJO, TINTA, PAPEL, GRILLA = "#C71E1D", "#001219", "#FAF8F3", "#E2DDD3"
PIZARRA, GRIS, TINT_ROJO = "#5C6B70", "#8A9699", "#F8E5E3"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
W, H, S = 1200, 630, 2          # se dibuja al doble y se reduce: bordes suaves sin depender del antialias de PIL


def fuente(archivo, px):
    return ImageFont.truetype(str(FUENTES / archivo), int(px * S))


def num(x, dec=2, signo=False):
    t = f"{abs(x):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("−" if x < 0 else "+" if signo and x > 0 else "") + t


def serie(lista):
    return {p["fecha"]: p["valor"] for p in lista if p.get("valor") is not None}


def main():
    gen = json.loads((DATA / "ipc_general.json").read_text(encoding="utf-8"))
    ia, ms = serie(gen["var_interanual"]), serie(gen["var_mensual"])
    u = max(ia)
    a, m = int(u[:4]), int(u[5:7])
    hace = f"{a - 1}-{m:02d}"
    nuc = None
    p = DATA / "ipc_descomposicion.json"
    if p.exists():
        n = serie(json.loads(p.read_text(encoding="utf-8")).get("nucleo", []))
        if u in n and hace in n:
            nuc = (n[u] / n[hace] - 1) * 100

    img = Image.new("RGB", (W * S, H * S), PAPEL)
    d = ImageDraw.Draw(img)
    X = lambda v: int(v * S)
    d.rectangle([0, 0, X(W), X(8)], fill=ROJO)
    # firma (la de la tarjeta del Banco)
    d.text((X(72), X(56)), "P", font=fuente("PlayfairDisplay.ttf", 54), fill=ROJO)
    d.text((X(104), X(66)), "opuli", font=fuente("PlayfairDisplay-Italic.ttf", 40), fill=TINTA)
    d.text((X(72), X(142)), "ANÁLISIS DE INFLACIÓN · BOLIVIA", font=fuente("Inter-Bold.ttf", 17), fill=ROJO)
    d.text((X(72), X(172)), "La inflación", font=fuente("PlayfairDisplay.ttf", 50), fill=TINTA)
    d.text((X(72), X(232)), "mes a mes", font=fuente("PlayfairDisplay-Italic.ttf", 50), fill=ROJO)
    # la cifra
    # la cifra en Inter Bold: en JetBrains Mono, a este tamaño, la coma deja un hueco («5 , 70%»)
    d.text((X(68), X(300)), num(ia[u]) + "%", font=fuente("Inter-Bold.ttf", 128), fill=ROJO)
    d.text((X(72), X(452)), f"inflación interanual · {MESES[m - 1]} de {a}", font=fuente("Inter.ttf", 23), fill=PIZARRA)
    otros = [f"Mensual {num(ms[u], 2, True)}%"]
    if nuc is not None:
        otros.append(f"Núcleo {num(nuc)}%")
    if hace in ia:
        otros.append(f"Un año antes {num(ia[hace], 1)}%")
    d.text((X(72), X(486)), "  ·  ".join(otros), font=fuente("Inter-Bold.ttf", 20), fill=TINTA)
    d.text((X(72), X(556)), "Centro de Estudios POPULI · datos del INE", font=fuente("Inter-Bold.ttf", 18), fill=TINTA)
    d.text((X(72), X(582)), "centro-de-estudios-populi.github.io/populi-inflacion", font=fuente("JetBrainsMono-Regular.ttf", 16), fill=GRIS)

    # gráfico: la interanual desde el inicio de la serie, con relleno tenue y el último punto
    K = sorted(ia)
    x0, x1, y0, y1 = 700, 1130, 150, 470
    lo, hi = min(0, min(ia.values())), max(ia.values())
    hi = (int(hi / 5) + 1) * 5
    px = lambda i: x0 + (x1 - x0) * i / (len(K) - 1)
    py = lambda v: y1 - (y1 - y0) * (v - lo) / (hi - lo)
    d.text((X(x0), X(y0 - 46)), "Inflación interanual (%)", font=fuente("Inter-Bold.ttf", 16), fill=TINTA)
    for t in range(0, int(hi) + 1, 5 if hi <= 40 else 10):
        y = py(t)
        d.line([X(x0), X(y), X(x1), X(y)], fill=GRILLA if t else PIZARRA, width=X(1))
        d.text((X(x1 + 10), X(y - 9)), str(t), font=fuente("JetBrainsMono-Regular.ttf", 14), fill=GRIS)
    puntos = [(X(px(i)), X(py(ia[f]))) for i, f in enumerate(K)]
    d.polygon(puntos + [(X(x1), X(py(0))), (X(x0), X(py(0)))], fill=TINT_ROJO)
    d.line(puntos, fill=ROJO, width=X(3.2), joint="curve")
    ux, uy = puntos[-1]
    r = X(7)
    d.ellipse([ux - r, uy - r, ux + r, uy + r], fill=ROJO, outline=PAPEL, width=X(2))
    for anio in (K[0][:4], K[-1][:4]):
        i = next(j for j, f in enumerate(K) if f[:4] == anio)
        d.text((X(px(i)), X(y1 + 12)), anio, font=fuente("JetBrainsMono-Regular.ttf", 14), fill=PIZARRA, anchor="ma")

    img = img.resize((W, H), Image.LANCZOS)
    img.save(SALIDA, optimize=True)
    print(f"[OK] {SALIDA.name}: {u}, interanual {num(ia[u])}% ({SALIDA.stat().st_size // 1024} KB)")

    # el mes del dato en la URL de la imagen (index.html): la tarjeta nueva no queda tapada por la vieja en caché
    if INDEX.exists():
        html = INDEX.read_text(encoding="utf-8")
        nuevo, k = re.subn(r"og\.png\?v=[0-9-]+", f"og.png?v={u}", html)
        if k and nuevo != html:
            INDEX.write_text(nuevo, encoding="utf-8")
            print(f"[OK] index.html: og.png?v={u} ({k} etiquetas)")


if __name__ == "__main__":
    main()
