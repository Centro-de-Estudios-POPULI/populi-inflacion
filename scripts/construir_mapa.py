"""
Construye data/bolivia.geojson (los 9 departamentos del mapa de la sección 06) desde la
georreferencia madre de POPULI: bo-geo-maestro/geo/departamentos.geojson.

No corre en el robot: se usa una vez, cuando cambia la geometría madre.
    python scripts/construir_mapa.py [ruta a bo-geo-maestro/geo/departamentos.geojson]

Por qué no el archivo anterior: sus 9 polígonos eran inválidos (bordes que se cruzan) y sus
coordenadas venían redondeadas a 0,01° (~1 km).

Pasos:
1. Los límites municipales del INE dejan fuera los salares (Uyuni, Coipasa): son huecos en la
   unión de los departamentos. En un mapa de coropletas se leerían como «sin dato», así que cada
   hueco se asigna al departamento con el que comparte más borde. Coipasa, que toca la frontera
   con Chile, es una ensenada y no un hueco: se cierra sólo en su recuadro.
2. Se descartan las islas y los fragmentos de menos de MIN_PARTE grados² (los de los lagos);
   la península de Copacabana (La Paz) queda.
3. Simplificación con topología compartida (topojson): los bordes comunes se simplifican una sola
   vez, así no se abren rendijas ni se pisan departamentos vecinos.
4. Coordenadas con DECIMALES decimales (~11 m) y la propiedad `name` que leen las páginas.
"""
import json
import sys
from pathlib import Path

import geopandas as gpd
import shapely
import topojson as tp

RAIZ = Path(__file__).resolve().parent.parent
MADRE = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ.parent / "bo-geo-maestro" / "geo" / "departamentos.geojson"
SALIDA = RAIZ / "data" / "bolivia.geojson"
MIN_HUECO = 0.01     # grados²: los huecos menores se rellenan igual (eran rendijas), los mayores son salares
MIN_PARTE = 0.01     # grados²: fragmentos menores se descartan
TOLERANCIA = 0.006   # grados: Douglas-Peucker sobre la topología compartida
CIERRE = 0.15        # grados: radio del cierre que tapa la boca del Salar de Coipasa
COIPASA = shapely.box(-68.9, -19.75, -67.9, -19.0)
DECIMALES = 4


def partes(g):
    return list(g.geoms) if g.geom_type == "MultiPolygon" else [g]


def main():
    dep = gpd.read_file(MADRE)[["dpto", "cod_dep", "geometry"]]
    assert len(dep) == 9 and dep.is_valid.all(), "la geometría madre debe traer 9 departamentos válidos"

    # 1. huecos de la unión (salares) → al departamento que más borde comparte con cada uno
    union = shapely.union_all(dep.geometry.values)
    huecos = [shapely.Polygon(r) for p in partes(union) for r in p.interiors]
    geoms = list(dep.geometry)
    for h in huecos:
        borde = h.exterior
        largos = [g.boundary.intersection(borde.buffer(1e-6)).length for g in geoms]
        k = max(range(len(geoms)), key=lambda i: largos[i])
        geoms[k] = shapely.union_all([geoms[k], h])
        if h.area >= MIN_HUECO:
            print(f"  hueco de {h.area:.3f} grados² → {dep.dpto.iloc[k]}")
    # agujeros internos de cada departamento (no son de otro: el solape de la madre es 0)
    geoms = [shapely.union_all([shapely.Polygon(p.exterior) for p in partes(g)]) for g in geoms]
    # El Salar de Coipasa toca la frontera con Chile: no es un hueco sino una ensenada abierta. Se cierra con
    # un «cierre» morfológico (agrandar y achicar CIERRE grados) SÓLO dentro de su recuadro, para no rellenar
    # las bahías del Titicaca, y va al departamento con el que comparte más borde.
    lleno = shapely.union_all(geoms)
    cerrado = lleno.buffer(CIERRE, join_style="mitre").buffer(-CIERRE, join_style="mitre")
    for h in partes(shapely.make_valid(cerrado).intersection(COIPASA).difference(lleno)):
        if h.geom_type != "Polygon" or h.area < MIN_HUECO:
            continue
        largos = [g.boundary.intersection(h.exterior.buffer(1e-6)).length for g in geoms]
        k = max(range(len(geoms)), key=lambda i: largos[i])
        geoms[k] = shapely.union_all([geoms[k], h])
        print(f"  ensenada de {h.area:.3f} grados² (Coipasa) → {dep.dpto.iloc[k]}")
    geoms = [shapely.MultiPolygon(partes(shapely.union_all([shapely.Polygon(p.exterior) for p in partes(g)]))) for g in geoms]

    # 2. fuera las islas y los fragmentos chicos
    geoms = [shapely.MultiPolygon([p for p in partes(g) if p.area >= MIN_PARTE]) for g in geoms]
    dep = dep.set_geometry(geoms)
    assert abs(shapely.union_all(dep.geometry.values).area - sum(g.area for g in geoms)) < 1e-6, "los departamentos se pisan"

    # 3. simplificación con topología compartida
    topo = tp.Topology(dep, prequantize=1e6, toposimplify=TOLERANCIA, simplify_algorithm="dp")
    out = topo.to_gdf()
    out["geometry"] = out.geometry.apply(shapely.make_valid)
    assert out.is_valid.all()

    # 4. salida
    redondear = lambda a: [redondear(x) for x in a] if isinstance(a[0], (list, tuple)) else [round(a[0], DECIMALES), round(a[1], DECIMALES)]
    feats = []
    for _, r in out.sort_values("cod_dep").iterrows():
        g = shapely.geometry.mapping(r.geometry)
        feats.append({"type": "Feature", "properties": {"name": r.dpto, "cod_dep": r.cod_dep},
                      "geometry": {"type": g["type"], "coordinates": redondear(g["coordinates"])}})
    fc = {"type": "FeatureCollection",
          "fuente": "bo-geo-maestro/geo/departamentos.geojson (límites del INE); salares asignados al departamento que los rodea; simplificado con topología compartida",
          "features": feats}
    SALIDA.write_text(json.dumps(fc, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    n = sum(len(p.exterior.coords) for g in out.geometry for p in partes(g))
    print(f"[OK] {SALIDA.name}: {len(feats)} departamentos, ~{n} vértices, {SALIDA.stat().st_size // 1024} KB")


if __name__ == "__main__":
    main()
