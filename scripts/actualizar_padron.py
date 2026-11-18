"""Descarga de SUNAT el padrón SSCO y la relación art. 12 y los guarda en data/ (GitHub Actions)."""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
import ssco_core as core  # noqa: E402

TAREAS = [
    ("padrón SSCO", core.descargar_padron, core.leer_padron, "sujesincapacidadOperativa.xlsx", 100),
    ("relación art. 12", core.descargar_relacion, core.leer_relacion, "relacion-ssco.xlsx", 10),
]

errores = 0
for nombre, descargar, leer, archivo, minimo in TAREAS:
    try:
        data = descargar()
        filas = len(leer(data))
        if filas < minimo:
            raise ValueError(f"solo {filas} filas; no se actualiza")
        (RAIZ / "data" / archivo).write_bytes(data)
        print(f"{nombre}: actualizado ({filas} registros)")
    except Exception as e:  # noqa: BLE001
        errores += 1
        print(f"{nombre}: NO actualizado – {e}")
sys.exit(1 if errores == len(TAREAS) else 0)
