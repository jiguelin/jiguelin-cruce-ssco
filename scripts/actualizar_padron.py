"""Descarga el padrón SSCO de SUNAT y lo guarda en data/ (lo usa GitHub Actions)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ssco_core as core  # noqa: E402

destino = Path(__file__).resolve().parents[1] / "data" / "sujesincapacidadOperativa.xlsx"
data = core.descargar_padron()
padron = core.leer_padron(data)
if len(padron) < 100:
    sys.exit(f"Padrón sospechosamente pequeño ({len(padron)} filas); no se actualiza.")
destino.write_bytes(data)
print(f"Padrón actualizado: {len(padron)} sujetos")
