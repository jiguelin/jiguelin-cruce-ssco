import io
import zipfile
from datetime import datetime

import openpyxl
import pandas as pd

import ssco_core as core

PADRON_COLS = ["RUC", "Razón social", "Resolución de atribución como SSCO", "Fecha de publicación (2)"]


def _xlsx(filas):
    wb = openpyxl.Workbook()
    for f in filas:
        wb.active.append(f)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _padron():
    return core.leer_padron(_xlsx([
        PADRON_COLS,
        [20610837621, "COMERCIALIZADORA JESDRI S.R.L.", "RI 192-024-0000189/SUNAT", datetime(2024, 12, 31)],
        [10001276510, "TAPIA GONZALES MONICA", "RI 024-024-0090639/SUNAT", datetime(2026, 10, 31)],
    ]))


def test_ruc():
    assert core.limpiar_ruc(20610837621) == "20610837621"
    assert core.limpiar_ruc("206-108-37621") == "20610837621"
    assert core.limpiar_ruc("00000000114") == ""
    assert core.ruc_digito_valido("20610837621")
    assert not core.ruc_digito_valido("20610837622")


def test_padron():
    p = _padron()
    assert list(p["RUC"]) == ["20610837621", "10001276510"]
    assert core.columna_razon_social(p) == "Razón social"


def test_formato_81_contable():
    cab = [None, "CÓDIGO UNICO", None, None, "FECHA DE EMISIÓN", None, "COMPROBANTE"]
    fila = [None, "10", "01", 5, datetime(2026, 10, 3), None, "01", "00E001", "    ", "0000000000040",
            "6 ", "20610837621", "JESDRI", 100.0, 18.0, 0, 0, 0, 0, 0, 0, 0, 118.0]
    fila2 = list(fila); fila2[4] = datetime(2022, 1, 5); fila2[11] = "20555555555"
    r = core.leer_archivo_compras("COMPRAS.xlsx", _xlsx([cab, [], fila, fila2]))
    assert len(r.compras) == 2
    a = core.cruzar(r.compras, _padron())
    assert len(a) == 1
    f = a.iloc[0]
    assert (f["Serie"], f["Número"], f["Total"], f["IGV"]) == ("E001", "40", 118.0, 18.0)
    assert f["Observación"].startswith("PROVEEDOR SSCO")


def test_resultado_sin_encabezado_y_fuera_de_alcance():
    fila = [datetime(2022, 2, 1), None, "01", "F011", None, "156428", "6", "20610837621", "JESDRI",
            107.6, 19.37, 0, 0, 0, 0, 0, 0, 0, 126.97]
    r = core.leer_archivo_compras("resultado_1.xlsx", _xlsx([[None], [None], [None], fila]))
    a = core.cruzar(r.compras, _padron())
    assert a.iloc[0]["Total"] == 126.97
    assert a.iloc[0]["Observación"].startswith("Fuera de alcance")


def test_sire_zip():
    txt = ("RUC|Apellidos y Nombres o Razón social|Periodo|CAR SUNAT|Fecha de emisión|Fecha Vcto/Pago|"
           "Tipo CP/Doc.|Serie del CDP|Año|Nro CP o Doc. Nro Inicial (Rango)|Nro Final (Rango)|"
           "Tipo Doc Identidad|Nro Doc Identidad|Apellidos Nombres/ Razón  Social|BI Gravado DG|"
           "IGV / IPM DG|Total CP|Moneda\n"
           "20602768440|MI EMPRESA|202610|X|24/09/2026||01|E001||1012||6|10001276510|TAPIA|100.00|18.00|118.00|PEN\n"
           "20602768440|MI EMPRESA|202610|X|24/09/2026||01|E001||1013||6|20100047218|BCP|10.00|1.80|11.80|PEN\n")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("propuesta.txt", txt.encode("utf-8"))
    r = core.leer_archivo_compras("propuesta.zip", buf.getvalue())
    assert list(r.compras["RUC proveedor"]) == ["10001276510", "20100047218"]
    a = core.cruzar(r.compras, _padron())
    assert list(a["RUC proveedor"]) == ["10001276510"]
    resumen = core.resumen_por_proveedor(a)
    assert resumen.iloc[0]["IGV"] == 18.0
    xls = core.exportar_excel(a, resumen, r.compras, {"x": 1})
    assert set(pd.read_excel(io.BytesIO(xls), sheet_name=None)) >= {"Alertas SSCO", "Información"}


def test_busqueda_manual():
    res = core.buscar_rucs("20610837621, 20100047218 abc", _padron())
    assert res["Resultado"].tolist()[0].startswith("🚨")
    assert res["Resultado"].tolist()[1].startswith("✅")
    assert "inválido" in res["Resultado"].tolist()[2]


def test_varios_ejercicios_y_publicacion():
    titulo = [None, "PERIODO NOVIEMBRE 2023"]
    fila = [None, "11", "01", 5, datetime(2023, 11, 3), None, "01", "00E001", "", "40",
            "6", "20610837621", "JESDRI", 100.0, 18.0, 0, 0, 0, 0, 0, 0, 0, 118.0]
    r = core.leer_archivo_compras("COMPRAS2023.xlsx", _xlsx([[], titulo, [], fila]))
    assert r.compras.iloc[0]["Periodo"] == "202311"
    fila2 = list(fila); fila2[4] = datetime(2025, 3, 1)
    r2 = core.leer_archivo_compras("x.xlsx", _xlsx([[], [], [], fila2]))
    assert r2.compras.iloc[0]["Periodo"] == "202503"
    compras = pd.concat([r.compras, r2.compras], ignore_index=True)
    a = core.cruzar(compras, _padron())
    momentos = a["Emisión vs. publicación SSCO"].tolist()
    assert momentos[0].startswith("Emitido antes") and momentos[1].startswith("Emitido después")
    pe = core.resumen_por_ejercicio(compras, a)
    assert pe["Ejercicio"].tolist() == ["2023", "2025"]
