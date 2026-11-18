"""Lógica de negocio: lectura de registros de compras, padrón SSCO y cruce por RUC.

Formatos de compras soportados (se detectan automáticamente):
  * Excel del software contable "FORMATO 8.1: REGISTRO DE COMPRAS".
  * Excel "resultado_*.xlsx" (comprobantes sin encabezados, misma disposición 8.1).
  * Propuesta RCE del SIRE de SUNAT: .zip, .csv (comas) o .txt (pipes).
  * Cualquier otro Excel/CSV: se detecta la columna que contiene RUCs.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd
import requests

SSCO_URL = "https://www.sunat.gob.pe/padronesnotificaciones/ssco/sujesincapacidadOperativa.xlsx"
# Relación art. 12 DL 1532 (EIRL y sociedades vinculadas a SSCO). SUNAT no documenta una URL
# fija para el Excel; se prueban estas rutas y, si fallan, la app usa la copia en data/.
RELACION_PAGINA = "https://www.sunat.gob.pe/padronesnotificaciones/relacion-ssco.html"
RELACION_URLS = [
    "https://www.sunat.gob.pe/padronesnotificaciones/ssco/descarga-relacion-ssco.xlsx",
    "https://www.sunat.gob.pe/padronesnotificaciones/descarga-relacion-ssco.xlsx",
]
FECHA_INICIO_DL1532 = date(2022, 3, 20)

RUC_RE = re.compile(r"^(10|15|16|17|20)\d{9}$")

TIPOS_CP = {
    "01": "Factura", "02": "Recibo por honorarios", "03": "Boleta", "07": "Nota de crédito",
    "08": "Nota de débito", "12": "Ticket", "14": "Servicios públicos", "30": "Doc. entidades financieras",
    "50": "DUA", "53": "Declaración mensajería", "54": "Liquidación cobranza", "91": "Comprobante no domiciliado",
}

COLUMNAS_NORMALIZADAS = [
    "Origen", "Hoja", "Fila", "Periodo", "RUC proveedor", "Razón social (compras)", "Fecha emisión",
    "Tipo CP", "Serie", "Número", "Base imponible", "IGV", "Total", "Moneda",
]


# --------------------------------------------------------------------------- utilidades
def limpiar_ruc(valor) -> str:
    """Devuelve el RUC como texto de 11 dígitos, o '' si no lo es."""
    if valor is None:
        return ""
    if isinstance(valor, float):
        if pd.isna(valor):
            return ""
        if valor.is_integer():
            valor = int(valor)
    s = re.sub(r"[\s\-\.]", "", str(valor))
    s = s.lstrip("0") if len(s) > 11 else s
    return s if RUC_RE.match(s) else ""


def ruc_digito_valido(ruc: str) -> bool:
    """Valida el dígito verificador (módulo 11) del RUC."""
    if not RUC_RE.match(ruc or ""):
        return False
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(d) * p for d, p in zip(ruc[:10], pesos))
    dv = 11 - (suma % 11)
    dv = {10: 0, 11: 1}.get(dv, dv)
    return dv == int(ruc[10])


def _norm(texto) -> str:
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def _a_fecha(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return pd.NaT
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v)
    return pd.to_datetime(str(v).strip(), dayfirst=True, errors="coerce")


def _a_num(v):
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def _txt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _sin_ceros(v) -> str:
    s = _txt(v)
    return s.lstrip("0") or s


# --------------------------------------------------------------------------- lectura de compras
@dataclass
class ResultadoLectura:
    compras: pd.DataFrame
    avisos: list[str] = field(default_factory=list)


def _leer_texto(data: bytes) -> str:
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _tabla_desde_texto(data: bytes) -> pd.DataFrame:
    texto = _leer_texto(data)
    primera = texto.splitlines()[0] if texto else ""
    sep = "|" if primera.count("|") > primera.count(",") else ("," if "," in primera else ";")
    return pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, keep_default_na=False,
                       quoting=csv.QUOTE_MINIMAL, index_col=False)


def _es_sire(df: pd.DataFrame) -> bool:
    cols = {_norm(c) for c in df.columns}
    return "NRO DOC IDENTIDAD" in cols and "TOTAL CP" in cols


def _parse_sire(df: pd.DataFrame, origen: str) -> pd.DataFrame:
    c = {_norm(col): col for col in df.columns}

    def col(*nombres):
        for n in nombres:
            if n in c:
                return df[c[n]]
        return pd.Series([""] * len(df))

    razon = next((c[k] for k in c if k.startswith("APELLIDOS NOMBRES/")), None)
    out = pd.DataFrame({
        "Origen": origen,
        "Hoja": "",
        "Fila": range(2, len(df) + 2),
        "Periodo": col("PERIODO").map(_txt),
        "RUC proveedor": col("NRO DOC IDENTIDAD").map(limpiar_ruc),
        "Razón social (compras)": df[razon] if razon else "",
        "Fecha emisión": col("FECHA DE EMISION").map(_a_fecha),
        "Tipo CP": col("TIPO CP/DOC.").map(_txt),
        "Serie": col("SERIE DEL CDP").map(_txt),
        "Número": col("NRO CP O DOC. NRO INICIAL (RANGO)").map(_sin_ceros),
        "Base imponible": col("BI GRAVADO DG").map(_a_num),
        "IGV": col("IGV / IPM DG").map(_a_num),
        "Total": col("TOTAL CP").map(_a_num),
        "Moneda": col("MONEDA").map(_txt),
    })
    return out


MESES = {m: i for i, m in enumerate(
    ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
     "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"], 1)}
MESES["SEPTIEMBRE"] = 9


def _periodo_en_titulo(raw: pd.DataFrame) -> str:
    """Busca 'PERIODO NOVIEMBRE 2025' (o 'PERIODO 202511') en las primeras filas."""
    for v in raw.iloc[:10].to_numpy().ravel():
        t = _norm(v)
        m = re.search(r"PERIODO\s*:?\s*([A-Z]+)\s+(?:DE(?:L)?\s+)?(\d{4})", t)
        if m and m.group(1) in MESES:
            return f"{m.group(2)}{MESES[m.group(1)]:02d}"
        m = re.search(r"PERIODO\s*:?\s*(\d{4})(\d{2})\b", t)
        if m:
            return m.group(1) + m.group(2)
    return ""


def _columna_ruc(raw: pd.DataFrame) -> int | None:
    """Columna con más RUCs válidos; descarta columnas con un solo RUC repetido
    (suele ser el RUC del propio contribuyente) si hay otra alternativa."""
    mejores = []
    for j in range(raw.shape[1]):
        rucs = raw.iloc[:, j].map(limpiar_ruc)
        validos = rucs[rucs != ""]
        if len(validos):
            cab = any(_norm(v) == "RUC" for v in raw.iloc[:15, j])
            # prioridad: columna titulada "RUC", RUCs variados, cantidad, más a la izquierda
            mejores.append((cab, validos.nunique() > 1, len(validos), -j))
    if not mejores:
        return None
    mejores.sort(reverse=True)
    return -mejores[0][3]


def _parse_tabla_generica(raw: pd.DataFrame, origen: str, hoja: str) -> pd.DataFrame | None:
    """Lee una hoja sin asumir encabezados. Si la disposición es la del Formato 8.1
    (contable o 'resultado'), toma fecha, tipo, serie, número y montos por posición
    relativa a la columna del RUC; si no, solo RUC y razón social."""
    raw = raw.T.reset_index(drop=True).T  # columnas 0..n (sin eliminar vacías: importan las posiciones)
    j = _columna_ruc(raw)
    if j is None:
        return None
    rucs = raw.iloc[:, j].map(limpiar_ruc)
    filas = raw[rucs != ""]
    if filas.empty:
        return None

    periodo = _periodo_en_titulo(raw)

    def celda(fila, k):
        return fila.iloc[k] if 0 <= k < len(fila) else None

    # ¿Disposición 8.1? tipo doc identidad a la izquierda y fecha 7 columnas antes.
    muestra = filas.head(30)
    tipo_doc_ok = muestra.iloc[:, j - 1].map(lambda v: _txt(v) in {"6", "06", "0", "1", "4", "7", "A"}).mean() > 0.7 if j >= 1 else False
    fecha_ok = muestra.iloc[:, j - 7].map(lambda v: not pd.isna(_a_fecha(v))).mean() > 0.7 if j >= 7 else False
    disposicion_81 = bool(tipo_doc_ok and fecha_ok)

    registros = []
    for idx, fila in filas.iterrows():
        reg = {
            "Origen": origen, "Hoja": hoja, "Fila": int(idx) + 1, "Periodo": periodo,
            "RUC proveedor": limpiar_ruc(fila.iloc[j]),
            "Razón social (compras)": _txt(celda(fila, j + 1)),
        }
        if disposicion_81:
            reg.update({
                "Fecha emisión": _a_fecha(celda(fila, j - 7)),
                "Tipo CP": _txt(celda(fila, j - 5)).zfill(2),
                "Serie": _txt(celda(fila, j - 4)).lstrip("0") or _txt(celda(fila, j - 4)),
                "Número": _sin_ceros(celda(fila, j - 2)),
                "Base imponible": _a_num(celda(fila, j + 2)),
                "IGV": _a_num(celda(fila, j + 3)),
                "Total": _a_num(celda(fila, j + 11)),
            })
        else:
            fecha = next((_a_fecha(v) for v in fila if isinstance(v, (datetime, pd.Timestamp))), pd.NaT)
            reg["Fecha emisión"] = fecha
        registros.append(reg)
    return pd.DataFrame(registros)


def leer_archivo_compras(nombre: str, data: bytes) -> ResultadoLectura:
    """Lee un archivo de compras de cualquier formato soportado."""
    nombre_l = nombre.lower()
    avisos: list[str] = []
    partes: list[pd.DataFrame] = []

    if nombre_l.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                sub = leer_archivo_compras(f"{nombre}/{info.filename}", z.read(info))
                partes.append(sub.compras)
                avisos += sub.avisos
    elif nombre_l.endswith((".csv", ".txt")):
        df = _tabla_desde_texto(data)
        if _es_sire(df):
            partes.append(_parse_sire(df, nombre))
        else:
            raw = pd.read_csv(io.StringIO(_leer_texto(data)), sep=None, engine="python",
                              header=None, dtype=str, keep_default_na=False)
            p = _parse_tabla_generica(raw, nombre, "")
            if p is not None:
                partes.append(p)
    elif nombre_l.endswith((".xlsx", ".xlsm", ".xls")):
        hojas = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)
        for hoja, raw in hojas.items():
            # ¿Es una propuesta SIRE guardada como Excel?
            cab = next((i for i in range(min(15, len(raw)))
                        if "NRO DOC IDENTIDAD" in {_norm(v) for v in raw.iloc[i]}), None)
            if cab is not None:
                df = raw.iloc[cab + 1:].copy()
                df.columns = [_txt(v) for v in raw.iloc[cab]]
                partes.append(_parse_sire(df.astype(object).where(df.notna(), ""), nombre))
                continue
            p = _parse_tabla_generica(raw, nombre, str(hoja))
            if p is not None:
                partes.append(p)
    else:
        avisos.append(f"{nombre}: formato no soportado (use .xlsx, .csv, .txt o .zip).")

    partes = [p for p in partes if p is not None and not p.empty]
    if not partes and not avisos:
        avisos.append(f"{nombre}: no se encontraron RUCs de proveedores.")
    compras = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=COLUMNAS_NORMALIZADAS)
    for c in COLUMNAS_NORMALIZADAS:
        if c not in compras.columns:
            compras[c] = None
    compras = compras[compras["RUC proveedor"] != ""].copy()
    fechas = pd.to_datetime(compras["Fecha emisión"], errors="coerce")
    sin_periodo = compras["Periodo"].fillna("").astype(str).str.strip() == ""
    compras.loc[sin_periodo, "Periodo"] = fechas[sin_periodo].dt.strftime("%Y%m").fillna("")
    return ResultadoLectura(compras[COLUMNAS_NORMALIZADAS].reset_index(drop=True), avisos)


# --------------------------------------------------------------------------- padrón SSCO
def descargar_excel(urls: list[str], timeout: int = 60) -> bytes:
    """Descarga el primer Excel válido de la lista de URLs."""
    errores = []
    for url in urls:
        try:
            r = requests.get(url, timeout=timeout, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
            })
            r.raise_for_status()
            if not r.content.startswith(b"PK"):
                raise ValueError("SUNAT no devolvió un archivo Excel válido (posible bloqueo o mantenimiento).")
            return r.content
        except Exception as e:  # noqa: BLE001
            errores.append(f"{url}: {e}")
    raise RuntimeError("\n".join(errores))


def descargar_padron(timeout: int = 60) -> bytes:
    """Descarga el Excel oficial del padrón SSCO desde SUNAT."""
    return descargar_excel([SSCO_URL], timeout)


def descargar_relacion(timeout: int = 60) -> bytes:
    """Descarga el Excel de la relación del art. 12 DL 1532 (vinculados a SSCO)."""
    return descargar_excel(RELACION_URLS, timeout)


def _fecha_txt(v) -> str:
    return f"{v:%d/%m/%Y}" if pd.notna(v) else ""


def leer_relacion(data: bytes, nombre: str = "relacion.xlsx") -> pd.DataFrame:
    """Lee la relación del art. 12 DL 1532 (una fila por empresa vinculada y SSCO que la
    origina) y devuelve una fila por RUC vinculado."""
    if nombre.lower().endswith((".csv", ".txt")):
        hojas = {"": pd.read_csv(io.StringIO(_leer_texto(data)), sep=None, engine="python",
                                 header=None, dtype=str, keep_default_na=False)}
    else:
        hojas = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)

    partes = []
    for _, raw in hojas.items():
        cab = next((i for i in range(min(20, len(raw)))
                    if any(_norm(v) in ("NUMERO DE RUC", "RUC") for v in raw.iloc[i])), None)
        if cab is None:
            continue
        nombres = [_norm(v) for v in raw.iloc[cab]]
        df = raw.iloc[cab + 1:].reset_index(drop=True)

        def col(pred):
            j = next((k for k, n in enumerate(nombres) if pred(n)), None)
            return df.iloc[:, j] if j is not None else pd.Series([None] * len(df))

        ruc = col(lambda n: n in ("NUMERO DE RUC", "RUC")).map(limpiar_ruc)
        partes.append(pd.DataFrame({
            "RUC": ruc,
            "Razón social": col(lambda n: n.startswith("DENOMINACION O RAZON SOCIAL")
                                and "REPRESENTANTE" not in n and "SSCO" not in n).map(_txt),
            "Titular": col(lambda n: "DEL TITULAR" in n and "NOMBRE" in n).map(_txt),
            "Representante legal": col(lambda n: "REPRESENTANTE" in n and "RAZON" in n).map(_txt),
            "Supuesto de inclusión": col(lambda n: n.startswith("SUPUESTO")).map(
                lambda v: re.sub(r"^\d+\s*-\s*", "", _txt(v))),
            "SSCO que origina": col(lambda n: "SSCO QUE ORIGINA" in n and "RAZON" in n).map(_txt),
            "RUC SSCO que origina": col(lambda n: "RUC DEL SSCO" in n).map(limpiar_ruc),
            "Fecha límite restricción": col(lambda n: n.startswith("FECHA LIMITE")).map(_a_fecha),
            "Fecha de publicación": col(lambda n: n.startswith("FECHA DE PUBLICACION")).map(_a_fecha),
        }))
    if not partes:
        raise ValueError("No se encontró la columna 'Número de RUC' en la relación.")
    det = pd.concat(partes, ignore_index=True)
    det = det[det["RUC"] != ""]

    def unir(s):
        return " / ".join(dict.fromkeys(v for v in s if v))

    det["SSCO que origina"] = det.apply(
        lambda f: f"{f['SSCO que origina']} ({f['RUC SSCO que origina']})" if f["RUC SSCO que origina"]
        else f["SSCO que origina"], axis=1)
    rel = det.groupby("RUC", as_index=False, sort=False).agg(**{
        "Razón social": ("Razón social", "first"),
        "Titular": ("Titular", unir),
        "Representante legal": ("Representante legal", unir),
        "Supuesto(s) de inclusión": ("Supuesto de inclusión", unir),
        "SSCO que origina la inclusión": ("SSCO que origina", unir),
        "Fecha de publicación": ("Fecha de publicación", "min"),
        "Fecha límite restricción": ("Fecha límite restricción", "max"),
    })
    for c in ("Fecha de publicación", "Fecha límite restricción"):
        rel[c] = pd.to_datetime(rel[c], errors="coerce").dt.date
    return rel


def cruzar_vinculados(compras: pd.DataFrame, relacion: pd.DataFrame) -> pd.DataFrame:
    """Cruce con la relación art. 12: empresas vinculadas a SSCO que, desde el día siguiente a
    su publicación y hasta la fecha límite, solo pueden emitir boletas de venta y notas
    de crédito/débito electrónicas."""
    rel = relacion.rename(columns={c: f"Art.12 · {c}" for c in relacion.columns if c != "RUC"})
    al = compras.merge(rel, left_on="RUC proveedor", right_on="RUC", how="inner").drop(columns="RUC")
    if al.empty:
        al.insert(0, "Nivel", [])
        al.insert(1, "Observación", [])
        return al

    def evaluar(f):
        fe = pd.to_datetime(f["Fecha emisión"], errors="coerce")
        pub = pd.to_datetime(f["Art.12 · Fecha de publicación"], errors="coerce")
        lim = pd.to_datetime(f["Art.12 · Fecha límite restricción"], errors="coerce")
        periodo = f"desde el {_fecha_txt(pub + pd.Timedelta(days=1)) if pd.notna(pub) else '?'}" + \
                  (f" hasta el {_fecha_txt(lim)}" if pd.notna(lim) else "")
        if pd.isna(fe) or pd.isna(pub):
            return "Medio", "Proveedor vinculado a un SSCO (art. 12 DL 1532). Revisar el comprobante."
        if fe > pub and (pd.isna(lim) or fe <= lim):
            if str(f["Tipo CP"]).zfill(2) == "01":
                return ("Alto", f"FACTURA EMITIDA DURANTE LA RESTRICCIÓN: {periodo} el proveedor solo puede "
                                "emitir boletas de venta y notas de crédito/débito electrónicas (art. 12 DL 1532). "
                                "Verifique la validez del comprobante y del crédito fiscal.")
            return ("Medio", f"Proveedor en periodo de restricción ({periodo}): solo puede emitir boletas de "
                             "venta y notas electrónicas (art. 12 DL 1532).")
        return ("Informativo", f"Proveedor vinculado a un SSCO (art. 12 DL 1532), publicado el {_fecha_txt(pub)}. "
                               "Comprobante emitido fuera del periodo de restricción; evalúe el riesgo del proveedor.")

    res = al.apply(evaluar, axis=1, result_type="expand")
    al.insert(0, "Nivel", res[0])
    al.insert(1, "Observación", res[1])
    orden = {"Alto": 0, "Medio": 1, "Informativo": 2}
    return al.sort_values("Nivel", key=lambda s: s.map(orden), kind="stable").reset_index(drop=True)


def leer_padron(data: bytes, nombre: str = "padron.xlsx") -> pd.DataFrame:
    """Devuelve el padrón con una columna 'RUC' normalizada + columnas originales."""
    if nombre.lower().endswith((".csv", ".txt")):
        hojas = {"": pd.read_csv(io.StringIO(_leer_texto(data)), sep=None, engine="python",
                                 header=None, dtype=str, keep_default_na=False)}
    else:
        hojas = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object)

    partes = []
    for _, raw in hojas.items():
        raw = raw.dropna(how="all", axis=1).dropna(how="all", axis=0)
        if raw.empty:
            continue
        j = _columna_ruc(raw)
        if j is None:
            continue
        raw = raw.T.reset_index(drop=True).T
        es_ruc = raw.iloc[:, j].map(limpiar_ruc) != ""
        primera = es_ruc.idxmax()
        pos = raw.index.get_loc(primera)
        # Encabezado: fila(s) de texto inmediatamente encima de los datos.
        cab = [""] * raw.shape[1]
        for k in range(max(0, pos - 3), pos):
            for c in range(raw.shape[1]):
                v = _txt(raw.iloc[k, c])
                if v and not limpiar_ruc(v):
                    cab[c] = (cab[c] + " " + v).strip()
        nombres, vistos = [], set()
        for c, n in enumerate(cab):
            n = n or f"Columna {c + 1}"
            while n in vistos:
                n += "_"
            vistos.add(n)
            nombres.append(n)
        datos = raw[es_ruc].copy()
        datos.columns = nombres
        valores_ruc = datos.iloc[:, j].map(limpiar_ruc).values
        datos = datos.drop(columns=[nombres[j]])
        datos.insert(0, "RUC", valores_ruc)
        for c in datos.columns:
            if datos[c].map(lambda v: isinstance(v, (datetime, pd.Timestamp))).any():
                datos[c] = pd.to_datetime(datos[c], errors="coerce").dt.date
            elif c != "RUC":
                datos[c] = datos[c].map(_txt)
        partes.append(datos)
    if not partes:
        raise ValueError("No se encontró ninguna columna con RUCs en el padrón.")
    padron = pd.concat(partes, ignore_index=True)
    padron = padron.dropna(axis=1, how="all")
    return padron.drop_duplicates().reset_index(drop=True)


def columna_razon_social(padron: pd.DataFrame) -> str | None:
    for c in padron.columns:
        n = _norm(c)
        if any(k in n for k in ("RAZON", "DENOMINACION", "NOMBRE")) and "REPRESENT" not in n:
            return c
    otras = [c for c in padron.columns if c != "RUC" and padron[c].map(lambda v: isinstance(v, str)).mean() > 0.8]
    return otras[0] if otras else None


# --------------------------------------------------------------------------- cruce
def cruzar(compras: pd.DataFrame, padron: pd.DataFrame) -> pd.DataFrame:
    """Coincidencia exacta por RUC entre compras y padrón SSCO."""
    pad = padron.rename(columns={c: f"SSCO · {c}" for c in padron.columns if c != "RUC"})
    pad = pad.groupby("RUC", as_index=False).agg(
        lambda s: " / ".join(dict.fromkeys(_txt(v) for v in s if _txt(v))))
    alertas = compras.merge(pad, left_on="RUC proveedor", right_on="RUC", how="inner").drop(columns="RUC")
    if alertas.empty:
        alertas["Observación"] = []
        return alertas

    def observacion(f):
        fe = f["Fecha emisión"]
        if pd.notna(fe) and pd.Timestamp(fe).date() < FECHA_INICIO_DL1532:
            return "Fuera de alcance: comprobante anterior al 20/03/2022 (1ra DCF DL 1532)."
        return ("PROVEEDOR SSCO: el comprobante no otorga crédito fiscal ni sustenta costo/gasto "
                "para IR (arts. 8 y 9 DL 1532).")

    alertas.insert(0, "Observación", alertas.apply(observacion, axis=1))

    col_pub = next((c for c in padron.columns if "PUBLICACI" in _norm(c)), None)
    if col_pub:
        def momento(f):
            fe = pd.to_datetime(f["Fecha emisión"], errors="coerce")
            fechas_pub = [pd.to_datetime(x.strip(), errors="coerce")
                          for x in str(f[f"SSCO · {col_pub}"]).split(" / ")]
            fechas_pub = [x for x in fechas_pub if pd.notna(x)]
            if pd.isna(fe) or not fechas_pub:
                return ""
            pub = min(fechas_pub)
            if fe <= pub:
                return f"Emitido antes de la publicación ({pub:%d/%m/%Y}) – art. 9"
            return f"Emitido después de la publicación ({pub:%d/%m/%Y}) – art. 8"
        alertas.insert(1, "Emisión vs. publicación SSCO", alertas.apply(momento, axis=1))
    return alertas


def resumen_por_proveedor(alertas: pd.DataFrame) -> pd.DataFrame:
    if alertas.empty:
        return pd.DataFrame(columns=["RUC proveedor", "Razón social (compras)", "Comprobantes", "Base imponible", "IGV", "Total"])
    return (alertas.groupby("RUC proveedor", as_index=False)
            .agg(**{"Razón social (compras)": ("Razón social (compras)", "first"),
                    "Comprobantes": ("RUC proveedor", "size"),
                    "Base imponible": ("Base imponible", "sum"),
                    "IGV": ("IGV", "sum"),
                    "Total": ("Total", "sum")})
            .sort_values("Total", ascending=False))


def resumen_por_ejercicio(compras: pd.DataFrame, alertas: pd.DataFrame) -> pd.DataFrame:
    def ejercicio(df):
        return df["Periodo"].fillna("").astype(str).str[:4].replace("", "Sin periodo")
    rev = compras.assign(Ejercicio=ejercicio(compras)).groupby("Ejercicio").size().rename("Comprobantes revisados")
    if alertas.empty:
        obs = pd.DataFrame(columns=["Comprobantes con SSCO", "IGV observado", "Total observado"])
    else:
        obs = (alertas.assign(Ejercicio=ejercicio(alertas)).groupby("Ejercicio")
               .agg(**{"Comprobantes con SSCO": ("RUC proveedor", "size"),
                       "IGV observado": ("IGV", "sum"), "Total observado": ("Total", "sum")}))
    out = pd.concat([rev, obs], axis=1).fillna(0).reset_index().rename(columns={"index": "Ejercicio"})
    out["Comprobantes con SSCO"] = out["Comprobantes con SSCO"].astype(int)
    return out.sort_values("Ejercicio")


def buscar_rucs(texto: str, padron: pd.DataFrame, relacion: pd.DataFrame | None = None) -> pd.DataFrame:
    """Búsqueda manual de uno o varios RUCs (separados por espacios, comas o saltos de línea)."""
    tokens = [t for t in re.split(r"[\s,;]+", texto or "") if t]
    filas = []
    indice = padron.set_index("RUC")
    ind_rel = relacion.set_index("RUC") if relacion is not None else None
    for t in dict.fromkeys(tokens):
        ruc = limpiar_ruc(t)
        if not ruc:
            filas.append({"RUC consultado": t, "Resultado": "⚠️ Formato de RUC inválido"})
            continue
        fila = {"RUC consultado": ruc}
        if ruc in indice.index:
            datos = indice.loc[[ruc]].iloc[0].to_dict()
            fila["Resultado"] = "🚨 FIGURA EN EL PADRÓN SSCO"
            fila.update({k: _txt(v) for k, v in datos.items()})
        else:
            fila["Resultado"] = "✅ No figura en el padrón SSCO"
        if ind_rel is not None:
            if ruc in ind_rel.index:
                r = ind_rel.loc[ruc]
                fila["Resultado"] += " · ⚠️ VINCULADO A SSCO (art. 12)"
                if not fila.get("Razón social"):
                    fila["Razón social"] = r["Razón social"]
                fila.update({f"Art.12 · {k}": (_fecha_txt(v) if isinstance(v, date) else _txt(v))
                             for k, v in r.items() if k != "Razón social"})
            else:
                fila["Resultado"] += " · no figura en la relación art. 12"
        if not ruc_digito_valido(ruc):
            fila["Resultado"] += " (dígito verificador no válido)"
        filas.append(fila)
    return pd.DataFrame(filas)


def exportar_excel(alertas: pd.DataFrame, resumen: pd.DataFrame, compras: pd.DataFrame,
                   info: dict, por_ejercicio: pd.DataFrame | None = None,
                   vinculados: pd.DataFrame | None = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as xw:
        hojas = [("Alertas SSCO", alertas), ("Resumen proveedores", resumen)]
        if vinculados is not None:
            hojas.append(("Vinculados art. 12", vinculados))
        if por_ejercicio is not None:
            hojas.append(("Resumen por ejercicio", por_ejercicio))
        hojas.append(("Compras revisadas", compras))
        for nombre, df in hojas:
            df.to_excel(xw, sheet_name=nombre, index=False)
            hoja = xw.sheets[nombre]
            hoja.freeze_panes(1, 0)
            for i, c in enumerate(df.columns):
                ancho = min(60, max(10, len(str(c)), *(len(_txt(v)) for v in df[c].head(200))) if len(df) else len(str(c)) + 2)
                hoja.set_column(i, i, ancho + 1)
        pd.DataFrame(list(info.items()), columns=["Dato", "Valor"]).to_excel(xw, sheet_name="Información", index=False)
        xw.sheets["Información"].set_column(0, 1, 45)
    return buf.getvalue()
