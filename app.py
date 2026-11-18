"""App Streamlit: cruce del Registro de Compras con el padrón SUNAT de Sujetos sin Capacidad Operativa (SSCO)."""
from __future__ import annotations

import hmac
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import ssco_core as core

st.set_page_config(page_title="Cruce SSCO – Registro de Compras", page_icon="🔎", layout="wide")

PADRON_LOCAL = Path(__file__).parent / "data" / "sujesincapacidadOperativa.xlsx"
TZ = ZoneInfo("America/Lima")


# --------------------------------------------------------------------------- acceso opcional
def verificar_acceso() -> bool:
    """Si se define APP_PASSWORD en los secrets de Streamlit, se pide contraseña."""
    try:
        clave = st.secrets.get("APP_PASSWORD")
    except Exception:
        clave = None
    if not clave or st.session_state.get("autenticado"):
        return True
    st.title("🔒 Cruce SSCO")
    pwd = st.text_input("Contraseña", type="password")
    if pwd:
        if hmac.compare_digest(pwd, str(clave)):
            st.session_state["autenticado"] = True
            st.rerun()
        st.error("Contraseña incorrecta")
    return False


if not verificar_acceso():
    st.stop()


# --------------------------------------------------------------------------- padrón
@st.cache_data(ttl=6 * 3600, show_spinner="Descargando padrón SSCO desde SUNAT…")
def padron_sunat() -> tuple[pd.DataFrame, str]:
    data = core.descargar_padron()
    return core.leer_padron(data), datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


@st.cache_data(show_spinner=False)
def padron_desde_bytes(data: bytes, nombre: str) -> pd.DataFrame:
    return core.leer_padron(data, nombre)


def obtener_padron() -> tuple[pd.DataFrame | None, str]:
    with st.sidebar:
        st.header("📋 Padrón SSCO")
        fuente = st.radio(
            "Fuente del padrón",
            ["Descargar de SUNAT (automático)", "Copia guardada en la app", "Subir archivo manualmente"],
            help="Si SUNAT no responde, use la copia guardada o suba el Excel descargado de "
                 "SUNAT › Padrones y Notificaciones › SSCO.",
        )
        padron, origen = None, ""
        if fuente.startswith("Descargar"):
            try:
                padron, hora = padron_sunat()
                origen = f"SUNAT en línea (descargado {hora})"
                if st.button("🔄 Volver a descargar"):
                    padron_sunat.clear()
                    st.rerun()
            except Exception as e:  # noqa: BLE001
                st.warning("SUNAT no respondió; se usa la copia guardada en la app.")
                with st.expander("Detalle técnico"):
                    st.code(str(e))
                fuente = "Copia guardada en la app"
        if fuente.startswith("Copia"):
            if PADRON_LOCAL.exists():
                padron = padron_desde_bytes(PADRON_LOCAL.read_bytes(), PADRON_LOCAL.name)
                origen = f"Copia guardada ({PADRON_LOCAL.name})"
            else:
                st.error("No hay copia guardada; suba el archivo manualmente.")
        if fuente.startswith("Subir"):
            archs = st.file_uploader(
                "Excel(es) del padrón SSCO", type=["xlsx", "xls", "csv"], key="padron",
                accept_multiple_files=True,
                help="Puede subir varias versiones (p. ej. padrones de años anteriores): se combinan.")
            if archs:
                padron = pd.concat([padron_desde_bytes(a.getvalue(), a.name) for a in archs], ignore_index=True)
                padron = padron.astype(str).replace({"nan": "", "NaT": "", "None": ""}).drop_duplicates()
                origen = "Archivo(s) subido(s): " + ", ".join(a.name for a in archs)

        if padron is not None:
            st.success(f"{len(padron):,} sujetos SSCO cargados")
            pub = next((c for c in padron.columns if "publicaci" in c.lower()), None)
            if pub:
                ultima = pd.to_datetime(padron[pub], errors="coerce").max()
                if pd.notna(ultima):
                    st.caption(f"Última publicación incluida: **{ultima:%d/%m/%Y}**")
            st.caption(f"Fuente: {origen}")
        st.divider()
        st.caption("Cruce por coincidencia **exacta** de RUC. Comprobantes anteriores al 20/03/2022 "
                   "se marcan como fuera de alcance del D. Leg. 1532.")
    return padron, origen


padron, origen_padron = obtener_padron()

st.title("🔎 Cruce de compras vs. Sujetos sin Capacidad Operativa (SSCO)")
if padron is None:
    st.info("Cargue el padrón SSCO desde la barra lateral para comenzar.")
    st.stop()

tab_cruce, tab_ruc, tab_padron, tab_ayuda = st.tabs(
    ["📂 Cruzar registro de compras", "🔍 Buscar RUC", "📋 Ver padrón", "ℹ️ Ayuda"])


def formato_tabla(df: pd.DataFrame) -> dict:
    cfg = {}
    for c in ("Base imponible", "IGV", "Total"):
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f")
    if "Fecha emisión" in df.columns:
        cfg["Fecha emisión"] = st.column_config.DateColumn("Fecha emisión", format="DD/MM/YYYY")
    return cfg


# --------------------------------------------------------------------------- cruce masivo
with tab_cruce:
    st.markdown(
        "Suba uno o varios archivos: **Excel del software contable** (Formato 8.1), "
        "**Excel `resultado_…`**, o la **propuesta RCE del SIRE** (`.zip`, `.csv` o `.txt`).")
    archivos = st.file_uploader("Archivos de compras", accept_multiple_files=True,
                                type=["xlsx", "xls", "xlsm", "csv", "txt", "zip"], key="compras")
    if archivos:
        partes, avisos = [], []
        for a in archivos:
            try:
                r = core.leer_archivo_compras(a.name, a.getvalue())
                partes.append(r.compras)
                avisos += r.avisos
            except Exception as e:  # noqa: BLE001
                avisos.append(f"{a.name}: error al leer ({e})")
        for av in avisos:
            st.warning(av)
        compras = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=core.COLUMNAS_NORMALIZADAS)

        if compras.empty:
            st.error("No se encontraron comprobantes con RUC de proveedor en los archivos.")
        else:
            ejercicios = sorted({p[:4] for p in compras["Periodo"].astype(str) if p[:4].isdigit()})
            if len(ejercicios) > 1:
                elegidos = st.multiselect("Ejercicios a revisar", ejercicios, default=ejercicios)
                compras = compras[compras["Periodo"].astype(str).str[:4].isin(elegidos)
                                  | ~compras["Periodo"].astype(str).str[:4].str.isdigit()]
            alertas = core.cruzar(compras, padron)
            resumen = core.resumen_por_proveedor(alertas)
            por_ejercicio = core.resumen_por_ejercicio(compras, alertas)
            en_alcance = alertas[alertas["Observación"].str.startswith("PROVEEDOR")] if not alertas.empty else alertas

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Comprobantes revisados", f"{len(compras):,}")
            c2.metric("Proveedores distintos", f"{compras['RUC proveedor'].nunique():,}")
            c3.metric("Comprobantes con SSCO", f"{len(alertas):,}")
            c4.metric("IGV en riesgo", f"S/ {en_alcance['IGV'].fillna(0).sum():,.2f}" if len(en_alcance) else "S/ 0.00")

            if alertas.empty:
                st.success("✅ Ningún proveedor del registro de compras figura en el padrón SSCO.")
            else:
                st.error(f"🚨 {resumen.shape[0]} proveedor(es) figuran en el padrón SSCO.")
                if len(por_ejercicio) > 1:
                    st.subheader("Resumen por ejercicio")
                    st.dataframe(por_ejercicio, hide_index=True, use_container_width=True,
                                 column_config={c: st.column_config.NumberColumn(c, format="%.2f")
                                                for c in ("IGV observado", "Total observado")})
                st.subheader("Resumen por proveedor")
                st.dataframe(resumen, hide_index=True, use_container_width=True, column_config=formato_tabla(resumen))
                st.subheader("Detalle de comprobantes observados")
                st.dataframe(alertas, hide_index=True, use_container_width=True, column_config=formato_tabla(alertas))

            info = {
                "Fecha y hora del cruce": datetime.now(TZ).strftime("%d/%m/%Y %H:%M"),
                "Fuente del padrón SSCO": origen_padron,
                "Sujetos en el padrón": len(padron),
                "Archivos procesados": ", ".join(a.name for a in archivos),
                "Ejercicios revisados": ", ".join(sorted(por_ejercicio["Ejercicio"].astype(str))),
                "Comprobantes revisados": len(compras),
                "Comprobantes observados": len(alertas),
                "Criterio": "Coincidencia exacta de RUC; comprobantes desde el 20/03/2022 (D. Leg. 1532).",
            }
            st.download_button(
                "⬇️ Descargar reporte Excel",
                data=core.exportar_excel(alertas, resumen, compras, info, por_ejercicio),
                file_name=f"cruce_ssco_{datetime.now(TZ):%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            with st.expander(f"Ver los {len(compras):,} comprobantes leídos"):
                st.dataframe(compras, hide_index=True, use_container_width=True, column_config=formato_tabla(compras))

# --------------------------------------------------------------------------- búsqueda manual
with tab_ruc:
    st.markdown("Ingrese uno o varios RUC (separados por espacio, coma o salto de línea).")
    texto = st.text_area("RUC(s)", placeholder="20123456789\n10456789012", height=120)
    if st.button("Buscar", type="primary") or texto:
        res = core.buscar_rucs(texto, padron)
        if res.empty:
            st.info("Ingrese al menos un RUC.")
        else:
            for _, f in res.iterrows():
                msg = f"**{f['RUC consultado']}** — {f['Resultado']}"
                if "FIGURA" in f["Resultado"]:
                    st.error(msg)
                elif "inválido" in f["Resultado"]:
                    st.warning(msg)
                else:
                    st.success(msg)
            hallados = res[res["Resultado"].str.contains("FIGURA")]
            if not hallados.empty:
                st.dataframe(hallados.dropna(axis=1, how="all"), hide_index=True, use_container_width=True)

# --------------------------------------------------------------------------- padrón completo
with tab_padron:
    filtro = st.text_input("Filtrar por RUC o razón social")
    vista = padron
    if filtro:
        mask = padron.astype(str).apply(lambda s: s.str.contains(filtro, case=False, na=False)).any(axis=1)
        vista = padron[mask]
    st.caption(f"{len(vista):,} de {len(padron):,} registros")
    st.dataframe(vista, hide_index=True, use_container_width=True)

# --------------------------------------------------------------------------- ayuda
with tab_ayuda:
    st.markdown("""
**¿Qué hace esta app?** Cruza los RUC de los proveedores de su registro de compras con la
*Relación de Sujetos sin Capacidad Operativa con resoluciones de atribución firmes* que publica SUNAT.

**Formatos de compras aceptados (se detectan solos):**
- Excel del software contable – *Formato 8.1 Registro de Compras*.
- Excel `resultado_AAAAMMDD_HHMMSS.xlsx` (hoja *Comprobantes*, sin encabezados).
- Propuesta RCE del SIRE: el `.zip` descargado tal cual, o el `.csv` / `.txt` que contiene.
- Cualquier otro Excel/CSV con una columna de RUC (solo se toman RUC y razón social).

**Años anteriores:** puede subir registros de cualquier ejercicio (varios archivos o un Excel con una
hoja por mes). La app identifica el periodo, resume por ejercicio y permite filtrar. Para cada
comprobante observado indica si se emitió **antes** (art. 9) o **después** (art. 8) de la publicación
del SSCO. Como SUNAT mantiene cada publicación en su web solo 4 años, para revisiones antiguas
conviene subir también versiones anteriores del padrón (opción *Subir archivo manualmente*, admite varios).

**Efectos (D. Leg. 1532, arts. 8 y 9):** los comprobantes emitidos por un SSCO desde el 20/03/2022
no otorgan crédito fiscal ni sustentan costo o gasto para el Impuesto a la Renta. El adquirente puede
presentar una *Solicitud de revisión* dentro de los 30 días hábiles siguientes a la publicación.

**Padrón:** se descarga de
[SUNAT]({url}) y se guarda en caché por 6 horas. Si SUNAT no responde, se usa la copia guardada
en el repositorio (`data/`), que un proceso de GitHub Actions intenta actualizar cada mes.

> Herramienta de apoyo; no sustituye la revisión de un especialista tributario.
""".replace("{url}", core.SSCO_URL))
