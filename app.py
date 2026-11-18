"""App Streamlit: cruce del Registro de Compras con el padrón SUNAT de Sujetos sin Capacidad Operativa (SSCO)."""
from __future__ import annotations

import hmac
import importlib
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import ssco_core as core

# Streamlit Cloud puede volver a ejecutar app.py tras un `git pull` sin recargar los módulos
# importados; recargarlo evita usar una versión desactualizada de ssco_core.
core = importlib.reload(core)

st.set_page_config(page_title="Cruce SSCO – Registro de Compras", page_icon="🔎", layout="wide")

PADRON_LOCAL = Path(__file__).parent / "data" / "sujesincapacidadOperativa.xlsx"
RELACION_LOCAL = Path(__file__).parent / "data" / "relacion-ssco.xlsx"
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


# --------------------------------------------------------------------------- listas SUNAT
LISTAS = {
    "padron": dict(
        titulo="📋 Padrón SSCO", corto="padrón SSCO", unidad="sujetos SSCO cargados",
        descargar=core.descargar_padron, leer=core.leer_padron, local=PADRON_LOCAL, pagina=core.SSCO_PAGINA,
        ayuda="SUNAT › Padrones y Notificaciones › Sujetos sin Capacidad Operativa."),
    "relacion": dict(
        titulo="🔗 Vinculados a SSCO (art. 12)", corto="relación art. 12", unidad="empresas vinculadas cargadas",
        descargar=core.descargar_relacion, leer=core.leer_relacion, local=RELACION_LOCAL, pagina=core.RELACION_PAGINA,
        ayuda="SUNAT › Padrones y Notificaciones › Relación art. 12 DL 1532 (EIRL y sociedades "
              "cuyos titulares o socios son SSCO)."),
}


@st.cache_data(ttl=6 * 3600, show_spinner="Descargando de SUNAT…")
def lista_sunat(clave: str) -> tuple[pd.DataFrame, str]:
    cfg = LISTAS[clave]
    return cfg["leer"](cfg["descargar"]()), datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


@st.cache_data(show_spinner=False)
def lista_desde_bytes(clave: str, data: bytes, nombre: str) -> pd.DataFrame:
    return LISTAS[clave]["leer"](data, nombre)


def obtener_lista(clave: str) -> tuple[pd.DataFrame | None, str]:
    cfg = LISTAS[clave]
    st.header(cfg["titulo"])
    st.link_button("🌐 Abrir página de SUNAT", cfg["pagina"], use_container_width=True,
                   help="Allí puede ver la lista y descargar el Excel con «Descargue la información aquí».")
    fuente = st.radio(
        "Fuente", ["Descargar de SUNAT (automático)", "Copia guardada en la app", "Subir archivo manualmente"],
        key=f"fuente_{clave}",
        help=f"Si SUNAT no responde, use la copia guardada o suba el Excel descargado de {cfg['ayuda']}")
    df, origen = None, ""
    if fuente.startswith("Descargar"):
        try:
            df, hora = lista_sunat(clave)
            origen = f"SUNAT en línea (descargado {hora})"
            if st.button("🔄 Volver a descargar", key=f"recargar_{clave}"):
                lista_sunat.clear()
                st.rerun()
        except Exception as e:  # noqa: BLE001
            st.warning("SUNAT no respondió; se usa la copia guardada en la app.")
            with st.expander("Detalle técnico"):
                st.code(str(e))
            fuente = "Copia guardada en la app"
    if fuente.startswith("Copia"):
        if cfg["local"].exists():
            df = lista_desde_bytes(clave, cfg["local"].read_bytes(), cfg["local"].name)
            origen = f"Copia guardada ({cfg['local'].name})"
        else:
            st.error("No hay copia guardada; suba el archivo manualmente.")
    if fuente.startswith("Subir"):
        st.caption(f"Descargue el Excel desde la [página de SUNAT]({cfg['pagina']}) "
                   "(enlace «Descargue la información aquí») y súbalo aquí.")
        archs = st.file_uploader(
            f"Excel(es) de la {cfg['corto']}", type=["xlsx", "xls", "csv"], key=f"subir_{clave}",
            accept_multiple_files=True,
            help="Puede subir varias versiones (p. ej. de años anteriores): se combinan.")
        if archs:
            df = pd.concat([lista_desde_bytes(clave, a.getvalue(), a.name) for a in archs], ignore_index=True)
            if clave == "padron":
                df = df.astype(str).replace({"nan": "", "NaT": "", "None": ""})
            df = df.drop_duplicates(subset="RUC" if clave == "relacion" else None)
            origen = "Archivo(s) subido(s): " + ", ".join(a.name for a in archs)

    if df is not None:
        st.success(f"{len(df):,} {cfg['unidad']}")
        pub = next((c for c in df.columns if "publicaci" in c.lower()), None)
        if pub:
            ultima = pd.to_datetime(df[pub], errors="coerce").max()
            if pd.notna(ultima):
                st.caption(f"Última publicación incluida: **{ultima:%d/%m/%Y}**")
        st.caption(f"Fuente: {origen}")
    return df, origen


with st.sidebar:
    padron, origen_padron = obtener_lista("padron")
    st.divider()
    relacion, origen_relacion = obtener_lista("relacion")
    st.divider()
    st.caption("Cruce por coincidencia **exacta** de RUC. Comprobantes anteriores al 20/03/2022 "
               "se marcan como fuera de alcance del D. Leg. 1532.")

st.title("🔎 Cruce de compras vs. Sujetos sin Capacidad Operativa (SSCO)")
if padron is None:
    st.info("Cargue el padrón SSCO desde la barra lateral para comenzar.")
    st.stop()

tab_cruce, tab_ruc, tab_padron, tab_ayuda = st.tabs(
    ["📂 Cruzar registro de compras", "🔍 Buscar RUC", "📋 Ver listas", "ℹ️ Ayuda"])


def formato_tabla(df: pd.DataFrame) -> dict:
    cfg = {}
    for c in ("Base imponible", "IGV", "Total"):
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f")
    for c in df.columns:
        if c == "Fecha emisión" or c.startswith("Art.12 · Fecha"):
            cfg[c] = st.column_config.DateColumn(c, format="DD/MM/YYYY")
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
            vinculados = core.cruzar_vinculados(compras, relacion) if relacion is not None else None

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Comprobantes revisados", f"{len(compras):,}")
            c2.metric("Proveedores distintos", f"{compras['RUC proveedor'].nunique():,}")
            c3.metric("Comprobantes con SSCO", f"{len(alertas):,}")
            c4.metric("IGV en riesgo (SSCO)", f"S/ {en_alcance['IGV'].fillna(0).sum():,.2f}" if len(en_alcance) else "S/ 0.00")
            c5.metric("Con vinculados art. 12", f"{len(vinculados):,}" if vinculados is not None else "—")

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

            st.subheader("🔗 Proveedores vinculados a SSCO (art. 12 DL 1532)")
            if vinculados is None:
                st.info("Cargue la relación art. 12 en la barra lateral para incluir este control.")
            elif vinculados.empty:
                st.success("✅ Ningún proveedor figura en la relación de empresas vinculadas a SSCO (art. 12).")
            else:
                altos = int((vinculados["Nivel"] == "Alto").sum())
                msg = f"{vinculados['RUC proveedor'].nunique()} proveedor(es) figuran como vinculados a un SSCO."
                (st.error if altos else st.warning)(
                    msg + (f" **{altos} factura(s) emitidas durante el periodo de restricción.**" if altos else ""))
                st.dataframe(vinculados, hide_index=True, use_container_width=True,
                             column_config=formato_tabla(vinculados))

            info = {
                "Fecha y hora del cruce": datetime.now(TZ).strftime("%d/%m/%Y %H:%M"),
                "Fuente del padrón SSCO": origen_padron,
                "Sujetos en el padrón": len(padron),
                "Fuente de la relación art. 12": origen_relacion or "No cargada",
                "Empresas en la relación art. 12": len(relacion) if relacion is not None else 0,
                "Comprobantes con vinculados art. 12": len(vinculados) if vinculados is not None else 0,
                "Archivos procesados": ", ".join(a.name for a in archivos),
                "Ejercicios revisados": ", ".join(sorted(por_ejercicio["Ejercicio"].astype(str))),
                "Comprobantes revisados": len(compras),
                "Comprobantes observados": len(alertas),
                "Criterio": "Coincidencia exacta de RUC; comprobantes desde el 20/03/2022 (D. Leg. 1532).",
            }
            st.download_button(
                "⬇️ Descargar reporte Excel",
                data=core.exportar_excel(alertas, resumen, compras, info, por_ejercicio, vinculados),
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
        res = core.buscar_rucs(texto, padron, relacion)
        if res.empty:
            st.info("Ingrese al menos un RUC.")
        else:
            for _, f in res.iterrows():
                msg = f"**{f['RUC consultado']}** — {f['Resultado']}"
                if "FIGURA" in f["Resultado"]:
                    st.error(msg)
                elif "VINCULADO" in f["Resultado"] or "inválido" in f["Resultado"]:
                    st.warning(msg)
                else:
                    st.success(msg)
            hallados = res[res["Resultado"].str.contains("FIGURA|VINCULADO")]
            if not hallados.empty:
                st.dataframe(hallados.dropna(axis=1, how="all"), hide_index=True, use_container_width=True)

# --------------------------------------------------------------------------- padrón completo
with tab_padron:
    opciones = {"Padrón SSCO": padron}
    if relacion is not None:
        opciones["Vinculados a SSCO (art. 12)"] = relacion
    elegida = st.radio("Lista", list(opciones), horizontal=True)
    base = opciones[elegida]
    filtro = st.text_input("Filtrar por RUC, razón social u otro dato")
    vista = base
    if filtro:
        mask = base.astype(str).apply(lambda s: s.str.contains(filtro, case=False, na=False)).any(axis=1)
        vista = base[mask]
    st.caption(f"{len(vista):,} de {len(base):,} registros")
    st.dataframe(vista, hide_index=True, use_container_width=True)

# --------------------------------------------------------------------------- ayuda
with tab_ayuda:
    st.markdown("""
**¿Qué hace esta app?** Cruza los RUC de los proveedores de su registro de compras con la
*Relación de Sujetos sin Capacidad Operativa con resoluciones de atribución firmes* que publica SUNAT.

**Páginas oficiales de SUNAT (para ver o descargar las listas):**
- Padrón SSCO: [{pag_ssco}]({pag_ssco})
- Relación art. 12 (vinculados a SSCO): [{pag_rel}]({pag_rel})

**Descarga directa de los Excel:**
- Padrón SSCO: [{url}]({url})
- Relación art. 12: [{url_rel}]({url_rel})

**Relación art. 12 (vinculados a SSCO):** SUNAT también publica las EIRL y sociedades cuyos
titulares o socios son SSCO. Desde el día siguiente a su publicación y hasta la fecha límite indicada,
solo pueden emitir boletas de venta y notas de crédito/débito electrónicas. La app las reporta por
separado con un nivel de riesgo: **Alto** (factura emitida durante la restricción), **Medio**
(otro comprobante durante la restricción) o **Informativo** (comprobante fuera de ese periodo).
A diferencia del padrón SSCO, la norma no anula automáticamente el crédito fiscal de estos proveedores;
la alerta es para revisión.

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
""".replace("{url}", core.SSCO_URL).replace("{pag_ssco}", core.SSCO_PAGINA)
    .replace("{pag_rel}", core.RELACION_PAGINA).replace("{url_rel}", core.RELACION_URL))
