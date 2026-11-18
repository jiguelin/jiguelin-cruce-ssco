# 🔎 Cruce SSCO – Registro de Compras vs. Sujetos sin Capacidad Operativa

App web (Streamlit) que cruza el **Registro de Compras** con el padrón SUNAT de
**Sujetos sin Capacidad Operativa (SSCO)** – D. Leg. 1532 – y alerta los comprobantes
que no otorgan crédito fiscal ni sustentan gasto.

## Funciones
- **Cruce masivo**: suba uno o varios archivos de compras de cualquier ejercicio:
  - Excel del software contable (*Formato 8.1: Registro de Compras*).
  - Excel `resultado_AAAAMMDD_HHMMSS.xlsx` (hoja *Comprobantes*, sin encabezados).
  - Propuesta RCE del **SIRE** (`.zip` tal cual se descarga, o el `.csv` / `.txt` interno).
  - Cualquier Excel/CSV con una columna de RUC.
- **Años anteriores**: detecta el periodo de cada comprobante, resume por ejercicio, filtra por año
  e indica si el comprobante se emitió antes (art. 9) o después (art. 8) de la publicación SSCO.
- **Vinculados a SSCO (art. 12 DL 1532)**: cruza también la relación de EIRL y sociedades cuyos
  titulares o socios son SSCO ([página SUNAT](https://www.sunat.gob.pe/padronesnotificaciones/relacion-ssco.html)).
  Se reporta aparte con nivel de riesgo: **Alto** (factura emitida durante la restricción, cuando solo
  pueden emitir boletas y notas electrónicas), **Medio** u **Informativo**.
- **Búsqueda manual** de uno o varios RUC (valida dígito verificador).
- **Padrón SSCO**: descarga automática desde SUNAT (caché 6 h); si SUNAT no responde usa la copia
  en `data/`; o se pueden subir uno o varios padrones manualmente (útil para versiones históricas).
- **Reporte Excel** con alertas, resumen por proveedor, resumen por ejercicio, compras revisadas y
  datos de auditoría (fecha, fuente del padrón, archivos).

Criterios: coincidencia **exacta** de RUC; comprobantes anteriores al 20/03/2022 se marcan
como fuera de alcance.

## Publicar en Streamlit Community Cloud
1. Entre a <https://share.streamlit.io> e inicie sesión con GitHub.
2. **Create app → Deploy a public app from GitHub**: elija este repositorio, rama `main`,
   archivo principal `app.py`.
3. (Recomendado) En **Advanced settings → Secrets** agregue una contraseña de acceso:
   ```toml
   APP_PASSWORD = "su-clave"
   ```
4. **Deploy**. Cada `git push` a `main` actualiza la app automáticamente.

> Si el repositorio es privado, Streamlit pedirá autorizar el acceso a repos privados.

## Actualización del padrón
- La app intenta descargar el padrón de SUNAT cada vez que se abre (caché 6 h).
- El workflow `.github/workflows/actualizar-padron.yml` intenta actualizar `data/sujesincapacidadOperativa.xlsx`
  y `data/relacion-ssco.xlsx`
  del 1 al 5 de cada mes (también se puede ejecutar a mano en *Actions*).
- SUNAT puede bloquear servidores fuera del Perú; en ese caso descargue el Excel desde
  *SUNAT › Padrones y Notificaciones › SSCO* y súbalo en la app, o reemplace el archivo en `data/`.

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```
Pruebas: `pip install pytest && pytest`

> Herramienta de apoyo; no sustituye la revisión de un especialista tributario.
