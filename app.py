import io
import json
import os
import urllib.parse
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
import streamlit as st

# Google Drive es opcional. Si no se configuran las credenciales, el programa
# conserva la posibilidad de descargar Excel y guardar un histórico local.
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
    GOOGLE_DRIVE_DISPONIBLE = True
except ImportError:
    GOOGLE_DRIVE_DISPONIBLE = False

# ============================================================
# 1. CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Diagnóstico Wialon | Korban Global Solutions",
    page_icon="📡",
    layout="wide",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
HISTORIAL_LOCAL_PATH = os.path.join(BASE_DIR, "registro_observaciones.xlsx")
HISTORIAL_DRIVE_NAME = "registro_observaciones_korban.xlsx"
WIALON_URL = "https://hst-api.wialon.com/wialon/ajax.html"

COLUMNAS_DIAGNOSTICO = [
    "Unidad", "Cliente / Grupo", "Último Reporte", "Días Sin Reporte",
    "Estado", "Conexión", "Novedad Batería / Voltaje", "Diagnóstico",
    "WhatsApp", "Observación",
]
COLUMNAS_HISTORIAL = [
    "Fecha Registro", "Unidad", "Cliente / Grupo", "Estado", "Conexión",
    "Observación", "Novedad Batería / Voltaje", "WhatsApp",
    "Último Reporte", "Días Sin Reporte", "Diagnóstico",
]

# ============================================================
# 2. ESTILOS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] {font-family: Inter, sans-serif;}
.stApp {background-color:#f8fafc;}
[data-testid="stSidebar"] {background-color:#0f172a;}
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {color:#f1f5f9 !important;}
[data-testid="stMetric"] {
    background:#fff; border:1px solid #e2e8f0; border-radius:10px;
    padding:14px 18px; box-shadow:0 2px 4px rgba(0,0,0,.02);
}
.stButton > button, .stLinkButton > a, .stDownloadButton > button {
    border-radius:8px; font-weight:600;
}
h1 {color:#0f172a;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# 3. UTILIDADES DE WIALON
# ============================================================
def wialon_call(servicio, parametros, sid=None):
    """Ejecuta una llamada a Wialon Remote API y valida errores HTTP/API."""
    datos = {
        "svc": servicio,
        "params": json.dumps(parametros, ensure_ascii=False),
    }
    if sid:
        datos["sid"] = sid

    respuesta = requests.get(WIALON_URL, params=datos, timeout=45)
    respuesta.raise_for_status()
    resultado = respuesta.json()

    if isinstance(resultado, dict) and "error" in resultado:
        raise RuntimeError(
            f"Wialon devolvió el error {resultado['error']} en {servicio}."
        )
    return resultado


def obtener_mapa_grupos(sid):
    """Relaciona el ID de cada unidad con los grupos a los que pertenece."""
    resultado = wialon_call("core/search_items", {
        "spec": {
            "itemsType": "avl_unit_group",
            "propName": "sys_name",
            "propValueMask": "*",
            "sortType": "sys_name",
        },
        "force": 1,
        "flags": 1,
        "from": 0,
        "to": 0,
    }, sid)

    mapa = {}
    for grupo in resultado.get("items", []) or []:
        nombre = grupo.get("nm", "Sin nombre")
        for unidad_id in grupo.get("u", []) or []:
            mapa.setdefault(unidad_id, []).append(nombre)
    return mapa


def obtener_unidades(sid):
    """Carga unidades con propiedades, sensores y último mensaje."""
    flags = 1 + 1024 + 4096 + 1048576 + 2097152
    resultado = wialon_call("core/search_items", {
        "spec": {
            "itemsType": "avl_unit",
            "propName": "sys_name",
            "propValueMask": "*",
            "sortType": "sys_name",
        },
        "force": 1,
        "flags": flags,
        "from": 0,
        "to": 0,
    }, sid)
    return resultado.get("items", []) or []


def analizar_bateria(unidad):
    """
    Lee sensores cuyo nombre menciona batería, voltaje, alimentación o power.
    Solo marca posible batería en cero cuando un sensor devuelve valor 0.
    """
    mensaje = unidad.get("lmsg") or {}
    parametros = mensaje.get("p") or {}
    sensores = unidad.get("sens") or {}

    if isinstance(sensores, dict):
        lista_sensores = list(sensores.values())
    elif isinstance(sensores, list):
        lista_sensores = sensores
    else:
        lista_sensores = []

    encontrados = []
    sin_bateria = False

    for sensor in lista_sensores:
        if not isinstance(sensor, dict):
            continue

        nombre = str(sensor.get("n", "")).strip()
        nombre_lower = nombre.lower()
        palabras_clave = ("bater", "volt", "power", "aliment")

        if not any(palabra in nombre_lower for palabra in palabras_clave):
            continue

        valor = sensor.get("v")
        if valor is None and sensor.get("p") is not None:
            valor = parametros.get(sensor.get("p"))

        if valor is None:
            encontrados.append(f"{nombre}: sin dato")
            continue

        try:
            numero = float(valor)
            encontrados.append(f"{nombre}: {numero:g}")
            if numero == 0:
                sin_bateria = True
        except (ValueError, TypeError):
            encontrados.append(f"{nombre}: {valor}")

    texto = (
        " | ".join(encontrados)
        if encontrados
        else "Sin sensores de batería/voltaje configurados"
    )
    return sin_bateria, texto


def diagnosticar_unidades(unidades, mapa_grupos, horas_limite):
    """Genera una fila por unidad con su estado, último mensaje y diagnóstico."""
    ahora = datetime.now(timezone.utc)
    limite = ahora - timedelta(hours=int(horas_limite))
    filas = []

    for unidad in unidades:
        unidad_id = unidad.get("id")
        nombre = unidad.get("nm", f"Unidad {unidad_id}")
        grupos = ", ".join(mapa_grupos.get(unidad_id, [])) or "Sin Grupo / General"
        mensaje = unidad.get("lmsg") or {}
        timestamp = mensaje.get("t")
        sin_bateria, texto_bateria = analizar_bateria(unidad)

        if timestamp:
            fecha_utc = datetime.fromtimestamp(float(timestamp), tz=timezone.utc)
            dias_sin_reporte = round(
                max(0, (ahora - fecha_utc).total_seconds()) / 86400, 1
            )
            ultimo_reporte = fecha_utc.astimezone().strftime("%Y-%m-%d %H:%M:%S")
            supera_umbral = fecha_utc <= limite
        else:
            fecha_utc = None
            dias_sin_reporte = "N/A"
            ultimo_reporte = "SIN DATOS"
            supera_umbral = True

        # IMPORTANTE: netconn representa conectividad en Wialon, no es una
        # lectura inequívoca del estado real del motor.
        conectado = bool(unidad.get("netconn"))

        if sin_bateria:
            estado = "SIN BATERÍA (sensor 0)"
        elif not timestamp:
            estado = "SIN DATOS"
        elif conectado:
            estado = "ENCENDIDO / CONECTADO"
        else:
            estado = "APAGADO / SIN CONEXIÓN"

        if not timestamp:
            conexion = "🔴 SIN HISTORIAL"
            diagnostico = (
                "No hay último mensaje disponible. Revisar alta, configuración "
                "del equipo y recepción de datos."
            )
        elif supera_umbral and sin_bateria:
            conexion = "🪫 POSIBLE SIN BATERÍA"
            diagnostico = (
                f"Sensor de batería/voltaje en 0. Último reporte hace "
                f"{dias_sin_reporte} días."
            )
        elif supera_umbral and conectado:
            conexion = "📶 SIN REPORTE RECIENTE"
            diagnostico = (
                f"Wialon indica conexión, pero el último mensaje supera "
                f"{int(horas_limite)} horas."
            )
        elif supera_umbral:
            conexion = "🔴 INACTIVA"
            diagnostico = (
                f"Sin comunicación por más de {int(horas_limite)} horas. "
                "Verificar alimentación, cobertura, antena y equipo GPS."
            )
        else:
            conexion = "🟢 ACTIVA"
            diagnostico = "Último mensaje dentro del umbral configurado."

        filas.append({
            "Unidad": nombre,
            "Cliente / Grupo": grupos,
            "Último Reporte": ultimo_reporte,
            "Días Sin Reporte": dias_sin_reporte,
            "Estado": estado,
            "Conexión": conexion,
            "Novedad Batería / Voltaje": texto_bateria,
            "Diagnóstico": diagnostico,
            "WhatsApp": "",
            "Observación": "",
            "_timestamp": float(timestamp or 0),
            "_inactiva": bool(supera_umbral),
            "_sin_bateria": bool(sin_bateria),
            "_conectado": conectado,
        })

    return pd.DataFrame(filas)


# ============================================================
# 4. EXCEL E HISTÓRICO
# ============================================================
def generar_excel(df, nombre_hoja="Diagnostico"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        df.to_excel(escritor, index=False, sheet_name=nombre_hoja)
        hoja = escritor.sheets[nombre_hoja]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(
                (len(str(celda.value or "")) for celda in columna),
                default=12,
            )
            hoja.column_dimensions[columna[0].column_letter].width = min(
                max(largo + 2, 12), 55
            )
    buffer.seek(0)
    return buffer.getvalue()


def obtener_servicio_drive():
    """Crea el cliente Drive desde Streamlit Secrets si están configurados."""
    if not GOOGLE_DRIVE_DISPONIBLE:
        raise RuntimeError(
            "Faltan dependencias de Google Drive. Instala google-api-python-client "
            "y google-auth mediante requirements.txt."
        )
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "No se configuró [gcp_service_account] en Streamlit Secrets."
        )

    credenciales_dict = dict(st.secrets["gcp_service_account"])
    if isinstance(credenciales_dict.get("private_key"), str):
        credenciales_dict["private_key"] = credenciales_dict["private_key"].replace(
            "\\n", "\n"
        )

    credenciales = service_account.Credentials.from_service_account_info(
        credenciales_dict,
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=credenciales)


def guardar_historico_local(df_nuevas):
    """Acumula observaciones en el Excel local."""
    if os.path.exists(HISTORIAL_LOCAL_PATH):
        try:
            anterior = pd.read_excel(HISTORIAL_LOCAL_PATH)
        except Exception:
            anterior = pd.DataFrame(columns=COLUMNAS_HISTORIAL)
    else:
        anterior = pd.DataFrame(columns=COLUMNAS_HISTORIAL)

    combinado = pd.concat([anterior, df_nuevas], ignore_index=True)
    combinado.to_excel(HISTORIAL_LOCAL_PATH, index=False)
    return len(combinado)


def sincronizar_historico_drive(df_nuevas):
    """
    Agrega observaciones al Excel de Google Drive.
    Requiere DRIVE_FOLDER_ID y gcp_service_account en Streamlit Secrets.
    """
    if "DRIVE_FOLDER_ID" not in st.secrets:
        raise RuntimeError("Falta DRIVE_FOLDER_ID en Streamlit Secrets.")

    folder_id = str(st.secrets["DRIVE_FOLDER_ID"]).strip()
    if not folder_id or "COLOCA" in folder_id.upper():
        raise RuntimeError("Configura un DRIVE_FOLDER_ID válido.")

    servicio = obtener_servicio_drive()
    query = (
        f"'{folder_id}' in parents and name = '{HISTORIAL_DRIVE_NAME}' "
        "and trashed = false"
    )
    resultado = servicio.files().list(
        q=query,
        fields="files(id,name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    archivos = resultado.get("files", [])

    if archivos:
        file_id = archivos[0]["id"]
        solicitud = servicio.files().get_media(fileId=file_id)
        buffer_descarga = io.BytesIO()
        descargador = MediaIoBaseDownload(buffer_descarga, solicitud)
        terminado = False
        while not terminado:
            _, terminado = descargador.next_chunk()
        buffer_descarga.seek(0)
        anterior = pd.read_excel(buffer_descarga)
        completo = pd.concat([anterior, df_nuevas], ignore_index=True)
    else:
        file_id = None
        completo = df_nuevas.copy()

    buffer_subida = io.BytesIO()
    with pd.ExcelWriter(buffer_subida, engine="openpyxl") as escritor:
        completo.to_excel(escritor, index=False, sheet_name="Observaciones")
    buffer_subida.seek(0)

    media = MediaIoBaseUpload(
        buffer_subida,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )

    if file_id:
        servicio.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True,
        ).execute()
    else:
        servicio.files().create(
            body={"name": HISTORIAL_DRIVE_NAME, "parents": [folder_id]},
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()

    # Guarda también una copia local cuando el entorno permite escribirla.
    try:
        completo.to_excel(HISTORIAL_LOCAL_PATH, index=False)
    except Exception:
        pass

    return len(completo)


# ============================================================
# 5. ENCABEZADO
# ============================================================
if os.path.exists(LOGO_PATH):
    col_logo, col_titulo = st.columns([1, 4])
    with col_logo:
        st.image(LOGO_PATH, width=190)
    with col_titulo:
        st.title("📡 Panel de Diagnóstico Wialon")
        st.caption(
            "Korban Global Solutions — Barrido de flota, diagnóstico "
            "y gestión de novedades"
        )
else:
    st.title("📡 Panel de Diagnóstico Wialon")
    st.caption(
        "Korban Global Solutions — Barrido de flota, diagnóstico "
        "y gestión de novedades"
    )
st.divider()

# ============================================================
# 6. BARRA LATERAL
# ============================================================
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    st.header("⚙️ Parámetros de conexión")
    token = st.text_input(
        "Token de Wialon",
        type="password",
        help="Pega aquí un token vigente. No lo escribas directamente en el código.",
    )
    horas_limite = st.number_input(
        "Horas sin reporte",
        min_value=1,
        max_value=720,
        value=48,
        step=12,
    )
    st.caption("El umbral se calcula usando la fecha del último mensaje recibido.")
    consultar = st.button("🔍 Ejecutar barrido", use_container_width=True)

if consultar:
    if not token.strip():
        st.error("Introduce tu token de Wialon en la barra lateral.")
    else:
        sid = None
        try:
            with st.spinner("Conectando con Wialon y analizando las unidades..."):
                login = wialon_call("token/login", {"token": token.strip()})
                sid = login.get("eid")
                if not sid:
                    raise RuntimeError("Wialon no devolvió el identificador de sesión.")

                mapa_grupos = obtener_mapa_grupos(sid)
                unidades = obtener_unidades(sid)
                if not unidades:
                    raise RuntimeError(
                        "La API no devolvió unidades. Revisa permisos del token."
                    )

                df_resultado = diagnosticar_unidades(
                    unidades, mapa_grupos, int(horas_limite)
                )
                st.session_state["df_completo"] = df_resultado
                st.session_state["total_unidades"] = len(unidades)
                st.session_state["fecha_barrido"] = datetime.now().astimezone().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                st.session_state["horas_limite_usadas"] = int(horas_limite)
        except Exception as error:
            st.error(f"No se pudo completar el barrido: {error}")
        finally:
            if sid:
                try:
                    wialon_call("core/logout", {}, sid)
                except Exception:
                    pass

# ============================================================
# 7. DASHBOARD Y DIAGNÓSTICO
# ============================================================
if "df_completo" not in st.session_state:
    st.info(
        "Introduce el token de Wialon en la barra lateral y pulsa "
        "**Ejecutar barrido** para cargar el dashboard."
    )
else:
    df_all = st.session_state["df_completo"].copy()
    umbral_actual = st.session_state.get("horas_limite_usadas", int(horas_limite))
    criticos = df_all[df_all["_inactiva"]].copy()

    total = len(df_all)
    sin_bateria = int(df_all["_sin_bateria"].sum())
    conectadas = int(df_all["_conectado"].sum())
    sin_conexion = total - conectadas
    sin_reporte = int(df_all["_inactiva"].sum())

    st.caption(
        f"Último barrido: {st.session_state['fecha_barrido']} · "
        f"Umbral utilizado: {umbral_actual} horas"
    )
    st.subheader("📊 Dashboard general de flota")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Unidades evaluadas", total)
    c2.metric("Sin batería*", sin_bateria)
    c3.metric("Con conexión Wialon", conectadas)
    c4.metric("Sin conexión Wialon", sin_conexion)
    c5.metric("Sin reporte en el umbral", sin_reporte)

    st.caption(
        "*“Sin batería” se cuenta solo cuando un sensor configurado de batería/"
        "voltaje devuelve exactamente 0. Con/sin conexión usa el indicador "
        "netconn de Wialon; no confirma por sí solo si el motor está encendido. "
        "Para el estado real del motor debe validarse un sensor IGN/RPM."
    )

    with st.expander("📈 Distribución de unidades", expanded=False):
        grafico = pd.DataFrame({
            "Categoría": [
                "Sin batería (sensor en 0)",
                "Con conexión Wialon",
                "Sin conexión Wialon",
                "Sin reporte según umbral",
            ],
            "Cantidad": [sin_bateria, conectadas, sin_conexion, sin_reporte],
        }).set_index("Categoría")
        st.bar_chart(grafico)

    st.divider()
    st.subheader("🧰 Diagnóstico de unidades sin reporte")
    f1, f2, f3 = st.columns([2, 1, 1])
    with f1:
        buscar = st.text_input(
            "Buscar unidad o cliente",
            placeholder="Escribe nombre de unidad o cliente...",
        )
    with f2:
        filtro_estado = st.selectbox(
            "Estado",
            [
                "Todos",
                "SIN BATERÍA (sensor 0)",
                "SIN DATOS",
                "ENCENDIDO / CONECTADO",
                "APAGADO / SIN CONEXIÓN",
            ],
        )
    with f3:
        ordenar = st.selectbox(
            "Ordenar por",
            ["Más días sin reporte", "Unidad", "Cliente / Grupo"],
        )

    vista = criticos.copy()
    if buscar.strip():
        buscar_mask = (
            vista["Unidad"].astype(str).str.contains(buscar, case=False, na=False)
            | vista["Cliente / Grupo"].astype(str).str.contains(
                buscar, case=False, na=False
            )
        )
        vista = vista[buscar_mask]
    if filtro_estado != "Todos":
        vista = vista[vista["Estado"] == filtro_estado]

    if ordenar == "Más días sin reporte":
        vista["_dias_num"] = pd.to_numeric(vista["Días Sin Reporte"], errors="coerce")
        vista = vista.sort_values("_dias_num", ascending=False, na_position="first")
    else:
        vista = vista.sort_values(ordenar, na_position="last")

    st.info(
        f"Se muestran {len(vista)} unidad(es) que superan el umbral "
        f"de {umbral_actual} horas."
    )
    vista_editor = vista[COLUMNAS_DIAGNOSTICO].copy()
    editor = st.data_editor(
        vista_editor,
        key="editor_diagnostico",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=[
            "Unidad", "Cliente / Grupo", "Último Reporte", "Días Sin Reporte",
            "Estado", "Conexión", "Novedad Batería / Voltaje", "Diagnóstico",
        ],
        column_config={
            "WhatsApp": st.column_config.TextColumn(
                "WhatsApp (código de país + número, solo dígitos)"
            ),
            "Observación": st.column_config.TextColumn(
                "Observación / seguimiento"
            ),
        },
    )

    # ========================================================
    # 8. GUARDAR, DESCARGAR Y SINCRONIZAR
    # ========================================================
    st.divider()
    st.subheader("💾 Acciones del registro")
    b1, b2, b3 = st.columns(3)

    with b1:
        if st.button(
            "💾 Guardar observaciones",
            use_container_width=True,
            disabled=editor.empty,
        ):
            observadas = editor[
                editor["Observación"].fillna("").astype(str).str.strip().ne("")
            ].copy()
            if observadas.empty:
                st.warning("Escribe al menos una observación antes de guardar.")
            else:
                observadas["Fecha Registro"] = datetime.now().astimezone().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                nuevas = observadas[COLUMNAS_HISTORIAL]
                try:
                    cantidad = guardar_historico_local(nuevas)
                    st.success(
                        f"Se guardaron {len(nuevas)} observaciones. "
                        f"Total acumulado local: {cantidad}."
                    )
                except Exception as error:
                    st.error(f"No se pudo guardar el histórico local: {error}")

    with b2:
        st.download_button(
            "📥 Descargar diagnóstico Excel",
            data=generar_excel(vista_editor),
            file_name=f"Diagnostico_Wialon_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with b3:
        if st.button("☁️ Guardar observaciones en Google Drive", use_container_width=True):
            observadas = editor[
                editor["Observación"].fillna("").astype(str).str.strip().ne("")
            ].copy()
            if observadas.empty:
                st.warning("Escribe al menos una observación antes de sincronizar.")
            else:
                observadas["Fecha Registro"] = datetime.now().astimezone().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                nuevas = observadas[COLUMNAS_HISTORIAL]
                try:
                    with st.spinner("Sincronizando con Google Drive..."):
                        cantidad = sincronizar_historico_drive(nuevas)
                    st.success(
                        f"Google Drive actualizado. Registros acumulados: {cantidad}."
                    )
                except Exception as error:
                    st.error(f"No se pudo sincronizar con Google Drive: {error}")

    if os.path.exists(HISTORIAL_LOCAL_PATH):
        with open(HISTORIAL_LOCAL_PATH, "rb") as archivo:
            st.download_button(
                "📚 Descargar histórico local acumulado",
                data=archivo.read(),
                file_name="registro_observaciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    st.caption(
        "Nota: el almacenamiento local de Streamlit Community Cloud puede ser "
        "efímero. Para conservar el histórico entre reinicios, configura Google Drive."
    )

    # ========================================================
    # 9. WHATSAPP
    # ========================================================
    st.divider()
    st.subheader("📲 Preparar aviso por WhatsApp")

    if not editor.empty:
        opciones = [
            f"{fila['Unidad']} — {fila['Cliente / Grupo']}"
            for _, fila in editor.iterrows()
        ]
        seleccion = st.selectbox("Selecciona la unidad", opciones)
        posicion = opciones.index(seleccion)
        fila = editor.iloc[posicion]

        observacion = str(fila["Observación"]).strip()
        mensaje_default = (
            f"Estimado cliente ({fila['Cliente / Grupo']}), desde Korban Global "
            f"Solutions le informamos que la unidad {fila['Unidad']} presenta "
            f"la siguiente novedad: {fila['Conexión']}. "
            f"Diagnóstico: {fila['Diagnóstico']}\n\n"
            f"Último reporte: {fila['Último Reporte']}.\n"
            f"Tiempo sin reporte: {fila['Días Sin Reporte']} días.\n"
            f"Batería/voltaje: {fila['Novedad Batería / Voltaje']}."
        )
        if observacion and observacion.lower() != "nan":
            mensaje_default += f"\nObservación de soporte: {observacion}"

        mensaje = st.text_area(
            "Editar mensaje antes de abrir WhatsApp",
            value=mensaje_default,
            height=170,
            key="mensaje_whatsapp",
        )
        numero = "".join(
            caracter for caracter in str(fila["WhatsApp"]) if caracter.isdigit()
        )

        if numero and numero.lower() != "nan":
            enlace = f"https://wa.me/{numero}?text={urllib.parse.quote(mensaje)}"
            st.link_button("💬 Abrir WhatsApp con mensaje preparado", enlace)
        else:
            st.warning(
                "Completa el número WhatsApp en la tabla. Incluye el código "
                "de país, sin signo + ni espacios."
            )
    else:
        st.info("No hay unidades críticas que coincidan con los filtros actuales.")

    # ========================================================
    # 10. TODAS LAS UNIDADES
    # ========================================================
    with st.expander("🔎 Ver todas las unidades evaluadas"):
        st.dataframe(
            df_all[COLUMNAS_DIAGNOSTICO],
            use_container_width=True,
            hide_index=True,
        )
