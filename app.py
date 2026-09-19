import json
import os
import urllib.parse
from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import streamlit as st

# 1. Configuración de página
st.set_page_config(
    page_title="Diagnóstico Wialon | Korban Global Solutions",
    page_icon="📡",
    layout="wide",
)

# 2. Obtener ruta absoluta del directorio del script (resuelve el problema del Logo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
EXCEL_HISTORIAL_PATH = os.path.join(BASE_DIR, "registro_observaciones.xlsx")

# 3. Estilos CSS Personalizados
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"], .stMarkdown, .stText, p, div, span {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }

    .stApp {
        background-color: #f8fafc;
    }

    h1 {
        color: #0f172a !important;
        font-weight: 700 !important;
        letter-spacing: -0.5px;
    }
    
    h2, h3 {
        color: #1e293b !important;
        font-weight: 600 !important;
    }

    /* Barra lateral */
    [data-testid="stSidebar"] {
        background-color: #0f172a !important;
    }
    
    /* Aplicar color blanco solo a textos descriptivos de la barra lateral, evitando romper iconos del sistema */
    [data-testid="stSidebar"] .stMarkdown, 
    [data-testid="stSidebar"] label, 
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #f1f5f9 !important;
    }

    /* Ocultar texto crudo de iconos del sistema en barra lateral */
    [data-testid="stSidebarCollapseButton"] span {
        font-size: 0px !important;
    }

    /* Tarjetas de Métricas */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 15px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    
    [data-testid="stMetricLabel"] {
        color: #64748b !important;
        font-weight: 500;
        font-size: 0.85rem;
    }
    
    [data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-weight: 700;
    }

    /* Botones principales */
    .stButton > button {
        background-color: #0284c7 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 500 !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.2s ease-in-out;
    }

    .stButton > button:hover {
        background-color: #0369a1 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    /* Botón WhatsApp */
    .stLinkButton > a {
        background-color: #16a34a !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
    }
    
    .stLinkButton > a:hover {
        background-color: #15803d !important;
    }

    /* Cajas de texto */
    .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 8px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
    }

    .stAlert {
        border-radius: 8px !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Encabezado con Logo y Título
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
  if os.path.exists(LOGO_PATH):
    st.image(LOGO_PATH, width=190)
  else:
    st.caption("📷 *Guarda la imagen como 'logo.png' junto a app.py*")

with col_titulo:
  st.title("📡 Panel de Diagnóstico Wialon")
  st.caption(
      "Korban Global Solutions — Monitoreo ejecutivo de flotilla y gestión de"
      " novedades"
  )

st.divider()

# Barra lateral
if os.path.exists(LOGO_PATH):
  st.sidebar.image(LOGO_PATH, use_container_width=True)

st.sidebar.header("⚙️ Parámetros de Conexión")
token_input = st.sidebar.text_input(
    "Token de Wialon",
    value="ec232dab9897c9e6885aedd524a874560CA7ADC8CA1BE88705D0E08E268678A64030B042",
    type="password",
)
horas_limite = st.sidebar.number_input("Horas sin reporte", value=48, step=12)

if st.sidebar.button("🔍 Consultar Unidades"):
  with st.spinner("Conectando con servidores Wialon y procesando estado..."):
    try:
      url = "https://hst-api.wialon.com/wialon/ajax.html"

      # Login
      login_res = requests.get(
          url,
          params={
              "svc": "token/login",
              "params": json.dumps({"token": token_input}),
          },
          timeout=30,
      ).json()

      sid = login_res.get("eid") or (
          login_res.get("session")
          if isinstance(login_res.get("session"), str)
          else None
      )
      if not sid and isinstance(login_res.get("session"), dict):
        sid = login_res["session"].get("eid")

      if not sid:
        st.error(f"Error de sesión: {login_res}")
        st.stop()

      # Consultar Grupos
      params_grupos = {
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
      }

      res_grupos = requests.get(
          url,
          params={
              "svc": "core/search_items",
              "params": json.dumps(params_grupos),
              "sid": sid,
          },
          timeout=30,
      ).json()

      mapa_grupos = {}
      for grupo in res_grupos.get("items", []):
        nombre_grupo = grupo.get("nm", "Sin Cliente")
        u_ids = grupo.get("u", [])
        if isinstance(u_ids, list):
          for u_id in u_ids:
            if u_id not in mapa_grupos:
              mapa_grupos[u_id] = []
            mapa_grupos[u_id].append(nombre_grupo)

      # Consultar Unidades
      flags_unidades = 1 + 1024 + 4096 + 1048576 + 2097152
      params_unidades = {
          "spec": {
              "itemsType": "avl_unit",
              "propName": "sys_name",
              "propValueMask": "*",
              "sortType": "sys_name",
          },
          "force": 1,
          "flags": flags_unidades,
          "from": 0,
          "to": 0,
      }

      res = requests.get(
          url,
          params={
              "svc": "core/search_items",
              "params": json.dumps(params_unidades),
              "sid": sid,
          },
          timeout=30,
      ).json()
      unidades = res.get("items", [])

      ahora = datetime.now(timezone.utc)
      limite = ahora - timedelta(hours=horas_limite)
      data = []

      for u in unidades:
        u_id = u.get("id")
        cliente_grupo = ", ".join(
            mapa_grupos.get(u_id, ["Sin Grupo / General"])
        )

        lmsg = u.get("lmsg", {})
        t = lmsg.get("t")

        sensores = u.get("sens", {})
        info_bateria = []
        bateria_cero = False

        if isinstance(sensores, dict):
          sensores_list = sensores.values()
        elif isinstance(sensores, list):
          sensores_list = sensores
        else:
          sensores_list = []

        for s in sensores_list:
          nombre_sens = str(s.get("n", "")).strip()
          if "bater" in nombre_sens.lower():
            val = s.get("v")
            if val is None and "p" in s:
              param_key = s.get("p")
              val = lmsg.get("p", {}).get(param_key, "N/A")

            try:
              val_num = float(val)
              if val_num == 0:
                bateria_cero = True
                info_bateria.append(f"⚠️ {nombre_sens}: 0V")
              else:
                info_bateria.append(f"{nombre_sens}: {val_num:.1f}V")
            except (ValueError, TypeError):
              info_bateria.append(
                  f"{nombre_sens}: {val if val is not None else 'Sin dato'}"
              )

        texto_bateria = (
            " | ".join(info_bateria)
            if info_bateria
            else "Sin sensores de batería"
        )

        if not t:
          estado_conexion = "🔴 FALLA FÍSICA"
          diagnostico = "Sin registros o equipo no configurado"
          data.append({
              "Unidad": u.get("nm"),
              "Cliente / Grupo": cliente_grupo,
              "Último Reporte": "SIN DATOS",
              "Días Sin Reporte": "N/A",
              "Conexión": estado_conexion,
              "Novedad Batería / Voltaje": texto_bateria,
              "Diagnóstico": diagnostico,
              "WhatsApp": "",
              "Observación": "",
              "_timestamp": 0,
          })
        else:
          fecha_u = datetime.fromtimestamp(t, tz=timezone.utc)
          if fecha_u <= limite:
            dias_inactivo = round(
                (ahora - fecha_u).total_seconds() / 86400, 1
            )

            if bateria_cero:
              estado_conexion = "🪫 SIN BATERÍA"
              diagnostico = (
                  f"Sensor en 0V. Inactivo por más de {horas_limite}h"
              )
            elif u.get("netconn"):
              estado_conexion = "📶 SIN COBERTURA"
              diagnostico = "Equipo encendido pero sin reporte GPS/GPRS"
            else:
              estado_conexion = "🔴 INACTIVO"
              diagnostico = f"Sin comunicación por más de {horas_limite}h"

            data.append({
                "Unidad": u.get("nm"),
                "Cliente / Grupo": cliente_grupo,
                "Último Reporte": fecha_u.strftime("%Y-%m-%d %H:%M:%S"),
                "Días Sin Reporte": dias_inactivo,
                "Conexión": estado_conexion,
                "Novedad Batería / Voltaje": texto_bateria,
                "Diagnóstico": diagnostico,
                "WhatsApp": "",
                "Observación": "",
                "_timestamp": t,
            })

      df = pd.DataFrame(data)
      if not df.empty:
        df = df.sort_values(by="_timestamp", ascending=True)
        df = df.drop(columns=["_timestamp"])

      st.session_state["data_unidades"] = df
      st.session_state["total_evaluadas"] = len(unidades)

    except Exception as e:
      st.error(f"Ocurrió un error inesperado: {e}")

# Resultados
if "data_unidades" in st.session_state:
  df = st.session_state["data_unidades"]

  col1, col2, col3 = st.columns(3)
  col1.metric("Total Unidades Evaluadas", st.session_state["total_evaluadas"])
  col2.metric("Unidades Críticas", len(df))
  col3.metric("Filtro Aplicado", f">{horas_limite} Horas")

  st.write("")
  st.subheader("📊 Listado de Unidades Críticas")

  edited_df = st.data_editor(
      df,
      use_container_width=True,
      num_rows="fixed",
      column_config={
          "Cliente / Grupo": st.column_config.TextColumn("Cliente / Grupo"),
          "Días Sin Reporte": st.column_config.NumberColumn(
              "Días Sin Reporte", format="%.1f días"
          ),
          "Conexión": st.column_config.TextColumn("Conexión"),
          "WhatsApp": st.column_config.TextColumn(
              "WhatsApp (Ej: 584121234567)", help="Ingrese el número con código de país"
          ),
          "Observación": st.column_config.TextColumn(
              "Observación / Notas", help="Escriba aquí los comentarios del caso"
          ),
      },
  )

  st.divider()

  # Historial
  st.subheader("💾 Registro Histórico de Observaciones")
  col_guardar, col_descargar = st.columns([1, 1])

  with col_guardar:
    if st.button("💾 Guardar Observaciones en Excel"):
      filas_con_obs = edited_df[
          edited_df["Observación"].astype(str).str.strip().ne("")
          & edited_df["Observación"].notna()
      ].copy()

      if filas_con_obs.empty:
        st.warning("No hay observaciones escritas para guardar.")
      else:
        filas_con_obs["Fecha Registro"] = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
        columnas_historial = [
            "Fecha Registro",
            "Unidad",
            "Cliente / Grupo",
            "Conexión",
            "Observación",
            "Novedad Batería / Voltaje",
            "WhatsApp",
            "Último Reporte",
            "Días Sin Reporte",
            "Diagnóstico",
        ]
        df_guardar = filas_con_obs[columnas_historial]

        try:
          if os.path.exists(EXCEL_HISTORIAL_PATH):
            df_existente = pd.read_excel(EXCEL_HISTORIAL_PATH)
            df_final = pd.concat([df_existente, df_guardar], ignore_index=True)
          else:
            df_final = df_guardar

          df_final.to_excel(EXCEL_HISTORIAL_PATH, index=False)
          st.success(f"¡Se guardaron {len(df_guardar)} observación(es)!")
        except Exception as ex:
          st.error(
              f"Error al guardar: {ex}. Cierra el archivo Excel si está abierto."
          )

  with col_descargar:
    if os.path.exists(EXCEL_HISTORIAL_PATH):
      with open(EXCEL_HISTORIAL_PATH, "rb") as file_excel:
        st.download_button(
            label="📥 Descargar Excel Histórico",
            data=file_excel,
            file_name="registro_observaciones_korban.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

  st.divider()

  # WhatsApp
  st.subheader("📲 Envío y Personalización de Alertas por WhatsApp")

  unidad_seleccionada = st.selectbox(
      "Selecciona la unidad para preparar la notificación:",
      edited_df["Unidad"].tolist(),
  )

  if unidad_seleccionada:
    fila = edited_df[edited_df["Unidad"] == unidad_seleccionada].iloc[0]
    numero_wa = str(fila["WhatsApp"]).strip()
    obs_text = str(fila["Observación"]).strip()

    mensaje_predeterminado = (
        f"Estimado cliente ({fila['Cliente / Grupo']}), de parte de Korban"
        f" Global Solutions le informamos que la unidad *{fila['Unidad']}*"
        f" presenta el siguiente estado: *{fila['Conexión']}*"
        f" ({fila['Diagnóstico']}).\n\n- Último reporte:"
        f" {fila['Último Reporte']} ({fila['Días Sin Reporte']} días)."
    )
    if obs_text and obs_text != "nan":
      mensaje_predeterminado += f"\n- Observación: {obs_text}"

    mensaje_editado = st.text_area(
        "✏️ Puedes modificar el borrador del mensaje antes de enviarlo:",
        value=mensaje_predeterminado,
        height=140,
    )

    msg_encoded = urllib.parse.quote(mensaje_editado)

    col_info, col_btn_wa = st.columns([2, 1])
    with col_info:
      if not numero_wa or numero_wa == "nan":
        st.warning(
            "⚠️ Escribe el número telefónico en la columna WhatsApp para"
            " habilitar el botón."
        )
      else:
        st.info(f"📱 Número destino: **+{numero_wa}**")

    with col_btn_wa:
      if numero_wa and numero_wa != "nan":
        link_wa = f"https://wa.me/{numero_wa}?text={msg_encoded}"
        st.link_button("💬 Enviar WhatsApp Personalizado", link_wa)