import io
import json
import os
import urllib.parse
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Diagnóstico Wialon | Korban Global Solutions",
    page_icon="📡",
    layout="wide",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
HISTORIAL_LOCAL = os.path.join(BASE_DIR, "registro_observaciones.xlsx")
URL_WIALON = "https://hst-api.wialon.com/wialon/ajax.html"
COLUMNAS = [
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
# ESTILO
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] {font-family: Inter, sans-serif;}
.stApp {background:#f8fafc;}
[data-testid="stSidebar"] {background:#0f172a;}
[data-testid="stSidebar"] * {color:#f1f5f9;}
[data-testid="stMetric"] {background:white;border:1px solid #e2e8f0;border-radius:12px;padding:14px 18px;}
.stButton button, .stDownloadButton button {border-radius:8px;font-weight:600;}
h1 {color:#0f172a;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# FUNCIONES WIALON
# ============================================================
def wialon_call(svc, params, sid=None):
    payload = {"svc": svc, "params": json.dumps(params, ensure_ascii=False)}
    if sid:
        payload["sid"] = sid
    response = requests.get(URL_WIALON, params=payload, timeout=45)
    response.raise_for_status()
    result = response.json()
    if isinstance(result, dict) and "error" in result:
        raise RuntimeError(f"Wialon devolvió error {result['error']} en {svc}")
    return result

def extraer_grupos(sid):
    result = wialon_call("core/search_items", {
        "spec": {"itemsType": "avl_unit_group", "propName": "sys_name",
                 "propValueMask": "*", "sortType": "sys_name"},
        "force": 1, "flags": 1, "from": 0, "to": 0
    }, sid)
    mapa = {}
    for grupo in result.get("items", []) or []:
        for uid in grupo.get("u", []) or []:
            mapa.setdefault(uid, []).append(grupo.get("nm", "Sin grupo"))
    return mapa

def obtener_unidades(sid):
    # Incluye propiedades básicas, sensores y último mensaje.
    flags = 1 + 1024 + 4096 + 1048576 + 2097152
    result = wialon_call("core/search_items", {
        "spec": {"itemsType": "avl_unit", "propName": "sys_name",
                 "propValueMask": "*", "sortType": "sys_name"},
        "force": 1, "flags": flags, "from": 0, "to": 0
    }, sid)
    return result.get("items", []) or []

def analizar_bateria(unidad):
    mensaje = unidad.get("lmsg") or {}
    parametros = mensaje.get("p") or {}
    sensores = unidad.get("sens") or {}
    sensores = sensores.values() if isinstance(sensores, dict) else sensores if isinstance(sensores, list) else []
    encontrados = []
    sin_bateria = False

    for sensor in sensores:
        if not isinstance(sensor, dict):
