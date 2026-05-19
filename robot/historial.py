# =====================================================
# 📜 MÓDULO: HISTORIAL DE DESCARGAS Y AUDITORÍA
# =====================================================
# Guarda y recupera el historial de ejecuciones del robot.
# Estructura de archivo: historial_descargas.json
# =====================================================

import json
import re
from datetime import datetime
from pathlib import Path
import pandas as pd

# Ruta base (Render / Docker / local)
BASE_DIR = Path(__file__).resolve().parent.parent
HIST_DIR = BASE_DIR / "historiales"
HIST_DIR.mkdir(parents=True, exist_ok=True)
LEGACY_PATH = BASE_DIR / "historial_descargas.json"


def _sanear_nombre(nombre: str) -> str:
    if not nombre:
        return "general"
    limpio = re.sub(r"[^a-zA-Z0-9_-]", "", nombre)
    return limpio or "general"


def _historial_path(device_id: str | None) -> Path:
    if not device_id:
        return LEGACY_PATH
    return HIST_DIR / f"historial_{_sanear_nombre(device_id)}.json"

# =====================================================
# 🧾 REGISTRAR DESCARGA
# =====================================================
def registrar_descarga(ruc, origen, anio, mes, dia, tipo, resultado, device_id=None):
    """
    Registra una ejecución del robot en el archivo JSON.
    Crea el historial si no existe.
    """
    registro = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ruc": ruc,
        "origen": origen,
        "anio": int(anio),
        "mes": int(mes),
        "tipo": tipo,
        "dia": int(dia) if isinstance(dia, (int, float)) else dia,
        "estado": resultado.get("estado", "finalizado"),
        "n_xml": resultado.get("n_xml", 0),
        "n_pdf": resultado.get("n_pdf", 0),
        "n_registros": resultado.get("n_registros", 0),
    }
    if isinstance(resultado, dict):
        for campo in (
            "fecha_filtro",
            "estado_autorizacion",
            "establecimiento",
            "punto_emision",
            "tipo_visible",
        ):
            valor = resultado.get(campo)
            if valor:
                registro[campo] = valor

        reporte_path = resultado.get("reporte")
        if reporte_path:
            try:
                registro["reporte"] = str(Path(reporte_path).resolve())
            except Exception:
                registro["reporte"] = str(reporte_path)
        reporte_pdf_path = resultado.get("reporte_pdf")
        if reporte_pdf_path:
            try:
                registro["reporte_pdf"] = str(Path(reporte_pdf_path).resolve())
            except Exception:
                registro["reporte_pdf"] = str(reporte_pdf_path)
        reporte_xml_path = resultado.get("reporte_xml")
        if reporte_xml_path:
            try:
                registro["reporte_xml"] = str(Path(reporte_xml_path).resolve())
            except Exception:
                registro["reporte_xml"] = str(reporte_xml_path)
        reporte_pdf_anual = resultado.get("reporte_pdf_anual")
        if reporte_pdf_anual:
            try:
                registro["reporte_pdf_anual"] = str(Path(reporte_pdf_anual).resolve())
            except Exception:
                registro["reporte_pdf_anual"] = str(reporte_pdf_anual)
        reporte_xml_anual = resultado.get("reporte_xml_anual")
        if reporte_xml_anual:
            try:
                registro["reporte_xml_anual"] = str(Path(reporte_xml_anual).resolve())
            except Exception:
                registro["reporte_xml_anual"] = str(reporte_xml_anual)
        reportes_xml = resultado.get("reportes_xml")
        if reportes_xml:
            registro["reportes_xml"] = reportes_xml
        reportes_pdf = resultado.get("reportes_pdf")
        if reportes_pdf:
            registro["reportes_pdf"] = reportes_pdf


    # Leer historial existente o iniciar lista vacía
    hist_path = _historial_path(device_id)

    if hist_path.exists():
        try:
            with open(hist_path, "r", encoding="utf-8") as f:
                historial = json.load(f)
                if not isinstance(historial, list):
                    historial = []
        except Exception:
            historial = []
    else:
        historial = []

    # Agregar nuevo registro y guardar
    historial.append(registro)
    hist_path.parent.mkdir(parents=True, exist_ok=True)
    with open(hist_path, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)

    return registro


# =====================================================
# 📊 OBTENER HISTORIAL
# =====================================================
def obtener_historial(device_id=None):
    """
    Devuelve el historial como DataFrame ordenado (más recientes primero).
    Si no existe, devuelve un DataFrame vacío.
    """
    hist_path = _historial_path(device_id)

    if not hist_path.exists():
        return pd.DataFrame()

    try:
        with open(hist_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return pd.DataFrame()
        df = pd.DataFrame(data)
        if not df.empty:
            df = df.sort_values(by="timestamp", ascending=False).reset_index(drop=True)
        return df
    except Exception:
        # Si el JSON está corrupto, devolver vacío
        return pd.DataFrame()
