import os
import json
import math
import shutil
import re
import pandas as pd
import gdown
from datetime import datetime
from zoneinfo import ZoneInfo

# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1xXeea_F6HTsI-Wfj7HP2KdCnzGtPIUQq?usp=sharing"
DATA_DIR = "./data"
OUTPUT_HTML = "index.html"

# [Las funciones auxiliares: buscar_columna_linea, buscar_columna_equipo, etc., permanecen igual]
# Para brevedad, asumo que mantienes las funciones de ayuda del script original.

def buscar_columna_linea(df):
    for c in df.columns:
        if 'linea' in str(c).strip().lower().replace('í', 'i') and 'aux' not in str(c).lower():
            return c
    return None

def buscar_columna_equipo(df, planta_nombre):
    planta_lower = str(planta_nombre).lower()
    if "mercadeo" in planta_lower:
        for c in df.columns:
            if str(c).strip().lower() == 'detalle':
                return c
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ['equipo', 'componente', 'detalle']:
            return c
    return None

def buscar_columna_semana(df):
    for c in df.columns:
        if 'semana' in str(c).lower() and 'aux' not in str(c).lower():
            if not df[c].isnull().all():
                return c
    return None

def buscar_tiempo_detencion_hr(df, super_planta):
    for c in df.columns:
        cl = str(c).lower()
        if super_planta == 'Carnes' and 'tpo detenciones' in cl and 'hr' in cl:
            return c
    for c in df.columns:
        if 'detencion' in str(c).lower() and 'hr' in str(c).lower():
            return c
    return None

def buscar_columna_tiempo_oper(df):
    cols_lower = [str(c).lower().replace(' ', ' ').strip() for c in df.columns]
    for i, c in enumerate(cols_lower):
        if ' operativo' in c and 'hr' in c:
            return df.columns[i]
    return None

def limpiar_semana(serie):
    resultado = pd.to_numeric(serie, errors='coerce')
    mask_nan = resultado.isna()
    if mask_nan.any():
        extraidos = serie[mask_nan].astype(str).str.extract(r'(\d{1,2})')[0]
        resultado = resultado.copy()
        resultado[mask_nan] = pd.to_numeric(extraidos, errors='coerce')
    return resultado.fillna(-1).astype(int)

def buscar_columna_tiempo_plan(df, super_planta):
    for c in df.columns:
        cl = str(c).lower().strip()
        if super_planta == 'Carnes' and 'tpo hr plan' in cl:
            return c
        if super_planta == 'Masas' and 'tpo disponible' in cl and 'hr' in cl:
            return c
    cols_lower = [str(c).lower().replace(' ', ' ').strip() for c in df.columns]
    for i, c in enumerate(cols_lower):
        if (' plan' in c or 'disponible' in c) and 'hr' in c:
            return df.columns[i]
    return None

def filtrar_semanas(df):
    df = df[~((df['Linea_Clean'].str.contains('L2', na=False)) & (df['Semana_Clean'] < 15))]
    df = df[~((df['Linea_Clean'].str.contains('L4', na=False)) & (df['Semana_Clean'] < 19))]
    return df

# [Se omite el cuerpo de procesar_datos_confiabilidad por brevedad, es el mismo que tenías]
def procesar_datos_confiabilidad():
    # Lógica de extracción ETL igual a la original
    return {"equipos": [], "lineas": [], "acciones": {}}

# ==========================================
# GENERADOR HTML (CON LOS CAMBIOS SOLICITADOS)
# ==========================================
def generar_html_moderno(db_json):
    # (El bloque de estilos CSS sigue igual hasta .bottom-title)
    html_template = """... [AQUÍ VA TODO TU CSS ORIGINAL] ...
.bottom-title { font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:1px; color:rgba(255,255,255,0.7); margin-bottom:15px; display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center;}
.bottom-title strong { font-size: 11px; color:#fff; background: rgba(0,0,0,0.2); padding: 2px 6px; border-radius: 4px; margin-right: 5px;}
... [EL RESTO DEL HTML Y JS] ...
function renderResumen() {
  // ... (código de filtros y agrupaciones igual hasta lineas)
  
  // Calcular promedios globales de planta (NUEVA LÓGICA)
  const totalEquipos = eqs.length;
  const promMTBF = totalEquipos > 0 ? eqs.reduce((s, e) => {
      const lnOp = lineas[e.linea]?.op || 0;
      return s + (e.det > 0 ? (lnOp / e.det) : 0);
  }, 0) / totalEquipos : 0;

  const promMTTR = totalEquipos > 0 ? eqs.reduce((s, e) => s + (e.det > 0 ? (e.hrs / e.det) : 0), 0) / totalEquipos : 0;
  
  const promConf = totalEquipos > 0 ? eqs.reduce((s, e) => {
      const lnOp = lineas[e.linea]?.op || 0;
      const mtbf = e.det > 0 ? (lnOp / e.det) : 0;
      return s + (mtbf > 0 ? Math.exp(-120 / mtbf) * 100 : (e.det === 0 ? 100 : 0));
  }, 0) / totalEquipos : 0;

  const promMant = totalEquipos > 0 ? eqs.reduce((s, e) => {
      const mttr = e.det > 0 ? (e.hrs / e.det) : 0;
      return s + (mttr > 0 ? (1 - Math.exp(-1 / mttr)) * 100 : 100);
  }, 0) / totalEquipos : 0;
  
  const promProb = 100 - promConf;

  // En la generación de HTML:
  html += `
    <div class="bottom-section ${tema}">
      <div class="bottom-title">
        <span>Estado por Línea de Producción</span>
        <div style="display:flex; gap:12px;">
            <span>MTBF: <strong>${promMTBF.toFixed(1)}h</strong></span>
            <span>MTTR: <strong>${promMTTR.toFixed(2)}h</strong></span>
            <span>Conf: <strong>${promConf.toFixed(1)}%</strong></span>
            <span>Mant: <strong>${promMant.toFixed(1)}%</strong></span>
            <span>Pb. Falla: <strong>${promProb.toFixed(1)}%</strong></span>
        </div>
      </div>
      <div class="lines-grid"> ${barras} </div>
    </div>`;
}
"""
    # ... resto de la lógica de escritura de archivo
