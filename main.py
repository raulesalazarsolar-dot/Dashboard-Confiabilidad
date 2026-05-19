import os
import json
import math
import shutil
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

# ==========================================
# 2. BUSCADORES UNIVERSALES E INTELIGENTES
# ==========================================
def buscar_columna_linea(df):
    for c in df.columns:
        if 'linea' in str(c).strip().lower().replace('í', 'i') and 'aux' not in str(c).lower(): return c
    return None

def buscar_columna_equipo(df, planta_nombre):
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl in ['equipo', 'componente', 'detalle']: return c
    return None

def buscar_columna_semana(df):
    for c in df.columns:
        if 'semana' in str(c).lower() and 'aux' not in str(c).lower():
            if not df[c].isnull().all(): return c
    return None

def buscar_tiempo_detencion_hr(df):
    for c in df.columns:
        if 'detencion' in str(c).lower() and 'hr' in str(c).lower(): return c
    return None

def buscar_columna_tiempo_oper(df):
    for c in df.columns:
        if 'operativo' in str(c).lower() and 'hr' in str(c).lower(): return c
    return None

def buscar_columna_tiempo_plan(df):
    for c in df.columns:
        cl = str(c).lower()
        if ('plan' in cl or 'disponible' in cl) and 'hr' in cl: return c
    return None

# ==========================================
# 3. EXTRACCIÓN Y TRANSFORMACIÓN (ETL)
# ==========================================
def procesar_datos_confiabilidad():
    if os.path.exists(DATA_DIR): shutil.rmtree(DATA_DIR)
    os.makedirs(DATA_DIR, exist_ok=True)
    gdown.download_folder(url=DRIVE_FOLDER_URL, output=DATA_DIR, quiet=True, use_cookies=False)

    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]
    datos_equipos = []
    datos_lineas = []
    
    for archivo_nombre in archivos:
        try:
            excel = pd.ExcelFile(os.path.join(DATA_DIR, archivo_nombre))
            hoja_det = next((h for h in excel.sheet_names if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in excel.sheet_names if h.endswith('_Tiempos_Planificados')), None)
            
            if not hoja_det or not hoja_tpo: continue
                
            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)
            
            # Limpieza básica
            df_det['Linea_Clean'] = df_det[buscar_columna_linea(df_det)].astype(str).str.strip().str.upper()
            df_det['Semana_Clean'] = pd.to_numeric(df_det[buscar_columna_semana(df_det)].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(-1).astype(int)
            
            # Filtrado L2 (Sem 15+) y L4 (Sem 19+)
            df_det = df_det[~((df_det['Linea_Clean'].str.contains('L2', na=False)) & (df_det['Semana_Clean'] < 15))]
            df_det = df_det[~((df_det['Linea_Clean'].str.contains('L4', na=False)) & (df_det['Semana_Clean'] < 19))]
            
            # Procesar tiempos
            col_tpo_oper = buscar_columna_tiempo_oper(df_tpo)
            col_tpo_plan = buscar_columna_tiempo_plan(df_tpo)
            df_tpo['Linea_Clean'] = df_tpo[buscar_columna_linea(df_tpo)].astype(str).str.strip().str.upper()
            df_tpo['Semana_Clean'] = pd.to_numeric(df_tpo[buscar_columna_semana(df_tpo)].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(-1).astype(int)
            
            # Filtro Tiempos
            df_tpo = df_tpo[~((df_tpo['Linea_Clean'].str.contains('L2', na=False)) & (df_tpo['Semana_Clean'] < 15))]
            df_tpo = df_tpo[~((df_tpo['Linea_Clean'].str.contains('L4', na=False)) & (df_tpo['Semana_Clean'] < 19))]
            
            # Consolidar
            agrup_det = df_det.groupby(['Linea_Clean', 'Semana_Clean']).agg(det=('Hrs_Perdidas', 'count'), tpop=('Hrs_Perdidas', 'sum')).reset_index()
            df_tpo['Hrs_O'] = df_tpo[col_tpo_oper] if col_tpo_oper else df_tpo[col_tpo_plan]
            agrup_tpo = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(op=('Hrs_O', 'sum'), pl=('Hrs_Plan', 'sum')).reset_index()
            
            merged = pd.merge(agrup_tpo, agrup_det, on=['Linea_Clean', 'Semana_Clean'], how='outer').fillna(0)
            
            for _, row in merged.iterrows():
                datos_lineas.append({
                    "planta": archivo_nombre, "linea": row['Linea_Clean'], "semana": int(row['Semana_Clean']),
                    "tpo_operativo_linea": float(row['op']), "tpo_plan_linea": float(row['pl'])
                })
        except: continue
        
    return { "equipos": datos_equipos, "lineas": datos_lineas }

# [El resto del generador HTML se mantiene igual, integrando los nuevos KPIs en el dashboard]
# ...
