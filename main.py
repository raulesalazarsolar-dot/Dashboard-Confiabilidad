import os
import json
import math
import shutil
import re
import base64
import pandas as pd
import gdown
from datetime import datetime
from zoneinfo import ZoneInfo

# ==========================================
# 1. CONFIGURACIÓN (GOOGLE DRIVE)
# ==========================================
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1xXeea_F6HTsI-Wfj7HP2KdCnzGtPIUQq?usp=sharing"
DATA_DIR = "./data"
OUTPUT_HTML = "index.html"
AVATAR_IMG = "avatar_ways.png" # Guarda tu imagen sin fondo con este nombre en la misma carpeta

# ==========================================
# 2. BUSCADORES UNIVERSALES E INTELIGENTES Y MAPEO
# ==========================================
def mapear_linea(linea_str, planta_str):
    l = str(linea_str).strip().lower()
    p = str(planta_str).strip().lower()
    
    if "panader" in p:
        if "l1" in l or "hallulla" in l and "marraqueta" not in l and "l2" not in l: return "L1"
        if "l3" in l or "surtido" in l and "l5" not in l: return "L3"
        if "l4" in l or ("hallulla" in l and "marraqueta" in l): return "L4"
        if "l5" in l: return "L5"
        if "l2" in l or "marraquetas" in l: return "L2"
        if "multivac 1" in l or "m1" in l: return "M1"
        if "multivac 3" in l or "m3" in l: return "M3"
        if "variovac" in l: return "VPN"
    elif "boller" in p:
        if "empanadas" in l or "fritsch empanadas" in l: return "Empanada"
        if "bolleria" in l: return "Bollería"
        if "pizzas" in l or "fritsch pizzas" in l: return "Pizza"
        if "variovac pizza" in l or "vpz" in l: return "VPZ"
        if "flow pack 3" in l or "fp3" in l: return "FP3"
        if "flow pack" in l or "fp" in l: return "FP"
        
    return str(linea_str).strip().upper()

def buscar_columna_linea(df):
    for c in df.columns:
        if 'linea' in str(c).strip().lower().replace('í', 'i') and 'aux' not in str(c).lower():
            return c
    return None

def buscar_columna_equipo(df, planta_nombre):
    planta_lower = str(planta_nombre).lower()
    if "carne" in planta_lower or "mercadeo" in planta_lower or "molida" in planta_lower:
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

# ==========================================
# 4. EXTRACCIÓN Y TRANSFORMACIÓN (ETL)
# ==========================================
def procesar_datos_confiabilidad():
    print("🔗 Conectando a Google Drive...")
    if os.path.exists(DATA_DIR):
        shutil.rmtree(DATA_DIR)
    os.makedirs(DATA_DIR, exist_ok=True)

    try:
        gdown.download_folder(url=DRIVE_FOLDER_URL, output=DATA_DIR, quiet=False, use_cookies=False)
    except Exception as e:
        print(f"❌ Error al descargar de Drive: {e}")

    print("\n📂 INICIANDO EXTRACCIÓN DE DATOS BASE...")
    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]

    datos_equipos = []
    datos_lineas = []
    datos_acciones = {}

    archivo_acciones = next((f for f in archivos if 'acciones' in f.lower()), None)
    if archivo_acciones:
        print(f"\n📝 Leyendo Acciones Correctivas desde: {archivo_acciones}")
        try:
            ruta_acc = os.path.join(DATA_DIR, archivo_acciones)
            df_acc = pd.read_excel(ruta_acc)
            df_acc.columns = [str(c).lower().strip() for c in df_acc.columns]

            col_top = next((c for c in df_acc.columns if 'top' in c), None)
            if col_top:
                for _, row in df_acc.iterrows():
                    top_rank = pd.to_numeric(row[col_top], errors='coerce')
                    if pd.isna(top_rank): continue
                    top_rank = int(top_rank)

                    for col in df_acc.columns:
                        if col != col_top:
                            if col not in datos_acciones: datos_acciones[col] = {}
                            val = row[col]
                            datos_acciones[col][top_rank] = str(val).strip() if pd.notna(val) else "N/A"
                print("  ✅ Acciones correctivas extraídas correctamente.")
            else:
                print("  ⚠️ No se encontró columna 'top' en el archivo de acciones.")
        except Exception as e:
            print(f"  ❌ Error leyendo archivo de acciones: {e}")

    for archivo_nombre in archivos:
        if 'acciones' in archivo_nombre.lower(): continue

        ruta_completa = os.path.join(DATA_DIR, archivo_nombre)
        print(f"\n🔍 Analizando: {archivo_nombre}")
        try:
            excel = pd.ExcelFile(ruta_completa)
            hoja_det = next((h for h in excel.sheet_names if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in excel.sheet_names if h.endswith('_Tiempos_Planificados')), None)

            if not hoja_det or not hoja_tpo:
                print("  ⚠️ Faltan pestañas base. Saltando...")
                continue

            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)

            planta_nombre = re.sub(r'(?i)confiabilidad', '', archivo_nombre)
            planta_nombre = re.sub(r'(?i)\.xlsx', '', planta_nombre).strip()

            planta_lower = planta_nombre.lower()
            if "carne" in planta_lower or "mercadeo" in planta_lower or "molida" in planta_lower:
                super_planta = "Carnes"
            else:
                super_planta = "Masas"

            col_equipo      = buscar_columna_equipo(df_det, planta_nombre)
            col_semana_det  = buscar_columna_semana(df_det)
            col_linea_det   = buscar_columna_linea(df_det)
            col_tpo_det     = buscar_tiempo_detencion_hr(df_det, super_planta)

            if not all([col_equipo, col_semana_det, col_linea_det, col_tpo_det]):
                print(f"  ❌ Faltan columnas vitales en FEM para {planta_nombre}. Saltando...")
                continue

            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas']  = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            df_det['Linea_Clean']   = df_det[col_linea_det].apply(lambda x: mapear_linea(x, planta_nombre))
            df_det['Equipo_Clean']  = df_det[col_equipo].astype(str).str.strip().str.title()
            df_det['Semana_Clean']  = limpiar_semana(df_det[col_semana_det])
            df_det = df_det[df_det['Semana_Clean'] > 0]

            df_det = filtrar_semanas(df_det)

            agrup_det_linea = df_det.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_perdido_linea=('Hrs_Perdidas', 'sum')
            ).reset_index()

            col_semana_tpo  = buscar_columna_semana(df_tpo)
            col_linea_tpo   = buscar_columna_linea(df_tpo)
            col_tpo_oper    = buscar_columna_tiempo_oper(df_tpo)
            col_tpo_plan    = buscar_columna_tiempo_plan(df_tpo, super_planta)

            if not all([col_semana_tpo, col_linea_tpo]) or (not col_tpo_oper and not col_tpo_plan):
                print(f"  ❌ Faltan columnas de tiempo en {archivo_nombre}. Saltando...")
                continue

            df_tpo = df_tpo.dropna(subset=[col_linea_tpo, col_semana_tpo])
            df_tpo['Hrs_Oper']      = pd.to_numeric(df_tpo[col_tpo_oper], errors='coerce').fillna(0) if col_tpo_oper else 0
            df_tpo['Hrs_Plan']      = pd.to_numeric(df_tpo[col_tpo_plan], errors='coerce').fillna(0) if col_tpo_plan else 0
            df_tpo['Linea_Clean']   = df_tpo[col_linea_tpo].apply(lambda x: mapear_linea(x, planta_nombre))
            df_tpo['Semana_Clean']  = limpiar_semana(df_tpo[col_semana_tpo])
            df_tpo = df_tpo[df_tpo['Semana_Clean'] > 0]

            df_tpo = filtrar_semanas(df_tpo)

            agrup_tpo_linea = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_operativo_linea=('Hrs_Oper', 'sum'),
                tpo_plan_linea=('Hrs_Plan', 'sum')
            ).reset_index()

            linea_merged = pd.merge(agrup_tpo_linea, agrup_det_linea, on=['Linea_Clean', 'Semana_Clean'], how='outer')
            linea_merged['tpo_perdido_linea']   = linea_merged.get('tpo_perdido_linea',   pd.Series([0] * len(linea_merged))).fillna(0)
            linea_merged['tpo_operativo_linea'] = linea_merged.get('tpo_operativo_linea', pd.Series([0] * len(linea_merged))).fillna(0)
            linea_merged['tpo_plan_linea']      = linea_merged.get('tpo_plan_linea',      pd.Series([0] * len(linea_merged))).fillna(0)

            tipo_tiempo = 'operativo' if col_tpo_oper else 'plan'

            for idx, row in linea_merged.iterrows():
                plan    = row['tpo_plan_linea']
                oper    = row['tpo_operativo_linea']
                perdido = row['tpo_perdido_linea']

                if super_planta == 'Carnes':
                    linea_merged.at[idx, 'tpo_operativo_linea'] = max(0, plan - perdido)
                else:
                    if tipo_tiempo == 'operativo':
                        if oper == 0 and plan > 0:
                            linea_merged.at[idx, 'tpo_operativo_linea'] = plan
                    else:
                        if oper == 0 and plan > 0:
                            linea_merged.at[idx, 'tpo_operativo_linea'] = max(0, plan - perdido)
                        elif plan == 0 and oper > 0:
                            linea_merged.at[idx, 'tpo_plan_linea'] = oper + perdido

            tiempos_lineas_madre = {}
            for idx, row_lm in linea_merged.iterrows():
                l_str  = str(row_lm['Linea_Clean']).upper()
                sem    = row_lm['Semana_Clean']
                tpo_op = row_lm['tpo_operativo_linea']
                tpo_pl = row_lm['tpo_plan_linea']
                if tpo_op > 0:
                    for pref in ['L1', 'L2', 'L3', 'L4', 'L5']:
                        if pref in l_str:
                            tiempos_lineas_madre[(pref, sem)] = {'op': tpo_op, 'pl': tpo_pl}

            for idx, row_lm in linea_merged.iterrows():
                l_str = str(row_lm['Linea_Clean']).upper()
                if row_lm['tpo_operativo_linea'] == 0 or 'MULTIVAC' in l_str or 'VARIOVAC' in l_str or 'M1' in l_str or 'M3' in l_str or 'VPN' in l_str:
                    sem        = row_lm['Semana_Clean']
                    target_key = None
                    if   'MULTIVAC 1' in l_str or 'M1' in l_str:                       target_key = 'L1'
                    elif 'MULTIVAC 2' in l_str or 'M2' in l_str or 'VARIOVAC' in l_str or 'VPN' in l_str: target_key = 'L2'
                    elif 'MULTIVAC 3' in l_str or 'M3' in l_str:                       target_key = 'L3'
                    if target_key and (target_key, sem) in tiempos_lineas_madre:
                        linea_merged.at[idx, 'tpo_operativo_linea'] = tiempos_lineas_madre[(target_key, sem)]['op']
                        linea_merged.at[idx, 'tpo_plan_linea']       = tiempos_lineas_madre[(target_key, sem)]['pl']

            for idx, row in linea_merged.iterrows():
                if row['tpo_operativo_linea'] <= 0:
                    herencia = linea_merged[
                        (linea_merged['Semana_Clean'] == row['Semana_Clean']) &
                        (linea_merged['Linea_Clean'].str.contains('L2', na=False))
                    ]
                    if not herencia.empty:
                        linea_merged.at[idx, 'tpo_operativo_linea'] = herencia['tpo_operativo_linea'].max()

            for _, row in linea_merged.iterrows():
                datos_lineas.append({
                    "super_planta":         super_planta,
                    "planta":               planta_nombre,
                    "linea":                row['Linea_Clean'],
                    "semana":               int(row['Semana_Clean']),
                    "tpo_operativo_linea":  float(row.get('tpo_operativo_linea', 0)),
                    "tpo_plan_linea":       float(row.get('tpo_plan_linea', 0)),
                })

            for _, row in df_det.iterrows():
                datos_equipos.append({
                    "super_planta":  super_planta,
                    "planta":        planta_nombre,
                    "linea":         row['Linea_Clean'],
                    "equipo":        row['Equipo_Clean'],
                    "semana":        int(row['Semana_Clean']),
                    "detenciones":   1,
                    "tpo_perdido_eq": float(row['Hrs_Perdidas']),
                    "fecha":         str(row['Fecha'])[:10] if 'Fecha' in df_det.columns else 'N/A',
                    "componente":    str(row['Componente']) if 'Componente' in df_det.columns else 'N/A',
                    "tipo":          str(row['Tipo Detención']) if 'Tipo Detención' in df_det.columns else 'N/A',
                })

            print(f"  ✅ Procesado con éxito. Extraídas detenciones de {planta_nombre}.")

        except Exception as e:
            print(f"  ❌ Error fatal procesando {archivo_nombre}: {e}")

    db_json = {"equipos": datos_equipos, "lineas": datos_lineas, "acciones": datos_acciones}
    print("\n✅ Extracción finalizada. Datos listos para el Dashboard.")
    return db_json

# ==========================================
# 5. GENERADOR HTML DASHBOARD
# ==========================================
def generar_html_moderno(db_json):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    
    # Procesar imagen del Avatar si existe
    avatar_base64 = ""
    if os.path.exists(AVATAR_IMG):
        try:
            with open(AVATAR_IMG, "rb") as image_file:
                avatar_base64 = base64.b64encode(image_file.read()).decode('utf-8')
        except Exception as e:
            print(f"⚠️ Error al leer imagen de avatar: {e}")

    html_template = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Confiabilidad</title>
<link rel="icon" type="image/x-icon" href="https://www.walmart.com/favicon.ico">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800;900&display=swap');
:root {
  --primary: #0D2C54; --secondary: #3A4A5C; --accent: #0071CE;
  --bg: #F2F6FC; --surface: #ffffff; --border: #DDE6F2;
  --text: #0D2C54; --text-muted: #8899AA;
  --success: #27AE60; --danger: #C0392B; --warning: #E67E22;
}
body.theme-carnes {
  --primary: #4A0E0E; --secondary: #7f1d1d; --accent: #A93226;
  --bg: #fef2f2; --border: #fecaca;
}
* { box-sizing: border-box; outline: none; font-family: 'DM Sans', system-ui, sans-serif; -webkit-font-smoothing: antialiased; }
body { background: var(--bg); color: var(--text); margin: 0; display: flex; flex-direction: column; min-height: 100vh; transition: background 0.4s; }
.top-bar { background: var(--primary); color: white; padding: 0 25px; height: 65px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); z-index: 10; transition: background 0.4s; border-bottom:4px solid var(--accent); }
.brand { display: flex; align-items: center; gap: 15px; }
.brand h2 { margin: 0; font-size: 1.3rem; font-weight: 700; letter-spacing: 0.5px; }
.brand span { opacity: 0.7; font-weight: 400; font-size: 1rem; border-left: 1px solid rgba(255,255,255,0.3); padding-left: 15px; }
.planta-switch { display: flex; align-items: center; gap: 12px; font-
