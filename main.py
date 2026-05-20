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
# 1. CONFIGURACIÓN (GOOGLE DRIVE)
# ==========================================
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1xXeea_F6HTsI-Wfj7HP2KdCnzGtPIUQq?usp=sharing"
DATA_DIR = "./data"
OUTPUT_HTML = "index.html"

# ==========================================
# 2. BUSCADORES UNIVERSALES E INTELIGENTES Y MAPEO
# ==========================================
def mapear_linea(linea_str, planta_str):
    l = str(linea_str).strip().lower()
    p = str(planta_str).strip().lower()
    
    # Mapeos directos según la imagen de referencia para asegurar que se apliquen siempre
    if "flow pack 3" in l or "fp3" in l: return "FP3"
    if "flow pack" in l or "fp" in l and "fp3" not in l: return "FP"
    if "multivac 2" in l or "m2" in l: return "M2"
    if "variovac pizza" in l or "vpz" in l: return "VPZ"
    if "fritsch empanadas" in l or "empanadas" in l: return "Empanada"
    if "fritsch pizzas" in l or "pizzas" in l: return "Pizza"

    if "panader" in p:
        if "l1" in l or ("hallulla" in l and "marraqueta" not in l and "l2" not in l): return "L1"
        if "l3" in l or ("surtido" in l and "l5" not in l): return "L3"
        if "l4" in l or ("hallulla" in l and "marraqueta" in l): return "L4"
        if "l5" in l: return "L5"
        if "l2" in l or "marraquetas" in l: return "L2"
        if "multivac 1" in l or "m1" in l: return "M1"
        if "multivac 3" in l or "m3" in l: return "M3"
        if "variovac" in l and "pizza" not in l: return "VPN"
    elif "boller" in p:
        if "bolleria" in l: return "Bollería"
        
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

def buscar_columna_mes(df):
    for c in df.columns:
        if str(c).strip().lower() == 'mes':
            return c
    # Si no la encuentra como 'mes', busca en la columna C (índice 2)
    if len(df.columns) > 2 and 'mes' in str(df.columns[2]).lower():
        return df.columns[2]
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

    meses_dict = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 
                  7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

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
            col_mes         = buscar_columna_mes(df_det)

            if not all([col_equipo, col_semana_det, col_linea_det, col_tpo_det]):
                print(f"  ❌ Faltan columnas vitales en FEM para {planta_nombre}. Saltando...")
                continue

            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas']  = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            df_det['Linea_Clean']   = df_det[col_linea_det].apply(lambda x: mapear_linea(x, planta_nombre))
            df_det['Equipo_Clean']  = df_det[col_equipo].astype(str).str.strip().str.title()
            df_det['Semana_Clean']  = limpiar_semana(df_det[col_semana_det])
            
            if col_mes:
                df_det['Mes_Clean'] = pd.to_numeric(df_det[col_mes], errors='coerce').map(meses_dict).fillna('N/A')
            else:
                df_det['Mes_Clean'] = 'N/A'

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
                    "mes":           str(row['Mes_Clean']) if 'Mes_Clean' in df_det.columns else 'N/A',
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
.planta-switch { display: flex; align-items: center; gap: 12px; font-weight: 700; font-size: 0.95rem; background: rgba(0,0,0,0.25); padding: 6px 20px; border-radius: 30px; border: 1px solid rgba(255,255,255,0.1); }
.planta-switch span { opacity: 0.5; transition: 0.3s; }
.planta-switch span.active { opacity: 1; color: #fff; }
.switch { position: relative; display: inline-block; width: 48px; height: 24px; }
.switch input { opacity: 0; width: 0; height: 0; }
.slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: rgba(255,255,255,0.3); transition: .4s; border-radius: 34px; }
.slider:before { position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px; background-color: white; transition: .4s; border-radius: 50%; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }
input:checked + .slider { background-color: #ef4444; }
input:checked + .slider:before { transform: translateX(24px); }
.main-container { display: flex; flex: 1; overflow: hidden; }
.sidebar { width: 280px; background: #ffffff; border-right: 1px solid var(--border); display: flex; flex-direction: column; }
.filter-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center; }
.filter-body { padding: 20px; flex: 1; overflow-y: auto; }
.f-group { margin-bottom: 20px; }
.f-group label { display: block; font-size: 0.75rem; font-weight: 700; color: var(--text-muted); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }
select { width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.9rem; color: var(--text); background: var(--bg); cursor: pointer; transition: 0.2s; }
select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(0,113,206,0.2); }
.rango-semanas { display: flex; gap: 10px; align-items: center; }
.rango-semanas select { flex: 1; }
.rango-semanas span { font-weight: 700; color: var(--text-muted); }
.btn-export { background: #ffffff; border: 2px solid var(--success); color: var(--success); padding: 10px; border-radius: 8px; cursor: pointer; font-weight: 700; width: 100%; display: flex; justify-content: center; gap: 8px; transition: 0.2s; margin-top: 10px; }
.btn-export:hover { background: var(--success); color: white; }

/* ESTILOS DE TEAM BRANDING FLUIDOS */
.team-branding { text-align: center; display: flex; flex-direction: column; align-items: center; gap: 10px; margin-bottom: 18px; width: 100%; background: #F8FAFD; padding: 12px; border-radius: 10px; border: 1px dashed var(--border); }
body.theme-carnes .team-branding { background: #FFF5F5; }
.team-branding p { font-size: 0.75rem; font-weight: 800; color: var(--text); text-transform: uppercase; letter-spacing: 0.5px; margin: 0; line-height: 1.3; }

.content { flex: 1; padding: 30px; overflow-y: auto; display: flex; flex-direction: column; gap: 25px; }
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 15px; }
.kpi-card { background: #ffffff; padding: 15px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; gap: 5px; position: relative; overflow: hidden; }
.kpi-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--accent); }
.kpi-card.c-red::before { background: var(--danger); }
.kpi-card.c-green::before { background: var(--success); }
.kpi-card.c-warn::before { background: var(--warning); }
.kpi-card span { font-size: 0.75rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
.kpi-card h3 { margin: 0; font-size: 1.8rem; color: var(--secondary); font-weight: 800; }
.charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(500px, 1fr)); gap: 25px; }
.chart-container { background: #ffffff; padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); height: 350px; display: flex; flex-direction: column; }
.chart-header { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; display: flex; justify-content: space-between; }
.canvas-wrapper { flex: 1; position: relative; min-height: 0; }
.table-container { background: #ffffff; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); overflow: hidden; display: flex; flex-direction: column; }
.table-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center; }
.table-header input { padding: 8px 15px; border: 1px solid var(--border); border-radius: 20px; font-size: 0.9rem; width: 300px; background: var(--bg); }
.table-wrapper { overflow-x: auto; max-height: 500px; }
table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.85rem; }
th { background: #F8FAFD; padding: 12px 15px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; position: sticky; top: 0; z-index: 2; border-bottom: 2px solid var(--border); cursor: pointer; }
body.theme-carnes th { background: #fef2f2; }
td { padding: 12px 15px; border-bottom: 1px solid var(--border); color: var(--secondary); font-weight: 500; }
tr.clickable-row { cursor: pointer; transition: background 0.15s; }
tr.clickable-row:hover td { background: rgba(0,113,206,0.04) !important; }
body.theme-carnes tr.clickable-row:hover td { background: rgba(169,50,38,0.04) !important; }
.badge { padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; }
.b-ok { background: #C8E6C9; color: #1B5E20; }
.b-warn { background: #FFF9C4; color: #827717; }
.b-danger { background: #FFCDD2; color: #B71C1C; }
.modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(13,44,84,0.6); backdrop-filter: blur(5px); display: none; justify-content: center; align-items: center; z-index: 100; padding: 20px; }
.modal-content { background: #ffffff; width: 100%; max-width: 850px; border-radius: 16px; border: 1px solid var(--border); box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1); display: flex; flex-direction: column; max-height: 80vh; animation: modalIn 0.2s ease-out; }
@keyframes modalIn { from { transform: translateY(15px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
.modal-header { padding: 20px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; background: #F8FAFD; border-top-left-radius: 16px; border-top-right-radius: 16px; }
body.theme-carnes .modal-header { background: #fef2f2; }
.modal-header h3 { margin: 0; font-size: 1.2rem; font-weight: 800; color: var(--secondary); }
.modal-header p { margin: 2px 0 0 0; font-size: 0.85rem; color: var(--text-muted); font-weight: 500; }
.modal-body { padding: 20px; overflow-y: auto; }
.close-btn { background: none; border: none; font-size: 1.5rem; color: var(--text-muted); cursor: pointer; font-weight: 700; }
.close-btn:hover { color: var(--danger); }
.modal-body table th { position: static; background: #F8FAFD; border-bottom: 1px solid var(--border); }
body.theme-carnes .modal-body table th { background: #fee2e2; }

/* ── TABS ── */
.tab-nav { background: #ffffff; border-bottom: 2px solid var(--border); display: flex; gap: 2px; padding: 0 20px; flex-shrink: 0; }
.tab-btn { padding: 11px 22px; font-size: 0.85rem; font-weight: 700; color: var(--text-muted); border: none; background: none; cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; transition: 0.2s; }
.tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
body.theme-carnes .tab-btn.active { color: #A93226; border-bottom-color: #A93226; }
.tab-btn:hover:not(.active) { color: var(--text); background: var(--bg); border-radius: 6px 6px 0 0; }
.tab-panel { display: none; flex: 1; overflow: hidden; }
.tab-panel.active { display: flex; }

/* ── NUEVO DISEÑO RESUMEN V3 ── */
.resumen-panel { display:flex; flex-direction:column; padding: 20px 30px; overflow-y: auto; background: #F4F6FA; width:100%; gap: 25px;}
.fila-planta-v3 { display:flex; flex-direction:column; background: #fff; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); overflow:hidden;}
.top-section { padding: 20px 25px; }
.plant-header { display:flex; align-items:center; gap: 12px; margin-bottom: 20px; }
.plant-badge { color:#fff; padding:5px 12px; border-radius:4px; font-size:12px; font-weight:800; letter-spacing:1px; text-transform:uppercase; }
.plant-header h3 { margin:0; font-size:16px; color:var(--primary); font-weight:800; }
.grid-table { width: 100%; }

/* GRID ACTUALIZADA CON COLUMNA PARA EL GRAFICO MES */
.grid-th { display:grid; grid-template-columns: 25px 1.5fr 130px 85px 85px 80px 80px 2fr; gap:10px; padding-bottom:8px; border-bottom:2px solid var(--accent); margin-bottom:5px; }
.grid-th > div { font-size:9.5px; font-weight:800; text-transform:uppercase; letter-spacing:.7px; color:var(--text-muted); align-self: end; line-height: 1.2;}
.grid-th .center { text-align: center; }
.grid-th small { font-size: 7.5px; font-weight: 600; opacity: 0.8; display: block; margin-top: 2px;}
.grid-tr { display:grid; grid-template-columns: 25px 1.5fr 130px 85px 85px 80px 80px 2fr; gap:10px; align-items:center; padding:10px 0; border-bottom:1px solid var(--border); }
.grid-tr:last-child { border-bottom:none; }

.rank-circle { width:22px; height:22px; border-radius:50%; color:#fff; font-size:10px; font-weight:800; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.eq-name { font-size:12px; font-weight:700; color:var(--primary); line-height:1.2;}
.line-badge { padding:2px 6px; border-radius:4px; font-size:8px; font-weight:800; background:var(--danger); color:#fff; margin-left:6px; vertical-align:middle; }
.stat-box { text-align:center; border-radius:4px; padding:6px 4px; display:flex; flex-direction:column; justify-content:center; align-items:center; }
.stat-val { font-size:14px; font-weight:800; line-height:1; }
.stat-trend { font-size:9px; font-weight:800; margin-top: 3px; display:flex; align-items:center; gap:2px; }
.trend-up { color: #C0392B; } 
.trend-dn { color: #27AE60; } 
.trend-neu { color: #95A5A6; }
.bg-red-light { background: #FEF0F0; }
.bg-blue-light { background: #EEF4FF; }
.bg-gray-light { background: #F8FAFD; }

/* Bottom section (Bars) */
.bottom-section { padding: 15px 25px; color: #fff; }
.bottom-section.tema-masas { background: #1A3A5C; }
.bottom-section.tema-carnes { background: #4A0E0E; }
.bottom-section.tema-dely { background: #2C4A3E; }
.bottom-section.tema-molida { background: #4A2511; }

.section-label { font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:1px; color:rgba(255,255,255,0.7); margin-bottom:8px; display:block; }
.lines-grid { display: flex; flex-wrap: wrap; gap: 20px; }
.line-item { flex: 1; min-width: 220px; }

/* --- CONTENEDORES PARA ENCABEZADO DIVIDIDO Y ROTATIVO --- */
.indicators-header { display: flex; flex-wrap: wrap; gap: 20px; margin-bottom: 20px; align-items: stretch; }
.zona-block { flex: 1 1 320px; display: flex; flex-direction: column; }
.line-rotator-block { flex: 1 1 320px; display: flex; flex-direction: column; border-left: 2px solid rgba(255,255,255,0.1); padding-left: 20px; }
.kpis-planta { display:flex; gap:12px; font-size:10px; align-items:center; background: rgba(0,0,0,0.25); padding: 8px 12px; border-radius: 6px; flex-wrap: wrap; flex:1; }
.kpis-planta span strong { font-size: 12px; color:#fff; margin-left: 3px; }

.kpis-rotator { display:flex; gap:12px; font-size:10px; align-items:center; background: rgba(255,255,255,0.1); padding: 8px 12px; border-radius: 6px; flex-wrap: wrap; transition: opacity 0.5s ease; flex:1; }
.kpis-rotator span strong { font-size: 12px; color:#fff; margin-left: 3px; }

@media (max-width: 768px) {
    .line-rotator-block { border-left: none; border-top: 2px solid rgba(255,255,255,0.1); padding-left: 0; padding-top: 15px; }
}
</style>
</head>
<body>
<div class="top-bar">
  <div class="brand">
    <img src="https://upload.wikimedia.org/wikipedia/commons/b/b1/Walmart_logo_%282008%29.svg" alt="Walmart" style="height:30px;filter:brightness(0) invert(1);">
    <h2>Dashboard Confiabilidad <span>Libro Confiabilidad Walmart 2026</span></h2>
  </div>
  <div class="planta-switch">
    <span id="lbl_masas" class="active">Masas</span>
    <label class="switch">
      <input type="checkbox" id="theme_toggle" onchange="toggleTheme()">
      <span class="slider"></span>
    </label>
    <span id="lbl_carnes">Carnes</span>
  </div>
</div>

<div class="tab-nav">
  <button class="tab-btn active" id="tbtn_resumen" onclick="switchTab('resumen')">&#128202; Resumen Semanal</button>
  <button class="tab-btn" id="tbtn_analisis" onclick="switchTab('analisis')">&#128270; An&#225;lisis Detallado</button>
</div>

<div class="main-container">
  <div class="sidebar">
    <div class="filter-header"><span>Filtros Acumulados</span></div>
    <div class="filter-body">
      <div class="f-group">
        <label>📅 Rango de Semanas</label>
        <div class="rango-semanas">
          <select id="f_sem_desde" onchange="applyFilters()"></select>
          <span>a</span>
          <select id="f_sem_hasta" onchange="applyFilters()"></select>
        </div>
      </div>
      <div id="filters_dynamic"></div>
    </div>
    <div style="padding:20px; border-top:1px solid var(--border); display:flex; flex-direction:column; align-items:center; width:100%;">
      
      <div class="team-branding" id="branding_div">
          <p>Equipo Planificación<br>WAYS 2026</p>
      </div>

      <button class="btn-export" onclick="descargarExcel()">⬇️ Exportar Data a Excel</button>
      <div style="text-align:center;font-size:0.7rem;color:var(--text-muted);margin-top:15px;font-weight:600;">
        Actualizado: __FECHA_ACTUAL__
      </div>
      
    </div>
  </div>

  <div class="tab-panel active" id="tab_resumen">
    <div class="resumen-panel" id="resumen_content"></div>
  </div>

  <div class="tab-panel" id="tab_analisis">
  <div class="content">
    <div class="kpi-row">
      <div class="kpi-card c-red"><span>Equipos con Fallas</span><h3 id="k_equipos">0</h3></div>
      <div class="kpi-card c-red"><span>Detenciones MTTO</span><h3 id="k_detenciones">0</h3></div>
      <div class="kpi-card c-red"><span>Tpo. Perdido Total (Hrs)</span><h3 id="k_hrs">0.0</h3></div>
      <div class="kpi-card c-green"><span>Tiempo Planificado</span><h3 id="k_plan">0.0</h3></div>
      <div class="kpi-card c-green"><span>Tiempo de Operación</span><h3 id="k_oper">0.0</h3></div>
      <div class="kpi-card c-warn"><span>Mantenibilidad Global (M)</span><h3 id="k_mant">0%</h3></div>
    </div>

    <div class="charts-row">
      <div class="chart-container">
        <div class="chart-header">🔪 Jackknife: Durabilidad (Hrs) vs Repetición de Falla</div>
        <div class="canvas-wrapper"><canvas id="chart_jackknife"></canvas></div>
      </div>
      <div class="chart-container">
        <div class="chart-header">⏱️ Evolución Tiempos Medios (MTBF vs MTTR)</div>
        <div class="canvas-wrapper"><canvas id="chart_trend_mtbf"></canvas></div>
      </div>
    </div>

    <div class="table-container">
      <div class="table-header">
        <span>📋 Matriz Acumulada de KPIs (Haz clic en cualquier fila para ver el desglose)</span>
        <input type="text" id="search_input" placeholder="🔍 Buscar equipo o línea..." onkeyup="renderTable()">
      </div>
      <div class="table-wrapper">
        <table id="data_table">
          <thead>
            <tr>
              <th onclick="sortTable(0)">Planta ↕</th>
              <th onclick="sortTable(1)">Línea ↕</th>
              <th onclick="sortTable(2)">Equipo ↕</th>
              <th onclick="sortTable(3)">Detenciones ↕</th>
              <th onclick="sortTable(4)">Tpo Perdido (Hrs) ↕</th>
              <th onclick="sortTable(5)">MTBF ↕</th>
              <th onclick="sortTable(6)">MTTR ↕</th>
              <th onclick="sortTable(7)">Confiabilidad ↕</th>
              <th onclick="sortTable(8)">Mantenibilidad ↕</th>
              <th onclick="sortTable(9)">Prob. Falla ↕</th>
            </tr>
          </thead>
          <tbody id="table_body"></tbody>
        </table>
      </div>
    </div>
  </div>
  </div>
</div>

<div id="modal_overlay" class="modal-overlay" onclick="cerrarModalExterno(event)">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h3 id="modal_titulo">Historial de Detenciones</h3>
        <p id="modal_subtitulo">Planta | Línea</p>
      </div>
      <button class="close-btn" onclick="cerrarModal()">&times;</button>
    </div>
    <div class="modal-body">
      <table id="modal_table">
        <thead>
          <tr>
            <th>Fecha</th>
            <th>Componente Afectado</th>
            <th>Tipo de Falla</th>
            <th>Tiempo Perdido (Hrs)</th>
          </tr>
        </thead>
        <tbody id="modal_table_body"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
Chart.defaults.font.family = "'DM Sans', system-ui, sans-serif";
Chart.defaults.color = '#8899AA';

const dbRaw       = __DB_JSON_DATA__;
const recordsEq   = dbRaw.equipos;
const recordsLn   = dbRaw.lineas;

let isCarnesTheme  = false;
let currentEqData  = [];
let currentLnData  = [];
let tableDataFull  = [];
let chartInstances = {};

window.rotatorIntervals = [];
window.rotatorData = {};
window.sparklineTasks = [];

function toggleTheme() {
  isCarnesTheme = document.getElementById('theme_toggle').checked;
  if (isCarnesTheme) {
    document.body.classList.add('theme-carnes');
    document.getElementById('lbl_carnes').classList.add('active');
    document.getElementById('lbl_masas').classList.remove('active');
  } else {
    document.body.classList.remove('theme-carnes');
    document.getElementById('lbl_masas').classList.add('active');
    document.getElementById('lbl_carnes').classList.remove('active');
  }
  buildFilters();
  applyFilters();
}

function buildFilters() {
  let baseLn = recordsLn.filter(d => isCarnesTheme ? d.super_planta === 'Carnes' : d.super_planta === 'Masas');
  let semanasUnicas = [...new Set(baseLn.map(x => x.semana))].sort((a, b) => a - b);
  let htmlDesde = '', htmlHasta = '';
  semanasUnicas.forEach((s, idx) => {
    htmlDesde += `<option value="${s}" ${idx === 0 ? 'selected' : ''}>Sem ${s}</option>`;
    htmlHasta += `<option value="${s}" ${idx === semanasUnicas.length - 1 ? 'selected' : ''}>Sem ${s}</option>`;
  });
  document.getElementById('f_sem_desde').innerHTML = htmlDesde;
  document.getElementById('f_sem_hasta').innerHTML = htmlHasta;

  let plantas = [...new Set(baseLn.map(x => x.planta))].sort();
  let lineas = [...new Set(baseLn.map(x => x.linea))].sort();

  let htmlDyn = `
    <div class="f-group">
      <label>🏭 Planta</label>
      <select id="f_planta" onchange="actualizarFiltroLinea()">
        <option value="ALL">Todas</option>
        ${plantas.map(p => `<option value="${p}">${p}</option>`).join('')}
      </select>
    </div>
    <div class="f-group">
      <label>🔧 Línea</label>
      <select id="f_linea" onchange="applyFilters()">
        <option value="ALL">Todas</option>
        ${lineas.map(l => `<option value="${l}">${l}</option>`).join('')}
      </select>
    </div>
  `;
  document.getElementById('filters_dynamic').innerHTML = htmlDyn;
}

function actualizarFiltroLinea() {
  let baseLn = recordsLn.filter(d => isCarnesTheme ? d.super_planta === 'Carnes' : d.super_planta === 'Masas');
  const plantaSeleccionada = document.getElementById('f_planta').value;

  if (plantaSeleccionada !== 'ALL') {
    baseLn = baseLn.filter(d => d.planta === plantaSeleccionada);
  }

  let lineas = [...new Set(baseLn.map(x => x.linea))].sort();
  let comboLinea = document.getElementById('f_linea');
  let seleccionActual = comboLinea.value;

  let html = '<option value="ALL">Todas</option>';
  lineas.forEach(l => html += `<option value="${l}">${l}</option>`);
  comboLinea.innerHTML = html;

  if (lineas.includes(seleccionActual)) {
    comboLinea.value = seleccionActual;
  } else {
    comboLinea.value = 'ALL';
  }

  applyFilters();
}

function applyFilters() {
  const sDesde = parseInt(document.getElementById('f_sem_desde').value);
  const sHasta = parseInt(document.getElementById('f_sem_hasta').value);
  const fPla   = document.getElementById('f_planta').value;
  const fLin   = document.getElementById('f_linea').value;

  const baseFilter = d => {
    if (isCarnesTheme ? d.super_planta !== 'Carnes' : d.super_planta !== 'Masas') return false;
    if (d.semana < sDesde || d.semana > sHasta) return false;
    if (fPla !== 'ALL' && d.planta !== fPla) return false;
    if (fLin !== 'ALL' && d.linea  !== fLin) return false;
    return true;
  };

  currentLnData = recordsLn.filter(baseFilter);
  currentEqData = recordsEq.filter(baseFilter);

  let opTimeByLine = {};
  let totalPlanificadoGlobal  = 0;

  currentLnData.forEach(d => {
    let k = d.planta + "|" + d.linea;
    if (!opTimeByLine[k]) opTimeByLine[k] = { op: 0, pl: 0 };
    opTimeByLine[k].op      += d.tpo_operativo_linea;
    opTimeByLine[k].pl      += d.tpo_plan_linea;
    totalPlanificadoGlobal  += d.tpo_plan_linea;
  });

  let eqMap = {};
  currentEqData.forEach(d => {
    let eqKey = d.planta + "|" + d.linea + "|" + d.equipo;
    if (!eqMap[eqKey]) eqMap[eqKey] = { p: d.planta, l: d.linea, e: d.equipo, det: 0, tpop: 0, eventos: [] };
    eqMap[eqKey].det  += d.detenciones;
    eqMap[eqKey].tpop += d.tpo_perdido_eq;
    eqMap[eqKey].eventos.push({ fecha: d.fecha, componente: d.componente, tipo: d.tipo, hrs: d.tpo_perdido_eq, mes: d.mes });
  });

  let lostByLine = {};
  currentEqData.forEach(d => {
    let k = d.planta + "|" + d.linea;
    if (!lostByLine[k]) lostByLine[k] = 0;
    lostByLine[k] += d.tpo_perdido_eq;
  });

  tableDataFull = Object.values(eqMap).map(d => {
    let planTime = opTimeByLine[d.p + "|" + d.l] ? opTimeByLine[d.p + "|" + d.l].pl : 0;
    let opTime   = Math.max(0, planTime - (lostByLine[d.p + "|" + d.l] || 0));
    let mtbf     = d.det > 0 ? (opTime / d.det) : 0;
    let mttr   = d.det > 0 ? (d.tpop / d.det) : 0;
    let conf   = mtbf > 0 ? Math.exp(-120 / mtbf) * 100 : (d.det === 0 ? 100 : 0);
    let mant   = mttr > 0 ? (1 - Math.exp(-1 / mttr)) * 100 : 100;
    let prob   = 100 - conf;
    return { ...d, opTime, mtbf, mttr, conf, mant, prob };
  });

  let eqCount           = tableDataFull.length;
  let sumPerdidoGlobal  = tableDataFull.reduce((s, d) => s + d.tpop, 0);
  let sumFallasGlobal   = tableDataFull.reduce((s, d) => s + d.det, 0);
  let mttrGlobal        = sumFallasGlobal > 0 ? (sumPerdidoGlobal / sumFallasGlobal) : 0;
  let mantGlobal        = mttrGlobal > 0 ? (1 - Math.exp(-1 / mttrGlobal)) * 100 : 100;

  document.getElementById('k_equipos').innerText    = eqCount;
  document.getElementById('k_detenciones').innerText = sumFallasGlobal;
  document.getElementById('k_hrs').innerText         = sumPerdidoGlobal.toFixed(1);
  document.getElementById('k_plan').innerText        = totalPlanificadoGlobal.toFixed(1);
  document.getElementById('k_oper').innerText        = Math.max(0, totalPlanificadoGlobal - sumPerdidoGlobal).toFixed(1);
  document.getElementById('k_mant').innerText        = mantGlobal.toFixed(1) + "%";

  drawCharts(sDesde, sHasta);
  renderTable();
  renderResumen();
}

function drawCharts(sDesde, sHasta) {
  if (currentLnData.length === 0) return;

  const secColor    = isCarnesTheme ? '#991b1b' : '#3A4A5C';
  let weeks        = [];
  let mtbfTrend    = [];
  let mttrTrend    = [];

  for (let i = sDesde; i <= sHasta; i++) weeks.push(i);

  weeks.forEach(w => {
    let dLn   = currentLnData.filter(d => d.semana === w);
    let dEq   = currentEqData.filter(d => d.semana === w);
    let sPlan = dLn.reduce((s, d) => s + d.tpo_plan_linea, 0);
    let sPerd = dEq.reduce((s, d) => s + d.tpo_perdido_eq, 0);
    let sDet  = dEq.reduce((s, d) => s + d.detenciones, 0);
    let sOper = Math.max(0, sPlan - sPerd);
    mtbfTrend.push(sDet > 0 ? (sOper / sDet).toFixed(2) : (sOper > 0 ? sOper : 0));
    mttrTrend.push(sDet > 0 ? (sPerd / sDet).toFixed(2) : 0);
  });

  if (chartInstances['jackknife']) chartInstances['jackknife'].destroy();
  chartInstances['jackknife'] = new Chart(document.getElementById('chart_jackknife'), {
    type: 'scatter',
    data: {
      datasets: [{
        label: 'Equipos',
        data: tableDataFull.filter(d => d.det > 0).map(d => ({ x: d.det, y: d.tpop, e: d.e, l: d.l })),
        backgroundColor: isCarnesTheme ? '#C0392B' : '#0071CE',
        borderColor:     isCarnesTheme ? '#7f1d1d' : '#0D2C54',
        borderWidth: 1, pointRadius: 6, pointHoverRadius: 8,
      }]
    },
    options: {
      maintainAspectRatio: false,
      plugins: {
        tooltip: { callbacks: { label: ctx => `${ctx.raw.e} (${ctx.raw.l}): ${ctx.raw.x} Detenciones, ${ctx.raw.y.toFixed(1)} Hrs` } },
        legend: { display: false },
      },
      scales: {
        x: { type: 'linear', position: 'bottom', title: { display: true, text: 'Repetición (N° Detenciones)', font: { weight: 'bold' } } },
        y: { type: 'linear', position: 'left',   title: { display: true, text: 'Durabilidad (Tiempo Perdido Hrs)', font: { weight: 'bold' } } },
      },
    },
  });

  if (chartInstances['trend_mtbf']) chartInstances['trend_mtbf'].destroy();
  chartInstances['trend_mtbf'] = new Chart(document.getElementById('chart_trend_mtbf'), {
    type: 'line',
    data: {
      labels: weeks.map(w => 'Semana ' + w),
      datasets: [
        { label: 'MTBF (Hrs)', data: mtbfTrend, borderColor: secColor,    borderWidth: 3, tension: 0.3, yAxisID: 'y' },
        { label: 'MTTR (Hrs)', data: mttrTrend, borderColor: '#E67E22',   borderWidth: 3, borderDash: [5, 5], tension: 0.3, yAxisID: 'y1' },
      ],
    },
    options: {
      maintainAspectRatio: false,
      scales: {
        y:  { type: 'linear', position: 'left',  title: { display: true, text: 'MTBF (h)' } },
        y1: { type: 'linear', position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'MTTR (h)' } },
      },
    },
  });
}

function renderTable() {
  const search = document.getElementById('search_input').value.toLowerCase();
  const tbody  = document.getElementById('table_body');
  tbody.innerHTML = '';

  let tbl = [...tableDataFull];
  if (search) tbl = tbl.filter(d => `${d.p} ${d.l} ${d.e}`.toLowerCase().includes(search));
  tbl.sort((a, b) => a.conf - b.conf);

  tbl.forEach(d => {
    let badgeClass = d.conf >= 80 ? 'b-ok' : (d.conf >= 50 ? 'b-warn' : 'b-danger');
    let tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.setAttribute('onclick', `abrirModalEventos('${d.p}','${d.l}','${d.e}')`);
    tr.innerHTML = `
      <td>${d.p}</td>
      <td>${d.l}</td>
      <td style="font-weight:700;color:var(--text);">${d.e}</td>
      <td>${d.det}</td>
      <td style="font-weight:600;">${d.tpop.toFixed(2)}</td>
      <td>${d.mtbf.toFixed(1)}</td>
      <td>${d.mttr.toFixed(2)}</td>
      <td><span class="badge ${badgeClass}">${d.conf.toFixed(1)}%</span></td>
      <td>${d.mant.toFixed(1)}%</td>
      <td>${d.prob.toFixed(1)}%</td>
    `;
    tbody.appendChild(tr);
  });
}

function abrirModalEventos(planta, linea, equipo) {
  const eqData = tableDataFull.find(x => x.p === planta && x.l === linea && x.e === equipo);
  if (!eqData) return;
  document.getElementById('modal_titulo').innerText    = `Historial de Detenciones: ${eqData.e}`;
  document.getElementById('modal_subtitulo').innerText = `Planta: ${eqData.p} | Línea de Proceso: ${eqData.l}`;

  const mBody = document.getElementById('modal_table_body');
  mBody.innerHTML = '';
  [...eqData.eventos].sort((a, b) => b.hrs - a.hrs).forEach(ev => {
    let tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:600;color:var(--secondary);">${ev.fecha}</td>
      <td>${ev.componente}</td>
      <td><span style="font-size:0.8rem;font-weight:500;">${ev.tipo}</span></td>
      <td style="font-weight:700;color:var(--danger);">${ev.hrs.toFixed(2)} hrs</td>
    `;
    mBody.appendChild(tr);
  });
  document.getElementById('modal_overlay').style.display = 'flex';
}

function cerrarModal()              { document.getElementById('modal_overlay').style.display = 'none'; }
function cerrarModalExterno(e)      { if (e.target.id === 'modal_overlay') cerrarModal(); }

let sortAsc = true, lastCol = -1;
function sortTable(colIdx) {
  const tbody = document.getElementById('data_table').querySelector('tbody');
  const rows  = Array.from(tbody.querySelectorAll('tr'));
  sortAsc = (lastCol === colIdx) ? !sortAsc : true;
  lastCol = colIdx;
  rows.sort((a, b) => {
    let vA = a.cells[colIdx].innerText.replace('%', '').trim();
    let vB = b.cells[colIdx].innerText.replace('%', '').trim();
    let nA = parseFloat(vA), nB = parseFloat(vB);
    if (!isNaN(nA) && !isNaN(nB)) return sortAsc ? nA - nB : nB - nA;
    return sortAsc ? vA.localeCompare(vB) : vB.localeCompare(vA);
  });
  tbody.innerHTML = '';
  rows.forEach(r => tbody.appendChild(r));
}

function descargarExcel() {
  if (tableDataFull.length === 0) return alert("No hay datos para exportar");
  const exportData = tableDataFull.map(d => ({
    "Planta":              d.p,
    "Línea":               d.l,
    "Equipo":              d.e,
    "Cant. Detenciones":   d.det,
    "Tpo Perdido (Hrs)":   d.tpop,
    "MTBF":                d.mtbf,
    "MTTR":                d.mttr,
    "Confiabilidad (%)":   d.conf,
    "Mantenibilidad (%)":  d.mant,
    "Prob Falla (%)":      d.prob,
  }));
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Base_Acumulada");
  XLSX.writeFile(wb, "Matriz_Confiabilidad_Acumulada.xlsx");
}

window.onload = () => {
    toggleTheme();
};

function switchTab(tab) {
  ['resumen','analisis'].forEach(t => {
    document.getElementById('tab_' + t).classList.toggle('active', t === tab);
    document.getElementById('tbtn_' + t).classList.toggle('active', t === tab);
  });
}

function drawSparkline(canvasId, dataPoints, labels) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    
    if (!dataPoints || dataPoints.length === 0) return;
    const maxVal = Math.max(...dataPoints, 1);
    const padX = 12;
    const padY = 12;
    const stepX = dataPoints.length > 1 ? (w - padX * 2) / (dataPoints.length - 1) : 0;
    
    // Dibujar línea de conexión
    ctx.beginPath();
    ctx.strokeStyle = '#BDC3C7';
    ctx.lineWidth = 1.5;
    dataPoints.forEach((val, i) => {
        const x = dataPoints.length > 1 ? padX + i * stepX : w / 2;
        const y = h - padY - (val / maxVal) * (h - padY * 2);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
    });
    if(dataPoints.length > 1) ctx.stroke();
    
    // Dibujar puntos y etiquetas
    dataPoints.forEach((val, i) => {
        const x = dataPoints.length > 1 ? padX + i * stepX : w / 2;
        const y = h - padY - (val / maxVal) * (h - padY * 2);
        
        ctx.beginPath();
        ctx.arc(x, y, 3.5, 0, 2 * Math.PI);
        ctx.fillStyle = val > 0 ? '#E74C3C' : '#3498DB';
        ctx.fill();
        
        // Etiqueta de valor encima del punto
        if (val > 0) {
            ctx.font = 'bold 8px "DM Sans"';
            ctx.fillStyle = '#C0392B';
            ctx.textAlign = 'center';
            ctx.fillText(val.toFixed(1), x, y - 6);
        }
        
        // Etiqueta del mes debajo
        ctx.font = '8px "DM Sans"';
        ctx.fillStyle = '#7F8C8D';
        ctx.textAlign = 'center';
        ctx.fillText(labels[i], x, h - 2);
    });
}

function renderResumen() {
  if (window.rotatorIntervals) {
      window.rotatorIntervals.forEach(clearInterval);
  }
  window.rotatorIntervals = [];
  window.rotatorData = {};
  window.sparklineTasks = [];

  const el = document.getElementById('resumen_content');
  if (!currentLnData.length) { el.innerHTML = '<p style="padding:30px;color:var(--text-muted);">Sin datos para el rango seleccionado.</p>'; return; }

  const sDesde = parseInt(document.getElementById('f_sem_desde').value);
  const sHasta = parseInt(document.getElementById('f_sem_hasta').value);
  const sHastaPrev = sHasta > sDesde ? sHasta - 1 : sDesde;

  const fPla = document.getElementById('f_planta').value;
  const fLin = document.getElementById('f_linea').value;

  const prevFilter = d => {
    if (isCarnesTheme ? d.super_planta !== 'Carnes' : d.super_planta !== 'Masas') return false;
    if (d.semana < sDesde || d.semana > sHastaPrev) return false;
    if (fPla !== 'ALL' && d.planta !== fPla) return false;
    if (fLin !== 'ALL' && d.linea  !== fLin) return false;
    return true;
  };

  const prevLnData = recordsLn.filter(prevFilter);
  const prevEqData = recordsEq.filter(prevFilter);

  const lnByPlanta = {};
  currentLnData.forEach(d => {
    if (!lnByPlanta[d.planta]) lnByPlanta[d.planta] = {};
    if (!lnByPlanta[d.planta][d.linea]) lnByPlanta[d.planta][d.linea] = {op:0, pl:0};
    lnByPlanta[d.planta][d.linea].op += d.tpo_operativo_linea;
    lnByPlanta[d.planta][d.linea].pl += d.tpo_plan_linea;
  });

  const eqByPlanta = {};
  currentEqData.forEach(d => {
    if (!eqByPlanta[d.planta]) eqByPlanta[d.planta] = {};
    const k = d.linea + '|||' + d.equipo;
    if (!eqByPlanta[d.planta][k]) eqByPlanta[d.planta][k] = {linea:d.linea, equipo:d.equipo, det:0, hrs:0};
    eqByPlanta[d.planta][k].det += d.detenciones;
    eqByPlanta[d.planta][k].hrs += d.tpo_perdido_eq;
  });

  const lnByPlantaPrev = {};
  prevLnData.forEach(d => {
    if (!lnByPlantaPrev[d.planta]) lnByPlantaPrev[d.planta] = {};
    if (!lnByPlantaPrev[d.planta][d.linea]) lnByPlantaPrev[d.planta][d.linea] = {op:0, pl:0};
    lnByPlantaPrev[d.planta][d.linea].op += d.tpo_operativo_linea;
    lnByPlantaPrev[d.planta][d.linea].pl += d.tpo_plan_linea;
  });

  const eqByPlantaPrev = {};
  prevEqData.forEach(d => {
    if (!eqByPlantaPrev[d.planta]) eqByPlantaPrev[d.planta] = {};
    const k = d.linea + '|||' + d.equipo;
    if (!eqByPlantaPrev[d.planta][k]) eqByPlantaPrev[d.planta][k] = {linea:d.linea, equipo:d.equipo, det:0, hrs:0};
    eqByPlantaPrev[d.planta][k].det += d.detenciones;
    eqByPlantaPrev[d.planta][k].hrs += d.tpo_perdido_eq;
  });

  // Identificar los meses que tienen actividad en este rango de datos
  const monthOrder = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  let activeMonths = new Set();
  currentEqData.forEach(d => { if (d.mes && d.mes !== 'N/A') activeMonths.add(d.mes); });
  let displayMonths = monthOrder.filter(m => activeMonths.has(m));
  if (displayMonths.length === 0) displayMonths = ['Ene']; // fallback

  let html = '';
  Object.keys(lnByPlanta).sort().forEach(planta => {
    const pKeyLowerCase = planta.toLowerCase();
    
    let tema = 'tema-masas';
    let temaPillColor = '#0071CE';
    
    const esCarnesTema = pKeyLowerCase.includes('carne') || pKeyLowerCase.includes('mercadeo') || pKeyLowerCase.includes('molida');
    const esMasasTema = !esCarnesTema;
    const limiteTop = esCarnesTema ? 6 : 3;
    
    if (esCarnesTema) { 
        tema = 'tema-carnes'; 
        temaPillColor = '#C0392B'; 
    } 
    if (pKeyLowerCase.includes('dely')) { 
        tema = 'tema-dely'; 
        temaPillColor = '#00897B'; 
    } else if (pKeyLowerCase.includes('molida')) { 
        tema = 'tema-molida'; 
        temaPillColor = '#D35400'; 
    }

    const lineas = lnByPlanta[planta];
    const eqs    = Object.values(eqByPlanta[planta] || {});

    const lineaStats = Object.entries(lineas).map(([ln, d]) => {
      const lnDet  = eqs.filter(e => e.linea === ln).reduce((s,e) => s + e.det, 0);
      const mtbf   = lnDet > 0 ? d.op / lnDet : 0;
      const prob   = mtbf > 0 ? (1 - Math.exp(-120 / mtbf)) * 100 : (lnDet > 0 ? 100 : 0);
      const perdido = eqs.filter(e => e.linea === ln).reduce((s,e) => s + e.hrs, 0);
      
      const mttr = lnDet > 0 ? perdido / lnDet : 0;
      const conf = mtbf > 0 ? Math.exp(-120 / mtbf) * 100 : (lnDet === 0 ? 100 : 0);
      const mant = mttr > 0 ? (1 - Math.exp(-1 / mttr)) * 100 : 100;

      return { ln, op: d.op, pl: d.pl, perdido, prob, mtbf, mttr, conf, mant };
    }).sort((a,b) => b.prob - a.prob);

    const pId = planta.replace(/\s+/g,'_');
    window.rotatorData[pId] = lineaStats.map(s => {
        let lnShort = s.ln.trim().substring(0,8);
        return `<span style="color:#A9CCE3; font-weight:900; font-size:11px; margin-right:8px; border-right:1px solid rgba(255,255,255,0.2); padding-right:10px; display:inline-block;">${lnShort}</span>` +
               `<span>MTBF: <strong>${s.mtbf.toFixed(1)}h</strong></span>` +
               `<span>MTTR: <strong>${s.mttr.toFixed(2)}h</strong></span>` +
               `<span>Conf: <strong>${s.conf.toFixed(1)}%</strong></span>` +
               `<span>Mant: <strong>${s.mant.toFixed(1)}%</strong></span>` +
               `<span>Pb. Falla: <strong>${s.prob.toFixed(1)}%</strong></span>`;
    });

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
    const maxProb = Math.max(...lineaStats.map(x => x.prob), 1);

    const topEquiposArray = eqs.map(e => {
      const lnOp  = lineas[e.linea]?.op || 0;
      const mtbf  = e.det > 0 ? lnOp / e.det : 0;
      const mttr  = e.det > 0 ? e.hrs / e.det : 0;
      const prob  = mtbf > 0 ? (1 - Math.exp(-120/mtbf))*100 : (e.det>0?100:0);

      let detPrev = 0, hrsPrev = 0, opPrev = 0;
      if (eqByPlantaPrev[planta] && eqByPlantaPrev[planta][e.linea + '|||' + e.equipo]) {
          detPrev = eqByPlantaPrev[planta][e.linea + '|||' + e.equipo].det;
      }
      if (lnByPlantaPrev[planta] && lnByPlantaPrev[planta][e.linea]) {
          opPrev = lnByPlantaPrev[planta][e.linea].op;
      }
      const mtbfPrev = detPrev > 0 ? opPrev / detPrev : 0;
      const probPrev = mtbfPrev > 0 ? (1 - Math.exp(-120/mtbfPrev))*100 : (detPrev>0?100:0);

      // Reconstruir la sumatoria del equipo para el sparkline (horas perdidas por mes)
      const evs = currentEqData.filter(d => d.planta === planta && d.linea === e.linea && d.equipo === e.equipo);
      let tpoMes = {};
      evs.forEach(ev => {
          if (!tpoMes[ev.mes]) tpoMes[ev.mes] = 0;
          tpoMes[ev.mes] += ev.tpo_perdido_eq;
      });

      return { ...e, prob, mtbf, mttr, probPrev, mtbfPrev, tpoMes };
    }).filter(e => {
        if (e.det <= 0) return false;
        let eqStr = String(e.equipo).trim().toLowerCase();
        if (esMasasTema && eqStr === '0') return false;
        if (esCarnesTema && (eqStr === '' || eqStr === 'nan' || eqStr === 'null' || eqStr === 'n/a')) return false;
        return true;
    }).sort((a,b)=>b.prob-a.prob).slice(0, limiteTop);

    const barras = lineaStats.map((s,i) => {
      const pct = (s.prob / Math.max(maxProb, 1) * 100).toFixed(1);
      const color = s.prob > 60 ? '#C0392B' : s.prob > 35 ? '#E67E22' : '#27AE60';
      const plTotal = Math.max(s.pl, s.op + s.perdido);
      const wOp = plTotal > 0 ? (s.op / plTotal * 100) : 0;
      const wPerdido = plTotal > 0 ? (s.perdido / plTotal * 100) : 0;
      
      return `
        <div class="line-item">
          <div style="display:flex;align-items:center;gap:6px;">
            <div style="width:40px;font-size:10px;font-weight:800;color:#fff;flex-shrink:0;opacity:.9;">${s.ln}</div>
            <div style="flex:1;background:rgba(255,255,255,0.15);border-radius:3px;height:12px;overflow:hidden;">
              <div style="width:${pct}%;background:${color};height:100%;border-radius:3px;opacity:0.9;"></div>
            </div>
            <div style="width:40px;text-align:right;font-size:10px;font-weight:800;color:${color};flex-shrink:0;">${s.prob.toFixed(1).replace('.',',')}%</div>
          </div>
          <div style="margin-top:4px;">
            <div style="position:relative;height:6px;background:rgba(255,255,255,0.15);border-radius:2px;overflow:hidden;width:100%;max-width:100%;margin-bottom:4px;">
              <div style="position:absolute;left:0;top:0;height:100%;width:${wOp}%;background:rgba(100,180,255,0.8);border-radius:2px;opacity:0.85;"></div>
              <div style="position:absolute;right:0;top:0;height:100%;width:${wPerdido}%;background:rgba(255,80,60,0.85);border-radius:2px;opacity:0.8;"></div>
            </div>
            <div style="display:flex;gap:5px;align-items:center;">
              <span style="width:6px;height:6px;background:rgba(100,180,255,0.8);border-radius:1px;display:inline-block;opacity:.85;flex-shrink:0;"></span>
              <span style="font-size:8px;color:rgba(255,255,255,0.8);">${s.op.toFixed(1)}h</span>
              <span style="font-size:8px;color:rgba(255,255,255,0.3);">/</span>
              <span style="width:6px;height:6px;background:rgba(255,255,255,0.2);border-radius:1px;display:inline-block;flex-shrink:0;"></span>
              <span style="font-size:8px;color:rgba(255,255,255,0.8);">${s.pl.toFixed(1)}h</span>
              <span style="font-size:8px;color:rgba(255,255,255,0.3);">·</span>
              <span style="width:6px;height:6px;background:rgba(255,80,60,0.85);border-radius:1px;display:inline-block;opacity:.8;flex-shrink:0;"></span>
              <span style="font-size:8px;font-weight:700;color:rgba(255,140,120,1);">${s.perdido.toFixed(1)}h</span>
            </div>
          </div>
        </div>`;
    }).join('');

    const rankColors = ['#CC2222','#E65100','#F9A825','#F1C40F','#3498DB','#9B59B6'];
    const filas = topEquiposArray.map((e, i) => {
      const probColor = e.prob > 60 ? '#C0392B' : e.prob > 35 ? '#E67E22' : '#27AE60';
      const rankTop = i + 1;
      
      let accionStr = "-";
      if (dbRaw.acciones) {
        for (let col in dbRaw.acciones) {
            if (pKeyLowerCase.includes(col) || col.includes(pKeyLowerCase)) {
                accionStr = dbRaw.acciones[col][rankTop] || "-";
                break;
            }
        }
      }

      let probDif = e.prob - e.probPrev;
      let probArrow = "";
      if (e.probPrev > 0 || e.mtbfPrev > 0) {
          if (probDif > 0.1) probArrow = `<span class="stat-trend trend-up">🔺 +${probDif.toFixed(1)}pp</span>`;
          else if (probDif < -0.1) probArrow = `<span class="stat-trend trend-dn">🔻 ${probDif.toFixed(1)}pp</span>`;
          else probArrow = `<span class="stat-trend trend-neu">▬ 0.0pp</span>`;
      } else {
          probArrow = `<span class="stat-trend trend-neu">N/A Ant.</span>`;
      }

      let mtbfDif = e.mtbf - e.mtbfPrev;
      let mtbfArrow = "";
      if (e.probPrev > 0 || e.mtbfPrev > 0) {
          if (mtbfDif > 0.1) mtbfArrow = `<span class="stat-trend trend-dn">🔺 +${mtbfDif.toFixed(1)}h</span>`;
          else if (mtbfDif < -0.1) mtbfArrow = `<span class="stat-trend trend-up">🔻 ${mtbfDif.toFixed(1)}h</span>`;
          else mtbfArrow = `<span class="stat-trend trend-neu">▬ 0.0h</span>`;
      } else {
          mtbfArrow = `<span class="stat-trend trend-neu">N/A Ant.</span>`;
      }

      // Prepara datos del gráfico de puntos (Sparkline)
      let dataPoints = displayMonths.map(m => e.tpoMes[m] || 0);
      let canvasId = `spark_${pId}_${i}`;
      window.sparklineTasks.push({ id: canvasId, data: dataPoints, labels: displayMonths });

      return `
        <div class="grid-tr">
          <div class="rank-circle" style="background:${rankColors[i]}">${rankTop}</div>
          <div class="eq-name">
            ${e.equipo}<span class="line-badge">${e.linea}</span>
          </div>
          
          <div class="stat-box" style="background:#fff; border: 1px dashed var(--border);">
             <canvas id="${canvasId}" width="120" height="40" style="display:block; margin:auto;"></canvas>
          </div>

          <div class="stat-box bg-red-light">
            <span class="stat-val" style="color:${probColor}">${e.prob.toFixed(1)}%</span>
            ${probArrow}
          </div>
          <div class="stat-box bg-gray-light">
            <span class="stat-val" style="color:#7F8C8D">${e.probPrev.toFixed(1)}%</span>
          </div>

          <div class="stat-box bg-blue-light">
            <span class="stat-val" style="color:#1A5276">${e.mtbf.toFixed(1)}h</span>
            ${mtbfArrow}
          </div>
          <div class="stat-box bg-gray-light">
            <span class="stat-val" style="color:#7F8C8D">${e.mtbfPrev.toFixed(1)}h</span>
          </div>

          <div style="padding-left:10px; border-left:2px solid #EAF0F8;">
            <div style="font-size:10px; color:#2C3E50; line-height:1.4;">${accionStr}</div>
          </div>
        </div>`;
    }).join('');

    html += `
      <div class="fila-planta-v3">
        <div class="top-section">
          <div class="plant-header">
            <span class="plant-badge" style="background:${temaPillColor}">${planta.toUpperCase()}</span>
            <h3>Top ${limiteTop} Equipos Críticos</h3>
          </div>
          <div class="grid-table">
            <div class="grid-th">
              <div>#</div>
              <div>EQUIPO</div>
              <div class="center">TPO PERDIDO / MES<small>Evolución</small></div>
              <div class="center">PB. FALLA<small>Sem ${sDesde} a ${sHasta}</small></div>
              <div class="center" style="opacity:0.6;">PB. FALLA (Ant.)<small>Sem ${sDesde} a ${sHastaPrev}</small></div>
              <div class="center">MTBF<small>Sem ${sDesde} a ${sHasta}</small></div>
              <div class="center" style="opacity:0.6;">MTBF (Ant.)<small>Sem ${sDesde} a ${sHastaPrev}</small></div>
              <div>ACCIONES CORRECTIVAS</div>
            </div>
            ${topEquiposArray.length ? filas : '<p style="color:var(--text-muted);font-size:0.9rem;padding:15px 0;">Sin detenciones críticas en el período.</p>'}
          </div>
        </div>

        <div class="bottom-section ${tema}">
          <div class="indicators-header">
              <div class="zona-block">
                  <span class="section-label">Indicadores zona 2026</span>
                  <div class="kpis-planta">
                      <span>MTBF: <strong>${promMTBF.toFixed(1)}h</strong></span>
                      <span>MTTR: <strong>${promMTTR.toFixed(2)}h</strong></span>
                      <span>Conf: <strong>${promConf.toFixed(1)}%</strong></span>
                      <span>Mant: <strong>${promMant.toFixed(1)}%</strong></span>
                      <span>Pb. Falla: <strong>${promProb.toFixed(1)}%</strong></span>
                  </div>
              </div>
              <div class="line-rotator-block">
                  <span class="section-label">Indicadores por Línea (Rotativo)</span>
                  <div class="kpis-rotator" id="rotator_${pId}"></div>
              </div>
          </div>
          <span class="section-label">ESTADO POR LÍNEA DE PRODUCCIÓN</span>
          <div class="lines-grid">
            ${barras}
          </div>
        </div>
      </div>`;
  });

  el.innerHTML = html;

  // Renderizar gráficos de puntos sparkline
  window.sparklineTasks.forEach(task => drawSparkline(task.id, task.data, task.labels));

  Object.keys(window.rotatorData).forEach(pid => {
      const container = document.getElementById('rotator_' + pid);
      const data = window.rotatorData[pid];
      if (container && data.length > 0) {
          let i = 0;
          container.innerHTML = data[i];
          if (data.length > 1) {
              const intv = setInterval(() => {
                  container.style.opacity = '0';
                  setTimeout(() => {
                      i = (i + 1) % data.length;
                      container.innerHTML = data[i];
                      container.style.opacity = '1';
                  }, 300);
              }, 5000);
              window.rotatorIntervals.push(intv);
          }
      } else if (container) {
          container.innerHTML = "<span>Sin datos de líneas</span>";
      }
  });
}
</script>
</body>
</html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    full_html = full_html.replace("__FECHA_ACTUAL__", fecha_actual)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(full_html)

    print(f"✅ Dashboard generado de forma exitosa: {OUTPUT_HTML}")

if __name__ == "__main__":
    db = procesar_datos_confiabilidad()
    if db:
        generar_html_moderno(db)
