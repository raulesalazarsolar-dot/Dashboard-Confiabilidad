import os
import json
import math
import shutil
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
# 2. BUSCADORES UNIVERSALES E INTELIGENTES
# ==========================================
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

def buscar_tiempo_detencion_hr(df):
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

def buscar_columna_tiempo_plan(df):
    cols_lower = [str(c).lower().replace(' ', ' ').strip() for c in df.columns]
    for i, c in enumerate(cols_lower):
        if (' plan' in c or 'disponible' in c) and 'hr' in c:
            return df.columns[i]
    return None

# ==========================================
# 3. FILTROS DE NEGOCIO (CENTRALIZADOS)
# ==========================================
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

    # --- LECTURA DEL EXCEL DE ACCIONES CORRECTIVAS ---
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

    # --- LECTURA DE LOS EXCEL DE CONFIABILIDAD ---
    for archivo_nombre in archivos:
        if 'acciones' in archivo_nombre.lower(): continue # Saltar el de acciones

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

            planta_nombre = archivo_nombre.replace("Confiabilidad ", " ").replace(" .xlsx", "").strip()
            super_planta = "Carnes" if "carne" in planta_nombre.lower() else "Masas"

            # --- LIMPIEZA DETENCIONES ---
            col_equipo      = buscar_columna_equipo(df_det, planta_nombre)
            col_semana_det  = buscar_columna_semana(df_det)
            col_linea_det   = buscar_columna_linea(df_det)
            col_tpo_det     = buscar_tiempo_detencion_hr(df_det)

            if not all([col_equipo, col_semana_det, col_linea_det, col_tpo_det]):
                print("  ❌ Faltan columnas vitales en FEM. Saltando...")
                continue

            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas']  = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            df_det['Linea_Clean']   = df_det[col_linea_det].astype(str).str.replace(' ', ' ').str.strip().str.upper()
            df_det['Equipo_Clean']  = df_det[col_equipo].astype(str).str.strip().str.title()
            df_det['Semana_Clean']  = limpiar_semana(df_det[col_semana_det])
            df_det = df_det[df_det['Semana_Clean'] > 0]

            df_det = filtrar_semanas(df_det)

            agrup_det_linea = df_det.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_perdido_linea=('Hrs_Perdidas', 'sum')
            ).reset_index()

            # --- LIMPIEZA TIEMPOS PLANIFICADOS Y OPERATIVOS ---
            col_semana_tpo  = buscar_columna_semana(df_tpo)
            col_linea_tpo   = buscar_columna_linea(df_tpo)
            col_tpo_oper    = buscar_columna_tiempo_oper(df_tpo)
            col_tpo_plan    = buscar_columna_tiempo_plan(df_tpo)

            if not all([col_semana_tpo, col_linea_tpo]) or (not col_tpo_oper and not col_tpo_plan):
                print(f"  ❌ Faltan columnas de tiempo en {archivo_nombre}. Saltando...")
                continue

            df_tpo = df_tpo.dropna(subset=[col_linea_tpo, col_semana_tpo])
            df_tpo['Hrs_Oper']      = pd.to_numeric(df_tpo[col_tpo_oper], errors='coerce').fillna(0) if col_tpo_oper else 0
            df_tpo['Hrs_Plan']      = pd.to_numeric(df_tpo[col_tpo_plan], errors='coerce').fillna(0) if col_tpo_plan else 0
            df_tpo['Linea_Clean']   = df_tpo[col_linea_tpo].astype(str).str.replace(' ', ' ').str.strip().str.upper()
            df_tpo['Semana_Clean']  = limpiar_semana(df_tpo[col_semana_tpo])
            df_tpo = df_tpo[df_tpo['Semana_Clean'] > 0]

            df_tpo = filtrar_semanas(df_tpo)

            agrup_tpo_linea = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_operativo_linea=('Hrs_Oper', 'sum'),
                tpo_plan_linea=('Hrs_Plan', 'sum')
            ).reset_index()

            # --- CRUCE MAESTRO ---
            linea_merged = pd.merge(agrup_tpo_linea, agrup_det_linea, on=['Linea_Clean', 'Semana_Clean'], how='outer')
            linea_merged['tpo_perdido_linea']   = linea_merged.get('tpo_perdido_linea',   pd.Series([0] * len(linea_merged))).fillna(0)
            linea_merged['tpo_operativo_linea'] = linea_merged.get('tpo_operativo_linea', pd.Series([0] * len(linea_merged))).fillna(0)
            linea_merged['tpo_plan_linea']      = linea_merged.get('tpo_plan_linea',      pd.Series([0] * len(linea_merged))).fillna(0)

            tipo_tiempo = 'operativo' if col_tpo_oper else 'plan'

            for idx, row in linea_merged.iterrows():
                plan    = row['tpo_plan_linea']
                oper    = row['tpo_operativo_linea']
                perdido = row['tpo_perdido_linea']

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
                if row_lm['tpo_operativo_linea'] == 0 or 'MULTIVAC' in l_str or 'VARIOVAC' in l_str:
                    sem        = row_lm['Semana_Clean']
                    target_key = None
                    if   'MULTIVAC 1' in l_str or 'M1' in l_str:                       target_key = 'L1'
                    elif 'MULTIVAC 2' in l_str or 'M2' in l_str or 'VARIOVAC' in l_str: target_key = 'L2'
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
:root {
  --primary: #0f172a; --secondary: #1e293b; --accent: #0ea5e9;
  --bg: #f1f5f9; --surface: #ffffff; --border: #e2e8f0;
  --text: #0f172a; --text-muted: #64748b;
  --success: #10b981; --danger: #ef4444; --warning: #f59e0b;
}
body.theme-carnes {
  --primary: #450a0a; --secondary: #7f1d1d; --accent: #dc2626;
  --bg: #fef2f2; --border: #fecaca;
}
* { box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }
body { background: var(--bg); color: var(--text); margin: 0; display: flex; flex-direction: column; min-height: 100vh; transition: background 0.4s; }
.top-bar { background: var(--primary); color: white; padding: 0 25px; height: 65px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); z-index: 10; transition: background 0.4s; }
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
.sidebar { width: 280px; background: var(--surface); border-right: 1px solid var(--border); display: flex; flex-direction: column; }
.filter-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center; }
.filter-body { padding: 20px; flex: 1; overflow-y: auto; }
.f-group { margin-bottom: 20px; }
.f-group label { display: block; font-size: 0.75rem; font-weight: 700; color: var(--text-muted); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }
select { width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.9rem; color: var(--text); background: var(--bg); cursor: pointer; transition: 0.2s; }
select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(14,165,233,0.2); }
.rango-semanas { display: flex; gap: 10px; align-items: center; }
.rango-semanas select { flex: 1; }
.rango-semanas span { font-weight: 700; color: var(--text-muted); }
.btn-export { background: var(--surface); border: 2px solid var(--success); color: var(--success); padding: 10px; border-radius: 8px; cursor: pointer; font-weight: 700; width: 100%; display: flex; justify-content: center; gap: 8px; transition: 0.2s; margin-top: 10px; }
.btn-export:hover { background: var(--success); color: white; }
.content { flex: 1; padding: 30px; overflow-y: auto; display: flex; flex-direction: column; gap: 25px; }
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 15px; }
.kpi-card { background: var(--surface); padding: 15px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; gap: 5px; position: relative; overflow: hidden; }
.kpi-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--accent); }
.kpi-card.c-red::before { background: var(--danger); }
.kpi-card.c-green::before { background: var(--success); }
.kpi-card.c-warn::before { background: var(--warning); }
.kpi-card span { font-size: 0.75rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
.kpi-card h3 { margin: 0; font-size: 1.8rem; color: var(--secondary); font-weight: 800; }
.charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(500px, 1fr)); gap: 25px; }
.chart-container { background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); height: 350px; display: flex; flex-direction: column; }
.chart-header { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; display: flex; justify-content: space-between; }
.canvas-wrapper { flex: 1; position: relative; min-height: 0; }
.table-container { background: var(--surface); border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); overflow: hidden; display: flex; flex-direction: column; }
.table-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center; }
.table-header input { padding: 8px 15px; border: 1px solid var(--border); border-radius: 20px; font-size: 0.9rem; width: 300px; background: var(--bg); }
.table-wrapper { overflow-x: auto; max-height: 500px; }
table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.85rem; }
th { background: #f8fafc; padding: 12px 15px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; position: sticky; top: 0; z-index: 2; border-bottom: 2px solid var(--border); cursor: pointer; }
body.theme-carnes th { background: #fef2f2; }
td { padding: 12px 15px; border-bottom: 1px solid var(--border); color: var(--secondary); font-weight: 500; }
tr.clickable-row { cursor: pointer; transition: background 0.15s; }
tr.clickable-row:hover td { background: rgba(14,165,233,0.04) !important; }
body.theme-carnes tr.clickable-row:hover td { background: rgba(220,38,38,0.04) !important; }
.badge { padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; }
.b-ok { background: #d1fae5; color: #047857; }
.b-warn { background: #fef3c7; color: #b45309; }
.b-danger { background: #fee2e2; color: #b91c1c; }
.modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(15,23,42,0.6); backdrop-filter: blur(5px); display: none; justify-content: center; align-items: center; z-index: 100; padding: 20px; }
.modal-content { background: var(--surface); width: 100%; max-width: 850px; border-radius: 16px; border: 1px solid var(--border); box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1); display: flex; flex-direction: column; max-height: 80vh; animation: modalIn 0.2s ease-out; }
@keyframes modalIn { from { transform: translateY(15px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
.modal-header { padding: 20px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; background: #f8fafc; border-top-left-radius: 16px; border-top-right-radius: 16px; }
body.theme-carnes .modal-header { background: #fef2f2; }
.modal-header h3 { margin: 0; font-size: 1.2rem; font-weight: 800; color: var(--secondary); }
.modal-header p { margin: 2px 0 0 0; font-size: 0.85rem; color: var(--text-muted); font-weight: 500; }
.modal-body { padding: 20px; overflow-y: auto; }
.close-btn { background: none; border: none; font-size: 1.5rem; color: var(--text-muted); cursor: pointer; font-weight: 700; }
.close-btn:hover { color: var(--danger); }
.modal-body table th { position: static; background: #f1f5f9; border-bottom: 1px solid var(--border); }
body.theme-carnes .modal-body table th { background: #fee2e2; }
/* ── TABS ── */
.tab-nav { background: var(--surface); border-bottom: 2px solid var(--border); display: flex; gap: 2px; padding: 0 20px; flex-shrink: 0; }
.tab-btn { padding: 11px 22px; font-size: 0.85rem; font-weight: 700; color: var(--text-muted); border: none; background: none; cursor: pointer; border-bottom: 3px solid transparent; margin-bottom: -2px; transition: 0.2s; }
.tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
body.theme-carnes .tab-btn.active { color: #dc2626; border-bottom-color: #dc2626; }
.tab-btn:hover:not(.active) { color: var(--text); background: var(--bg); border-radius: 6px 6px 0 0; }
.tab-panel { display: none; flex: 1; overflow: hidden; }
.tab-panel.active { display: flex; }

/* ── RESUMEN (RESTAURADO A COMO ESTABA AL PRINCIPIO) ── */
.resumen-panel { flex-direction: row; align-items: flex-start; flex-wrap: wrap; gap: 14px; padding: 18px 24px; overflow-y: auto; background: var(--bg); }
.plant-card { flex: 1; min-width: 420px; background: var(--surface); border-radius: 10px; border: 1px solid var(--border); overflow: hidden; display: flex; box-shadow: 0 2px 6px rgba(0,0,0,0.03); }
.plant-bar-col { width: 200px; flex-shrink: 0; padding: 12px 14px; display: flex; flex-direction: column; }
.plant-bar-col.tema-masas { background: #1A3A5C; }
.plant-bar-col.tema-carnes { background: #4A0E0E; }
.plant-bar-title { font-size: 8px; font-weight: 800; color: #fff; letter-spacing:.8px; text-transform: uppercase; margin-bottom: 8px; padding: 2px 8px; border-radius: 3px; display: inline-block; }
.plant-bar-col.tema-masas .plant-bar-title { background: #0071CE; }
.plant-bar-col.tema-carnes .plant-bar-title { background: #dc2626; }
.plant-sub-lbl { font-size: 7.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: rgba(255,255,255,0.45); margin-bottom: 8px; padding-bottom: 5px; border-bottom: 1px solid rgba(255,255,255,0.15); }
.line-bar-item { margin-bottom: 8px; }
.line-bar-top { display: flex; align-items: center; gap: 5px; margin-bottom: 3px; }
.line-bar-lbl { width: 34px; font-size: 8.5px; font-weight: 800; color: #fff; flex-shrink: 0; opacity:.9; }
.line-bar-track { flex: 1; background: rgba(255,255,255,0.15); border-radius: 3px; height: 11px; overflow: hidden; }
.line-bar-fill { height: 100%; border-radius: 3px; opacity: 0.9; }
.line-bar-pct { width: 38px; text-align: right; font-size: 9px; font-weight: 800; flex-shrink: 0; }
.line-hrs { display: flex; gap: 3px; align-items: center; font-size: 6.5px; color: rgba(255,255,255,0.75); flex-wrap: wrap; }
.hrs-dot { width: 5px; height: 5px; border-radius: 1px; flex-shrink: 0; }
.plant-avg { margin-top: auto; text-align: center; padding-top: 10px; border-top: 1px solid rgba(255,255,255,0.15); margin-top: 6px; }
.plant-avg-lbl { font-size: 7px; font-weight: 700; color: rgba(255,255,255,0.45); text-transform: uppercase; letter-spacing:.8px; margin-bottom: 2px; }
.plant-avg-val { font-size: 28px; font-weight: 800; color: #fff; line-height: 1; }
.plant-eq-col { flex: 1; padding: 12px 16px; overflow-y: auto; }
.plant-eq-col h4 { font-size: 9px; font-weight: 800; text-transform: uppercase; letter-spacing:.8px; color: var(--text-muted); margin-bottom: 10px; }

/* ── GRID MODIFICADA CON LA COLUMNA ACCION CORRECTIVA ── */
.eq-rank-row { display: grid; grid-template-columns: 22px 1fr 1.5fr 88px 88px; gap: 7px; align-items: center; padding: 6px 0; border-bottom: 1px solid var(--border); }
.eq-rank-num { width: 20px; height: 20px; border-radius: 50%; font-size: 8px; font-weight: 800; color: #fff; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
.rn1{background:#CC2222;}.rn2{background:#E65100;}.rn3{background:#F9A825;color:#333;}.rn4{background:#5D6D7E;}.rn5{background:#95A5A6;}
.eq-rank-name { font-size: 10px; font-weight: 700; color: var(--secondary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.eq-rank-badge { font-size: 7px; padding: 1px 5px; border-radius: 3px; font-weight: 800; background: var(--accent); color: #fff; display: inline-block; margin-left: 4px; vertical-align: middle; }
body.theme-carnes .eq-rank-badge { background: #dc2626; }
.eq-rank-accion { font-size: 9px; font-weight: 500; font-style: italic; color: #64748b; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.eq-rank-kpi { text-align: center; background: #f8fafc; border-radius: 5px; padding: 4px 3px; }
body.theme-carnes .eq-rank-kpi { background: #fef2f2; }
.eq-rank-kpi .rval { font-size: 12px; font-weight: 800; line-height: 1; display: block; }
.eq-rank-kpi .rsub { font-size: 6.5px; color: var(--text-muted); margin-top: 1px; display: block; }
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
    <div style="padding:20px;border-top:1px solid var(--border);">
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
Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
Chart.defaults.color = '#64748b';

const dbRaw       = __DB_JSON_DATA__;
const recordsEq   = dbRaw.equipos;
const recordsLn   = dbRaw.lineas;

let isCarnesTheme  = false;
let currentEqData  = [];
let currentLnData  = [];
let tableDataFull  = [];
let chartInstances = {};

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

  const createSelect = (id, label, options) => {
    let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()"><option value="ALL">Todas</option>`;
    options.sort().forEach(o => sel += `<option value="${o}">${o}</option>`);
    return sel + `</select></div>`;
  };

  let htmlDyn = '';
  htmlDyn += createSelect('f_planta', '🏭 Planta',  [...new Set(baseLn.map(x => x.planta))]);
  htmlDyn += createSelect('f_linea',  '🔧 Línea',   [...new Set(baseLn.map(x => x.linea))]);
  document.getElementById('filters_dynamic').innerHTML = htmlDyn;
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
    eqMap[eqKey].eventos.push({ fecha: d.fecha, componente: d.componente, tipo: d.tipo, hrs: d.tpo_perdido_eq });
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

  const secColor    = isCarnesTheme ? '#991b1b' : '#334155';
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
        backgroundColor: isCarnesTheme ? '#dc2626' : '#ea580c',
        borderColor:     isCarnesTheme ? '#991b1b' : '#b45309',
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
        { label: 'MTTR (Hrs)', data: mttrTrend, borderColor: '#f59e0b',   borderWidth: 3, borderDash: [5, 5], tension: 0.3, yAxisID: 'y1' },
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

window.onload = () => toggleTheme();

function switchTab(tab) {
  ['resumen','analisis'].forEach(t => {
    document.getElementById('tab_' + t).classList.toggle('active', t === tab);
    document.getElementById('tbtn_' + t).classList.toggle('active', t === tab);
  });
}

function renderResumen() {
  const el = document.getElementById('resumen_content');
  if (!currentLnData.length) { el.innerHTML = '<p style="padding:30px;color:var(--text-muted);">Sin datos para el rango seleccionado.</p>'; return; }

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

  const tema = isCarnesTheme ? 'tema-carnes' : 'tema-masas';
  
  let html = '';
  Object.keys(lnByPlanta).sort().forEach(planta => {
    const lineas = lnByPlanta[planta];
    const eqs    = Object.values(eqByPlanta[planta] || {});

    const lineaStats = Object.entries(lineas).map(([ln, d]) => {
      const lnDet  = eqs.filter(e => e.linea === ln).reduce((s,e) => s + e.det, 0);
      const mtbf   = lnDet > 0 ? d.op / lnDet : 0;
      const prob   = mtbf > 0 ? (1 - Math.exp(-120 / mtbf)) * 100 : (lnDet > 0 ? 100 : 0);
      const perdido = eqs.filter(e => e.linea === ln).reduce((s,e) => s + e.hrs, 0);
      return { ln, op: d.op, pl: d.pl, perdido, prob };
    }).sort((a,b) => b.prob - a.prob);

    const avgProb = lineaStats.length ? lineaStats.reduce((s,x)=>s+x.prob,0)/lineaStats.length : 0;   
    const maxProb = Math.max(...lineaStats.map(x => x.prob), 1);

    const top5 = eqs.map(e => {
      const lnOp  = lineas[e.linea]?.op || 0;
      const mtbf  = e.det > 0 ? lnOp / e.det : 0;
      const mttr  = e.det > 0 ? e.hrs / e.det : 0;
      const prob  = mtbf > 0 ? (1 - Math.exp(-120/mtbf))*100 : (e.det>0?100:0);
      return { ...e, prob, mtbf, mttr };
    }).filter(e => e.det > 0).sort((a,b)=>b.prob-a.prob).slice(0,5);

    const barras = lineaStats.map((s,i) => {
      const pct = (s.prob / Math.max(maxProb, 1) * 100).toFixed(1);
      const color = s.prob > 60 ? '#C0392B' : s.prob > 35 ? '#E67E22' : '#27AE60';
      return `
        <div class="line-bar-item">
          <div class="line-bar-top">
            <div class="line-bar-lbl">${s.ln.replace('LINEA','L').replace('L\u00cdNEA','L').substring(0,5)}</div>
            <div class="line-bar-track"><div class="line-bar-fill" style="width:${pct}%;background:${color};"></div></div>
            <div class="line-bar-pct" style="color:${color};">${s.prob.toFixed(1).replace('.',',')}%</div>
          </div>
          <div class="line-hrs">
            <span class="hrs-dot" style="background:rgba(100,180,255,0.85);"></span><span>${s.op.toFixed(0)}h</span>
            <span style="opacity:.4;">/</span>
            <span class="hrs-dot" style="background:rgba(255,255,255,0.25);"></span><span>${s.pl.toFixed(0)}h</span>
            <span style="opacity:.4;">&middot;</span>
            <span class="hrs-dot" style="background:rgba(255,80,60,0.85);"></span><span style="color:rgba(255,140,120,1);font-weight:700;">${s.perdido.toFixed(1)}h</span>
          </div>
        </div>`;
    }).join('');

    const rankColors = ['rn1','rn2','rn3','rn4','rn5'];
    const pKeyLowerCase = planta.toLowerCase();

    const filas = top5.map((e, i) => {
      const probColor = e.prob > 60 ? '#b91c1c' : e.prob > 35 ? '#b45309' : '#047857';
      const rankTop = i + 1;
      
      // Lógica de coincidencia inteligente para sacar el texto de Acción Correctiva
      let accionStr = "-";
      if (dbRaw.acciones) {
        for (let col in dbRaw.acciones) {
            // Si el nombre de la planta en el excel coincide total o parcialmente con la planta real
            if (pKeyLowerCase.includes(col) || col.includes(pKeyLowerCase)) {
                accionStr = dbRaw.acciones[col][rankTop] || "-";
                break;
            }
        }
      }

      return `
        <div class="eq-rank-row">
          <div class="eq-rank-num ${rankColors[i]}">${rankTop}</div>
          <div class="eq-rank-name" title="${e.equipo}">${e.equipo}<span class="eq-rank-badge">${e.linea.substring(0,5)}</span></div>
          <div class="eq-rank-accion" title="${accionStr}">${accionStr}</div>
          <div class="eq-rank-kpi">
            <span class="rval" style="color:${probColor};">${e.prob.toFixed(1)}%</span>
            <span class="rsub">Prob. Falla</span>
          </div>
          <div class="eq-rank-kpi">
            <span class="rval" style="color:#1A5276;">${e.hrs.toFixed(1)}h</span>
            <span class="rsub">${e.det} deten.</span>
          </div>
        </div>`;
    }).join('');

    html += `
      <div class="plant-card">
        <div class="plant-bar-col ${tema}">
          <div><span class="plant-bar-title">${planta.toUpperCase()}</span></div>
          <div class="plant-sub-lbl">Pb. Falla &middot; Hrs</div>
          ${barras}
          <div class="plant-avg">
            <div class="plant-avg-lbl">Pf. Promedio</div>
            <div class="plant-avg-val">${avgProb.toFixed(1).replace('.',',')}%</div>
          </div>
        </div>
        <div class="plant-eq-col">
          <h4>&#128680; Top Equipos Cr&iacute;ticos &mdash; ${planta}</h4>
          ${top5.length ? filas : '<p style="color:var(--text-muted);font-size:0.85rem;padding:10px 0;">Sin detenciones en el per&iacute;odo.</p>'}
        </div>
      </div>`;
  });

  el.innerHTML = html;
}
</script>
</body>
</html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    full_html = full_html.replace("__FECHA_ACTUAL__", fecha_actual)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(full_html)

    print(f"✅ Dashboard generado: {OUTPUT_HTML}")


if __name__ == "__main__":
    db = procesar_datos_confiabilidad()
    if db:
        generar_html_moderno(db)
