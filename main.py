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
# 2. BUSCADORES ULTRA-ESTRICTOS (A PRUEBA DE ERRORES)
# ==========================================
def buscar_columna_linea(df):
    for c in df.columns:
        if str(c).strip().lower().replace('í', 'i') == 'linea': return c
    for c in df.columns:
        if 'linea' in str(c).strip().lower().replace('í', 'i') and 'aux' not in str(c).lower(): return c
    return None

def buscar_columna_equipo(df, super_planta):
    if super_planta == "Carnes":
        for c in df.columns:
            if str(c).strip().lower() == 'detalle': return c
    for c in df.columns:
        if str(c).strip().lower() == 'equipo': return c
    for c in df.columns:
        if str(c).strip().lower() == 'componente': return c
    return None

def buscar_columna_semana(df):
    # Busca todas las candidatas que tengan la palabra semana
    candidatas = [c for c in df.columns if 'semana' in str(c).lower() and 'aux' not in str(c).lower()]
    # Retorna la primera que realmente contenga datos (Esquiva la columna vacía de Carnes Molida)
    for c in candidatas:
        if not df[c].isnull().all(): return c
    return candidatas[0] if candidatas else None

def buscar_tiempo_detencion_hr(df):
    for c in df.columns:
        cl = str(c).lower().strip().replace('  ', ' ')
        if 'detencion' in cl and 'hr' in cl: return c
    return None

def buscar_tiempo_planificado_hr(df):
    for c in df.columns:
        cl = str(c).lower().strip().replace('  ', ' ')
        if 'plan' in cl and 'hr' in cl: return c
        if 'disponible' in cl and 'hr' in cl: return c
    return None

# ==========================================
# 3. EXTRACCIÓN Y TRANSFORMACIÓN (ETL)
# ==========================================
def procesar_datos_confiabilidad():
    print(f"📥 Conectando a Google Drive...")
    if os.path.exists(DATA_DIR): shutil.rmtree(DATA_DIR)
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        gdown.download_folder(url=DRIVE_FOLDER_URL, output=DATA_DIR, quiet=False, use_cookies=False)
    except Exception as e:
        print(f"❌ Error al descargar de Drive: {e}")

    print("\n🚀 INICIANDO EXTRACCIÓN DE DATOS BASE...")
    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]
    datos_equipos = []
    datos_lineas = []
    
    for archivo_nombre in archivos:
        ruta_completa = os.path.join(DATA_DIR, archivo_nombre)
        print(f"\n📊 Analizando: {archivo_nombre}")
        
        try:
            excel = pd.ExcelFile(ruta_completa)
            hoja_det = next((h for h in excel.sheet_names if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in excel.sheet_names if h.endswith('_Tiempos_Planificados')), None)
            
            if not hoja_det or not hoja_tpo:
                print(f"   ⚠️ Faltan pestañas base. Saltando...")
                continue
                
            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)
            planta_nombre = archivo_nombre.replace("Confiabilidad ", "").replace(".xlsx", "").strip()
            super_planta = "Carnes" if "carne" in planta_nombre.lower() else "Masas"
            
            # --- LIMPIEZA DETENCIONES ---
            col_equipo = buscar_columna_equipo(df_det, super_planta)
            col_semana_det = buscar_columna_semana(df_det)
            col_linea_det = buscar_columna_linea(df_det)
            col_tpo_det = buscar_tiempo_detencion_hr(df_det)
            
            if not all([col_equipo, col_semana_det, col_linea_det, col_tpo_det]):
                print(f"   ❌ Faltan columnas. Eq:{col_equipo}, Sem:{col_semana_det}, Lin:{col_linea_det}, Tpo:{col_tpo_det}")
                continue

            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas'] = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            
            df_det['Linea_Clean'] = df_det[col_linea_det].astype(str).str.replace('  ', ' ').str.strip().str.upper()
            df_det['Equipo_Clean'] = df_det[col_equipo].astype(str).str.strip().str.title()
            
            # Limpia letras como "S19" dejando solo "19"
            df_det['Semana_Clean'] = df_det[col_semana_det].astype(str).str.replace(r'[^\d]', '', regex=True)
            df_det['Semana_Clean'] = pd.to_numeric(df_det['Semana_Clean'], errors='coerce').fillna(-1).astype(int)
            df_det = df_det[df_det['Semana_Clean'] > 0]
            
            agrup_det_linea = df_det.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_perdido_linea=('Hrs_Perdidas', 'sum')
            ).reset_index()
            
            agrup_det_eq = df_det.groupby(['Linea_Clean', 'Semana_Clean', 'Equipo_Clean']).agg(
                detenciones=('Equipo_Clean', 'count'),
                tpo_perdido_eq=('Hrs_Perdidas', 'sum')
            ).reset_index()
            
            # --- LIMPIEZA TIEMPOS PLANIFICADOS ---
            col_semana_tpo = buscar_columna_semana(df_tpo)
            col_linea_tpo = buscar_columna_linea(df_tpo)
            col_tpo_plan = buscar_tiempo_planificado_hr(df_tpo)
            
            if not all([col_semana_tpo, col_linea_tpo, col_tpo_plan]):
                print(f"   ❌ Faltan columnas. Sem:{col_semana_tpo}, Lin:{col_linea_tpo}, Tpo:{col_tpo_plan}")
                continue

            df_tpo = df_tpo.dropna(subset=[col_linea_tpo, col_semana_tpo])
            df_tpo['Hrs_Plan'] = pd.to_numeric(df_tpo[col_tpo_plan], errors='coerce').fillna(0)
            
            df_tpo['Linea_Clean'] = df_tpo[col_linea_tpo].astype(str).str.replace('  ', ' ').str.strip().str.upper()
            
            df_tpo['Semana_Clean'] = df_tpo[col_semana_tpo].astype(str).str.replace(r'[^\d]', '', regex=True)
            df_tpo['Semana_Clean'] = pd.to_numeric(df_tpo['Semana_Clean'], errors='coerce').fillna(-1).astype(int)
            df_tpo = df_tpo[df_tpo['Semana_Clean'] > 0]
            
            agrup_tpo_linea = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_plan_linea=('Hrs_Plan', 'sum')
            ).reset_index()
            
            # --- CRUCE MAESTRO (OUTER JOIN PARA NO PERDER DATOS) ---
            linea_merged = pd.merge(agrup_tpo_linea, agrup_det_linea, on=['Linea_Clean', 'Semana_Clean'], how='outer')
            linea_merged['tpo_plan_linea'] = linea_merged['tpo_plan_linea'].fillna(0)
            linea_merged['tpo_perdido_linea'] = linea_merged['tpo_perdido_linea'].fillna(0)
            linea_merged['tpo_operativo_linea'] = (linea_merged['tpo_plan_linea'] - linea_merged['tpo_perdido_linea']).clip(lower=0)
            
            for _, row in linea_merged.iterrows():
                datos_lineas.append({
                    "super_planta": super_planta,
                    "planta": planta_nombre,
                    "linea": row['Linea_Clean'],
                    "semana": int(row['Semana_Clean']),
                    "tpo_operativo_linea": float(row.get('tpo_operativo_linea', 0))
                })

            df_final_eq = pd.merge(agrup_det_eq, linea_merged, on=['Linea_Clean', 'Semana_Clean'], how='left')
            for _, row in df_final_eq.iterrows():
                datos_equipos.append({
                    "super_planta": super_planta,
                    "planta": planta_nombre,
                    "linea": row['Linea_Clean'],
                    "equipo": row['Equipo_Clean'],
                    "semana": int(row['Semana_Clean']),
                    "detenciones": int(row['detenciones']),
                    "tpo_perdido_eq": float(row['tpo_perdido_eq'])
                })
                
            print(f"   ✅ Procesado con éxito. Extraídos {len(df_final_eq)} equipos con fallas.")
                
        except Exception as e:
            print(f"   ❌ Error fatal procesando {archivo_nombre}: {e}")

    db_json = { "equipos": datos_equipos, "lineas": datos_lineas }
    print(f"\n✅ Extracción finalizada. Datos listos para el Dashboard.")
    return db_json

# ==========================================
# 4. GENERADOR HTML DASHBOARD
# ==========================================
def generar_html_moderno(db_json):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    
    html_template = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Dashboard Confiabilidad</title>
    <link rel="icon" type="image/x-icon" href="https://www.walmart.com/favicon.ico">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root { --primary: #0f172a; --secondary: #1e293b; --accent: #0ea5e9; --bg: #f1f5f9; --surface: #ffffff; --border: #e2e8f0; --text: #0f172a; --text-muted: #64748b; --success: #10b981; --danger: #ef4444; --warning: #f59e0b; }
        body.theme-carnes { --primary: #450a0a; --secondary: #7f1d1d; --accent: #dc2626; --bg: #fef2f2; --border: #fecaca; }
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
        select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2); }
        
        .rango-semanas { display: flex; gap: 10px; align-items: center; }
        .rango-semanas select { flex: 1; }
        .rango-semanas span { font-weight: 700; color: var(--text-muted); }

        .btn-export { background: var(--surface); border: 2px solid var(--success); color: var(--success); padding: 10px; border-radius: 8px; cursor: pointer; font-weight: 700; width: 100%; display: flex; justify-content: center; gap: 8px; transition: 0.2s; margin-top: 10px; }
        .btn-export:hover { background: var(--success); color: white; }
        
        .content { flex: 1; padding: 30px; overflow-y: auto; display: flex; flex-direction: column; gap: 25px; }
        .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; }
        .kpi-card { background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; gap: 5px; position: relative; overflow: hidden; }
        .kpi-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--accent); }
        .kpi-card.c-red::before { background: var(--danger); }
        .kpi-card.c-green::before { background: var(--success); }
        .kpi-card span { font-size: 0.8rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
        .kpi-card h3 { margin: 0; font-size: 2rem; color: var(--secondary); font-weight: 800; }
        
        .charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(500px, 1fr)); gap: 25px; }
        .chart-container { background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); height: 350px; display: flex; flex-direction: column; }
        .chart-header { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; display: flex; justify-content: space-between;}
        .canvas-wrapper { flex: 1; position: relative; min-height: 0; }

        .table-container { background: var(--surface); border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); overflow: hidden; display: flex; flex-direction: column; }
        .table-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center;}
        .table-header input { padding: 8px 15px; border: 1px solid var(--border); border-radius: 20px; font-size: 0.9rem; width: 300px; background: var(--bg); }
        .table-wrapper { overflow-x: auto; max-height: 500px; }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.85rem; }
        th { background: #f8fafc; padding: 12px 15px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; position: sticky; top: 0; z-index: 2; border-bottom: 2px solid var(--border); cursor: pointer; }
        body.theme-carnes th { background: #fef2f2; }
        td { padding: 12px 15px; border-bottom: 1px solid var(--border); color: var(--secondary); font-weight: 500; }
        tr:hover td { background: var(--bg); }
        .badge { padding: 4px 8px; border-radius: 6px; font-weight: 700; font-size: 0.75rem; }
        .b-ok { background: #d1fae5; color: #047857; }
        .b-warn { background: #fef3c7; color: #b45309; }
        .b-danger { background: #fee2e2; color: #b91c1c; }
    </style>
</head>
<body>
    <div class="top-bar">
        <div class="brand">
            <img src="https://upload.wikimedia.org/wikipedia/commons/b/b1/Walmart_logo_%282008%29.svg" alt="Walmart" style="height: 30px; filter: brightness(0) invert(1);">
            <h2>Dashboard Confiabilidad <span>Subgerencia de mantenimiento 2026</span></h2>
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
    
    <div class="main-container">
        <div class="sidebar">
            <div class="filter-header">
                <span>Filtros Acumulados</span>
            </div>
            <div class="filter-body">
                <div class="f-group">
                    <label>📆 Rango de Semanas</label>
                    <div class="rango-semanas">
                        <select id="f_sem_desde" onchange="applyFilters()"></select>
                        <span>a</span>
                        <select id="f_sem_hasta" onchange="applyFilters()"></select>
                    </div>
                </div>
                <div id="filters_dynamic"></div>
            </div>
            <div style="padding: 20px; border-top: 1px solid var(--border);">
                <button class="btn-export" onclick="descargarExcel()">⬇️ Exportar Data a Excel</button>
                <div style="text-align:center; font-size:0.7rem; color:var(--text-muted); margin-top:15px; font-weight:600;">Actualizado: __FECHA_ACTUAL__</div>
            </div>
        </div>

        <div class="content">
            <div class="kpi-row">
                <div class="kpi-card">
                    <span>Equipos con Fallas</span>
                    <h3 id="k_equipos">0</h3>
                </div>
                <div class="kpi-card c-green">
                    <span>Confiabilidad Global (R)</span>
                    <h3 id="k_conf">0%</h3>
                </div>
                <div class="kpi-card c-red">
                    <span>Tpo. Perdido Total (Hrs)</span>
                    <h3 id="k_hrs">0.0</h3>
                </div>
                <div class="kpi-card">
                    <span>Mantenibilidad Global (M)</span>
                    <h3 id="k_mant">0%</h3>
                </div>
            </div>

            <div class="charts-row">
                <div class="chart-container">
                    <div class="chart-header">📈 Tendencia de Confiabilidad y Prob. Falla (Semanal)</div>
                    <div class="canvas-wrapper"><canvas id="chart_trend_conf"></canvas></div>
                </div>
                <div class="chart-container">
                    <div class="chart-header">⏱️ Evolución Tiempos Medios (MTBF vs MTTR)</div>
                    <div class="canvas-wrapper"><canvas id="chart_trend_mtbf"></canvas></div>
                </div>
            </div>

            <div class="table-container">
                <div class="table-header">
                    <span>📋 Matriz Acumulada de KPIs (Rango de semanas seleccionado)</span>
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

    <script>
    Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
    Chart.defaults.color = '#64748b';

    const dbRaw = __DB_JSON_DATA__;
    const recordsEq = dbRaw.equipos;
    const recordsLn = dbRaw.lineas;
    
    let isCarnesTheme = false;
    let currentEqData = [];
    let currentLnData = [];
    let tableDataFull = []; 
    let chartInstances = {};

    function toggleTheme() {
        isCarnesTheme = document.getElementById('theme_toggle').checked;
        if(isCarnesTheme) {
            document.body.classList.add('theme-carnes');
            document.getElementById('lbl_carnes').classList.add('active');
            document.getElementById('lbl_masas').classList.remove('active');
