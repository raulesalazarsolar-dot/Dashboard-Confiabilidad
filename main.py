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

# PARÁMETROS DEL MODELO (Igual a tu Excel)
TIEMPO_MISION_HRS = 120
TIEMPO_IDEAL_REP_HRS = 1

# ==========================================
# 2. FUNCIONES DE BÚSQUEDA INTELIGENTE
# ==========================================
def buscar_columna(columnas, palabras_clave, excluir=None):
    for c in columnas:
        cl = str(c).lower().replace('ó','o').replace('í','i').strip()
        if excluir and excluir in cl:
            continue
        for palabra in palabras_clave:
            if palabra in cl:
                return c
    return None

def buscar_tiempo_detencion(columnas):
    for c in columnas:
        cl = str(c).lower().replace('ó','o')
        if 'std' in cl: continue
        if 'total [min]' in cl: return c, True
        if 'detencion [hr]' in cl: return c, False
    for c in columnas:
        cl = str(c).lower().replace('ó','o')
        if 'std' in cl: continue
        if 'detencion' in cl and '[min]' in cl: return c, True
        if 'detencion' in cl and 'hr' in cl: return c, False
    return None, False

def buscar_tiempo_operativo(columnas):
    for c in columnas:
        cl = str(c).lower()
        if ('operativo' in cl or 'real' in cl) and '[hr]' in cl: return c, False
        if 'tpo hr real' in cl: return c, False
        if 'operativo' in cl and '[min]' in cl: return c, True
    return None, False

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

    print("\n🚀 INICIANDO PROCESAMIENTO INTELIGENTE DE EXCEL...")
    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]
    datos_equipos = []
    
    for archivo_nombre in archivos:
        ruta_completa = os.path.join(DATA_DIR, archivo_nombre)
        print(f"\n📊 Analizando: {archivo_nombre}")
        
        try:
            excel = pd.ExcelFile(ruta_completa)
            hojas = excel.sheet_names
            
            hoja_det = next((h for h in hojas if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in hojas if h.endswith('_Tiempos_Planificados')), None)
            
            if not hoja_det or not hoja_tpo:
                print(f"   ⚠️ Faltan pestañas base. Saltando...")
                continue
                
            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)
            planta_nombre = archivo_nombre.replace("Confiabilidad ", "").replace(".xlsx", "").strip()
            
            # Clasificación del Mundo: Carnes vs Masas
            super_planta = "Carnes" if "carne" in planta_nombre.lower() else "Masas"
            
            # --- LIMPIEZA DETENCIONES ---
            col_equipo = buscar_columna(df_det.columns, ['equipo', 'componente'])
            col_semana_det = buscar_columna(df_det.columns, ['semana'], excluir='aux')
            col_mes_det = buscar_columna(df_det.columns, ['mes', 'month'], excluir='aux')
            col_linea_det = buscar_columna(df_det.columns, ['linea'])
            col_tpo_det, is_det_min = buscar_tiempo_detencion(df_det.columns)
            
            if not all([col_equipo, col_semana_det, col_linea_det, col_tpo_det]):
                continue

            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas'] = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            if is_det_min: df_det['Hrs_Perdidas'] /= 60.0
            
            df_det['Linea_Clean'] = df_det[col_linea_det].astype(str).str.strip().str.upper()
            df_det['Semana_Clean'] = df_det[col_semana_det].astype(str).str.strip()
            df_det['Equipo_Clean'] = df_det[col_equipo].astype(str).str.strip()
            
            if col_mes_det:
                df_det['Mes_Clean'] = pd.to_numeric(df_det[col_mes_det], errors='coerce').fillna(0).astype(int).astype(str)
                df_det.loc[df_det['Mes_Clean'] == '0', 'Mes_Clean'] = 'N/A'
            else:
                df_det['Mes_Clean'] = 'N/A'
            
            agrup_det = df_det.groupby(['Linea_Clean', 'Semana_Clean', 'Equipo_Clean']).agg(
                detenciones=('Equipo_Clean', 'count'),
                tpo_perdido_hrs=('Hrs_Perdidas', 'sum'),
                Mes_Clean=('Mes_Clean', 'first')
            ).reset_index()
            
            # --- LIMPIEZA TIEMPOS PLANIFICADOS ---
            col_semana_tpo = buscar_columna(df_tpo.columns, ['semana'], excluir='aux')
            col_linea_tpo = buscar_columna(df_tpo.columns, ['linea'])
            col_tpo_op, is_op_min = buscar_tiempo_operativo(df_tpo.columns)
            
            if not all([col_semana_tpo, col_linea_tpo, col_tpo_op]):
                continue

            df_tpo = df_tpo.dropna(subset=[col_linea_tpo, col_semana_tpo])
            df_tpo['Tpo_Op_Num'] = pd.to_numeric(df_tpo[col_tpo_op], errors='coerce').fillna(0)
            if is_op_min: df_tpo['Tpo_Op_Num'] /= 60.0
            
            df_tpo['Linea_Clean'] = df_tpo[col_linea_tpo].astype(str).str.strip().str.upper()
            df_tpo['Semana_Clean'] = df_tpo[col_semana_tpo].astype(str).str.strip()
            
            agrup_tpo = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(
                tpo_operativo_linea_hrs=('Tpo_Op_Num', 'sum')
            ).reset_index()
            
            # --- MERGE (CRUCE EXACTO) ---
            df_final = pd.merge(agrup_det, agrup_tpo, on=['Linea_Clean', 'Semana_Clean'], how='left')
            
            for _, row in df_final.iterrows():
                datos_equipos.append({
                    "super_planta": super_planta,
                    "planta": planta_nombre,
                    "linea": str(row['Linea_Clean']),
                    "equipo": str(row['Equipo_Clean']),
                    "semana": str(row['Semana_Clean']),
                    "mes": str(row.get('Mes_Clean', 'N/A')),
                    "detenciones": int(row['detenciones']),
                    "tpo_perdido_hrs": float(row['tpo_perdido_hrs']),
                    "tpo_operativo_linea_hrs": float(row['tpo_operativo_linea_hrs']) if pd.notna(row['tpo_operativo_linea_hrs']) else 0.0
                })
            print(f"   ✅ Procesado correctamente. Datos extraídos: {len(df_final)}")
                
        except Exception as e:
            print(f"   ❌ Error fatal procesando {archivo_nombre}: {e}")

    # ==========================================
    # CÁLCULO MATEMÁTICO DE CONFIABILIDAD
    # ==========================================
    print("\n⏳ Calculando MTBF, MTTR, Confiabilidad y Mantenibilidad...")
    db_json = {}
    
    for idx, row in enumerate(datos_equipos):
        det = row['detenciones']
        tpo_perd = row['tpo_perdido_hrs']
        tpo_op = row['tpo_operativo_linea_hrs']
        
        # Mismas fórmulas exactas de tu Excel
        if det > 0:
            mttr = tpo_perd / det
            mtbf = tpo_op / det
            conf = math.exp(-TIEMPO_MISION_HRS / mtbf) * 100 if mtbf > 0 else 0
            mant = (1 - math.exp(-TIEMPO_IDEAL_REP_HRS / mttr)) * 100 if mttr > 0 else 100
        else:
            mttr = 0
            mtbf = float('inf')
            conf = 100.0
            mant = 100.0
            
        key_id = f"EQ_{idx+1}"
        db_json[key_id] = {
            "key_id": key_id,
            "super_planta": row['super_planta'],
            "planta": row['planta'],
            "linea": row['linea'],
            "equipo": row['equipo'].title(),
            "semana": int(float(row['semana'])) if row['semana'].replace('.','').isdigit() else row['semana'],
            "mes": int(float(row['mes'])) if row['mes'].replace('.','').isdigit() else row['mes'],
            "detenciones": det,
            "tpo_perdido_hrs": round(tpo_perd, 2),
            "tpo_operativo_hrs": round(tpo_op, 2),
            "mtbf": round(mtbf, 2) if det > 0 else 0,
            "mttr": round(mttr, 3) if det > 0 else 0,
            "confiabilidad": round(conf, 2),
            "mantenibilidad": round(mant, 2),
            "prob_falla": round(100 - conf, 2)
        }
        
    print(f"✅ Dashboard alimentado con {len(db_json)} equipos.")
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
        :root { 
            --primary: #0f172a; 
            --secondary: #1e293b; 
            --accent: #0ea5e9; 
            --bg: #f1f5f9; 
            --surface: #ffffff; 
            --border: #e2e8f0; 
            --text: #0f172a; 
            --text-muted: #64748b; 
            --success: #10b981; 
            --danger: #ef4444; 
            --warning: #f59e0b; 
        }
        body.theme-carnes {
            --primary: #450a0a; 
            --secondary: #7f1d1d; 
            --accent: #dc2626; 
            --bg: #fef2f2;
            --border: #fecaca;
        }
        * { box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }
        body { background: var(--bg); color: var(--text); margin: 0; display: flex; flex-direction: column; min-height: 100vh; transition: background 0.4s; }
        
        .top-bar { background: var(--primary); color: white; padding: 0 25px; height: 65px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); z-index: 10; transition: background 0.4s; }
        .brand { display: flex; align-items: center; gap: 15px; }
        .brand h2 { margin: 0; font-size: 1.3rem; font-weight: 700; letter-spacing: 0.5px; } 
        .brand span { opacity: 0.7; font-weight: 400; font-size: 1rem; border-left: 1px solid rgba(255,255,255,0.3); padding-left: 15px; }

        /* Switch UI */
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
        
        /* Sidebar Filters */
        .sidebar { width: 280px; background: var(--surface); border-right: 1px solid var(--border); display: flex; flex-direction: column; }
        .filter-header { padding: 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--secondary); display: flex; justify-content: space-between; align-items: center; }
        .filter-body { padding: 20px; flex: 1; overflow-y: auto; }
        .f-group { margin-bottom: 20px; }
        .f-group label { display: block; font-size: 0.75rem; font-weight: 700; color: var(--text-muted); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }
        select { width: 100%; padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.9rem; color: var(--text); background: var(--bg); cursor: pointer; transition: 0.2s; }
        select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2); }
        .btn-export { background: var(--surface); border: 2px solid var(--success); color: var(--success); padding: 10px; border-radius: 8px; cursor: pointer; font-weight: 700; width: 100%; display: flex; justify-content: center; gap: 8px; transition: 0.2s; margin-top: 10px; }
        .btn-export:hover { background: var(--success); color: white; }
        
        /* Content Area */
        .content { flex: 1; padding: 30px; overflow-y: auto; display: flex; flex-direction: column; gap: 25px; }
        
        /* KPI Cards */
        .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; }
        .kpi-card { background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; flex-direction: column; gap: 5px; position: relative; overflow: hidden; }
        .kpi-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--accent); }
        .kpi-card.c-red::before { background: var(--danger); }
        .kpi-card.c-green::before { background: var(--success); }
        .kpi-card span { font-size: 0.8rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
        .kpi-card h3 { margin: 0; font-size: 2rem; color: var(--secondary); font-weight: 800; }
        
        /* Charts */
        .charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(500px, 1fr)); gap: 25px; }
        .chart-container { background: var(--surface); padding: 20px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 2px 4px rgba(0,0,0,0.02); height: 350px; display: flex; flex-direction: column; }
        .chart-header { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; display: flex; justify-content: space-between;}
        .canvas-wrapper { flex: 1; position: relative; min-height: 0; }

        /* Table */
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
                <span>Filtros Globales</span>
            </div>
            <div class="filter-body" id="filters_dynamic">
                </div>
            <div style="padding: 20px; border-top: 1px solid var(--border);">
                <button class="btn-export" onclick="descargarExcel()">⬇️ Exportar Data a Excel</button>
                <div style="text-align:center; font-size:0.7rem; color:var(--text-muted); margin-top:15px; font-weight:600;">Actualizado: __FECHA_ACTUAL__</div>
            </div>
        </div>

        <div class="content">
            <div class="kpi-row">
                <div class="kpi-card">
                    <span>Equipos Analizados</span>
                    <h3 id="k_equipos">0</h3>
                </div>
                <div class="kpi-card c-green">
                    <span>Confiabilidad Promedio</span>
                    <h3 id="k_conf">0%</h3>
                </div>
                <div class="kpi-card c-red">
                    <span>Tpo. Perdido Total (Hrs)</span>
                    <h3 id="k_hrs">0.0</h3>
                </div>
                <div class="kpi-card">
                    <span>MTBF Global (Hrs)</span>
                    <h3 id="k_mtbf">0.0</h3>
                </div>
            </div>

            <div class="charts-row">
                <div class="chart-container">
                    <div class="chart-header">📈 Tendencia de Confiabilidad por Semana</div>
                    <div class="canvas-wrapper"><canvas id="chart_trend_conf"></canvas></div>
                </div>
                <div class="chart-container">
                    <div class="chart-header">⏱️ Evolución Tiempos Medios (MTBF vs MTTR)</div>
                    <div class="canvas-wrapper"><canvas id="chart_trend_mtbf"></canvas></div>
                </div>
            </div>

            <div class="table-container">
                <div class="table-header">
                    <span>📋 Matriz de Equipos y Fallas (Acumulado)</span>
                    <input type="text" id="search_input" placeholder="🔍 Buscar equipo o línea..." onkeyup="renderTable()">
                </div>
                <div class="table-wrapper">
                    <table id="data_table">
                        <thead>
                            <tr>
                                <th onclick="sortTable(0)">Planta ↕</th>
                                <th onclick="sortTable(1)">Línea ↕</th>
                                <th onclick="sortTable(2)">Equipo ↕</th>
                                <th onclick="sortTable(3)">Fallas ↕</th>
                                <th onclick="sortTable(4)">Tpo Perdido (Hrs) ↕</th>
                                <th onclick="sortTable(5)">MTBF ↕</th>
                                <th onclick="sortTable(6)">MTTR ↕</th>
                                <th onclick="sortTable(7)">Confiabilidad ↕</th>
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

    const db = __DB_JSON_DATA__;
    const records = Object.values(db);
    
    let isCarnesTheme = false;
    let baseData = []; // Datos filtrados por Mundo (Masas/Carnes)
    let currentData = []; // Datos finales filtrados por los selects
    let chartInstances = {};

    function toggleTheme() {
        isCarnesTheme = document.getElementById('theme_toggle').checked;
        if(isCarnesTheme) {
            document.body.classList.add('theme-carnes');
            document.getElementById('lbl_carnes').classList.add('active');
            document.getElementById('lbl_masas').classList.remove('active');
        } else {
            document.body.classList.remove('theme-carnes');
            document.getElementById('lbl_masas').classList.add('active');
            document.getElementById('lbl_carnes').classList.remove('active');
        }
        
        baseData = records.filter(d => isCarnesTheme ? d.super_planta === 'Carnes' : d.super_planta === 'Masas');
        buildFilters();
        applyFilters();
    }

    function buildFilters() {
        const createSelect = (id, label, options) => {
            let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()"><option value="ALL">Todos(as)</option>`;
            options.sort((a,b) => a-b).forEach(o => { 
                if(o !== 'N/A' && o != null) sel += `<option value="${o}">${o}</option>`;
            });
            return sel + `</select></div>`;
        };
        
        let html = '';
        html += createSelect('f_mes', '📅 Mes', [...new Set(baseData.map(x=>x.mes))]);
        html += createSelect('f_semana', '📆 Semana', [...new Set(baseData.map(x=>x.semana))]);
        html += createSelect('f_planta', '🏢 Planta', [...new Set(baseData.map(x=>x.planta))]);
        html += createSelect('f_linea', '🏭 Línea', [...new Set(baseData.map(x=>x.linea))]);
        
        document.getElementById('filters_dynamic').innerHTML = html;
    }

    function applyFilters() {
        const fMes = document.getElementById('f_mes').value;
        const fSem = document.getElementById('f_semana').value;
        const fPla = document.getElementById('f_planta').value;
        const fLin = document.getElementById('f_linea').value;
        const search = document.getElementById('search_input').value.toLowerCase();

        currentData = baseData.filter(d => {
            if(fMes !== 'ALL' && String(d.mes) !== fMes) return false;
            if(fSem !== 'ALL' && String(d.semana) !== fSem) return false;
            if(fPla !== 'ALL' && d.planta !== fPla) return false;
            if(fLin !== 'ALL' && d.linea !== fLin) return false;
            if(search && !`${d.equipo} ${d.linea}`.toLowerCase().includes(search)) return false;
            return true;
        });

        updateKPIs();
        drawCharts();
        renderTable();
    }

    function updateKPIs() {
        if(currentData.length === 0) {
            document.getElementById('k_equipos').innerText = "0";
            document.getElementById('k_conf').innerText = "0%";
            document.getElementById('k_hrs').innerText = "0.0";
            document.getElementById('k_mtbf').innerText = "0.0";
            return;
        }

        let unqEquipos = new Set(currentData.map(d => d.planta + d.linea + d.equipo)).size;
        let avgConf = currentData.reduce((s, d) => s + d.confiabilidad, 0) / currentData.length;
        let sumHrs = currentData.reduce((s, d) => s + d.tpo_perdido_hrs, 0);
        
        let sumOp = currentData.reduce((s, d) => s + d.tpo_operativo_hrs, 0);
        let sumDet = currentData.reduce((s, d) => s + d.detenciones, 0);
        let mtbfGlob = sumDet > 0 ? (sumOp / sumDet) : 0;

        document.getElementById('k_equipos').innerText = unqEquipos;
        document.getElementById('k_conf').innerText = avgConf.toFixed(1) + "%";
        document.getElementById('k_hrs').innerText = sumHrs.toFixed(1);
        document.getElementById('k_mtbf').innerText = mtbfGlob.toFixed(1);
    }

    function drawCharts() {
        if(currentData.length === 0) return;
        
        const accentColor = isCarnesTheme ? '#dc2626' : '#0ea5e9';
        const secColor = isCarnesTheme ? '#991b1b' : '#334155';
        
        let weeks = [...new Set(currentData.map(d => parseInt(d.semana) || d.semana))].sort((a,b) => a-b);
        
        let confTrend = [];
        let mtbfTrend = [];
        let mttrTrend = [];

        weeks.forEach(w => {
            let dw = currentData.filter(d => d.semana == w);
            let avgC = dw.reduce((s, d) => s + d.confiabilidad, 0) / dw.length;
            
            let sOp = dw.reduce((s, d) => s + d.tpo_operativo_hrs, 0);
            let sPerd = dw.reduce((s, d) => s + d.tpo_perdido_hrs, 0);
            let sDet = dw.reduce((s, d) => s + d.detenciones, 0);
            
            let avgMTBF = sDet > 0 ? (sOp / sDet) : 0;
            let avgMTTR = sDet > 0 ? (sPerd / sDet) : 0;

            confTrend.push(avgC.toFixed(2));
            mtbfTrend.push(avgMTBF.toFixed(2));
            mttrTrend.push(avgMTTR.toFixed(2));
        });

        // 1. Gráfico Confiabilidad
        if(chartInstances['trend_conf']) chartInstances['trend_conf'].destroy();
        chartInstances['trend_conf'] = new Chart(document.getElementById('chart_trend_conf'), {
            type: 'line',
            data: {
                labels: weeks.map(w => 'Semana ' + w),
                datasets: [{
                    label: 'Confiabilidad Promedio (%)',
                    data: confTrend,
                    borderColor: accentColor,
                    backgroundColor: accentColor + '20',
                    borderWidth: 3,
                    pointRadius: 4,
                    fill: true,
                    tension: 0.3
                }]
            },
            options: {
                maintainAspectRatio: false,
                scales: { y: { min: 0, max: 100, grid: { borderDash: [4, 4] } }, x: { grid: { display: false } } },
                plugins: { legend: { display: false } }
            }
        });

        // 2. Gráfico MTBF y MTTR
        if(chartInstances['trend_mtbf']) chartInstances['trend_mtbf'].destroy();
        chartInstances['trend_mtbf'] = new Chart(document.getElementById('chart_trend_mtbf'), {
            type: 'line',
            data: {
                labels: weeks.map(w => 'Semana ' + w),
                datasets: [
                    {
                        label: 'MTBF (Hrs)',
                        data: mtbfTrend,
                        borderColor: secColor,
                        borderWidth: 3,
                        tension: 0.3,
                        yAxisID: 'y'
                    },
                    {
                        label: 'MTTR (Hrs)',
                        data: mttrTrend,
                        borderColor: '#f59e0b',
                        borderWidth: 3,
                        borderDash: [5, 5],
                        tension: 0.3,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                maintainAspectRatio: false,
                scales: {
                    x: { grid: { display: false } },
                    y: { type: 'linear', display: true, position: 'left', title: {display:true, text:'MTBF (h)'} },
                    y1: { type: 'linear', display: true, position: 'right', grid: { drawOnChartArea: false }, title: {display:true, text:'MTTR (h)'} }
                },
                plugins: { legend: { position: 'top' } }
            }
        });
    }

    function renderTable() {
        const search = document.getElementById('search_input').value.toLowerCase();
        const tbody = document.getElementById('table_body');
        tbody.innerHTML = '';

        let eqMap = {};
        currentData.forEach(d => {
            let key = d.planta + "|" + d.linea + "|" + d.equipo;
            if(!eqMap[key]) {
                eqMap[key] = { p: d.planta, l: d.linea, e: d.equipo, det: 0, tpop: 0, tpoo: 0 };
            }
            eqMap[key].det += d.detenciones;
            eqMap[key].tpop += d.tpo_perdido_hrs;
            eqMap[key].tpoo += d.tpo_operativo_hrs;
        });

        let tableData = Object.values(eqMap).map(d => {
            let mtbf = d.det > 0 ? (d.tpoo / d.det) : 0;
            let mttr = d.det > 0 ? (d.tpop / d.det) : 0;
            let conf = mtbf > 0 ? Math.exp(-120 / mtbf) * 100 : (d.det === 0 ? 100 : 0);
            return { ...d, mtbf, mttr, conf };
        });

        if(search) {
            tableData = tableData.filter(d => `${d.p} ${d.l} ${d.e}`.toLowerCase().includes(search));
        }

        tableData.sort((a,b) => a.conf - b.conf);

        tableData.forEach(d => {
            let badgeClass = d.conf >= 80 ? 'b-ok' : (d.conf >= 50 ? 'b-warn' : 'b-danger');
            let tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${d.p}</td>
                <td>${d.l}</td>
                <td style="font-weight:700;">${d.e}</td>
                <td>${d.det}</td>
                <td>${d.tpop.toFixed(2)}</td>
                <td>${d.mtbf.toFixed(1)}</td>
                <td>${d.mttr.toFixed(2)}</td>
                <td><span class="badge ${badgeClass}">${d.conf.toFixed(1)}%</span></td>
            `;
            tbody.appendChild(tr);
        });
    }

    let sortAsc = true;
    let lastCol = -1;
    function sortTable(colIdx) {
        const table = document.getElementById("data_table");
        const tbody = table.querySelector("tbody");
        const rows = Array.from(tbody.querySelectorAll("tr"));
        
        sortAsc = (lastCol === colIdx) ? !sortAsc : true;
        lastCol = colIdx;

        rows.sort((a, b) => {
            let valA = a.cells[colIdx].innerText.replace('%','').trim();
            let valB = b.cells[colIdx].innerText.replace('%','').trim();
            
            let numA = parseFloat(valA);
            let numB = parseFloat(valB);
            
            if(!isNaN(numA) && !isNaN(numB)) {
                return sortAsc ? numA - numB : numB - numA;
            }
            return sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        });
        
        tbody.innerHTML = '';
        rows.forEach(r => tbody.appendChild(r));
    }

    function descargarExcel() {
        if(currentData.length === 0) return alert("No hay datos para exportar");
        const ws = XLSX.utils.json_to_sheet(currentData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Base_Confiabilidad");
        XLSX.writeFile(wb, `Dataset_Confiabilidad.xlsx`);
    }

    window.onload = () => {
        toggleTheme();
    };
    </script>
</body></html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    full_html = full_html.replace("__FECHA_ACTUAL__", fecha_actual)
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: 
        f.write(full_html)

if __name__ == "__main__":
    db = procesar_datos_confiabilidad()
    if db:
        generar_html_moderno(db)
