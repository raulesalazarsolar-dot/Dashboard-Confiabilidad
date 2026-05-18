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
# Aquí está tu link público de la carpeta de Drive
DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1xXeea_F6HTsI-Wfj7HP2KdCnzGtPIUQq?usp=sharing"

DATA_DIR = "./data"
OUTPUT_HTML = "index.html"

# PARÁMETROS DEL MODELO
TIEMPO_MISION_HRS = 120
TIEMPO_IDEAL_REP_HRS = 1

# ==========================================
# 2. DESCARGA DESDE GOOGLE DRIVE
# ==========================================
def descargar_desde_drive():
    print(f"📥 Conectando a Google Drive para descargar Excels...")
    
    # Limpiamos la carpeta si existe para asegurar que tenemos la última versión
    if os.path.exists(DATA_DIR):
        shutil.rmtree(DATA_DIR)
    os.makedirs(DATA_DIR, exist_ok=True)
    
    try:
        # Descarga el contenido de la carpeta pública
        gdown.download_folder(url=DRIVE_FOLDER_URL, output=DATA_DIR, quiet=False, use_cookies=False)
        print("✅ Descarga desde Google Drive completada.")
    except Exception as e:
        print(f"❌ Error al descargar de Drive: {e}")

# ==========================================
# 3. EXTRACCIÓN Y TRANSFORMACIÓN (ETL)
# ==========================================
def procesar_datos_confiabilidad():
    # Primero descargamos los archivos
    descargar_desde_drive()
    
    print("🚀 INICIANDO PROCESAMIENTO DE EXCEL...")
    
    if not os.path.exists(DATA_DIR):
        print(f"❌ Error: No se encontró la carpeta '{DATA_DIR}'.")
        return {}

    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]
    print(f"   ✅ Se encontraron {len(archivos)} archivos Excel para procesar.")

    datos_equipos = []
    
    for archivo_nombre in archivos:
        ruta_completa = os.path.join(DATA_DIR, archivo_nombre)
        print(f"   📊 Procesando archivo: {archivo_nombre}")
        
        try:
            excel = pd.ExcelFile(ruta_completa)
            hojas = excel.sheet_names
            
            # Buscar las hojas dinámicamente por sufijo
            hoja_det = next((h for h in hojas if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in hojas if h.endswith('_Tiempos_Planificados')), None)
            
            if not hoja_det or not hoja_tpo:
                print(f"      ⚠️ No se encontraron las pestañas base en {archivo_nombre}. Saltando...")
                continue
                
            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)
            
            planta_nombre = archivo_nombre.replace("Confiabilidad ", "").replace(".xlsx", "").strip()
            
            # ----------------------------------------------------
            # LIMPIEZA Y AGRUPACIÓN - DETENCIONES
            # ----------------------------------------------------
            col_semana = 'N° Semana' if 'N° Semana' in df_det.columns else 'Semana'
            col_tpo_det = 'Tpo detención total [min]' if 'Tpo detención total [min]' in df_det.columns else 'Tiempo Det total [min]'
            
            df_det = df_det.dropna(subset=['Equipo', col_semana, 'Línea'])
            df_det['Hrs_Perdidas'] = df_det[col_tpo_det] / 60.0
            
            agrup_det = df_det.groupby(['Línea', col_semana, 'Equipo']).agg(
                detenciones=('Equipo', 'count'),
                tpo_perdido_hrs=('Hrs_Perdidas', 'sum')
            ).reset_index()
            
            # ----------------------------------------------------
            # LIMPIEZA Y AGRUPACIÓN - TIEMPOS PLANIFICADOS (OPERACIONAL)
            # ----------------------------------------------------
            col_sem_tpo = 'N° Semana' if 'N° Semana' in df_tpo.columns else 'Semana'
            
            df_tpo = df_tpo.dropna(subset=['Línea', col_sem_tpo])
            agrup_tpo = df_tpo.groupby(['Línea', col_sem_tpo]).agg(
                tpo_operativo_linea_hrs=('Tpo Operativo [hr]', 'sum')
            ).reset_index()
            
            # ----------------------------------------------------
            # MERGE (CRUCE) DE DATOS
            # ----------------------------------------------------
            df_final = pd.merge(
                agrup_det, agrup_tpo, 
                left_on=['Línea', col_semana], 
                right_on=['Línea', col_sem_tpo], 
                how='left'
            )
            
            for _, row in df_final.iterrows():
                datos_equipos.append({
                    "planta": planta_nombre,
                    "linea": str(row['Línea']),
                    "equipo": str(row['Equipo']),
                    "semana": str(int(row[col_semana])),
                    "detenciones": int(row['detenciones']),
                    "tpo_perdido_hrs": float(row['tpo_perdido_hrs']),
                    "tpo_operativo_linea_hrs": float(row['tpo_operativo_linea_hrs']) if pd.notna(row['tpo_operativo_linea_hrs']) else 0.0
                })
                
        except Exception as e:
            print(f"      ❌ Error procesando {archivo_nombre}: {e}")

    # ==========================================
    # CÁLCULO MATEMÁTICO DE CONFIABILIDAD
    # ==========================================
    print("   ⏳ Calculando MTBF, MTTR, Confiabilidad y Mantenibilidad...")
    db_json = {}
    
    for idx, row in enumerate(datos_equipos):
        det = row['detenciones']
        tpo_perd = row['tpo_perdido_hrs']
        tpo_op = row['tpo_operativo_linea_hrs']
        
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
            
        prob_falla = 100 - conf
        
        key_id = f"EQ_{idx+1}"
        db_json[key_id] = {
            "key_id": key_id,
            "planta": row['planta'],
            "linea": row['linea'],
            "equipo": row['equipo'],
            "semana": row['semana'],
            "detenciones": det,
            "tpo_perdido_hrs": round(tpo_perd, 2),
            "tpo_operativo_hrs": round(tpo_op, 2),
            "mtbf": round(mtbf, 2) if det > 0 else 'N/A',
            "mttr": round(mttr, 3) if det > 0 else 'N/A',
            "confiabilidad": round(conf, 2),
            "mantenibilidad": round(mant, 2),
            "prob_falla": round(prob_falla, 2)
        }
        
    print(f"   ✅ Se procesaron {len(db_json)} registros consolidados exitosamente.")
    return db_json

# ==========================================
# 4. GENERADOR HTML DASHBOARD
# ==========================================
def generar_html_moderno(db_json):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    
    html_template = """<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Confiabilidad Walmart</title>
    <link rel="icon" type="image/x-icon" href="https://www.walmart.com/favicon.ico">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root { --primary: #0f172a; --secondary: #334155; --accent: #0ea5e9; --bg: #f8fafc; --border: #e2e8f0; --text: #1e293b; --muted: #64748b; --success: #10b981; --warn: #f59e0b; --danger: #ef4444; }
        * { box-sizing: border-box; outline: none; font-family: 'Segoe UI', system-ui, sans-serif; }
        body { background: transparent; color: var(--text); margin: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        
        .top-bar { background: var(--primary); color: white; padding: 0 20px; height: 60px; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; z-index: 10; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .brand { flex: 1; }
        .brand h2 { margin: 0; font-size: 1.2rem; display:flex; align-items:center; gap: 8px; } 
        .brand span { opacity: 0.7; font-weight: 300; font-size: 0.95rem; }

        .tabs-container { background: white; border-bottom: 1px solid var(--border); padding: 0 20px; flex-shrink: 0; display:flex; justify-content: space-between; align-items: center; z-index: 5; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .tabs-nav { display: flex; gap: 15px; }
        .tab-btn { background: none; border: none; padding: 15px 5px; font-weight: 600; color: var(--muted); cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 0.95rem; }
        .tab-btn:hover { color: var(--accent); } .tab-btn.active { color: var(--accent); border-bottom-color: var(--accent); }
        
        .app-layout { display: flex; height: calc(100vh - 110px); width: 100%; overflow: hidden; }
        
        .col-filters { width: 280px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; z-index: 5; }
        .filters-header { padding: 15px 20px; border-bottom: 1px solid var(--border); font-weight: 700; color: var(--primary); font-size: 0.9rem; text-transform: uppercase; background: #f8fafc; display: flex; justify-content: space-between; align-items: center; }
        .filters-body { flex: 1; overflow-y: auto; padding: 20px; } 
        
        .f-group { margin-bottom: 15px; }
        .f-group label { font-size: 0.75rem; font-weight: 700; color: var(--muted); display: block; margin-bottom: 6px; text-transform: uppercase; }
        select, input[type="text"] { width: 100%; padding: 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.85rem; color: var(--text); }
        
        .btn-action { background: white; border: 1px solid var(--success); color: var(--success); padding: 8px 15px; border-radius: 6px; cursor: pointer; font-weight: 700; transition: 0.2s; display: flex; align-items: center; gap: 8px; }
        .btn-action:hover { background: var(--success); color: white; }

        .col-list { width: 400px; background: #fff; border-right: 1px solid var(--border); display: flex; flex-direction: column; flex-shrink: 0; }
        .list-header { padding: 20px; border-bottom: 1px solid var(--border); background: #f8fafc; }
        .list-scroll-area { flex: 1; overflow-y: auto; }
        
        .list-item { padding: 15px 20px; border-bottom: 1px solid var(--border); cursor: pointer; transition: 0.2s; border-left: 4px solid transparent; }
        .list-item:hover { background: #f8fafc; } .list-item.selected { background: #f0f9ff; border-left-color: var(--accent); }
        .li-top { display: flex; justify-content: space-between; font-size: 0.75rem; color: var(--muted); font-weight: 600; margin-bottom: 5px; }
        .li-title { font-weight: 700; font-size: 1rem; color: var(--primary); margin-bottom: 8px; }
        .li-btm { display: flex; justify-content: space-between; font-size: 0.8rem; }
        .tag-conf { font-weight: 700; padding: 4px 8px; border-radius: 4px; }
        
        .col-detail { flex: 1; overflow-y: auto; padding: 40px; }
        .detail-card { background: white; border-radius: 12px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border: 1px solid var(--border); padding: 30px; max-width: 900px; margin: 0 auto; }
        
        .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-top: 20px; }
        .metric-box { background: var(--bg); padding: 20px; border-radius: 8px; border: 1px solid var(--border); text-align: center; }
        .metric-box small { display: block; color: var(--muted); font-weight: 700; text-transform: uppercase; font-size: 0.75rem; margin-bottom: 5px; }
        .metric-box strong { font-size: 1.8rem; color: var(--primary); }

        .graficos-layout { flex: 1; padding: 30px; display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 25px; overflow-y: auto; align-content: start; }
        .chart-card { background: white; padding: 25px; border-radius: 12px; border: 1px solid var(--border); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); height: 400px; display: flex; flex-direction: column; }
        .chart-card.wide { grid-column: 1 / -1; height: 450px; }
        .chart-title { font-size: 1rem; font-weight: 700; color: var(--secondary); margin-bottom: 15px; text-transform: uppercase; text-align: center; }
        .canvas-container { position: relative; flex: 1; min-height: 0; }
        
        .footer-date { font-size: 0.7rem; color: var(--muted); text-align: center; margin-top: 20px; font-weight: 600; }
    </style>
</head>
<body>
    <div class="top-bar">
        <div class="brand"><h2>📈 Dashboard Ingeniería de Mantenimiento <span>| Confiabilidad de Activos</span></h2></div>
        <div><img src="https://upload.wikimedia.org/wikipedia/commons/b/b1/Walmart_logo_%282008%29.svg" alt="Walmart Logo" style="height: 30px; filter: brightness(0) invert(1);"></div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs-nav">
            <button class="tab-btn active" onclick="setView('list', this)" id="btn_tab_list">📋 Matriz de Equipos</button>
            <button class="tab-btn" onclick="setView('charts', this)">📊 Análisis Gráfico</button>
        </div>
        <button onclick="descargarExcel()" class="btn-action">📊 Exportar Dataset</button>
    </div>
    
    <div class="app-layout">
        <div class="col-filters">
            <div class="filters-header"><span>🔍 Parámetros</span></div>
            <div class="filters-body" id="filters_dynamic"></div>
            <div class="footer-date">Última actualización:<br>__FECHA_ACTUAL__</div>
        </div>

        <div id="view_list" style="display:flex; flex:1; overflow:hidden;">
            <div class="col-list">
                <div class="list-header">
                    <input type="text" id="search_input" placeholder="🔍 Buscar equipo o línea..." onkeyup="applyFilters()">
                </div>
                <div id="list_container" class="list-scroll-area"></div>
            </div>
            <div class="col-detail">
                <div id="empty_state" style="text-align:center; padding-top:100px; color:var(--muted);">
                    <h1 style="font-size:4rem; margin:0;">⚙️</h1>
                    <h3>Selecciona un equipo de la lista</h3>
                </div>
                <div id="detail_view" class="detail-card" style="display:none">
                    <div style="border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 20px;">
                        <span id="d_semana" style="float:right; font-weight:700; color:var(--accent);"></span>
                        <h2 id="d_equipo" style="margin:0; font-size:1.8rem; color:var(--primary);"></h2>
                        <p id="d_planta_linea" style="margin:5px 0 0; color:var(--muted); font-weight:600;"></p>
                    </div>
                    
                    <div class="metric-grid">
                        <div class="metric-box"><small>Confiabilidad (R)</small><strong id="d_conf" style="color:var(--success);"></strong></div>
                        <div class="metric-box"><small>Mantenibilidad (M)</small><strong id="d_mant" style="color:var(--accent);"></strong></div>
                        <div class="metric-box"><small>Prob. de Falla (Pf)</small><strong id="d_prob" style="color:var(--danger);"></strong></div>
                    </div>
                    
                    <div class="metric-grid" style="margin-top:15px;">
                        <div class="metric-box" style="background:#fff;"><small>Cant. Detenciones (N)</small><strong id="d_det"></strong></div>
                        <div class="metric-box" style="background:#fff;"><small>Tpo. Perdido (Hrs)</small><strong id="d_tpop"></strong></div>
                        <div class="metric-box" style="background:#fff;"><small>Tpo. Operativo (Hrs)</small><strong id="d_tpoo"></strong></div>
                    </div>

                    <div class="metric-grid" style="margin-top:15px; grid-template-columns: 1fr 1fr;">
                        <div class="metric-box"><small>MTBF (Tpo. Medio Entre Fallas)</small><strong id="d_mtbf"></strong></div>
                        <div class="metric-box"><small>MTTR (Tpo. Medio Reparación)</small><strong id="d_mttr"></strong></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="view_charts" class="graficos-layout" style="display:none;">
            <div class="chart-card"><div class="chart-title">Confiabilidad Promedio por Planta</div><div class="canvas-container"><canvas id="chart1"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Distribución de Tiempos Perdidos (Hrs)</div><div class="canvas-container"><canvas id="chart2"></canvas></div></div>
            <div class="chart-card wide"><div class="chart-title">Top 10 Equipos con Mayor Probabilidad de Falla</div><div class="canvas-container"><canvas id="chart3"></canvas></div></div>
        </div>
    </div>

    <script>
    Chart.register(ChartDataLabels);
    Chart.defaults.plugins.datalabels.display = false;
    Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";

    const db = __DB_JSON_DATA__;
    const records = Object.values(db);
    let currentData = [];

    function buildFilters() {
        const createSelect = (id, label, options) => {
            let sel = `<div class="f-group"><label>${label}</label><select id="${id}" onchange="applyFilters()"><option value="ALL">Todos</option>`;
            options.sort().forEach(o => sel += `<option value="${o}">${o}</option>`);
            return sel + `</select></div>`;
        };
        
        let html = '';
        html += createSelect('f_semana', '📆 Semana', [...new Set(records.map(x=>x.semana))]);
        html += createSelect('f_planta', '🏢 Planta', [...new Set(records.map(x=>x.planta))]);
        html += createSelect('f_linea', '🏭 Línea', [...new Set(records.map(x=>x.linea))]);
        document.getElementById('filters_dynamic').innerHTML = html;
        
        let semanas = [...new Set(records.map(x=>parseInt(x.semana)))].filter(x => !isNaN(x)).sort((a,b)=>b-a);
        if(semanas.length > 0) {
            document.getElementById('f_semana').value = semanas[0].toString();
        }
    }

    function applyFilters() {
        const fSem = document.getElementById('f_semana').value;
        const fPla = document.getElementById('f_planta').value;
        const fLin = document.getElementById('f_linea').value;
        const search = document.getElementById('search_input').value.toLowerCase();

        currentData = records.filter(d => {
            if(fSem !== 'ALL' && d.semana !== fSem) return false;
            if(fPla !== 'ALL' && d.planta !== fPla) return false;
            if(fLin !== 'ALL' && d.linea !== fLin) return false;
            if(search && !`${d.equipo} ${d.linea}`.toLowerCase().includes(search)) return false;
            return true;
        });

        renderList();
        if(document.getElementById('view_charts').style.display !== 'none') drawCharts();
    }

    function renderList() {
        const cont = document.getElementById('list_container');
        cont.innerHTML = '';
        currentData.forEach(d => {
            let item = document.createElement('div');
            item.className = 'list-item';
            let colorConf = d.confiabilidad > 80 ? '#10b981' : (d.confiabilidad > 50 ? '#f59e0b' : '#ef4444');
            item.innerHTML = `
                <div class="li-top"><span>${d.planta} | ${d.linea}</span><span>Sem: ${d.semana}</span></div>
                <div class="li-title">${d.equipo}</div>
                <div class="li-btm">
                    <span>Fallas: <b>${d.detenciones}</b></span>
                    <span class="tag-conf" style="background:${colorConf}20; color:${colorConf};">R: ${d.confiabilidad}%</span>
                </div>
            `;
            item.onclick = () => {
                document.querySelectorAll('.list-item').forEach(i=>i.classList.remove('selected'));
                item.classList.add('selected');
                showDetail(d);
            };
            cont.appendChild(item);
        });
    }

    function showDetail(d) {
        document.getElementById('empty_state').style.display = 'none';
        document.getElementById('detail_view').style.display = 'block';
        
        document.getElementById('d_equipo').innerText = d.equipo;
        document.getElementById('d_planta_linea').innerText = `${d.planta} / ${d.linea}`;
        document.getElementById('d_semana').innerText = `Semana ${d.semana}`;
        
        document.getElementById('d_conf').innerText = `${d.confiabilidad}%`;
        document.getElementById('d_mant').innerText = `${d.mantenibilidad}%`;
        document.getElementById('d_prob').innerText = `${d.prob_falla}%`;
        
        document.getElementById('d_det').innerText = d.detenciones;
        document.getElementById('d_tpop').innerText = d.tpo_perdido_hrs;
        document.getElementById('d_tpoo').innerText = d.tpo_operativo_hrs;
        
        document.getElementById('d_mtbf').innerText = d.mtbf;
        document.getElementById('d_mttr').innerText = d.mttr;
    }

    function setView(view, btn) {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('view_list').style.display = view === 'list' ? 'flex' : 'none';
        document.getElementById('view_charts').style.display = view === 'charts' ? 'grid' : 'none';
        if(view === 'charts') drawCharts();
    }

    function descargarExcel() {
        if(currentData.length === 0) return alert("No hay datos para exportar");
        const ws = XLSX.utils.json_to_sheet(currentData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Base_Confiabilidad");
        XLSX.writeFile(wb, `Reporte_Confiabilidad.xlsx`);
    }

    function drawCharts() {
        if(currentData.length === 0) return;
        
        const opts = { maintainAspectRatio: false, responsive: true, plugins: { legend: { position: 'bottom' } } };
        
        const plantas = [...new Set(currentData.map(d=>d.planta))];
        const confProm = plantas.map(p => {
            let items = currentData.filter(d=>d.planta === p);
            return items.length > 0 ? items.reduce((sum, d)=>sum+d.confiabilidad, 0) / items.length : 0;
        });

        getFreshCanvas('chart1');
        new Chart(document.getElementById('chart1'), {
            type: 'bar',
            data: { labels: plantas, datasets: [{ label: '% Confiabilidad Promedio', data: confProm, backgroundColor: '#0ea5e9', borderRadius: 4 }] },
            options: { ...opts, scales: { y: { max: 100 } } }
        });

        const tpoPerdido = plantas.map(p => {
            let items = currentData.filter(d=>d.planta === p);
            return items.reduce((sum, d)=>sum+d.tpo_perdido_hrs, 0);
        });

        getFreshCanvas('chart2');
        new Chart(document.getElementById('chart2'), {
            type: 'doughnut',
            data: { labels: plantas, datasets: [{ data: tpoPerdido, backgroundColor: ['#ef4444', '#f59e0b', '#10b981', '#8b5cf6'], borderWidth: 2 }] },
            options: opts
        });

        let topFallas = [...currentData].sort((a,b)=> b.prob_falla - a.prob_falla).slice(0,10);
        getFreshCanvas('chart3');
        new Chart(document.getElementById('chart3'), {
            type: 'bar',
            data: { 
                labels: topFallas.map(d=> d.equipo.length > 20 ? d.equipo.substring(0,20)+'...' : d.equipo), 
                datasets: [{ label: 'Probabilidad de Falla (%)', data: topFallas.map(d=>d.prob_falla), backgroundColor: '#ef4444', borderRadius: 4 }] 
            },
            options: { ...opts, indexAxis: 'y', scales: { x: { max: 100 } } }
        });
    }

    function getFreshCanvas(id) {
        let old = document.getElementById(id);
        if(!old) return;
        let p = old.parentElement;
        p.innerHTML = `<canvas id="${id}"></canvas>`;
    }

    window.onload = () => {
        buildFilters();
        applyFilters();
    };
    </script>
</body></html>"""

    full_html = html_template.replace("__DB_JSON_DATA__", json.dumps(db_json))
    full_html = full_html.replace("__FECHA_ACTUAL__", fecha_actual)
    
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: 
        f.write(full_html)
        
    print(f"\n✅ REPORTE GENERADO CON ÉXITO EN: {OUTPUT_HTML}")

if __name__ == "__main__":
    db = procesar_datos_confiabilidad()
    if db:
        generar_html_moderno(db)
