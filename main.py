import os
import json
import math
import shutil
import re
import pandas as pd
import gdown
from datetime import datetime
from zoneinfo import ZoneInfo

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
    if os.path.exists(DATA_DIR): shutil.rmtree(DATA_DIR)
    os.makedirs(DATA_DIR, exist_ok=True)
    try:
        gdown.download_folder(url=DRIVE_FOLDER_URL, output=DATA_DIR, quiet=False, use_cookies=False)
    except Exception as e:
        print(f"❌ Error al descargar: {e}")

    archivos = [f for f in os.listdir(DATA_DIR) if f.endswith('.xlsx') and not f.startswith('~')]
    datos_equipos = []
    datos_lineas = []
    datos_acciones = {}

    archivo_acciones = next((f for f in archivos if 'acciones' in f.lower()), None)
    if archivo_acciones:
        try:
            df_acc = pd.read_excel(os.path.join(DATA_DIR, archivo_acciones))
            df_acc.columns = [str(c).lower().strip() for c in df_acc.columns]
            col_top = next((c for c in df_acc.columns if 'top' in c), None)
            if col_top:
                for _, row in df_acc.iterrows():
                    top_rank = pd.to_numeric(row[col_top], errors='coerce')
                    if pd.isna(top_rank): continue
                    for col in df_acc.columns:
                        if col != col_top:
                            if col not in datos_acciones: datos_acciones[col] = {}
                            datos_acciones[col][int(top_rank)] = str(row[col]).strip() if pd.notna(row[col]) else "N/A"
        except Exception as e: print(f"❌ Error acciones: {e}")

    for archivo_nombre in archivos:
        if 'acciones' in archivo_nombre.lower(): continue
        try:
            excel = pd.ExcelFile(os.path.join(DATA_DIR, archivo_nombre))
            hoja_det = next((h for h in excel.sheet_names if h.endswith('_Detenciones_FEM')), None)
            hoja_tpo = next((h for h in excel.sheet_names if h.endswith('_Tiempos_Planificados')), None)
            if not hoja_det or not hoja_tpo: continue

            df_det = pd.read_excel(excel, sheet_name=hoja_det)
            df_tpo = pd.read_excel(excel, sheet_name=hoja_tpo)
            planta_nombre = re.sub(r'(?i)confiabilidad', '', archivo_nombre).replace('.xlsx', '').strip()
            super_planta = "Carnes" if "carne" in planta_nombre.lower() else "Masas"

            # ETL básico
            col_equipo = buscar_columna_equipo(df_det, planta_nombre)
            col_semana_det = buscar_columna_semana(df_det)
            col_linea_det = buscar_columna_linea(df_det)
            col_tpo_det = buscar_tiempo_detencion_hr(df_det, super_planta)
            
            df_det = df_det.dropna(subset=[col_equipo, col_semana_det, col_linea_det])
            df_det['Hrs_Perdidas'] = pd.to_numeric(df_det[col_tpo_det], errors='coerce').fillna(0)
            df_det['Linea_Clean'] = df_det[col_linea_det].astype(str).str.upper().str.strip()
            df_det['Equipo_Clean'] = df_det[col_equipo].astype(str).str.strip().str.title()
            df_det['Semana_Clean'] = limpiar_semana(df_det[col_semana_det])
            df_det = filtrar_semanas(df_det)

            agrup_det_linea = df_det.groupby(['Linea_Clean', 'Semana_Clean']).agg(tpo_perdido_linea=('Hrs_Perdidas', 'sum')).reset_index()

            # Tiempos
            col_semana_tpo = buscar_columna_semana(df_tpo)
            col_linea_tpo = buscar_columna_linea(df_tpo)
            col_tpo_oper = buscar_columna_tiempo_oper(df_tpo)
            col_tpo_plan = buscar_columna_tiempo_plan(df_tpo, super_planta)

            df_tpo = df_tpo.dropna(subset=[col_linea_tpo, col_semana_tpo])
            df_tpo['Hrs_Oper'] = pd.to_numeric(df_tpo[col_tpo_oper], errors='coerce').fillna(0) if col_tpo_oper else 0
            df_tpo['Hrs_Plan'] = pd.to_numeric(df_tpo[col_tpo_plan], errors='coerce').fillna(0) if col_tpo_plan else 0
            df_tpo['Linea_Clean'] = df_tpo[col_linea_tpo].astype(str).str.upper().str.strip()
            df_tpo['Semana_Clean'] = limpiar_semana(df_tpo[col_semana_tpo])
            df_tpo = filtrar_semanas(df_tpo)
            agrup_tpo_linea = df_tpo.groupby(['Linea_Clean', 'Semana_Clean']).agg(tpo_operativo_linea=('Hrs_Oper', 'sum'), tpo_plan_linea=('Hrs_Plan', 'sum')).reset_index()

            # Merge y reconciliación
            linea_merged = pd.merge(agrup_tpo_linea, agrup_det_linea, on=['Linea_Clean', 'Semana_Clean'], how='outer').fillna(0)
            if super_planta == 'Carnes': linea_merged['tpo_operativo_linea'] = (linea_merged['tpo_plan_linea'] - linea_merged['tpo_perdido_linea']).clip(lower=0)
            
            for _, row in linea_merged.iterrows():
                datos_lineas.append({"super_planta": super_planta, "planta": planta_nombre, "linea": row['Linea_Clean'], "semana": int(row['Semana_Clean']), "tpo_operativo_linea": float(row['tpo_operativo_linea']), "tpo_plan_linea": float(row['tpo_plan_linea'])})
            for _, row in df_det.iterrows():
                datos_equipos.append({"super_planta": super_planta, "planta": planta_nombre, "linea": row['Linea_Clean'], "equipo": row['Equipo_Clean'], "semana": int(row['Semana_Clean']), "detenciones": 1, "tpo_perdido_eq": float(row['Hrs_Perdidas']), "fecha": str(row.get('Fecha', 'N/A'))[:10], "componente": str(row.get('Componente', 'N/A')), "tipo": str(row.get('Tipo Detención', 'N/A'))})
        except Exception as e: print(f"❌ Error procesando {archivo_nombre}: {e}")

    return {"equipos": datos_equipos, "lineas": datos_lineas, "acciones": datos_acciones}

def generar_html_moderno(db_json):
    fecha_actual = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")

    # Esta parte se encarga de la tabla de desglose con las nuevas columnas
    html_template = """<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><title>Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;700;900&display=swap');
body{background:#F4F6FA;font-family:'DM Sans',sans-serif;margin:0;padding:20px;}
.grid-table { width: 100%; border-collapse: collapse; }
.grid-th { display:grid; grid-template-columns: 2fr 1.5fr 1fr 1fr 1fr 1fr 1fr 1fr; gap:5px; background:#0D2C54; color:#fff; padding:10px; font-size:9px; font-weight:700; }
.grid-tr { display:grid; grid-template-columns: 2fr 1.5fr 1fr 1fr 1fr 1fr 1fr 1fr; gap:5px; background:#fff; padding:10px; border-bottom:1px solid #EAF0F8; font-size:10px; }
.stat-val{font-weight:700;}
</style></head>
<body>
    <div id="desglose_container"></div>
<script>
const db = __DB_JSON_DATA__;
function renderDesglose(planta) {
    const el = document.getElementById('desglose_container');
    // Filtramos datos de la planta, limpiando nombres si es necesario
    const datos = db.equipos.filter(e => e.planta === planta);
    
    let rows = datos.map(d => `
        <div class="grid-tr">
            <div>${d.equipo}</div>
            <div>${d.linea}</div>
            <div class="stat-val">${d.tpo_perdido_eq.toFixed(1)}</div>
            <div class="stat-val">MTBF</div>
            <div class="stat-val">MTTR</div>
            <div class="stat-val">Conf</div>
            <div class="stat-val">Mant</div>
            <div class="stat-val">${(100 - Math.exp(-120/(d.tpo_perdido_eq+1))*100).toFixed(1)}%</div>
        </div>`).join('');

    el.innerHTML = `
    <div class="fila-planta-v3">
        <h3>Detalle: ${planta}</h3>
        <div class="grid-th">
            <div>EQUIPO</div><div>LÍNEA</div>
            <div>PROM. TPO PERDIDO (Hrs)</div>
            <div>MTBF</div><div>MTTR</div>
            <div>CONFIAB.</div><div>MANTEN.</div><div>PROB. FALLA</div>
        </div>
        ${rows}
    </div>`;
}
// Llamar a renderDesglose cuando sea necesario...
</script></body></html>"""
    
    # Asegúrate de limpiar los nombres al pasar el JSON al HTML
    for item in db_json['equipos']:
        item['planta'] = re.sub(r'(?i)confiabilidad', '', item['planta']).replace('.xlsx', '').strip()

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_template.replace("__DB_JSON_DATA__", json.dumps(db_json)))
