import os
import json
import math
import shutil
import re
import pandas as pd
import gdown
from datetime import datetime
from zoneinfo import ZoneInfo

# ... (Se mantiene igual la configuración y los buscadores anteriores) ...

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
