#!/usr/bin/env python3
"""
genera_cronologico_desguaces.py
Genera el Archivo Cronológico de Retiradas de Residuos (Desguaces VFU)
a partir del Excel de Ingurunet.

Uso:
    python genera_cronologico_desguaces.py <ruta_excel> [--output <ruta_html>]

El HTML resultante se abre en el navegador y se imprime/exporta a PDF con Ctrl+P.
"""

import argparse
import base64
import os
import re
import sys
from datetime import datetime, date

import openpyxl

# ─────────────────────────────────────────────────────────
# Constantes Alvale
# ─────────────────────────────────────────────────────────
ALVALE_NOMBRE  = "Alvale Consulting Ingenieros, S.L."
ALVALE_DIR     = "Ribera de Axpe, 11, L-311-B. 48950 ERANDIO (Bizkaia)"
ALVALE_TEL     = "Tl.: 944 971 050"
ALVALE_WEB_URL = "https://alvaleconsulting.com"
ALVALE_WEB_TXT = "alvaleconsulting.com"
ALVALE_LI_URL  = "https://www.linkedin.com/company/alvale-consulting-ingenieros/"

ROWS_PER_PAGE  = 18   # filas por página interior (holgura suficiente)

# ─────────────────────────────────────────────────────────
# Tablas de abreviaciones
# ─────────────────────────────────────────────────────────
PROV_ABBR = {
    "álava":"ARA","araba":"ARA","araba/álava":"ARA",
    "albacete":"ALB","alicante":"ALI","alacant":"ALI","almería":"ALM",
    "asturias":"AST","ávila":"AVI","badajoz":"BAD","barcelona":"BCN",
    "bizkaia":"BIZ","vizcaya":"BIZ","burgos":"BUR",
    "cáceres":"CAC","cádiz":"CAD","cantabria":"CTB",
    "castellón":"CAS","castelló":"CAS","ciudad real":"CRE","córdoba":"COR",
    "a coruña":"ACO","coruña, a":"ACO","la coruña":"ACO",
    "cuenca":"CUE","girona":"GIR","gerona":"GIR","granada":"GRA",
    "guadalajara":"GUA","gipuzkoa":"GIP","guipúzcoa":"GIP",
    "huelva":"HUE","huesca":"HUS",
    "illes balears":"BAL","baleares":"BAL","islas baleares":"BAL",
    "jaén":"JAE","león":"LEO","lleida":"LLE","lérida":"LLE","lugo":"LUG",
    "madrid":"MAD","málaga":"MAL","murcia":"MUR",
    "navarra":"NAV","nafarroa":"NAV",
    "ourense":"OUR","orense":"OUR","palencia":"PAL","las palmas":"LPA",
    "pontevedra":"PON","la rioja":"RIO","rioja":"RIO",
    "salamanca":"SAL","santa cruz de tenerife":"TFE","segovia":"SEG",
    "sevilla":"SEV","soria":"SOR","tarragona":"TAR","teruel":"TER",
    "toledo":"TOL","valencia":"VAL","valència":"VAL","valladolid":"VLD",
    "zamora":"ZAM","zaragoza":"ZGZ","ceuta":"CEU","melilla":"MEL",
}

CCAA_ABBR = {
    "andalucía":"AND","andalucia":"AND",
    "aragón":"ARA","aragon":"ARA",
    "asturias":"AST","principado de asturias":"AST",
    "illes balears":"BAL","baleares":"BAL","islas baleares":"BAL",
    "canarias":"CAN","cantabria":"CTB",
    "castilla-la mancha":"CLM","castilla la mancha":"CLM",
    "castilla y león":"CYL","castilla y leon":"CYL",
    "cataluña":"CAT","catalunya":"CAT",
    "comunitat valenciana":"CVA","comunidad valenciana":"CVA",
    "extremadura":"EXT","galicia":"GAL",
    "la rioja":"RIO","rioja":"RIO",
    "madrid":"MAD","comunidad de madrid":"MAD",
    "murcia":"MUR","región de murcia":"MUR",
    "navarra":"NAV","comunidad foral de navarra":"NAV","nafarroa":"NAV",
    "país vasco":"PV","euskadi":"PV","pais vasco":"PV",
    "ceuta":"CEU","melilla":"MEL",
}

PAIS_ABBR = {
    "españa":"ES","espana":"ES","spain":"ES",
    "portugal":"PT","francia":"FR","france":"FR",
    "alemania":"DE","germany":"DE","italia":"IT","italy":"IT",
    "reino unido":"UK","united kingdom":"UK",
    "países bajos":"NL","holanda":"NL","netherlands":"NL",
    "bélgica":"BE","belgium":"BE","suiza":"CH","switzerland":"CH",
    "austria":"AT","polonia":"PL","poland":"PL",
    "rumanía":"RO","rumania":"RO","romania":"RO",
    "estados unidos":"US","usa":"US","united states":"US",
    "marruecos":"MA","morocco":"MA",
}

def _abbr_prov(s):
    if not s: return ""
    key = re.sub(r'^\d+\s*[-–]\s*','', str(s).strip().lower())
    return PROV_ABBR.get(key, str(s).strip()[:5])

def _abbr_ccaa(s):
    if not s: return ""
    key = re.sub(r'^\d+\s*[-–]\s*','', str(s).strip().lower())
    return CCAA_ABBR.get(key, str(s).strip()[:4])

def _abbr_pais(s):
    if not s: return ""
    return PAIS_ABBR.get(str(s).strip().lower(), str(s).strip()[:4])


# ─────────────────────────────────────────────────────────
# Logo
# ─────────────────────────────────────────────────────────
def _logo_tag():
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "alvale_logo.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        return f'<img src="data:image/png;base64,{b64}" alt="Alvale" style="height:25mm;width:auto;display:block;"/>'
    return '<div style="height:25mm;width:50mm;background:#EEEFF7;border-radius:4px;"></div>'


# ─────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────
MESES_ES = {
    "January":"enero","February":"febrero","March":"marzo","April":"abril",
    "May":"mayo","June":"junio","July":"julio","August":"agosto",
    "September":"septiembre","October":"octubre","November":"noviembre","December":"diciembre",
}

def _fmt_fecha_portada(v):
    if isinstance(v, (datetime, date)):
        s = v.strftime("%-d de %B de %Y")
        for en, es in MESES_ES.items():
            s = s.replace(en, es)
        return s
    return str(v) if v else ""

def _fmt_fecha_tabla(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return str(v) if v else ""

def _fmt_peso(t):
    try:
        s = f"{float(t):,.3f}"
        return s.replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return str(t) if t else ""

def _esc(s):
    if s is None: return ""
    return (str(s).replace("&","&amp;").replace("<","&lt;")
                  .replace(">","&gt;").replace('"',"&quot;"))

def _limpiar_municipio(mun):
    if not mun: return ""
    mun = re.sub(r'^\d+[-–]\s*', '', str(mun).strip())
    m = re.match(r'^(.+?)\s*-\s*[A-ZÁÉÍÓÚÜÑ]', mun)
    if m: mun = m.group(1).strip()
    return mun

def _title_case(s):
    if not s: return s
    acr = re.findall(r'\b(?:[A-ZÁÉÍÓÚÜÑ]\.){2,}', s)
    r = s.title()
    for a in acr:
        r = r.replace(a.title(), a)
    return r


# ─────────────────────────────────────────────────────────
# Lectura del Excel
# ─────────────────────────────────────────────────────────
def leer_datos_empresa(wb):
    ws = wb["DATOS EMPRESA"]
    campos = {
        "Razón Social":"razon_social","NIF":"nif","NIMA":"nima",
        "Nº Autorización":"num_aut","Tipo Autorización":"tipo_aut",
        "Dirección":"direccion","Municipio":"municipio",
        "Provincia":"provincia","Año cronológico":"ano",
    }
    datos = {}
    for row in ws.iter_rows(min_row=1, max_row=30, values_only=True):
        if row[0] and str(row[0]).strip() in campos:
            datos[campos[str(row[0]).strip()]] = row[1]
    return datos

def leer_registros(wb):
    ws = wb["REGISTRO SALIDAS"]
    registros = []
    for ri in range(3, ws.max_row + 1):
        fecha = ws.cell(row=ri, column=1).value
        if fecha is None:
            continue
        def cv(c): return ws.cell(row=ri, column=c).value

        peso_raw = cv(5)
        if peso_raw is None or (isinstance(peso_raw, str) and peso_raw.startswith("=")):
            try:   peso_raw = float(cv(4)) / 1000
            except: peso_raw = 0.0
        try:   peso_t = float(peso_raw)
        except: peso_t = 0.0

        registros.append({
            "fecha":       _fmt_fecha_tabla(fecha),
            "ler":         str(cv(2)) if cv(2) else "",
            "descripcion": _title_case(str(cv(3))) if cv(3) else "",
            "peso_t":      peso_t,
            "ndi":         str(cv(6)) if cv(6) else "",
            "nct":         str(cv(7)) if cv(7) else "",
            "dest_razon":  _title_case(str(cv(8))) if cv(8) else "",
            "dest_nif":    str(cv(9)) if cv(9) else "",
            "dest_nima":   str(cv(10)) if cv(10) else "",
            "dest_aut":    str(cv(11)) if cv(11) else "",
            "dest_tipo":   str(cv(12)) if cv(12) else "",
            "municipio":   _limpiar_municipio(cv(13)),
            "provincia":   _abbr_prov(cv(14)),
            "ccaa":        _abbr_ccaa(cv(15)),
            "pais":        _abbr_pais(cv(16)),
            "transp_razon":_title_case(str(cv(17))) if cv(17) else "",
            # col 18 NIF transp  → NO
            # col 19 NIMA transp → NO
            # col 20 Nº inscr    → ELIMINADO según feedback
            "metodo":      str(cv(21)) if cv(21) else "",
            "proceso":     str(cv(22)) if cv(22) else "",
        })
    return registros


# ─────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:ital,wght@0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500;600&family=Poppins:wght@400;600;700&display=swap');

:root {
  --brand-navy:     #2C2E88;
  --brand-green:    #238D3E;
  --brand-purple:   #5D549D;
  --brand-lavender: #8B83BA;
  --navy-50:  #EEEFF7;
  --navy-700: #1B1D58;
  --ink-25:   #FAFAFC;
  --ink-50:   #F4F4F8;
  --ink-100:  #E8E8EF;
  --ink-200:  #D2D2DE;
  --ink-400:  #7E8094;
  --ink-700:  #2A2C3D;
  --font-body: 'IBM Plex Sans', system-ui, sans-serif;
  --font-mono: 'IBM Plex Mono', 'Consolas', monospace;
  --font-brand: 'Poppins', 'IBM Plex Sans', system-ui, sans-serif;
}

*, *::before, *::after { box-sizing: border-box; }
h1,h2,h3 { margin:0; padding:0; font-size:inherit; font-weight:inherit; }
.pdf-bookmark { display:inline; font-size:inherit; font-weight:inherit;
  color:inherit; font-family:inherit; letter-spacing:inherit; }

html, body { margin:0; padding:0; font-family:var(--font-body);
  color:var(--ink-700); background:#d9dae0;
  -webkit-font-smoothing:antialiased; }

/* ── Sheets ── */
.sheet {
  width:297mm; height:210mm;
  background:#fff; margin:18px auto;
  box-shadow:0 2px 10px rgba(0,0,0,.10);
  page-break-after:always;
  position:relative; display:flex; flex-direction:column;
  overflow:hidden;
}
.sheet:last-child { page-break-after:auto; }
.sheet--cover { width:210mm; height:297mm; }

/* ── Cover decorative bar ── */
.cover__eyebrow { display:flex; gap:3px; margin-bottom:7mm; }
.cover__eyebrow span { display:block; width:20px; height:10px; border-radius:1px; }
.cover__eyebrow span:nth-child(1) { background:var(--brand-navy); }
.cover__eyebrow span:nth-child(2) { background:var(--brand-green); }
.cover__eyebrow span:nth-child(3) { background:var(--brand-purple); }
.cover__eyebrow span:nth-child(4) { background:var(--brand-lavender); }
.cover__eyebrow span:nth-child(5) { background:#C3C0DE; }

/* ── Header ── */
.sheet__head {
  display:flex; justify-content:space-between; align-items:center;
  background:#EEEFF7; flex-shrink:0;
}
.sheet__head-meta { font-size:7.5pt; color:var(--brand-navy); font-weight:600; padding:2.5mm 10mm; }
.sheet__head-meta b { font-weight:700; text-transform:uppercase; }
.sheet__head-right { font-size:7.5pt; color:var(--brand-navy); font-weight:600;
  padding:2.5mm 10mm; text-align:right; font-family:var(--font-brand); }

/* ── Body ── */
.sheet__body { flex:1; padding:2mm 8mm 2mm; display:flex;
  flex-direction:column; overflow:hidden; min-height:0; }

/* ── Footer ── */
.sheet__foot {
  border-top:0.5px solid var(--ink-200);
  padding:2mm 10mm 3mm;
  display:flex; justify-content:space-between;
  font-size:7pt; color:var(--ink-400); flex-shrink:0;
}

/* ── Section title ── */
.doc-section { margin-bottom:1mm; flex-shrink:0; }
.doc-section__title {
  font-family:var(--font-body); font-weight:700;
  font-size:9pt; color:var(--brand-navy); margin:0 0 1mm;
  line-height:1.1; text-transform:uppercase; letter-spacing:0.06em;
}

/* ══ TABLA ══
   18 columnas (sin Nº inscr. transportista):
   Fecha | LER | Desc | (t) | DI | CT | RS-dest | NIF | NIMA | Naut | Tipo |
   Mun | Prov | CCAA | País | RS-transp | Método | Proceso
*/
.crono-table {
  width:100%; border-collapse:collapse;
  font-size:5.6pt; font-family:var(--font-body);
  table-layout:fixed; line-height:1.2; color:var(--ink-700);
}

/* Group row */
.crono-table thead .group th {
  background:var(--brand-navy); color:#fff;
  font-weight:800; letter-spacing:0.04em; text-transform:uppercase;
  font-size:6pt; padding:1mm 0.8mm; text-align:center;
  border-right:0.25mm solid rgba(255,255,255,.3);
  border-bottom:0.25mm solid rgba(255,255,255,.4);
  vertical-align:middle; white-space:nowrap;
}
.crono-table thead .group th:last-child { border-right:0; }

/* Sub row */
.crono-table thead .sub th {
  background:#ADB1D9; color:#1B1D58;
  font-weight:700; font-size:5.4pt;
  padding:0.8mm 0.5mm; text-align:center;
  border-right:0.15mm solid rgba(255,255,255,.2);
  vertical-align:middle; line-height:1.15;
  white-space:normal; word-break:break-word;
}
.crono-table thead .sub th:last-child { border-right:0; }
.crono-table thead { display:table-header-group; }

/* Body */
.crono-table tbody td {
  padding:0.8mm 0.6mm;
  border-bottom:0.15mm solid var(--ink-100);
  border-right:0.1mm solid var(--ink-50);
  vertical-align:middle; text-align:center;
  word-wrap:break-word; overflow-wrap:break-word;
  font-size:5.6pt; line-height:1.2;
}
.crono-table tbody tr:nth-child(even) { background:var(--ink-25); }

/* Tipos de celda */
.crono-table td.num  { font-family:var(--font-mono); font-variant-numeric:tabular-nums; white-space:nowrap; }
.crono-table td.ler  { font-family:var(--font-mono); white-space:nowrap; }
.crono-table td.date { font-family:var(--font-mono); font-variant-numeric:tabular-nums; white-space:nowrap; }
.crono-table td.mono { font-family:var(--font-mono); font-size:5pt;
  white-space:normal; word-break:break-all; overflow-wrap:anywhere; }
.crono-table td.abbr { font-family:var(--font-mono); font-size:5.6pt; font-weight:600;
  white-space:normal; word-break:break-word; overflow-wrap:anywhere; text-align:center; }
.crono-table td.abbr-fixed { font-family:var(--font-mono); font-size:5.6pt; font-weight:600;
  white-space:nowrap; text-align:center; }
.crono-table td.izq  { text-align:center; }
.crono-table td.valor { white-space:normal; word-break:keep-all; line-height:1.15; }

/* Tfoot */
.crono-table tfoot td {
  font-family:var(--font-body); font-weight:800;
  background:#EEEFF7; color:var(--brand-navy);
  border-top:0.4mm solid var(--brand-navy);
  padding:1.2mm 0.8mm; font-size:6.5pt; text-align:center;
}
.crono-table tfoot td.num { font-family:var(--font-mono); font-size:7pt; }
.crono-table tfoot td:empty { background:transparent; border-top:0; padding:0; }
.th-razon { background:#ADB1D9; color:#1B1D58; }

/* ── Print bar ── */
.print-bar { position:fixed; top:16px; right:16px; z-index:10;
  background:#fff; border-radius:8px;
  box-shadow:0 6px 16px -6px rgba(11,12,36,0.12); padding:6px; }
.print-bar button { font-family:var(--font-body); font-weight:600; font-size:13px;
  padding:8px 14px; border-radius:5px; border:none;
  background:var(--brand-navy); color:#fff; cursor:pointer; }
.print-bar button:hover { background:var(--navy-700); }

/* ── Print ── */
@page { size:A4 landscape; margin:0; }
@media print {
  html,body { background:#fff; }
  .sheet { margin:0; box-shadow:none; }
  .print-bar { display:none !important; }
  .sheet--cover { page:cover; }
}
@page cover { size:A4 portrait; margin:0; }
"""

# ─────────────────────────────────────────────────────────
# Colgroup  — 18 columnas, anchos equilibrados
# Suma = 100 %
# ─────────────────────────────────────────────────────────
# Columnas:
#  1  Fecha        5.5
#  2  Cód LER      3.2
#  3  Descripción  7.5
#  4  (t)          3.5
#  5  Nº DI        6.5
#  6  Nº CT        9.5
#  7  RS destino   9.0
#  8  NIF destino  4.5
#  9  NIMA dest    5.5
# 10  Nº aut       5.0
# 11  Tipo insc    2.6
# 12  Mun          4.0
# 13  Prov         2.4
# 14  CCAA         2.4
# 15  País         2.2
# 16  RS transp    9.5   ← más ancha al quitar Nº inscr
# 17  Método       7.2
# 18  Proceso      9.5
# TOTAL = 100.0

COLGROUP = """<colgroup>
  <col style="width:5.3%"/>  <!-- Fecha       dd/mm/yyyy -->
  <col style="width:3.3%"/>  <!-- Cód LER     6 dígitos -->
  <col style="width:7.1%"/>  <!-- Descripción wrap -->
  <col style="width:3.6%"/>  <!-- (t)         numérico -->
  <col style="width:6.5%"/>  <!-- Nº DI       break-all -->
  <col style="width:10.1%"/> <!-- Nº CT       break-all -->
  <col style="width:9.5%"/>  <!-- RS destino  wrap -->
  <col style="width:4.8%"/>  <!-- NIF         9 chars -->
  <col style="width:5.7%"/>  <!-- NIMA        10 chars -->
  <col style="width:5.3%"/>  <!-- Nº aut      break -->
  <col style="width:2.7%"/>  <!-- Tipo insc   G01 -->
  <col style="width:5.0%"/>  <!-- Mun         Ormaiztegi -->
  <col style="width:2.6%"/>  <!-- Prov        GIP -->
  <col style="width:2.4%"/>  <!-- CCAA        PV -->
  <col style="width:2.3%"/>  <!-- País        ES -->
  <col style="width:9.5%"/>  <!-- RS transp   wrap -->
  <col style="width:6.5%"/>  <!-- Método      R1201/R1301 -->
  <col style="width:7.7%"/>  <!-- Proceso     Desmontaje VFU -->
</colgroup>"""

THEAD = """<thead>
  <tr class="group">
    <th>Fecha</th>
    <th colspan="2">Residuo</th>
    <th>Cant.</th>
    <th colspan="2">Documentación</th>
    <th colspan="5">Destino</th>
    <th colspan="4">Ubic. destino</th>
    <th>Transportista</th>
    <th>Trat.</th>
    <th>Proceso</th>
  </tr>
  <tr class="sub">
    <th>Fecha<br/>traslado</th>
    <th>Cód.<br/>LER</th>
    <th>Descripción</th>
    <th class="num">(t)</th>
    <th>Nº DI</th>
    <th>Nº CT</th>
    <th class="th-razon">Razón social</th>
    <th>NIF</th>
    <th>NIMA</th>
    <th>Nº aut.</th>
    <th>Tipo<br/>insc.</th>
    <th>Mun.</th>
    <th>Prov.</th>
    <th>CCAA</th>
    <th>País</th>
    <th class="th-razon">Razón social</th>
    <th>Método<br/>valor.</th>
    <th>Proceso</th>
  </tr>
</thead>"""

# Número de celdas colspan para tfoot
N_COLS = 18

def _fila_html(r):
    return (
        "<tr>"
        f'<td class="date">{_esc(r["fecha"])}</td>'
        f'<td class="ler">{_esc(r["ler"])}</td>'
        f'<td class="izq">{_esc(r["descripcion"])}</td>'
        f'<td class="num">{_esc(_fmt_peso(r["peso_t"]))}</td>'
        f'<td class="mono">{_esc(r["ndi"])}</td>'
        f'<td class="mono">{_esc(r["nct"])}</td>'
        f'<td class="izq">{_esc(r["dest_razon"])}</td>'
        f'<td class="mono">{_esc(r["dest_nif"])}</td>'
        f'<td class="mono">{_esc(r["dest_nima"])}</td>'
        f'<td class="mono">{_esc(r["dest_aut"])}</td>'
        f'<td class="abbr-fixed">{_esc(r["dest_tipo"])}</td>'
        f'<td class="abbr">{_esc(r["municipio"])}</td>'
        f'<td class="abbr-fixed">{_esc(r["provincia"])}</td>'
        f'<td class="abbr-fixed">{_esc(r["ccaa"])}</td>'
        f'<td class="abbr-fixed">{_esc(r["pais"])}</td>'
        f'<td class="izq">{_esc(r["transp_razon"])}</td>'
        f'<td class="valor">{_esc(r["metodo"])}</td>'
        f'<td class="valor">{_esc(r["proceso"])}</td>'
        "</tr>"
    )


# ─────────────────────────────────────────────────────────
# Portada
# ─────────────────────────────────────────────────────────
def _portada(datos, fecha_gen):
    empresa   = _esc(datos.get("razon_social",""))
    nima      = _esc(str(datos.get("nima","")))
    direccion = _esc(datos.get("direccion",""))
    num_aut   = _esc(str(datos.get("num_aut","")))
    tipo_aut  = _esc(str(datos.get("tipo_aut","")))
    ano       = _esc(str(datos.get("ano","")))
    fecha_str = _esc(_fmt_fecha_portada(fecha_gen))
    logo      = _logo_tag()

    filas = [
        ("Entidad",           empresa,  True),
        ("NIMA",              nima,     False),
        ("Planta",            direccion,False),
        ("Nº Autorización",   num_aut,  False),
        ("Tipo Autorización", tipo_aut, False),
        ("Periodo",           ano,      False),
    ]
    rows_html = ""
    for i, (label, val, _) in enumerate(filas):
        bdr = "" if i == len(filas)-1 else "border-bottom:0.2mm solid #E8E8EF;"
        bld = "font-weight:700;" if i == 0 else "font-weight:600;"
        rows_html += (
            f'<tr>'
            f'<td style="padding:2.6mm 4mm;{bdr}font-size:9.5pt;color:#2C2E88;font-weight:600;width:44mm;">{label}</td>'
            f'<td style="padding:2.6mm 4mm;{bdr}font-size:10pt;color:#0B0C16;{bld}">{val}</td>'
            f'</tr>'
        )

    return f"""
<section class="sheet sheet--cover">
  <div style="flex:1;display:flex;flex-direction:column;padding:28mm 22mm 0;">
    <div style="flex:0.4;"></div>
    <div class="cover__eyebrow">
      <span></span><span></span><span></span><span></span><span></span>
    </div>
    <h1 style="font-size:34pt;font-weight:700;color:#1B1D58;line-height:0.95;
      text-transform:uppercase;letter-spacing:-0.01em;margin:0 0 4mm;
      font-family:'IBM Plex Sans',system-ui,sans-serif;">
      Archivo<br/>Cronológico
    </h1>
    <div style="font-size:9.5pt;font-weight:700;color:#2C2E88;
      text-transform:uppercase;letter-spacing:0.06em;margin-bottom:14mm;
      font-family:'IBM Plex Sans',system-ui,sans-serif;">
      Gestión de Residuos de Vehículos al Final de su Vida Útil
    </div>
    <div style="border:1px solid #E8E8EF;border-left:3px solid #2C2E88;border-radius:3px;overflow:hidden;">
      <div style="padding:3mm 4mm;background:#EEEFF7;font-size:9pt;font-weight:700;
        color:#2C2E88;letter-spacing:0.08em;text-transform:uppercase;
        font-family:'IBM Plex Sans',system-ui,sans-serif;">
        Datos del centro
      </div>
      <table style="width:100%;border-collapse:collapse;font-family:'IBM Plex Sans',system-ui,sans-serif;">
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    <div style="flex:1;"></div>
    <div style="text-align:right;font-size:9pt;color:#7E8094;padding-bottom:6mm;
      margin-top:8mm;font-family:'IBM Plex Sans',system-ui,sans-serif;">
      Fecha: {fecha_str}
    </div>
  </div>
  <div style="margin:0 22mm;border-top:0.25mm solid #D2D2DE;"></div>
  <div style="padding:8mm 22mm 14mm;display:flex;justify-content:space-between;align-items:center;">
    <div style="display:flex;align-items:center;gap:6mm;">
      {logo}
      <div style="font-size:8.5pt;color:#2A2C3D;line-height:1.2;font-family:'Poppins',system-ui,sans-serif;">
        <b style="color:#2C2E88;font-weight:700;font-size:9pt;display:block;margin-bottom:0.5mm;">{ALVALE_NOMBRE}</b>
        <span style="display:block;font-family:'IBM Plex Sans',system-ui,sans-serif;">{ALVALE_DIR}</span>
        <span style="display:block;font-family:'IBM Plex Sans',system-ui,sans-serif;">{ALVALE_TEL}</span>
      </div>
    </div>
    <div style="font-size:9pt;text-align:right;line-height:1.8;font-family:'Poppins',system-ui,sans-serif;">
      <a href="{ALVALE_WEB_URL}" style="color:#2C2E88;text-decoration:none;font-weight:600;display:block;">{ALVALE_WEB_TXT}</a>
      <a href="{ALVALE_LI_URL}" style="color:#2A2C3D;text-decoration:none;display:inline-flex;align-items:center;gap:1.5mm;">
        <span style="display:inline-flex;align-items:center;justify-content:center;width:5mm;height:5mm;
          border-radius:0.6mm;background:#2C2E88;color:#fff;font-weight:800;font-size:7.5pt;">in</span>
        LinkedIn
      </a>
    </div>
  </div>
</section>"""


# ─────────────────────────────────────────────────────────
# Páginas interiores
# ─────────────────────────────────────────────────────────
def _pagina_interior(empresa, ano, filas_html, total_peso,
                     es_primera, es_ultima, page_num, total_pages):
    seccion = ""
    if es_primera:
        seccion = """<div class="doc-section">
      <div class="doc-section__title">
        Residuos retirados\u00a0<span style="font-weight:400;color:#5D549D;">(salidas)</span>
      </div></div>"""

    tfoot = ""

    emp = _esc(str(empresa))
    a   = _esc(str(ano))

    return f"""
<section class="sheet">
  <header class="sheet__head">
    <div class="sheet__head-meta"><b>{emp}</b> | {a} | Archivo Cronológico</div>
    <div class="sheet__head-right">Alvale Consulting Ingenieros, S.L.</div>
  </header>
  <div class="sheet__body">
    {seccion}
    <table class="crono-table">
      {COLGROUP}
      {THEAD}
      <tbody>{''.join(filas_html)}</tbody>
      {tfoot}
    </table>
  </div>
  <footer class="sheet__foot">
    <span>Archivo Cronológico \u2014 {emp} \u2014 {a}</span>
    <span>Página {page_num} de {total_pages}</span>
  </footer>
</section>"""


# ─────────────────────────────────────────────────────────
# Generación del HTML
# ─────────────────────────────────────────────────────────
def generar_html(ruta_excel, ruta_salida=None):
    wb        = openpyxl.load_workbook(ruta_excel, data_only=True)
    datos     = leer_datos_empresa(wb)
    registros = [r for r in leer_registros(wb) if r["fecha"]]

    empresa    = datos.get("razon_social", "EMPRESA")
    ano        = datos.get("ano", "")
    total_peso = sum(r["peso_t"] for r in registros)

    if ruta_salida is None:
        base = os.path.splitext(os.path.basename(ruta_excel))[0]
        ruta_salida = f"Archivo_Cronologico_{base}.html"

    chunks = [registros[i:i+ROWS_PER_PAGE]
              for i in range(0, max(len(registros), 1), ROWS_PER_PAGE)]
    if not chunks: chunks = [[]]
    total_pages = len(chunks)

    sections = [_portada(datos, datetime.today())]
    for idx, chunk in enumerate(chunks):
        filas_html = [_fila_html(r) for r in chunk] or [
            f'<tr><td colspan="{N_COLS}" style="text-align:center;color:#7E8094;'
            f'padding:6mm;font-style:italic;">No se han registrado movimientos.</td></tr>'
        ]
        sections.append(_pagina_interior(
            empresa, ano, filas_html, total_peso,
            es_primera=(idx == 0), es_ultima=(idx == len(chunks)-1),
            page_num=idx+1, total_pages=total_pages,
        ))

    titulo = f"Archivo Cronológico — {_esc(str(empresa))} · {_esc(str(ano))} · Alvale Consulting"
    html = (
        '<!doctype html>\n<html lang="es">\n<head>\n'
        '<meta charset="utf-8"/>\n'
        f'<title>{titulo}</title>\n'
        f'<style>{CSS}</style>\n'
        '</head>\n<body>\n\n'
        '<div class="print-bar">'
        '<button onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button>'
        '</div>\n\n'
        + '\n'.join(sections) +
        '\n\n</body>\n</html>'
    )

    with open(ruta_salida, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅  HTML generado: {ruta_salida}")
    print(f"    {len(registros)} registros · {total_pages} página(s) de datos")
    print(f"    Total retiradas: {_fmt_peso(total_peso)} t")
    print()
    print("    Abre en el navegador → Ctrl+P → Guardar como PDF.")
    return ruta_salida


# ─────────────────────────────────────────────────────────
# Punto de entrada
# ─────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Genera el Archivo Cronológico de Retiradas de Residuos para desguaces VFU"
    )
    parser.add_argument("excel", help="Ruta al Excel de Ingurunet")
    parser.add_argument("--output", "-o", default=None, help="Ruta de salida HTML")
    args = parser.parse_args()
    if not os.path.isfile(args.excel):
        print(f"❌  Archivo no encontrado: {args.excel}")
        sys.exit(1)
    generar_html(args.excel, args.output)

if __name__ == "__main__":
    main()
