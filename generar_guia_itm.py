#!/usr/bin/env python3
"""
FDE-074  Guía N°1 – Visita de Inicio – Prácticas Profesionales ITM  (V10)
Dependencias:  pip install python-docx lxml
"""

from docx import Document
from docx.shared import Pt, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── Paleta ───────────────────────────────────────────────────────────────────
AZUL_MARINO   = "102D69"
AZUL_CLARO    = "56ACDE"
AZUL_MARINO_L = "E8EDF5"
AZUL_CLARO_L  = "EAF4FB"
BLANCO        = "FFFFFF"
NEGRO         = "1A1A1A"
GRIS_L        = "F4F6F9"
GRIS_M        = "C5CDD8"

# ─── Dimensiones A4 ───────────────────────────────────────────────────────────
PAGE_W  = 11906
PAGE_H  = 16838
MARGIN  = 1080
CONTENT = PAGE_W - 2 * MARGIN   # 9 746 DXA

# ─── Bordes ───────────────────────────────────────────────────────────────────
BMAR  = dict(style="single", size=8, color=AZUL_MARINO)
BCLR  = dict(style="single", size=6, color=AZUL_CLARO)
BGRIS = dict(style="single", size=4, color=GRIS_M)
BNONE = dict(style="nil",    size=0, color=BLANCO)

BORD_M = dict(top=BMAR,  bottom=BMAR,  left=BMAR,  right=BMAR)
BORD_C = dict(top=BCLR,  bottom=BCLR,  left=BCLR,  right=BCLR)
BORD_G = dict(top=BGRIS, bottom=BGRIS, left=BGRIS, right=BGRIS)

# ══════════════════════════════════════════════════════════════════════════════
#  Primitivas XML
# ══════════════════════════════════════════════════════════════════════════════

def _tcPr(tc):
    p = tc.find(qn('w:tcPr'))
    if p is None:
        p = OxmlElement('w:tcPr')
        tc.insert(0, p)
    return p

def _shading(tc, fill):
    pr = _tcPr(tc)
    for s in pr.findall(qn('w:shd')): pr.remove(s)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    pr.append(shd)

def _borders(tc, bords):
    pr = _tcPr(tc)
    for old in pr.findall(qn('w:tcBorders')): pr.remove(old)
    tcB = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right'):
        b = bords.get(side)
        if b:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'),   b['style'])
            el.set(qn('w:sz'),    str(b['size']))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), b['color'])
            tcB.append(el)
    pr.append(tcB)

def _valign(tc, val='center'):
    pr = _tcPr(tc)
    for old in pr.findall(qn('w:vAlign')): pr.remove(old)
    v = OxmlElement('w:vAlign')
    v.set(qn('w:val'), val)
    pr.append(v)

def _margins(tc, top=65, bottom=65, left=110, right=110):
    pr = _tcPr(tc)
    for old in pr.findall(qn('w:tcMar')): pr.remove(old)
    m = OxmlElement('w:tcMar')
    for side, val in (('top',top),('bottom',bottom),('left',left),('right',right)):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'),    str(val))
        el.set(qn('w:type'), 'dxa')
        m.append(el)
    pr.append(m)

def _width(tc, w):
    pr = _tcPr(tc)
    for old in pr.findall(qn('w:tcW')): pr.remove(old)
    el = OxmlElement('w:tcW')
    el.set(qn('w:w'),    str(w))
    el.set(qn('w:type'), 'dxa')
    pr.insert(0, el)

def _colspan(tc, span):
    if span <= 1: return
    pr = _tcPr(tc)
    for old in pr.findall(qn('w:gridSpan')): pr.remove(old)
    gs = OxmlElement('w:gridSpan')
    gs.set(qn('w:val'), str(span))
    pr.append(gs)

def _vmerge(tc, restart=True):
    pr = _tcPr(tc)
    vm = OxmlElement('w:vMerge')
    if restart: vm.set(qn('w:val'), 'restart')
    pr.append(vm)

# ──────────────────────────────────────────────────────────────────────────────
#  Runs y Párrafos
# ──────────────────────────────────────────────────────────────────────────────

def _run(text, size=18, bold=False, color=NEGRO, italics=False, underline=False):
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    fonts = OxmlElement('w:rFonts')
    for attr in ('w:ascii','w:hAnsi','w:cs'):
        fonts.set(qn(attr), 'Arial')
    rPr.append(fonts)
    for tag in ('w:sz','w:szCs'):
        el = OxmlElement(tag); el.set(qn('w:val'), str(size)); rPr.append(el)
    if bold:
        rPr.append(OxmlElement('w:b')); rPr.append(OxmlElement('w:bCs'))
    if italics:
        rPr.append(OxmlElement('w:i'))
    if underline:
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    col = OxmlElement('w:color'); col.set(qn('w:val'), color); rPr.append(col)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    return r

def _para(runs, align='left', before=40, after=40, line=276):
    """Build a w:p element."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), align); pPr.append(jc)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'),  str(after))
    sp.set(qn('w:line'),   str(line))
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)
    p.append(pPr)
    if isinstance(runs, str):
        runs = [{'text': runs}]
    elif isinstance(runs, dict):
        runs = [runs]
    for spec in runs:
        p.append(_run(**spec))
    return p

# ──────────────────────────────────────────────────────────────────────────────
#  Celdas
# ──────────────────────────────────────────────────────────────────────────────

def _tc(w, colspan=1, rs_start=False, rs_cont=False):
    tc = OxmlElement('w:tc')
    pr = OxmlElement('w:tcPr')
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'),    str(w))
    tcW.set(qn('w:type'), 'dxa')
    pr.append(tcW)
    if colspan > 1:
        gs = OxmlElement('w:gridSpan'); gs.set(qn('w:val'), str(colspan)); pr.append(gs)
    if rs_start:
        vm = OxmlElement('w:vMerge'); vm.set(qn('w:val'), 'restart'); pr.append(vm)
    elif rs_cont:
        pr.append(OxmlElement('w:vMerge'))
    tc.append(pr)
    return tc

def celda_tit(text, w, colspan=1, fill=AZUL_MARINO, bord=None):
    bord = bord or BORD_M
    tc = _tc(w, colspan=colspan)
    _shading(tc, fill); _borders(tc, bord)
    _margins(tc, 90, 90, 130, 130); _valign(tc, 'center')
    tc.append(_para({'text':text,'bold':True,'color':BLANCO,'size':18}, 'center', 0, 0))
    return tc

def celda_lbl(text, w, fill=AZUL_MARINO_L, bord=None):
    bord = bord or BORD_G
    tc = _tc(w)
    _shading(tc, fill); _borders(tc, bord)
    _margins(tc, 65, 65, 110, 110); _valign(tc, 'center')
    tc.append(_para({'text':text,'bold':True,'size':17,'color':AZUL_MARINO}, 'left', 0, 0))
    return tc

def celda_in(text, w, fill=BLANCO, bord=None):
    bord = bord or BORD_G
    tc = _tc(w)
    _shading(tc, fill); _borders(tc, bord)
    _margins(tc, 65, 65, 110, 110); _valign(tc, 'center')
    tc.append(_para({'text':text,'size':17}, 'left', 0, 0))
    return tc

# ──────────────────────────────────────────────────────────────────────────────
#  Tablas
# ──────────────────────────────────────────────────────────────────────────────

def _tbl(col_widths):
    tbl = OxmlElement('w:tbl')
    tblPr = OxmlElement('w:tblPr')
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'),    str(sum(col_widths)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{side}')
        for k,v in (('w:val','none'),('w:sz','0'),('w:space','0'),('w:color','auto')):
            el.set(qn(k), v)
        tblBorders.append(el)
    tblPr.append(tblBorders)
    tbl.append(tblPr)
    tblGrid = OxmlElement('w:tblGrid')
    for w in col_widths:
        gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(w)); tblGrid.append(gc)
    tbl.append(tblGrid)
    return tbl

def _tr(*cells):
    tr = OxmlElement('w:tr')
    for c in cells: tr.append(c)
    return tr

# ──────────────────────────────────────────────────────────────────────────────
#  Utilidades de documento
# ──────────────────────────────────────────────────────────────────────────────

def spacer():
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '30'); sp.set(qn('w:after'), '30')
    pPr.append(sp); p.append(pPr)
    p.append(_run(''))
    return p

def nota_italica(text):
    return _para({'text':text,'size':15,'italics':True,'color':AZUL_MARINO},
                 'left', 60, 60)

def sec_titulo(text):
    tbl = _tbl([300, CONTENT-300])
    tc1 = _tc(300)
    _shading(tc1, AZUL_CLARO); _borders(tc1, BORD_M)
    _margins(tc1, 80, 80, 0, 0)
    tc1.append(_para('', 'center', 0, 0))
    tc2 = _tc(CONTENT-300)
    _shading(tc2, AZUL_MARINO); _borders(tc2, BORD_M)
    _margins(tc2, 80, 80, 160, 160); _valign(tc2, 'center')
    tc2.append(_para({'text':text,'bold':True,'color':BLANCO,'size':19}, 'left', 0, 0))
    tbl.append(_tr(tc1, tc2))
    return tbl

# ══════════════════════════════════════════════════════════════════════════════
#  Documento
# ══════════════════════════════════════════════════════════════════════════════

doc  = Document()
sec  = doc.sections[0]
sec.page_width    = Twips(PAGE_W)
sec.page_height   = Twips(PAGE_H)
sec.left_margin   = Twips(MARGIN)
sec.right_margin  = Twips(MARGIN)
sec.top_margin    = Twips(MARGIN)
sec.bottom_margin = Twips(MARGIN)

doc.styles['Normal'].font.name = 'Arial'
doc.styles['Normal'].font.size = Pt(9)

body = doc.element.body
# Eliminar párrafo vacío inicial
for p in list(body):
    if p.tag == qn('w:p'):
        body.remove(p); break

def add(el):
    sp = body.find(qn('w:sectPr'))
    if sp is not None:
        sp.addprevious(el)
    else:
        body.append(el)

# ══════════════════════════════════════════════════════════════════════════════
#  ENCABEZADO INSTITUCIONAL
# ══════════════════════════════════════════════════════════════════════════════
w1, w2, w3 = 2100, 5646, 2000
tbl_enc = _tbl([w1, w2, w3])

# Fila 1
tc_logo = _tc(w1, rs_start=True)
_shading(tc_logo, AZUL_MARINO); _borders(tc_logo, BORD_M)
_margins(tc_logo, 80, 80, 100, 100); _valign(tc_logo, 'center')
tc_logo.append(_para({'text':'ITM','bold':True,'size':40,'color':BLANCO},       'center', 20, 0))
tc_logo.append(_para({'text':'Institución','size':15,'color':AZUL_CLARO},        'center', 0,  0))
tc_logo.append(_para({'text':'Universitaria','size':15,'color':AZUL_CLARO},      'center', 0,  0))

tc_tit = _tc(w2)
_shading(tc_tit, BLANCO); _borders(tc_tit, BORD_M)
_margins(tc_tit, 70, 70, 120, 120); _valign(tc_tit, 'center')
tc_tit.append(_para({'text':'OFICINA DE PRÁCTICAS PROFESIONALES','bold':True,'size':16,'color':AZUL_CLARO}, 'center', 0, 0))
tc_tit.append(_para({'text':'GUÍA N°1 – VISITA DE INICIO','bold':True,'size':22,'color':AZUL_MARINO},       'center', 0, 0))
tc_tit.append(_para({'text':'MODALIDAD: PRÁCTICA PROFESIONAL','bold':True,'size':15,'color':NEGRO},          'center', 0, 0))

tc_cod = _tc(w3, rs_start=True)
_shading(tc_cod, AZUL_MARINO_L); _borders(tc_cod, BORD_M)
_margins(tc_cod, 70, 70, 100, 100); _valign(tc_cod, 'center')
tc_cod.append(_para({'text':'Código:','bold':True,'size':15,'color':AZUL_MARINO}, before=0, after=0))
tc_cod.append(_para({'text':'FDE-074','bold':True,'size':18,'color':AZUL_CLARO},  before=0, after=0))
tc_cod.append(_para({'text':'Versión: V10','size':15},                             before=0, after=0))
tc_cod.append(_para({'text':'Proceso: Formación','size':14},                       before=0, after=0))
tc_cod.append(_para({'text':'Decreto 0223/2026','size':13,'italics':True,'color':AZUL_MARINO}, before=0, after=0))

tbl_enc.append(_tr(tc_logo, tc_tit, tc_cod))

# Fila 2
def _rs_cont(w, fill, bord):
    tc = _tc(w, rs_cont=True)
    _shading(tc, fill); _borders(tc, bord)
    tc.append(_para('', before=0, after=0))
    return tc

tc_dir = _tc(w2)
_shading(tc_dir, AZUL_CLARO_L); _borders(tc_dir, BORD_M)
_margins(tc_dir, 45, 45, 120, 120)
tc_dir.append(_para([
    {'text':'Dirección de Gestión Académica  |  ','size':15},
    {'text':'Página 1 de 1','size':15,'bold':True,'color':AZUL_MARINO},
], 'center', 0, 0))

tbl_enc.append(_tr(_rs_cont(w1, AZUL_MARINO, BORD_M), tc_dir, _rs_cont(w3, AZUL_MARINO_L, BORD_M)))

# Fila 3
tc_ref = _tc(w2)
_shading(tc_ref, BLANCO); _borders(tc_ref, BORD_M)
_margins(tc_ref, 45, 45, 120, 120)
tc_ref.append(_para([
    {'text':'Ref. normativa: ','size':14,'bold':True,'color':AZUL_MARINO},
    {'text':'Política de Prácticas ITM · Decreto 0223/2026 Arts. 2.2.6.3.1.1 al 2.2.6.3.1.11','size':14},
], 'center', 0, 0))

tbl_enc.append(_tr(_rs_cont(w1, AZUL_MARINO, BORD_M), tc_ref, _rs_cont(w3, AZUL_MARINO_L, BORD_M)))

add(tbl_enc)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  1. OBJETIVO
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("1.  OBJETIVO DEL INSTRUMENTO"))
add(_para({'text':(
    "La Guía N°1 – Visita de Inicio tiene como propósito verificar y registrar las condiciones iniciales "
    "en las que el practicante ingresa al escenario de práctica o empresa patrocinadora durante el primer "
    "mes de vinculación. Sirve además como instrumento de formalización del Plan de Trabajo (Plan de Práctica) "
    "exigido por el Artículo 2.2.6.3.1.11 del Decreto 0223 de 2026 y la Política de Prácticas Profesionales ITM, "
    "garantizando la articulación tripartita entre el practicante, el Tutor y el Monitor."),'size':17},
    before=100, after=60))
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  2. IDENTIFICACIÓN DEL PRACTICANTE
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("2.  IDENTIFICACIÓN DEL PRACTICANTE"))

wL = 2700; wR = CONTENT - wL
tbl_p = _tbl([wL, wR])
pract_rows = [
    ("Nombre completo:",                  BLANCO, ""),
    ("Número de identificación:",         GRIS_L, ""),
    ("Correo electrónico institucional:", BLANCO, ""),
    ("Teléfono de contacto:",             GRIS_L, ""),
    ("Facultad / Programa académico:",    BLANCO, ""),
    ("Semestre en curso:",                GRIS_L, ""),
    ("Monitor asignado (ITM):",           BLANCO, ""),
    ("Fecha de inicio de la práctica:",   GRIS_L, ""),
    ("Fecha estimada de finalización:",   BLANCO, ""),
    ("Modalidad de práctica:",            GRIS_L,
     "☐  Práctica Profesional     ☐  Práctica Social     ☐  Práctica Internacional"),
]
for lbl, fill, txt in pract_rows:
    tbl_p.append(_tr(celda_lbl(lbl, wL), celda_in(txt, wR, fill=fill)))
add(tbl_p)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  3. IDENTIFICACIÓN DEL ESCENARIO
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("3.  IDENTIFICACIÓN DEL ESCENARIO DE PRÁCTICA / EMPRESA PATROCINADORA"))

tbl_e = _tbl([wL, wR])
emp_rows = [
    ("Razón social / Nombre del escenario:", BLANCO, ""),
    ("NIT / Documento de identificación:",   GRIS_L, ""),
    ("Tipo de escenario:",                   BLANCO,
     "☐  Empresa patrocinadora (Contrato aprendizaje)     ☐  Escenario de práctica laboral (Vinc. Formativa)"),
    ("Sector económico:",                    GRIS_L, ""),
    ("Dirección / Ciudad:",                  BLANCO, ""),
    ("Nombre del Representante Legal:",      GRIS_L, ""),
    ("Nombre del Tutor designado:",          BLANCO, ""),
    ("Cargo del Tutor:",                     GRIS_L, ""),
    ("Correo electrónico del Tutor:",        BLANCO, ""),
    ("Teléfono del Tutor:",                  GRIS_L, ""),
    ("Área / Dependencia de práctica:",      BLANCO, ""),
    ("Modalidad de trabajo:",                GRIS_L,
     "☐  Presencial     ☐  Híbrida     ☐  Virtual"),
]
for lbl, fill, txt in emp_rows:
    tbl_e.append(_tr(celda_lbl(lbl, wL), celda_in(txt, wR, fill=fill)))
add(tbl_e)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  4. VERIFICACIÓN DE DOCUMENTOS
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("4.  VERIFICACIÓN DE DOCUMENTOS DE VINCULACIÓN"))
add(nota_italica("(Decreto 0223/2026 – Arts. 2.2.6.3.2.3, 2.2.6.3.2.4 y 2.2.6.3.3.14)"))

wD1, wD2, wD3 = 5100, 1600, CONTENT - 5100 - 1600
tbl_d = _tbl([wD1, wD2, wD3])
tbl_d.append(_tr(
    celda_tit("Documento requerido",       wD1, fill=AZUL_CLARO, bord=BORD_C),
    celda_tit("Verificado",                wD2, fill=AZUL_CLARO, bord=BORD_C),
    celda_tit("Observación / N° radicado", wD3, fill=AZUL_CLARO, bord=BORD_C),
))
docs_data = [
    ("Contrato de aprendizaje  /  Acuerdo de voluntades  /  Acto administrativo", False),
    ("Afiliación a ARL — obligatoria (Art. 2.2.6.3.2.9 y 2.2.6.3.3.7 Decreto 0223/2026)", True),
    ("Afiliación a EPS (Salud)", False),
    ("Afiliación a Fondo de Pensión (solo fase práctica contrato aprendizaje)", True),
    ("Carta de inicio emitida por la OPP-ITM", False),
    ("Reglamento interno de trabajo entregado al aprendiz (contrato aprendizaje)", True),
    ("Aprobación de funciones por Monitor ITM", False),
    ("Acta de inicio suscrita por las tres partes", True),
    ("Autorización Inspector de Trabajo (solo si el practicante tiene 15–17 años — Art. 2.2.6.3.1.4)", False),
]
for lbl, alt in docs_data:
    fill = AZUL_CLARO_L if alt else BLANCO
    tc_chk = _tc(wD2)
    _shading(tc_chk, fill); _borders(tc_chk, BORD_G)
    _margins(tc_chk, 55, 55, 80, 80); _valign(tc_chk, 'center')
    tc_chk.append(_para({'text':'☐ Sí   ☐ No   ☐ N/A','size':16}, 'center', 0, 0))
    tbl_d.append(_tr(celda_in(lbl, wD1, fill=fill), tc_chk, celda_in("", wD3, fill=fill)))
add(tbl_d)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  5. CONDICIONES DEL ESCENARIO
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("5.  VERIFICACIÓN DE CONDICIONES DEL ESCENARIO DE PRÁCTICA"))
add(nota_italica("(Decreto 0223/2026 – Art. 2.2.6.3.1.10 | Política ITM – Cap. VI, num. 6.2)"))

wC1, wC2, wC3 = 5000, 1700, CONTENT - 5000 - 1700
tbl_c = _tbl([wC1, wC2, wC3])
tbl_c.append(_tr(
    celda_tit("Condición verificada", wC1, fill=AZUL_CLARO, bord=BORD_C),
    celda_tit("Estado",               wC2, fill=AZUL_CLARO, bord=BORD_C),
    celda_tit("Observación",          wC3, fill=AZUL_CLARO, bord=BORD_C),
))
cond_data = [
    ("El practicante cuenta con espacio físico/virtual adecuado para realizar sus funciones", False),
    ("Se le suministraron los EPP según la actividad (Art. 2.2.6.3.1.10 num. 7)", True),
    ("Se realizó inducción institucional y de seguridad al ingreso", False),
    ("El practicante conoce el reglamento y normas internas del escenario", True),
    ("El horario no supera la jornada ordinaria del escenario ni la máxima legal", False),
    ("El horario permite la asistencia a actividades académicas convocadas por la ITM", True),
    ("Las funciones asignadas corresponden al área de formación del practicante", False),
    ("El escenario cuenta con protocolo de prevención de acoso sexual (Ley 2365/2024)", True),
    ("El Tutor tiene experiencia/conocimiento en las actividades de la práctica", False),
    ("El practicante conoce los canales de atención ante situaciones de acoso u hostigamiento", True),
]
for lbl, alt in cond_data:
    fill = AZUL_CLARO_L if alt else BLANCO
    tc_est = _tc(wC2)
    _shading(tc_est, fill); _borders(tc_est, BORD_G)
    _margins(tc_est, 55, 55, 80, 80); _valign(tc_est, 'center')
    tc_est.append(_para({'text':'☐ Cumple   ☐ No cumple','size':16}, 'center', 0, 0))
    tbl_c.append(_tr(celda_in(lbl, wC1, fill=fill), tc_est, celda_in("", wC3, fill=fill)))
add(tbl_c)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  6. PLAN DE TRABAJO – PRIMER MES
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("6.  PLAN DE TRABAJO – PRIMER MES DE PRÁCTICA"))
add(nota_italica("(Art. 2.2.6.3.1.11 Decreto 0223/2026 | Política ITM Sección 5.4 | Cap. VI Num. 6.3)"))

# Bloque objetivos
tbl_obj = _tbl([CONTENT])
tc_oh = _tc(CONTENT)
_shading(tc_oh, AZUL_CLARO_L); _borders(tc_oh, BORD_C)
_margins(tc_oh, 80, 80, 130, 130)
tc_oh.append(_para({'text':'Objetivos formativos acordados para el primer mes (Practicante · Tutor · Monitor):',
                    'bold':True,'size':17,'color':AZUL_MARINO}, 'left', 0, 0))
tbl_obj.append(_tr(tc_oh))

tc_ob = _tc(CONTENT)
_shading(tc_ob, BLANCO); _borders(tc_ob, BORD_G)
_margins(tc_ob, 70, 140, 130, 130)
for n in range(1, 4):
    tc_ob.append(_para({'text':f'{n}. _________________________________________________________________________','size':17}, before=40, after=0))
tbl_obj.append(_tr(tc_ob))
add(tbl_obj)
add(spacer())

# Tabla semanal
wS = [850, 3200, 2346, 1450, 1900]
tbl_sem = _tbl(wS)
tbl_sem.append(_tr(
    celda_tit("Período",                               wS[0]),
    celda_tit("Actividades / Funciones a desarrollar", wS[1]),
    celda_tit("Resultados de aprendizaje esperados",   wS[2]),
    celda_tit("Tipo de producto",                      wS[3]),
    celda_tit("Observaciones Monitor / Tutor",         wS[4]),
))
for sem in range(1, 5):
    alt  = (sem % 2 == 0)
    fill = AZUL_CLARO_L if alt else BLANCO
    flbl = AZUL_MARINO_L if alt else GRIS_L
    tc_s = _tc(wS[0])
    _shading(tc_s, flbl); _borders(tc_s, BORD_G)
    _margins(tc_s, 60, 60, 80, 80); _valign(tc_s, 'center')
    tc_s.append(_para({'text':f'Semana {sem}','bold':True,'size':16,'color':AZUL_MARINO}, 'center', 0, 0))
    tc_tp = _tc(wS[3])
    _shading(tc_tp, fill); _borders(tc_tp, BORD_G)
    _margins(tc_tp, 55, 55, 80, 80); _valign(tc_tp, 'center')
    for chk in ('☐ Entregable','☐ Informe','☐ Ninguno'):
        tc_tp.append(_para({'text':chk,'size':15}, before=0, after=0))
    tbl_sem.append(_tr(tc_s, celda_in("",wS[1],fill=fill), celda_in("",wS[2],fill=fill),
                       tc_tp, celda_in("",wS[4],fill=fill)))
add(tbl_sem)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  7. HORARIO
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("7.  HORARIO DE PRÁCTICA"))
add(nota_italica("(Art. 2.2.6.3.2.2 | Art. 2.2.6.3.1.5 Decreto 0223/2026 – Política ITM Sec. 8.14 para adolescentes)"))

wH = [1200, 1500, 1500, 1600, CONTENT - 5800]
tbl_h = _tbl(wH)
tbl_h.append(_tr(
    celda_tit("Día",             wH[0]),
    celda_tit("Hora entrada",    wH[1]),
    celda_tit("Hora salida",     wH[2]),
    celda_tit("Total horas/día", wH[3]),
    celda_tit("Observación",     wH[4]),
))
for i, dia in enumerate(["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"]):
    alt  = (i % 2 == 0)
    fdia = AZUL_CLARO_L if alt else AZUL_MARINO_L
    fin  = BLANCO if alt else GRIS_L
    tbl_h.append(_tr(celda_lbl(dia,wH[0],fill=fdia),
                     celda_in("",wH[1],fill=fin), celda_in("",wH[2],fill=fin),
                     celda_in("",wH[3],fill=fin), celda_in("",wH[4],fill=fin)))

tc_t1 = _tc(wH[1]+wH[2], colspan=2)
_shading(tc_t1, AZUL_CLARO_L); _borders(tc_t1, BORD_G)
_margins(tc_t1, 60, 60, 100, 100); _valign(tc_t1, 'center')
tc_t1.append(_para({'text':'__________ horas por semana','size':17}, before=0, after=0))

tc_t2 = _tc(wH[3]+wH[4], colspan=2)
_shading(tc_t2, AZUL_CLARO_L); _borders(tc_t2, BORD_G)
_margins(tc_t2, 60, 60, 100, 100); _valign(tc_t2, 'center')
tc_t2.append(_para({'text':'Horario compatible con actividades académicas ITM: ☐ Sí   ☐ No','size':17}, before=0, after=0))

tbl_h.append(_tr(celda_lbl("TOTAL SEMANAL", wH[0], fill=AZUL_CLARO_L), tc_t1, tc_t2))
add(tbl_h)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  8. COMPROMISOS
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("8.  COMPROMISOS Y OBSERVACIONES INICIALES"))

def bloque_obs(titulo, n=5):
    tbl = _tbl([CONTENT])
    tc_h = _tc(CONTENT)
    _shading(tc_h, AZUL_CLARO_L); _borders(tc_h, BORD_C)
    _margins(tc_h, 70, 70, 130, 130)
    tc_h.append(_para({'text':titulo,'bold':True,'size':17,'color':AZUL_MARINO}, before=0, after=0))
    tbl.append(_tr(tc_h))
    tc_b = _tc(CONTENT)
    _shading(tc_b, BLANCO); _borders(tc_b, BORD_G)
    _margins(tc_b, 70, 100, 130, 130)
    for _ in range(n):
        tc_b.append(_para({'text':'____________________________________________________________________________','size':17}, before=20, after=28))
    tbl.append(_tr(tc_b))
    return tbl

add(bloque_obs("8.1  Compromisos del Practicante (Política ITM Cap. VI, num. 6.3):"))
add(spacer())
add(bloque_obs("8.2  Compromisos del Tutor (Escenario de práctica) – Art. 2.2.6.3.1.11 Decreto 0223/2026:"))
add(spacer())
add(bloque_obs("8.3  Compromisos del Monitor (Oficina de Prácticas ITM) – Art. 2.2.6.3.1.11:"))
add(spacer())
add(bloque_obs("8.4  Observaciones generales de la visita de inicio:"))
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  9. DERECHOS Y DEBERES
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("9.  DERECHOS Y DEBERES BÁSICOS DEL PRACTICANTE"))
add(_para({'text':'El practicante declara haber sido informado de sus derechos y deberes conforme a la normativa vigente.',
           'size':16,'italics':True,'color':AZUL_MARINO}, before=60, after=60))

half = CONTENT // 2; half2 = CONTENT - half
tbl_dd = _tbl([half, half2])
tbl_dd.append(_tr(
    celda_tit("DERECHOS DEL PRACTICANTE  (Política ITM Cap. VI | Decreto 0223/2026)", half,  fill=AZUL_MARINO, bord=BORD_M),
    celda_tit("DEBERES DEL PRACTICANTE  (Art. 2.2.6.3.1.8 Decreto 0223/2026)",       half2, fill=AZUL_CLARO,  bord=BORD_C),
))
dd_rows = [
    ("✔ Recibir inducción completa al inicio de la práctica",              "✔ Presentar plan de práctica aprobado por tutor y monitor",                        False),
    ("✔ Contar con Tutor designado en el escenario",                       "✔ Cumplir el horario de práctica acordado",                                         True),
    ("✔ Conocer el plan de práctica y las actividades a realizar",         "✔ Mantener confidencialidad sobre información del escenario",                       False),
    ("✔ Afiliación a ARL durante toda la práctica",                        "✔ Informar a la IES y al escenario cualquier situación que afecte la práctica",    True),
    ("✔ Entorno libre de acoso y discriminación (Ley 2365/2024)",          "✔ Realizar uso adecuado de los elementos suministrados",                            False),
    ("✔ Conservar derechos morales de autor (Art. 2.2.6.3.1.7)",          "✔ Procurar el cuidado integral de su salud",                                        True),
    ("✔ Recibir apoyo de sostenimiento si aplica (contrato aprendizaje)",  "✔ Asistir a los encuentros convocados por el Monitor ITM",                          False),
]
for der, deb, alt in dd_rows:
    fill = AZUL_CLARO_L if alt else BLANCO
    tbl_dd.append(_tr(celda_in(der, half, fill=fill), celda_in(deb, half2, fill=fill)))
add(tbl_dd)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
# 10. PRÓXIMO SEGUIMIENTO
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("10.  PRÓXIMO ENCUENTRO DE SEGUIMIENTO (Guía N°2)"))
tbl_px = _tbl([wL, wR])
for lbl, fill, txt in [
    ("Fecha acordada Guía N°2 (Visita de seguimiento):",  BLANCO, ""),
    ("Medio de contacto preferido:",                       GRIS_L, "☐  Presencial     ☐  Videollamada     ☐  Correo electrónico"),
    ("Temas a abordar en la próxima visita:",              BLANCO, "Revisión de avance del Plan de Trabajo | Verificación de compromisos | Entregables parciales"),
]:
    tbl_px.append(_tr(celda_lbl(lbl, wL), celda_in(txt, wR, fill=fill)))
add(tbl_px)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
# 11. FIRMAS
# ══════════════════════════════════════════════════════════════════════════════
add(sec_titulo("11.  FIRMAS Y VALIDACIÓN TRIPARTITA"))

wF = CONTENT // 3; wF3 = CONTENT - wF * 2
tbl_f = _tbl([wF, wF, wF3])
tbl_f.append(_tr(
    celda_tit("PRACTICANTE",                         wF,  fill=AZUL_MARINO, bord=BORD_M),
    celda_tit("TUTOR  (Escenario de Práctica)",      wF,  fill=AZUL_CLARO,  bord=BORD_C),
    celda_tit("MONITOR  (Oficina de Prácticas ITM)", wF3, fill=AZUL_CLARO,  bord=BORD_C),
))

def celda_firma(lines, w):
    tc = _tc(w)
    _shading(tc, BLANCO); _borders(tc, BORD_G)
    _margins(tc, 70, 100, 100, 100)
    for line in lines:
        tc.append(_para({'text':line,'size':17}, before=20, after=28))
    return tc

firma_lines = [
    "Nombre: ________________________________", "",
    "Firma:", "", "",
    "________________________________", "",
    "C.C.: ___________________________________",
    "Fecha: __________________________________",
]
tutor_lines = [
    "Nombre: ________________________________", "",
    "Cargo:  ________________________________", "",
    "Firma:", "", "",
    "________________________________",
    "Fecha: __________________________________",
]
tbl_f.append(_tr(celda_firma(firma_lines, wF), celda_firma(tutor_lines, wF), celda_firma(tutor_lines, wF3)))

tc_sello = _tc(CONTENT, colspan=3)
_shading(tc_sello, AZUL_CLARO_L); _borders(tc_sello, BORD_C)
_margins(tc_sello, 60, 60, 130, 130)
tc_sello.append(_para({'text':(
    "Sello del Escenario / Empresa Patrocinadora (opcional): ______________________   "
    "Visita realizada en modalidad: ☐ Presencial   ☐ Virtual"),'size':16}, before=0, after=0))
tbl_f.append(_tr(tc_sello))
add(tbl_f)
add(spacer())

# ══════════════════════════════════════════════════════════════════════════════
#  CONTROL DOCUMENTAL
# ══════════════════════════════════════════════════════════════════════════════
wCtl = CONTENT // 3; wCtl3 = CONTENT - wCtl * 2
tbl_ctrl = _tbl([wCtl, wCtl, wCtl3])

tc_ch = _tc(CONTENT, colspan=3)
_shading(tc_ch, AZUL_MARINO); _borders(tc_ch, BORD_M)
_margins(tc_ch, 70, 70, 130, 130)
tc_ch.append(_para({'text':'CONTROL DOCUMENTAL','bold':True,'color':BLANCO,'size':18}, 'center', 0, 0))
tbl_ctrl.append(_tr(tc_ch))

tbl_ctrl.append(_tr(
    celda_lbl("Código: FDE-074",    wCtl,  fill=AZUL_MARINO_L),
    celda_lbl("Versión: V10",       wCtl,  fill=AZUL_CLARO_L),
    celda_lbl("Vigente desde: 2026",wCtl3, fill=AZUL_CLARO_L),
))
tbl_ctrl.append(_tr(
    celda_in("Elaboró: OPP-ITM",                   wCtl),
    celda_in("Revisó: Dirección Gestión Académica", wCtl),
    celda_in("Aprobó: Vicerrectoría de Docencia",  wCtl3),
))

tc_norm = _tc(CONTENT, colspan=3)
_shading(tc_norm, AZUL_CLARO_L); _borders(tc_norm, BORD_C)
_margins(tc_norm, 60, 60, 130, 130)
tc_norm.append(_para({'text':(
    "Referencia normativa: Política de Prácticas Profesionales ITM · Decreto 0223 del 5 de marzo de 2026 (Min. Trabajo) · "
    "Ley 1780/2016 · Ley 2466/2025 (Reforma Laboral) · Ley 2365/2024 (Prevención acoso sexual) · "
    "Ley 2039/2020 (Experiencia profesional) · Art. 81 Código Sustantivo del Trabajo"),'size':14}, before=0, after=0))
tbl_ctrl.append(_tr(tc_norm))
add(tbl_ctrl)

# ─── Párrafo final requerido por OOXML ───────────────────────────────────────
body.append(_para(''))

# ─── Guardar ─────────────────────────────────────────────────────────────────
OUTPUT = "FDE_074_Guia_N1_Visita_Inicio_PP_ITM_V10.docx"
doc.save(OUTPUT)
print(f"OK  Documento generado: {OUTPUT}")
