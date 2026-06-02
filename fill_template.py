import sys, json, base64, os, shutil, subprocess
from docx import Document
from docx.oxml.ns import qn
import lxml.etree as etree

WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
W_NS   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

# Mapa: nombre del recuadro → significado
# Rectangle 3 = PEN, Rectangle 4 = USD
# Rectangle 5 = Ocasionales, Rectangle 6 = Permanentes, Rectangle 2 = Combustible

def get_shapes(cell):
    """Devuelve dict {nombre: wsp_element} para todos los recuadros flotantes de la celda."""
    shapes = {}
    for para in cell.paragraphs:
        for run in para.runs:
            for anchor in run._r.iter(f'{{{WP_NS}}}anchor'):
                docPr = anchor.find(f'{{{WP_NS}}}docPr')
                name  = docPr.get('name') if docPr is not None else ''
                wsp   = anchor.find(f'.//{{{WPS_NS}}}wsp')
                if wsp is not None:
                    shapes[name] = wsp
    return shapes

def add_x_to_shape(wsp_el):
    """Inserta una X centrada y en negrita dentro del recuadro flotante."""
    txbx    = etree.Element(f'{{{WPS_NS}}}txbx')
    content = etree.SubElement(txbx, f'{{{W_NS}}}txbxContent')
    p       = etree.SubElement(content, f'{{{W_NS}}}p')
    pPr     = etree.SubElement(p, f'{{{W_NS}}}pPr')
    jc      = etree.SubElement(pPr, f'{{{W_NS}}}jc')
    jc.set(f'{{{W_NS}}}val', 'center')
    r       = etree.SubElement(p, f'{{{W_NS}}}r')
    rPr     = etree.SubElement(r, f'{{{W_NS}}}rPr')
    b       = etree.SubElement(rPr, f'{{{W_NS}}}b')
    sz      = etree.SubElement(rPr, f'{{{W_NS}}}sz')
    sz.set(f'{{{W_NS}}}val', '16')
    t       = etree.SubElement(r, f'{{{W_NS}}}t')
    t.text  = 'X'
    bodyPr  = wsp_el.find(f'{{{WPS_NS}}}bodyPr')
    wsp_el.insert(list(wsp_el).index(bodyPr), txbx)

def remove_highlight(run):
    rPr = run._r.find(qn('w:rPr'))
    if rPr is not None:
        for tag in [qn('w:highlight'), qn('w:shd')]:
            el = rPr.find(tag)
            if el is not None:
                rPr.remove(el)

def set_run(run, text):
    run.text = text
    remove_highlight(run)

def numero_en_letras(monto):
    entero   = int(monto)
    centavos = round((monto - entero) * 100)
    unidades = ['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ',
        'ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE',
        'VEINTE','VEINTIUNO','VEINTIDÓS','VEINTITRÉS','VEINTICUATRO','VEINTICINCO','VEINTISÉIS',
        'VEINTISIETE','VEINTIOCHO','VEINTINUEVE']
    decenas  = ['','','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA']
    centenas = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
        'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']

    def grupo(n):
        if n == 0: return ''
        if n == 100: return 'CIEN'
        r = ''
        if n >= 100:
            r += centenas[n // 100] + ' '; n %= 100
        if n < 30: r += unidades[n]
        else:
            r += decenas[n // 10]
            if n % 10: r += ' Y ' + unidades[n % 10]
        return r.strip()

    def convertir(n):
        if n == 0: return 'CERO'
        r = ''
        if n >= 1000:
            miles = n // 1000
            r += ('MIL' if miles == 1 else grupo(miles) + ' MIL') + ' '
            n %= 1000
        r += grupo(n)
        return r.strip()

    return f"{convertir(entero)} {str(centavos).zfill(2)}/100 NUEVOS SOLES"

def fmt_date(iso):
    if not iso: return ''
    parts = iso.split('-')
    return f"{parts[2]}/{parts[1]}/{parts[0]}" if len(parts) == 3 else iso

def fill_and_convert(data):
    src = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plantilla.docx')
    dst = '/tmp/filled.docx'
    shutil.copy(src, dst)
    doc = Document(dst)

    nombre       = data.get('nombre', '')
    dni          = data.get('dni', '')
    proyecto     = data.get('proyecto', '')
    consec       = data.get('consec', '001')
    fecha_elab   = fmt_date(data.get('fechaElab', ''))
    ciudad       = data.get('ciudad', '')
    motivo       = data.get('motivo', '')
    comentarios  = data.get('comentarios', '')
    deductivo    = float(data.get('deductivo', 0))
    deductivo_desc = data.get('deductivoDesc', '')
    fecha_sal    = fmt_date(data.get('fechaSalida', ''))
    fecha_reg    = fmt_date(data.get('fechaRegreso', ''))
    monto        = float(data.get('monto', 0))
    tipo         = data.get('tipo', 'Ocasionales')
    moneda       = data.get('moneda', 'PEN')
    cuenta       = data.get('cuenta', '')
    cci          = data.get('cci', '')
    banco        = data.get('banco', '')
    ticket_de    = data.get('ticketDe', '')
    ticket_a     = data.get('ticketA', '')
    ticket_fecha = fmt_date(data.get('ticketFecha', ''))
    ticket_hora  = data.get('ticketHora', '')
    dias         = int(data.get('dias', 0))

    # Texto de comentarios con deductivo integrado
    comentarios_full = comentarios
    if deductivo > 0:
        ded_texto = f"Deductivo: S/. {deductivo:.2f}"
        if deductivo_desc:
            ded_texto += f" — {deductivo_desc}"
        comentarios_full = (comentarios_full + ' | ' + ded_texto).strip(' |')

    monto_letras = numero_en_letras(monto)
    monto_str    = f"{monto:.2f}".split('.')
    monto_ent    = f"{int(monto_str[0]):,}"
    monto_dec    = monto_str[1]

    tables = doc.tables

    # ── TABLE 1: Consecutivo y Fecha elaboración ──────────────────
    t1   = tables[1]
    runs = t1.rows[1].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], consec)
    if len(runs) > 3: set_run(runs[3], '')

    runs_fe = t1.rows[1].cells[1].paragraphs[0].runs
    dp = fecha_elab.split('/') if fecha_elab else ['','','']
    if len(runs_fe) > 2: set_run(runs_fe[2], dp[0] if len(dp) > 0 else '')
    if len(runs_fe) > 3: set_run(runs_fe[3], '/')
    if len(runs_fe) > 4: set_run(runs_fe[4], dp[1] if len(dp) > 1 else '')
    if len(runs_fe) > 5: set_run(runs_fe[5], '/' + (dp[2] if len(dp) > 2 else ''))

    # ── TABLE 2: Datos del empleado ───────────────────────────────
    t2 = tables[2]

    runs = t2.rows[0].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], nombre)

    runs = t2.rows[1].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], dni)

    runs = t2.rows[2].cells[0].paragraphs[0].runs
    if len(runs) > 3: set_run(runs[3], proyecto)

    runs = t2.rows[3].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], ciudad)

    runs = t2.rows[4].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], motivo)

    runs = t2.rows[5].cells[0].paragraphs[0].runs
    if len(runs) > 2: set_run(runs[2], comentarios_full)

    # Fechas
    runs = t2.rows[6].cells[0].paragraphs[0].runs
    sal  = fecha_sal.split('/') if fecha_sal else ['','','']
    reg  = fecha_reg.split('/') if fecha_reg else ['','','']
    if len(runs) > 7:  set_run(runs[7],  sal[0] if len(sal) > 0 else '')
    if len(runs) > 8:  set_run(runs[8],  '/')
    if len(runs) > 9:  set_run(runs[9],  sal[1] if len(sal) > 1 else '')
    if len(runs) > 10: set_run(runs[10], '/' + (sal[2] if len(sal) > 2 else ''))
    if len(runs) > 17: set_run(runs[17], reg[0] if len(reg) > 0 else '')
    if len(runs) > 18: set_run(runs[18], '/')
    if len(runs) > 19: set_run(runs[19], reg[1] if len(reg) > 1 else '')
    if len(runs) > 20: set_run(runs[20], '/' + (reg[2] if len(reg) > 2 else ''))

    # ── Monto ─────────────────────────────────────────────────────
    cell  = t2.rows[7].cells[0]
    runs0 = cell.paragraphs[0].runs

    if deductivo > 0:
        # Monto bruto (antes del deductivo)
        monto_bruto     = monto + deductivo
        bruto_str       = f"{monto_bruto:.2f}".split('.')
        bruto_ent       = f"{int(bruto_str[0]):,}"
        bruto_dec       = bruto_str[1]
        ded_ent         = f"{int(deductivo):,}"
        ded_dec         = f"{deductivo:.2f}".split('.')[1]

        # Formato: S/. 510.00 − S/. 170.00 = S/. 340.00 (LETRAS)
        ecuacion = (
            f"S/. {bruto_ent}.{bruto_dec} "
            f"\u2212 S/. {ded_ent}.{ded_dec} "
            f"= S/. {monto_ent}.{monto_dec} "
            f"({monto_letras})"
        )
        # Colocar toda la ecuación en run1, vaciar los demás runs de P0
        if len(runs0) > 0: set_run(runs0[0], 'Solicito la suma de:  ')
        if len(runs0) > 1: set_run(runs0[1], ecuacion)
        for i in range(2, len(runs0)):
            set_run(runs0[i], '')
    else:
        # Sin deductivo: formato original S/ X,XXX.00 (LETRAS)
        if len(runs0) > 1: set_run(runs0[1], 'S/ ')
        if len(runs0) > 2: set_run(runs0[2], monto_ent)
        if len(runs0) > 3: set_run(runs0[3], '.')
        if len(runs0) > 4: set_run(runs0[4], monto_dec)
        if len(runs0) > 5: set_run(runs0[5], '')
        if len(runs0) > 6: set_run(runs0[6], '')
        if len(runs0) > 7: set_run(runs0[7], ' (')
        if len(runs0) > 8: set_run(runs0[8], monto_letras)
        if len(runs0) > 9: set_run(runs0[9], ')')

    # Limpiar X residual de runs de texto en P1 y P4
    runs1 = cell.paragraphs[1].runs
    if len(runs1) > 10: runs1[10].text = '     '   # quitar X de texto de USD

    runs4 = cell.paragraphs[4].runs
    if len(runs4) > 10: runs4[10].text = '         '  # quitar X de texto de Permanentes

    # ── Marcar recuadros con X según selección ────────────────────
    shapes = get_shapes(cell)

    # Moneda
    if moneda == 'PEN':
        if 'Rectangle 3' in shapes: add_x_to_shape(shapes['Rectangle 3'])
    else:  # USD
        if 'Rectangle 4' in shapes: add_x_to_shape(shapes['Rectangle 4'])

    # Tipo de viático
    if tipo == 'Ocasionales':
        if 'Rectangle 5' in shapes: add_x_to_shape(shapes['Rectangle 5'])
    elif tipo == 'Permanentes':
        if 'Rectangle 6' in shapes: add_x_to_shape(shapes['Rectangle 6'])
    else:  # Combustible
        if 'Rectangle 2' in shapes: add_x_to_shape(shapes['Rectangle 2'])

    # ── Datos bancarios ───────────────────────────────────────────
    for run in t2.rows[9].cells[0].paragraphs[1].runs:
        if 'X' in run.text: set_run(run, cuenta)
    for run in t2.rows[9].cells[1].paragraphs[1].runs:
        if 'X' in run.text: set_run(run, cci)
    for run in t2.rows[9].cells[2].paragraphs[1].runs:
        if 'X' in run.text: set_run(run, banco)

    # ── TABLE 3: Tiquete aéreo ────────────────────────────────────
    if ticket_de or ticket_a:
        t3 = tables[3]
        row = t3.rows[2]
        def set_cell_text(c, text):
            p = c.paragraphs[0]
            if not p.runs: p.add_run(text)
            else: p.runs[0].text = text
        if len(row.cells) > 1: set_cell_text(row.cells[1], ticket_de)
        if len(row.cells) > 2: set_cell_text(row.cells[2], ticket_a)
        if len(row.cells) > 3: set_cell_text(row.cells[3], ticket_fecha)
        if len(row.cells) > 4: set_cell_text(row.cells[4], ticket_hora)

    doc.save(dst)

    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', '/tmp/', dst],
        capture_output=True, text=True, timeout=60
    )
    pdf_path = '/tmp/filled.pdf'
    if not os.path.exists(pdf_path):
        raise Exception(f"PDF conversion failed: {result.returncode} — {result.stderr or result.stdout}")

    with open(pdf_path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode('utf-8')

    os.remove(pdf_path)
    os.remove(dst)
    return b64

if __name__ == '__main__':
    data = json.loads(sys.argv[1])
    print(fill_and_convert(data))
