from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.styles import Alignment
from PIL import Image as PILImage
import io, base64, os, json

app = Flask(__name__)
CORS(app)

TEMPLATE = os.path.join(os.path.dirname(__file__), 'AVISO_DE_RECEPCION_.xlsx')

# Slots: (key, from_col, from_row, to_col, to_row) — índices 0-based
PHOTO_SLOTS_REG = [
    ('recepcion',      2,  54, 10, 62),   # C55:J62
    ('alimentaciones', 12, 54, 20, 62),   # M55:T62
    ('anclajes',       2,  63, 10, 71),   # C64:J71
    ('accesorios',     12, 63, 20, 71),   # M64:T71
]
GUIA_SLOT = ('guia_recepcion', 1, 26, 21, 52)  # B27:U52

def _insert_image(ws, img_data, from_col, from_row, to_col, to_row):
    if isinstance(img_data, str):
        if ',' in img_data: img_data = img_data.split(',')[1]
        img_bytes = base64.b64decode(img_data)
    else:
        img_bytes = img_data
    pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGB')
    buf = io.BytesIO()
    pil.save(buf, format='JPEG', quality=85)
    buf.seek(0)
    img = XLImage(buf)
    img.anchor = TwoCellAnchor(
        editAs='twoCell',
        _from=AnchorMarker(col=from_col, row=from_row, colOff=0, rowOff=0),
        to=AnchorMarker(col=to_col,   row=to_row,   colOff=0, rowOff=0)
    )
    ws.add_image(img)

def build_fsgi249(rec, photos):
    wb = load_workbook(TEMPLATE)
    ws = wb.active

    def w(addr, val):
        cell = ws[addr]
        if type(cell).__name__ == 'MergedCell': return
        cell.value = val

    ws['U2'].value = 'FSGI-249'
    ws['U3'].value = f"Revisión: 3  ·  {rec.get('codigo','')}"
    ws['U4'].value = f"Fecha: {rec.get('fecha','')}"
    w('D7',  rec.get('cliente',''))
    w('Q7',  f"{rec.get('fecha','')}  {rec.get('hora','')}")
    w('D8',  rec.get('oc','') or '')
    w('Q8',  rec.get('pat_cam','') or '')
    w('F9',  rec.get('guia',''))
    w('Q9',  rec.get('pat_rem','') or '')
    w('E10', rec.get('ot_ant','') or '')
    w('O10', rec.get('ot','') or 'Pendiente')
    ws['B11'].value = (
        f"COMPONENTE:  {rec.get('tipo','')}   Marca: {rec.get('marca','—')}   "
        f"Modelo: {rec.get('modelo','—')}   S/N: {rec.get('nserie','—')}   "
        f"P/N: {rec.get('nparte','—')}   Estado: {rec.get('estvis','—')}   "
        f"Embalaje: {rec.get('cemb','—')}"
    )
    conf = rec.get('estvis','') == 'Bueno'
    ws['G12'].value = 'CONFORME:' + ('   ✓' if conf else '')
    ws['M12'].value = 'NO CONFORME:' + ('   ✓' if not conf else '')

    ck = rec.get('ck', {})
    for key, row in {'11':16,'12':17,'13':18,'21':20,'22':21,'23':22,'31':23}.items():
        val = ck.get(key)
        w(f'M{row}', '✓' if val == 'si' else '')
        w(f'N{row}', '✓' if val == 'no' else '')
        oc = ws[f'O{row}']
        if type(oc).__name__ != 'MergedCell': oc.value = ck.get(f'obs_{key}','')

    # Guía Y/O Factura: foto o texto
    if photos.get('guia_recepcion'):
        _insert_image(ws, photos['guia_recepcion'], 1, 26, 21, 52)  # B27:U52
    else:
        lineas = [f"Transportista: {rec.get('transp','—')}     Bultos: {rec.get('bultos','1')}     Código: {rec.get('codigo','')}"]
        if rec.get('acc'):      lineas.append(f"Accesorios: {rec['acc']}")
        if rec.get('obs_gral'): lineas.append(f"Obs. generales: {rec['obs_gral']}")
        if rec.get('obs_comp'): lineas.append(f"Obs. componente: {rec['obs_comp']}")
        ws['B27'].value = '\n'.join(lineas)
        ws['B27'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

    # Registro fotográfico — todas las fotos con TwoCellAnchor
    for key, fc, fr, tc, tr in PHOTO_SLOTS_REG:
        if photos.get(key):
            _insert_image(ws, photos[key], fc, fr, tc, tr)

    w('D206', rec.get('receptor',''))
    w('N206', rec.get('cargo','Bodega'))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar-fsgi249', methods=['POST'])
def generar():
    try:
        rec = json.loads(request.form.get('rec','{}'))
        photos = {}
        for k in ['recepcion','alimentaciones','anclajes','guia_recepcion','accesorios']:
            f = request.files.get(k)
            if f: photos[k] = f.read()
        xlsx = build_fsgi249(rec, photos)
        codigo  = rec.get('codigo','RC')
        cliente = rec.get('cliente','').replace(' ','_')
        return send_file(xlsx,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{codigo}_{cliente}_FSGI249.xlsx")
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
