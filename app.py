from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from PIL import Image as PILImage
import io, base64, os, json

app = Flask(__name__)
CORS(app)

TEMPLATE = os.path.join(os.path.dirname(__file__), 'AVISO_DE_RECEPCION_.xlsx')

PHOTO_SLOTS = [
    ('general',    'C55', 728, 234),
    ('placa',      'M55', 728, 234),
    ('serie',      'C64', 728, 234),
    ('danos',      'M64', 728, 234),
    ('accesorios', 'C73', 728, 234),
    ('embalaje',   'M73', 728, 234),
    ('guia',       'C82', 728, 234),
]

def build_fsgi249(rec, photos):
    wb = load_workbook(TEMPLATE)
    ws = wb.active

    def w(addr, val):
        cell = ws[addr]
        if type(cell).__name__ == 'MergedCell': return
        cell.value = val

    # Header
    ws['U2'].value = 'FSGI-249'
    ws['U3'].value = f"Revisión: 3  ·  {rec.get('codigo','')}"
    ws['U4'].value = f"Fecha: {rec.get('fecha','')}"

    # Datos generales
    w('D7',  rec.get('cliente',''))
    w('Q7',  f"{rec.get('fecha','')}  {rec.get('hora','')}")
    w('D8',  rec.get('oc','') or '')
    w('Q8',  rec.get('pat_cam','') or '')
    w('F9',  rec.get('guia',''))
    w('Q9',  rec.get('pat_rem','') or '')
    w('E10', rec.get('ot_ant','') or '')
    w('O10', rec.get('ot','') or 'Pendiente')

    # Componente
    ws['B11'].value = (
        f"COMPONENTE:  {rec.get('tipo','')}   "
        f"Marca: {rec.get('marca','—')}   "
        f"Modelo: {rec.get('modelo','—')}   "
        f"S/N: {rec.get('nserie','—')}   "
        f"P/N: {rec.get('nparte','—')}   "
        f"Estado: {rec.get('estvis','—')}   "
        f"Embalaje: {rec.get('cemb','—')}"
    )

    # Conforme
    conf = rec.get('estvis','') == 'Bueno'
    ws['G12'].value = 'CONFORME:' + ('   ✓' if conf else '')
    ws['M12'].value = 'NO CONFORME:' + ('   ✓' if not conf else '')

    # Checklist
    ck = rec.get('ck', {})
    chk_map = {'11':16,'12':17,'13':18,'21':20,'22':21,'23':22,'31':23}
    for key, row in chk_map.items():
        val = ck.get(key)
        w(f'M{row}', '✓' if val == 'si' else '')
        w(f'N{row}', '✓' if val == 'no' else '')
        oc = ws[f'O{row}']
        if type(oc).__name__ != 'MergedCell':
            oc.value = ck.get(f'obs_{key}', '')

    # Sección guía/factura
    lineas = [
        f"Transportista: {rec.get('transp','—')}     "
        f"Bultos: {rec.get('bultos','1')}     "
        f"Código interno Swanson: {rec.get('codigo','')}",
    ]
    if rec.get('acc'):      lineas.append(f"Accesorios: {rec['acc']}")
    if rec.get('obs_gral'): lineas.append(f"Obs. generales: {rec['obs_gral']}")
    if rec.get('obs_comp'): lineas.append(f"Obs. componente: {rec['obs_comp']}")
    ws['B27'].value = '\n'.join(lineas)
    ws['B27'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

    # Fotografías — pegar en celdas combinadas en orden
    for idx, (key, anchor, w_px, h_px) in enumerate(PHOTO_SLOTS):
        img_data = photos.get(key)
        if not img_data:
            continue
        # Decode base64 if needed
        if isinstance(img_data, str):
            if ',' in img_data:
                img_data = img_data.split(',')[1]
            img_bytes = base64.b64decode(img_data)
        else:
            img_bytes = img_data

        # Resize to fit cell proportionally
        pil = PILImage.open(io.BytesIO(img_bytes))
        pil = pil.convert('RGB')
        pil.thumbnail((w_px, h_px), PILImage.LANCZOS)
        # Center in cell
        final = PILImage.new('RGB', (w_px, h_px), (255,255,255))
        offset_x = (w_px - pil.width) // 2
        offset_y = (h_px - pil.height) // 2
        final.paste(pil, (offset_x, offset_y))

        buf = io.BytesIO()
        final.save(buf, format='PNG', optimize=True)
        buf.seek(0)

        xl_img = XLImage(buf)
        xl_img.width  = w_px
        xl_img.height = h_px
        xl_img.anchor = anchor
        ws.add_image(xl_img)

    # Firma
    w('D206', rec.get('receptor',''))
    w('N206', 'Bodega')

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'Swanson RC Backend'})

@app.route('/generar-fsgi249', methods=['POST'])
def generar():
    try:
        # Parse multipart form
        rec_json = request.form.get('rec')
        if not rec_json:
            return jsonify({'error': 'Falta campo rec'}), 400
        rec = json.loads(rec_json)

        # Collect photos from files
        photos = {}
        for slot_key, anchor, w_px, h_px in PHOTO_SLOTS:
            f = request.files.get(slot_key)
            if f:
                photos[slot_key] = f.read()

        xlsx = build_fsgi249(rec, photos)

        codigo  = rec.get('codigo', 'RC')
        cliente = rec.get('cliente', '').replace(' ', '_')
        fname   = f"{codigo}_{cliente}_FSGI249.xlsx"

        return send_file(
            xlsx,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=fname
        )
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
