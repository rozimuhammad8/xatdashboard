import os
import json
import uuid
import datetime
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS

from docx_generator import build_letter_document, safe_filename

app = Flask(__name__, static_folder='public')
CORS(app)

# ============================================================
# KONFIGURATSIYA
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, 'public')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
DB_FILE = os.path.join(BASE_DIR, 'database.json')

os.makedirs(PUBLIC_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# MA'LUMOTLAR BAZASI
# ============================================================
def load_database():
    if not os.path.exists(DB_FILE):
        initial_data = {
            'letters': [],
            'stats': {
                'total': 0,
                'byTemplate': {},
                'byDate': {}
            }
        }
        with open(DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(initial_data, f, ensure_ascii=False, indent=2)
        return initial_data
    try:
        with open(DB_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if 'stats' not in data:
                data['stats'] = {'total': 0, 'byTemplate': {}, 'byDate': {}}
            if 'byTemplate' not in data['stats']:
                data['stats']['byTemplate'] = {}
            if 'byDate' not in data['stats']:
                data['stats']['byDate'] = {}
            return data
    except:
        return {'letters': [], 'stats': {'total': 0, 'byTemplate': {}, 'byDate': {}}}

def save_database(data):
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def add_letter(letter_data):
    db = load_database()
    letter = {
        'id': str(uuid.uuid4()),
        **letter_data,
        'createdAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'createdAtDate': datetime.datetime.now().strftime('%Y-%m-%d')
    }
    db['letters'].append(letter)
    db['stats']['total'] = len(db['letters'])
    
    template = letter_data.get('template', 'unknown')
    if 'byTemplate' not in db['stats']:
        db['stats']['byTemplate'] = {}
    db['stats']['byTemplate'][template] = db['stats']['byTemplate'].get(template, 0) + 1
    
    date_key = datetime.datetime.now().strftime('%Y-%m-%d')
    if 'byDate' not in db['stats']:
        db['stats']['byDate'] = {}
    db['stats']['byDate'][date_key] = db['stats']['byDate'].get(date_key, 0) + 1
    
    save_database(db)
    return letter

def get_letters(filters=None):
    db = load_database()
    letters = db['letters']
    
    if filters:
        if filters.get('search'):
            search = filters['search'].lower()
            letters = [l for l in letters if 
                      search in l.get('fio', '').lower() or
                      search in l.get('mfyNomi', '').lower() or
                      search in l.get('street', '').lower() or
                      search in l.get('murojaatRaqami', '').lower()]
        
        if filters.get('fromDate'):
            letters = [l for l in letters if l.get('createdAtDate', '') >= filters['fromDate']]
        
        if filters.get('toDate'):
            letters = [l for l in letters if l.get('createdAtDate', '') <= filters['toDate']]
        
        if filters.get('template'):
            letters = [l for l in letters if l.get('template') == filters['template']]
    
    letters.sort(key=lambda x: x.get('createdAt', ''), reverse=True)
    return letters

def get_letter_by_id(letter_id):
    db = load_database()
    for letter in db['letters']:
        if letter['id'] == letter_id:
            return letter
    return None

def delete_letter(letter_id):
    db = load_database()
    index = None
    for i, l in enumerate(db['letters']):
        if l['id'] == letter_id:
            index = i
            break
    
    if index is None:
        return False
    
    letter = db['letters'][index]
    template = letter.get('template', 'unknown')
    if 'byTemplate' in db['stats'] and template in db['stats']['byTemplate']:
        db['stats']['byTemplate'][template] = max(0, db['stats']['byTemplate'][template] - 1)
    
    db['letters'].pop(index)
    db['stats']['total'] = len(db['letters'])
    save_database(db)
    return True

# ============================================================
# HTML ROUTELAR (public papkasidan)
# ============================================================

@app.route('/')
def index():
    """Bosh sahifa - Xat yaratish"""
    return send_from_directory(PUBLIC_DIR, 'index.html')

@app.route('/admin')
def admin():
    """Admin panel"""
    return send_from_directory(PUBLIC_DIR, 'admin.html')

@app.route('/rad')
def rad():
    """Rad etish xati"""
    return send_from_directory(PUBLIC_DIR, 'rad.html')

@app.route('/tasdiqlandi')
def tasdiq():
    """Tasdiqlash xati"""
    return send_from_directory(PUBLIC_DIR, 'tasdiqlandi.html')

@app.route('/tayinlandi')
def tayin():
    """Tayinlash xati"""
    return send_from_directory(PUBLIC_DIR, 'tayinlandi.html')

@app.route('/muddat')
def muddat():
    """Muddat sorash xati"""
    return send_from_directory(PUBLIC_DIR, 'muddat.html')

@app.route('/arizaKiritilmagan')
def ariza_kiritilmagan():
    """Ariza kiritilmagan xati"""
    return send_from_directory(PUBLIC_DIR, 'arizaKiritilmagan.html')

@app.route('/arizaKiritilgan')
def ariza_kiritilgan():
    """Ariza kiritilgan xati"""
    return send_from_directory(PUBLIC_DIR, 'arizaKiritilgan.html')

# ============================================================
# API ENDPOINTLAR
# ============================================================

@app.route('/api/templates', methods=['GET'])
def get_templates():
    try:
        templates = [
            {'id': 'rad', 'name': 'Rad etish'},
            {'id': 'tasdiqlandi', 'name': 'Tasdiqlash'},
            {'id': 'tayinlandi', 'name': 'Tayinlash'},
            {'id': 'muddat', 'name': 'Muddat sorash'},
            {'id': 'arizaKiritilmagan', 'name': 'Ariza kiritilmagan'},
            {'id': 'arizaKiritilgan', 'name': 'Ariza kiritilgan'}
        ]
        return jsonify(templates)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/generate', methods=['POST'])
def generate_letter():
    try:
        form_data = request.json
        template = form_data.get('template', 'rad')

        required_fields = ['fio', 'mfyNomi', 'street', 'murojaatfrom', 'murojaatVaqti', 'murojaatRaqami']
        # "Ariza kiritilmagan" va "Muddat sorash" xatlari ariza maqsadi/ID/sanasini
        # umuman ishlatmaydi (docx_generator.py) — bu ikki shablonda talab qilinmaydi
        # (index.html validate() bilan mos).
        if template not in ('arizaKiritilmagan', 'muddat'):
            required_fields += ['arizaVaqti', 'arizaID']
        missing_fields = [field for field in required_fields if not form_data.get(field)]
        
        if missing_fields:
            return jsonify({
                'success': False, 
                'errors': [f"{field} maydoni to'ldirilishi shart!" for field in missing_fields]
            }), 400

        if template == 'rad':
            rad_sabablari = form_data.get('radSabablari', {})
            has_reason = (
                rad_sabablari.get('uydaEmasRad') or
                rad_sabablari.get('norasmiyRad') or
                (rad_sabablari.get('uyRad') and len(rad_sabablari['uyRad']) > 0) or
                (rad_sabablari.get('avtoRad') and len(rad_sabablari['avtoRad']) > 0) or
                (rad_sabablari.get('rasmiyRad') and len(rad_sabablari['rasmiyRad']) > 0)
            )
            if not has_reason:
                return jsonify({
                    'success': False,
                    'errors': ['Kamida bitta rad etish sababini tanlang!']
                }), 400
        
        letter = add_letter(form_data)
        
        json_file = os.path.join(OUTPUT_DIR, f"{letter['id']}.json")
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(letter, f, ensure_ascii=False, indent=2)
        
        return jsonify({
            'success': True,
            'message': 'Ma\'lumotlar muvaffaqiyatli saqlandi!',
            'letterId': letter['id'],
            'data': letter
        })
        
    except Exception as e:
        print(f'Xat yaratishda xatolik: {e}')
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': str(e) or 'Xat yaratishda xatolik yuz berdi'
        }), 500

@app.route('/api/letters', methods=['GET'])
def get_all_letters():
    try:
        search = request.args.get('search', '')
        from_date = request.args.get('fromDate', '')
        to_date = request.args.get('toDate', '')
        template = request.args.get('template', '')
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        filters = {}
        if search:
            filters['search'] = search
        if from_date:
            filters['fromDate'] = from_date
        if to_date:
            filters['toDate'] = to_date
        if template:
            filters['template'] = template
        
        letters = get_letters(filters)
        total = len(letters)
        paginated = letters[offset:offset + limit]
        
        return jsonify({
            'success': True,
            'data': paginated,
            'total': total,
            'limit': limit,
            'offset': offset
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/letters/<letter_id>', methods=['GET'])
def get_letter_by_id_route(letter_id):
    try:
        letter = get_letter_by_id(letter_id)
        if not letter:
            return jsonify({'error': 'Xat topilmadi'}), 404
        
        return jsonify({
            'success': True,
            'data': letter
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/letters/<letter_id>', methods=['DELETE'])
def delete_letter_route(letter_id):
    try:
        result = delete_letter(letter_id)
        if not result:
            return jsonify({'error': 'Xat topilmadi'}), 404
        
        json_file = os.path.join(OUTPUT_DIR, f"{letter_id}.json")
        if os.path.exists(json_file):
            os.remove(json_file)
        
        return jsonify({
            'success': True, 
            'message': "Xat o'chirildi"
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/letters/<letter_id>/download', methods=['GET'])
def download_letter_json(letter_id):
    try:
        letter = get_letter_by_id(letter_id)
        if not letter:
            return jsonify({'error': 'Xat topilmadi'}), 404
        
        json_file = os.path.join(OUTPUT_DIR, f"{letter_id}.json")
        if not os.path.exists(json_file):
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(letter, f, ensure_ascii=False, indent=2)
        
        return send_file(
            json_file, 
            as_attachment=True, 
            download_name=f"letter_{letter_id}.json"
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/letters/<letter_id>/download-docx', methods=['GET'])
def download_letter_docx(letter_id):
    """Xat ma'lumotlari asosida namunadagi kabi rasmiylashtirilgan Word (.docx) faylini yaratib beradi"""
    try:
        letter = get_letter_by_id(letter_id)
        if not letter:
            return jsonify({'error': 'Xat topilmadi'}), 404

        buffer = build_letter_document(letter)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=safe_filename(letter),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f'Word fayl yaratishda xatolik: {e}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e) or 'Word faylini yaratishda xatolik yuz berdi'}), 500

@app.route('/api/stats', methods=['GET'])
def get_stats():
    try:
        db = load_database()
        return jsonify({
            'success': True,
            'stats': db.get('stats', {}),
            'total': len(db.get('letters', []))
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health():
    try:
        db = load_database()
        return jsonify({
            'status': 'OK',
            'timestamp': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'totalLetters': len(db.get('letters', [])),
            'templates': ['rad', 'tasdiqlandi', 'tayinlandi', 'muddat', 'arizaKiritilmagan', 'arizaKiritilgan']
        })
    except Exception as e:
        return jsonify({'status': 'ERROR', 'error': str(e)}), 500

# ============================================================
# XATOLIKLARNI USHLASH
# ============================================================
@app.errorhandler(Exception)
def handle_exception(e):
    print(f'Server xatosi: {e}')
    import traceback
    traceback.print_exc()
    return jsonify({
        'success': False,
        'error': 'Serverda xatolik yuz berdi',
        'details': str(e)
    }), 500

# ============================================================
# SERVER ISHGA TUSHIRISH
# ============================================================
if __name__ == '__main__':
    print('\n' + '='*50)
    print('🚀 PYTHON SERVER ISHGA TUSHIDI')
    print('📡 PORT: 3000')
    print('🌐 URL: http://localhost:3000')
    print(f'📁 Public papka: {PUBLIC_DIR}')
    print(f'📁 Chiqish: {OUTPUT_DIR}')
    print('='*50)
    print('📌 SAHIFALAR:')
    print('   🌐 http://localhost:3000/          - Xat yaratish')
    print('   🌐 http://localhost:3000/admin    - Admin panel')
    print('   🌐 http://localhost:3000/rad?id=.. - Rad etish xati')
    print('   🌐 http://localhost:3000/tasdiq?id=.. - Tasdiqlash xati')
    print('   🌐 http://localhost:3000/tayin?id=.. - Tayinlash xati')
    print('   🌐 http://localhost:3000/muddat?id=.. - Muddat sorash xati')
    print('   🌐 http://localhost:3000/arizaKiritilmagan?id=.. - Ariza kiritilmagan xati')
    print('   🌐 http://localhost:3000/arizaKiritilgan?id=.. - Ariza kiritilgan xati')
    print('='*50)
    print('📌 API ENDPOINTLAR:')
    print('   POST   /api/generate        - Xat ma\'lumotlarini saqlash')
    print('   GET    /api/letters         - Barcha xatlarni olish')
    print('   GET    /api/letters/<id>    - ID bo\'yicha xat olish')
    print('   DELETE /api/letters/<id>    - Xatni o\'chirish')
    print('   GET    /api/letters/<id>/download - JSON faylni yuklab olish')
    print('   GET    /api/letters/<id>/download-docx - Word (.docx) faylni yuklab olish')
    print('   GET    /api/stats           - Statistika')
    print('   GET    /api/health          - Server holati')
    print('='*50 + '\n')
    
    app.run(host='0.0.0.0', port=3000, debug=True)