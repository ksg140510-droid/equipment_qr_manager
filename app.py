import os
import io
import json
import time
import shutil
import socket
import secrets
import logging
import calendar
import qrcode
import urllib.request
import urllib.error
from urllib.parse import urlencode
from datetime import datetime, date, timedelta
from flask import (Flask, render_template, request, redirect, url_for,
                   flash, send_file, jsonify, abort, session)
from database import get_db, init_db
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)

BASE_DIR    = os.path.dirname(__file__)

# ─── 세션 서명용 비밀키 (최초 실행 시 랜덤 생성 후 파일로 보관) ──
SECRET_KEY_PATH = os.path.join(BASE_DIR, 'secret_key.txt')
def _load_or_create_secret_key():
    if os.path.exists(SECRET_KEY_PATH):
        with open(SECRET_KEY_PATH, encoding='utf-8') as f:
            key = f.read().strip()
            if key:
                return key
    key = secrets.token_hex(32)
    with open(SECRET_KEY_PATH, 'w', encoding='utf-8') as f:
        f.write(key)
    return key

app.secret_key = _load_or_create_secret_key()
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 요청당 최대 50MB (사진 최대 10장 대비)

QR_DIR      = os.path.join(BASE_DIR, 'qr_codes')
UPLOAD_DIR  = os.path.join(BASE_DIR, 'uploads')
FAULT_DIR   = os.path.join(UPLOAD_DIR, 'fault')
ACTION_DIR  = os.path.join(UPLOAD_DIR, 'action')
PHOTOS_DIR  = os.path.join(UPLOAD_DIR, 'photos')

os.makedirs(PHOTOS_DIR, exist_ok=True)

ALLOWED_EXT = {'png', 'jpg', 'jpeg', 'gif', 'webp'}

SECTIONS = ['Section 1', 'Section 2', 'Section 3',
            'Section 4.1', 'Section 4.2', 'Section 4.3', 'Section 4.4']

SYMPTOMS = [
    '센서 감지불량', '실린더 동작불량', '에어누설 (Air Leak)',
    '배큠흡착불량', '히터 불량', '부품 파손', '체인 불량',
    '베어링 소음', '벨트 마모', '체결부 풀림',
    '제품 걸림 (JAM)', 'HMI 알람', 'PLC 통신이상', '누액 (Leak)',
    '밸브 동작불량', '공정 타임아웃 (진행 안됨)', '카메라 인식오류', '도어 개폐 불량',
    '기타'
]

PARTS = ['센서', '실린더', '히터', '베어링', '벨트', '멤브레인 시트', '포팅스타']

GRADES = [('A', 'A : 생산정지'), ('B', 'B : 품질영향'), ('C', 'C : 경미고장'), ('D', 'D : 기타')]
STATUSES = ['미조치', '진행중', '대기', '완료']

# ─── 관리자 비밀번호 (해시로 저장, 원래 값은 synopex) ──────────
ADMIN_CONFIG_PATH = os.path.join(BASE_DIR, 'admin_config.json')
_DEFAULT_ADMIN_PASSWORD = 'synopex'

def _load_admin_password_hash():
    try:
        with open(ADMIN_CONFIG_PATH, encoding='utf-8') as f:
            return json.load(f)['password_hash']
    except Exception:
        h = generate_password_hash(_DEFAULT_ADMIN_PASSWORD)
        with open(ADMIN_CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump({'password_hash': h}, f)
        return h

def check_admin_password(pw):
    return check_password_hash(_load_admin_password_hash(), pw or '')


# ─── 카카오톡 알림 연동 (A등급 고장 발생 시) ──────────────────
KAKAO_CONFIG_PATH   = os.path.join(BASE_DIR, 'kakao_config.json')
KAKAO_AUTHORIZE_URL = 'https://kauth.kakao.com/oauth/authorize'
KAKAO_TOKEN_URL     = 'https://kauth.kakao.com/oauth/token'
KAKAO_SEND_URL      = 'https://kapi.kakao.com/v2/api/talk/memo/default/send'
KAKAO_USER_ME_URL   = 'https://kapi.kakao.com/v2/user/me'

def _kakao_load():
    try:
        with open(KAKAO_CONFIG_PATH, encoding='utf-8') as f:
            cfg = json.load(f)
    except Exception:
        cfg = {}
    cfg.setdefault('accounts', [])
    return cfg

def _kakao_save(cfg):
    with open(KAKAO_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def _kakao_connected():
    return bool(_kakao_load().get('accounts'))

def _kakao_authorize_url():
    cfg = _kakao_load()
    params = {
        'client_id': cfg.get('rest_api_key', ''),
        'redirect_uri': cfg.get('redirect_uri', ''),
        'response_type': 'code',
        'scope': 'talk_message',
    }
    return KAKAO_AUTHORIZE_URL + '?' + urlencode(params)

def _kakao_post_form(url, data, headers=None):
    body = urlencode(data).encode('utf-8')
    req = urllib.request.Request(url, data=body, method='POST')
    req.add_header('Content-Type', 'application/x-www-form-urlencoded;charset=utf-8')
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try:
            return json.loads(e.read().decode('utf-8'))
        except Exception:
            return {'error': f'http_{e.code}'}

def _kakao_get_profile(access_token):
    req = urllib.request.Request(KAKAO_USER_ME_URL, method='GET')
    req.add_header('Authorization', f'Bearer {access_token}')
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        kakao_id = str(data.get('id', ''))
        nickname = (data.get('properties') or {}).get('nickname', '') or f'사용자 {kakao_id[-4:]}'
        return kakao_id, nickname
    except Exception:
        logging.exception('kakao profile fetch failed')
        return None, None

def _kakao_exchange_code(code):
    cfg = _kakao_load()
    data = {
        'grant_type': 'authorization_code',
        'client_id': cfg.get('rest_api_key', ''),
        'redirect_uri': cfg.get('redirect_uri', ''),
        'code': code,
    }
    if cfg.get('client_secret'):
        data['client_secret'] = cfg['client_secret']
    result = _kakao_post_form(KAKAO_TOKEN_URL, data)
    if 'access_token' not in result:
        return result, None

    kakao_id, nickname = _kakao_get_profile(result['access_token'])
    account = {
        'kakao_id': kakao_id or f'user_{int(time.time())}',
        'nickname': nickname or '이름 없음',
        'access_token': result['access_token'],
        'refresh_token': result.get('refresh_token'),
        'token_expires_at': time.time() + result.get('expires_in', 21599) - 60,
    }
    accounts = [a for a in cfg['accounts'] if a.get('kakao_id') != account['kakao_id']]
    accounts.append(account)
    cfg['accounts'] = accounts
    _kakao_save(cfg)
    return result, account

def _kakao_refresh_account(account):
    cfg = _kakao_load()
    if not account.get('refresh_token'):
        return None
    data = {
        'grant_type': 'refresh_token',
        'client_id': cfg.get('rest_api_key', ''),
        'refresh_token': account['refresh_token'],
    }
    if cfg.get('client_secret'):
        data['client_secret'] = cfg['client_secret']
    result = _kakao_post_form(KAKAO_TOKEN_URL, data)
    if 'access_token' not in result:
        logging.warning('kakao refresh failed for %s: %s', account.get('nickname'), result)
        return None
    account['access_token'] = result['access_token']
    if 'refresh_token' in result:
        account['refresh_token'] = result['refresh_token']
    account['token_expires_at'] = time.time() + result.get('expires_in', 21599) - 60
    for i, a in enumerate(cfg['accounts']):
        if a.get('kakao_id') == account.get('kakao_id'):
            cfg['accounts'][i] = account
    _kakao_save(cfg)
    return account['access_token']

def _kakao_valid_token_for(account):
    if time.time() >= (account.get('token_expires_at') or 0):
        return _kakao_refresh_account(account)
    return account.get('access_token')

def _kakao_remove_account(kakao_id):
    cfg = _kakao_load()
    cfg['accounts'] = [a for a in cfg['accounts'] if a.get('kakao_id') != kakao_id]
    _kakao_save(cfg)

def _kakao_send_to_account(account, text, link_url):
    token = _kakao_valid_token_for(account)
    if not token:
        return False, '토큰 갱신 실패'
    template = {
        'object_type': 'text',
        'text': text,
        'link': {'web_url': link_url, 'mobile_web_url': link_url},
    }
    result = _kakao_post_form(
        KAKAO_SEND_URL,
        {'template_object': json.dumps(template, ensure_ascii=False)},
        headers={'Authorization': f'Bearer {token}'}
    )
    if result.get('result_code') == 0:
        return True, result
    logging.warning('kakao send failed for %s: %s', account.get('nickname'), result)
    return False, result

def _kakao_send_to_all(text, link_url):
    cfg = _kakao_load()
    results = []
    for account in cfg.get('accounts', []):
        ok, result = _kakao_send_to_account(account, text, link_url)
        results.append((account.get('nickname'), ok, result))
    return results

GRADE_NOTIFY_LABEL = {
    'A': ('🚨', 'A등급 (생산정지)'),
    'B': ('⚠️', 'B등급 (품질영향)'),
    'C': ('🔧', 'C등급 (경미고장)'),
}

def notify_fault_by_grade(grade, fault_id, eq_number, eq_name, symptom, status=None, worker=None, action_detail=None):
    if grade not in GRADE_NOTIFY_LABEL or not _kakao_connected():
        return
    emoji, label = GRADE_NOTIFY_LABEL[grade]
    ip = get_local_ip()
    link_url = f'http://{ip}:5001/fault/{fault_id}'
    text = (f'{emoji} {label} 고장 발생\n'
            f'설비: {eq_number} {eq_name}\n'
            f'증상: {symptom or "-"}\n'
            f'처리상태: {status or "미조치"}\n'
            f'작업자: {worker or "-"}\n'
            f'조치내용: {action_detail or "-"}\n'
            f'상세보기에서 확인해주세요.')
    try:
        _kakao_send_to_all(text, link_url)
    except Exception:
        logging.exception('notify_fault_by_grade failed')


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


PHOTO_MAX_DIM = 1600
PHOTO_JPEG_QUALITY = 80

def save_and_compress_image(file_storage, dest_path):
    img = Image.open(file_storage.stream)
    img = img.convert('RGB') if img.mode != 'RGB' else img
    w, h = img.size
    if max(w, h) > PHOTO_MAX_DIM:
        scale = PHOTO_MAX_DIM / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    img.save(dest_path, 'JPEG', quality=PHOTO_JPEG_QUALITY, optimize=True)


def last_n_months(n=12):
    """오늘 기준 최근 n개월을 'YYYY-MM' 형식으로, 최신순으로 반환 (데이터 유무와 무관)."""
    months = []
    today = date.today()
    y, m = today.year, today.month
    for i in range(n):
        mm = m - i
        yy = y
        while mm <= 0:
            mm += 12
            yy -= 1
        months.append(f'{yy:04d}-{mm:02d}')
    return months


def get_workers(db):
    rows = db.execute("SELECT name FROM workers ORDER BY name").fetchall()
    return [r['name'] for r in rows]


def register_worker(db, name):
    if name:
        db.execute("INSERT OR IGNORE INTO workers (name) VALUES (?)", (name,))


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


def parse_occurred_at(raw):
    """datetime-local 입력값('YYYY-MM-DDTHH:MM')을 DB 저장 형식으로 변환. 값이 없으면 현재 시각."""
    raw = (raw or '').strip()
    if not raw:
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    raw = raw.replace('T', ' ')
    if len(raw) == 16:
        raw += ':00'
    return raw


def generate_qr(eq_id, host_url=None):
    # 항상 LAN IP 사용 — 모바일(휴대폰)에서도 스캔 가능하도록
    ip = get_local_ip()
    url = f"http://{ip}:5001/fault/register/{eq_id}"
    qr = qrcode.QRCode(version=1, box_size=10, border=4,
                        error_correction=qrcode.constants.ERROR_CORRECT_H)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color='black', back_color='white')
    path = os.path.join(QR_DIR, f'eq_{eq_id}.png')
    img.save(path)
    return path


# ──────────────────────────────────────────────
# 대시보드
# ──────────────────────────────────────────────
@app.route('/')
def index():
    db = get_db()

    total     = db.execute("SELECT COUNT(*) FROM fault_history").fetchone()[0]
    this_week = db.execute("""
        SELECT COUNT(*) FROM fault_history
        WHERE strftime('%Y-%W', occurred_at) = strftime('%Y-%W', 'now','localtime')
    """).fetchone()[0]
    this_month = db.execute("""
        SELECT COUNT(*) FROM fault_history
        WHERE strftime('%Y-%m', occurred_at) = strftime('%Y-%m', 'now','localtime')
    """).fetchone()[0]

    top_eq = db.execute("""
        SELECT eq_number, eq_name, COUNT(*) as cnt
        FROM fault_history GROUP BY equipment_id
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    top_symptom = db.execute("""
        SELECT symptom, COUNT(*) as cnt
        FROM fault_history GROUP BY symptom
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    top_worker = db.execute("""
        SELECT worker, COUNT(*) as cnt
        FROM fault_history WHERE worker IS NOT NULL AND worker != ''
        GROUP BY worker ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    top_parts = db.execute("""
        SELECT part_name, SUM(quantity) as cnt
        FROM used_parts GROUP BY part_name
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    completed_parts = db.execute("""
        SELECT up.part_name, SUM(up.quantity) as cnt, COUNT(DISTINCT up.fault_id) as cases
        FROM used_parts up
        JOIN fault_history f ON up.fault_id = f.id
        WHERE f.status = '완료'
        GROUP BY up.part_name ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    grade_cnt = db.execute("""
        SELECT grade, COUNT(*) as cnt
        FROM fault_history GROUP BY grade ORDER BY grade
    """).fetchall()

    weekly = db.execute("""
        SELECT strftime('%Y-W%W', occurred_at) as week, COUNT(*) as cnt
        FROM fault_history
        WHERE occurred_at >= date('now','localtime','-12 weeks')
        GROUP BY week ORDER BY week
    """).fetchall()

    monthly = db.execute("""
        SELECT strftime('%Y-%m', occurred_at) as month, COUNT(*) as cnt
        FROM fault_history
        WHERE occurred_at >= date('now','localtime','-12 months')
        GROUP BY month ORDER BY month
    """).fetchall()

    available_months = [
        {'value': mstr, 'label': f"{mstr.split('-')[0]}년 {int(mstr.split('-')[1])}월"}
        for mstr in last_n_months(12)
    ]

    selected_month = request.args.get('month', '').strip() or date.today().strftime('%Y-%m')
    sy, sm = selected_month.split('-')
    selected_month_label = f"{sy}년 {int(sm)}월"

    _, last_day = calendar.monthrange(int(sy), int(sm))
    day_rows = db.execute("""
        SELECT CAST(strftime('%d', occurred_at) AS INTEGER) as d, COUNT(*) as cnt
        FROM fault_history
        WHERE strftime('%Y-%m', occurred_at) = ?
        GROUP BY d
    """, (selected_month,)).fetchall()
    day_counts = {r['d']: r['cnt'] for r in day_rows}
    daily_trend = [{'day': f'{d}일', 'cnt': day_counts.get(d, 0)} for d in range(1, last_day + 1)]

    weekly_trend = db.execute("""
        SELECT strftime('%Y-W%W', occurred_at) as week, COUNT(*) as cnt
        FROM fault_history
        WHERE strftime('%Y-%m', occurred_at) = ?
        GROUP BY week ORDER BY week
    """, (selected_month,)).fetchall()

    eq_chart = db.execute("""
        SELECT eq_number || ' ' || eq_name as label, COUNT(*) as cnt
        FROM fault_history GROUP BY equipment_id
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    symptom_chart = db.execute("""
        SELECT symptom as label, COUNT(*) as cnt
        FROM fault_history GROUP BY symptom
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    parts_chart = db.execute("""
        SELECT part_name as label, SUM(quantity) as cnt
        FROM used_parts GROUP BY part_name
        ORDER BY cnt DESC LIMIT 10
    """).fetchall()

    status_cnt = db.execute("""
        SELECT status, COUNT(*) as cnt FROM fault_history
        GROUP BY status
    """).fetchall()

    db.close()

    today_dt = date.today()
    week_start = (today_dt - timedelta(days=today_dt.weekday())).isoformat()
    month_start = today_dt.replace(day=1).isoformat()

    return render_template('index.html',
        total=total, this_week=this_week, this_month=this_month,
        top_eq=top_eq, top_symptom=top_symptom,
        top_worker=top_worker, top_parts=top_parts,
        completed_parts=completed_parts,
        grade_cnt=grade_cnt, status_cnt=status_cnt,
        weekly=weekly, monthly=monthly,
        available_months=available_months,
        selected_month=selected_month, selected_month_label=selected_month_label,
        daily_trend=daily_trend, weekly_trend=weekly_trend,
        eq_chart=eq_chart, symptom_chart=symptom_chart,
        parts_chart=parts_chart,
        week_start=week_start, month_start=month_start,
        kakao_connected=_kakao_connected(), kakao_accounts=_kakao_load().get('accounts', []))


# ──────────────────────────────────────────────
# 통합검색 (설비 + 고장이력)
# ──────────────────────────────────────────────
@app.route('/search')
def global_search():
    q = request.args.get('q', '').strip()
    equipments, faults = [], []
    if q:
        db = get_db()
        like = f'%{q}%'
        equipments = db.execute("""
            SELECT * FROM equipment
            WHERE eq_number LIKE ? OR eq_name LIKE ? OR location LIKE ? OR section LIKE ?
            ORDER BY section, eq_number
        """, [like, like, like, like]).fetchall()
        faults = db.execute("""
            SELECT * FROM fault_history
            WHERE eq_number LIKE ? OR eq_name LIKE ? OR symptom LIKE ? OR worker LIKE ? OR cause LIKE ?
            ORDER BY occurred_at DESC LIMIT 50
        """, [like, like, like, like, like]).fetchall()
        db.close()
    return render_template('search.html', q=q, equipments=equipments, faults=faults)


# ──────────────────────────────────────────────
# 설비 관리
# ──────────────────────────────────────────────
@app.route('/equipment/list')
def equipment_list():
    q = request.args.get('q', '').strip()
    sec = request.args.get('section', '').strip()
    db = get_db()
    sql = "SELECT * FROM equipment WHERE 1=1"
    params = []
    if q:
        sql += " AND (eq_number LIKE ? OR eq_name LIKE ? OR location LIKE ?)"
        params += [f'%{q}%', f'%{q}%', f'%{q}%']
    if sec:
        sql += " AND section = ?"
        params.append(sec)
    sql += " ORDER BY section, eq_number"
    equipments = db.execute(sql, params).fetchall()
    db.close()
    return render_template('equipment/list.html',
                           equipments=equipments, sections=SECTIONS, q=q, sec=sec)


@app.route('/equipment/add', methods=['GET', 'POST'])
def equipment_add():
    if request.method == 'POST':
        eq_number = request.form.get('eq_number', '').strip()
        eq_name   = request.form.get('eq_name', '').strip()
        section   = request.form.get('section', '').strip()
        location  = request.form.get('location', '').strip()
        note      = request.form.get('note', '').strip()

        if not eq_number or not eq_name or not section:
            flash('설비번호, 설비명, 섹션은 필수입니다.', 'danger')
            return render_template('equipment/add.html', sections=SECTIONS)

        db = get_db()
        cur = db.execute(
            "INSERT INTO equipment (eq_number,eq_name,section,location,note) VALUES (?,?,?,?,?)",
            (eq_number, eq_name, section, location, note)
        )
        eq_id = cur.lastrowid
        db.commit()
        db.close()

        host_url = request.host_url
        generate_qr(eq_id, host_url)

        flash(f'설비 "{eq_name}" 등록 완료. QR코드 생성됨.', 'success')
        return redirect(url_for('equipment_list'))

    return render_template('equipment/add.html', sections=SECTIONS)


@app.route('/equipment/edit/<int:eq_id>', methods=['GET', 'POST'])
def equipment_edit(eq_id):
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    if not eq:
        db.close()
        abort(404)

    if request.method == 'POST':
        eq_number = request.form.get('eq_number', '').strip()
        eq_name   = request.form.get('eq_name', '').strip()
        section   = request.form.get('section', '').strip()
        location  = request.form.get('location', '').strip()
        note      = request.form.get('note', '').strip()

        if not eq_number or not eq_name or not section:
            flash('설비번호, 설비명, 섹션은 필수입니다.', 'danger')
            db.close()
            return render_template('equipment/edit.html', eq=eq, sections=SECTIONS)

        db.execute(
            "UPDATE equipment SET eq_number=?,eq_name=?,section=?,location=?,note=? WHERE id=?",
            (eq_number, eq_name, section, location, note, eq_id)
        )
        db.commit()
        db.close()

        host_url = request.host_url
        generate_qr(eq_id, host_url)

        flash(f'설비 정보가 수정되었습니다.', 'success')
        return redirect(url_for('equipment_list'))

    db.close()
    return render_template('equipment/edit.html', eq=eq, sections=SECTIONS)


@app.route('/equipment/delete/<int:eq_id>', methods=['POST'])
def equipment_delete(eq_id):
    if not session.get('edit_authorized'):
        flash('삭제 권한이 없습니다. 비밀번호를 입력해 주세요.', 'danger')
        return redirect(request.referrer or url_for('equipment_list'))
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    if not eq:
        db.close()
        abort(404)
    fault_count = db.execute(
        "SELECT COUNT(*) FROM fault_history WHERE equipment_id=?", (eq_id,)
    ).fetchone()[0]
    if fault_count > 0:
        db.close()
        flash(f'이 설비에는 고장이력이 {fault_count}건 있어 삭제할 수 없습니다. '
              f'먼저 고장이력을 삭제한 후 다시 시도해주세요.', 'danger')
        return redirect(url_for('equipment_detail', eq_id=eq_id))
    db.execute("DELETE FROM equipment WHERE id=?", (eq_id,))
    db.commit()
    db.close()
    qr_path = os.path.join(QR_DIR, f'eq_{eq_id}.png')
    if os.path.exists(qr_path):
        os.remove(qr_path)
    flash('설비가 삭제되었습니다.', 'warning')
    return redirect(url_for('equipment_list'))


@app.route('/equipment/<int:eq_id>')
def equipment_detail(eq_id):
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    if not eq:
        db.close()
        abort(404)
    recent_faults = db.execute(
        "SELECT * FROM fault_history WHERE equipment_id=? ORDER BY occurred_at DESC LIMIT 5",
        (eq_id,)
    ).fetchall()
    fault_count = db.execute(
        "SELECT COUNT(*) FROM fault_history WHERE equipment_id=?", (eq_id,)
    ).fetchone()[0]
    db.close()
    return render_template('equipment/detail.html',
                           eq=eq, recent_faults=recent_faults, fault_count=fault_count)


# ──────────────────────────────────────────────
# QR 생성 / 출력
# ──────────────────────────────────────────────
@app.route('/qr/generate/<int:eq_id>')
def qr_generate(eq_id):
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    db.close()
    if not eq:
        abort(404)
    host_url = request.host_url
    generate_qr(eq_id, host_url)
    flash('QR코드가 재생성되었습니다.', 'success')
    return redirect(url_for('qr_print', eq_id=eq_id))


@app.route('/qr/print/<int:eq_id>')
def qr_print(eq_id):
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    db.close()
    if not eq:
        abort(404)
    qr_path = os.path.join(QR_DIR, f'eq_{eq_id}.png')
    if not os.path.exists(qr_path):
        generate_qr(eq_id, request.host_url)
    return render_template('qr/print.html', eq=eq)


@app.route('/qr/image/<int:eq_id>')
def qr_image(eq_id):
    path = os.path.join(QR_DIR, f'eq_{eq_id}.png')
    if not os.path.exists(path):
        db = get_db()
        eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
        db.close()
        if not eq:
            abort(404)
        generate_qr(eq_id, request.host_url)
    return send_file(path, mimetype='image/png')


@app.route('/qr/all')
def qr_all():
    db = get_db()
    sec = request.args.get('section', '').strip()
    sql = "SELECT * FROM equipment"
    params = []
    if sec:
        sql += " WHERE section=?"
        params.append(sec)
    sql += " ORDER BY section, eq_number"
    equipments = db.execute(sql, params).fetchall()
    db.close()
    for eq in equipments:
        qr_path = os.path.join(QR_DIR, f'eq_{eq["id"]}.png')
        if not os.path.exists(qr_path):
            generate_qr(eq['id'], request.host_url)
    return render_template('qr/print_all.html', equipments=equipments, sections=SECTIONS, sec=sec)


# ──────────────────────────────────────────────
# 고장 등록 — 설비 선택 (대시보드 직접 접근용)
# ──────────────────────────────────────────────
@app.route('/fault/select')
def fault_select():
    q   = request.args.get('q', '').strip()
    sec = request.args.get('section', '').strip()
    db  = get_db()
    sql = "SELECT * FROM equipment WHERE 1=1"
    params = []
    if q:
        sql += " AND (eq_number LIKE ? OR eq_name LIKE ? OR location LIKE ?)"
        params += [f'%{q}%', f'%{q}%', f'%{q}%']
    if sec:
        sql += " AND section = ?"
        params.append(sec)
    sql += " ORDER BY section, eq_number"
    equipments = db.execute(sql, params).fetchall()
    db.close()
    return render_template('fault/select_equipment.html',
                           equipments=equipments, sections=SECTIONS, q=q, sec=sec)


# ──────────────────────────────────────────────
# 고장 등록
# ──────────────────────────────────────────────
@app.route('/fault/register/<int:eq_id>', methods=['GET', 'POST'])
def fault_register(eq_id):
    db = get_db()
    eq = db.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    if not eq:
        db.close()
        abort(404)

    if request.method == 'POST':
        symptom_list = request.form.getlist('symptom[]')
        symptom      = ', '.join([s for s in symptom_list if s])
        cause        = request.form.get('cause', '').strip()
        action_detail= request.form.get('action_detail', '').strip()
        worker_sel   = request.form.get('worker_select', '').strip()
        worker_custom= request.form.get('worker_custom', '').strip()
        worker       = worker_custom if worker_sel == '__new__' else worker_sel
        completed_at = request.form.get('completed_at', '').strip() or None
        grade        = request.form.get('grade', '').strip()
        grade_note   = request.form.get('grade_note', '').strip() if grade == 'D' else None
        status       = request.form.get('status', '미조치').strip()
        status_note  = request.form.get('status_note', '').strip() if status == '대기' else None
        occurred_at  = parse_occurred_at(request.form.get('occurred_at', ''))

        if worker:
            register_worker(db, worker)

        cur = db.execute("""
            INSERT INTO fault_history
            (equipment_id,eq_number,eq_name,occurred_at,symptom,cause,action_detail,worker,completed_at,grade,grade_note,photo_fault,photo_action,status,status_note)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (eq_id, eq['eq_number'], eq['eq_name'], occurred_at,
              symptom, cause, action_detail, worker, completed_at, grade, grade_note,
              None, None, status, status_note))
        fault_id = cur.lastrowid

        # 고장 사진 저장 (최대 5장)
        photo_save_failed = False
        def save_photos(files_key, labels_key, ptype):
            nonlocal photo_save_failed
            files  = request.files.getlist(files_key)
            labels = request.form.getlist(labels_key)
            for i, pf in enumerate(files[:5]):
                if pf and pf.filename and allowed_file(pf.filename):
                    label = (labels[i] if i < len(labels) else '').strip()
                    ts = datetime.now().strftime('%Y%m%d%H%M%S%f')[:17]
                    base = os.path.splitext(secure_filename(pf.filename))[0]
                    fn = f"p{eq_id}_{ptype}_{ts}_{i}_{base}.jpg"
                    try:
                        save_and_compress_image(pf, os.path.join(PHOTOS_DIR, fn))
                        db.execute(
                            "INSERT INTO fault_photos (fault_id,filename,label,photo_type) VALUES (?,?,?,?)",
                            (fault_id, fn, label, ptype)
                        )
                    except Exception:
                        logging.exception('photo save failed: %s', fn)
                        photo_save_failed = True

        save_photos('fault_photos[]', 'fault_photo_labels[]', 'fault')
        save_photos('action_photos[]', 'action_photo_labels[]', 'action')

        part_names = request.form.getlist('part_name[]')
        part_qtys  = request.form.getlist('part_qty[]')
        part_notes = request.form.getlist('part_note[]')
        for pname, pqty, pnote in zip(part_names, part_qtys, part_notes):
            pname = pname.strip()
            if pname:
                try:
                    qty = int(pqty) if pqty else 1
                except ValueError:
                    qty = 1
                db.execute(
                    "INSERT INTO used_parts (fault_id,part_name,quantity,note) VALUES (?,?,?,?)",
                    (fault_id, pname, qty, pnote.strip())
                )

        db.commit()
        db.close()
        notify_fault_by_grade(grade, fault_id, eq['eq_number'], eq['eq_name'], symptom, status, worker, action_detail)
        flash('고장이 등록되었습니다.', 'success')
        if photo_save_failed:
            flash('일부 사진 저장에 실패했습니다. 필요하면 수정 화면에서 다시 첨부해주세요.', 'warning')
        return redirect(url_for('equipment_detail', eq_id=eq_id))

    workers = get_workers(db)
    db.close()
    return render_template('fault/register.html',
                           eq=eq, symptoms=SYMPTOMS, grades=GRADES,
                           parts=PARTS, statuses=STATUSES, workers=workers,
                           today=date.today().isoformat(),
                           now_dt=datetime.now().strftime('%Y-%m-%dT%H:%M'))


# ──────────────────────────────────────────────
# 고장 이력 조회
# ──────────────────────────────────────────────
@app.route('/fault/list')
def fault_list():
    q         = request.args.get('q', '').strip()
    sym       = request.args.get('symptom', '').strip()
    worker    = request.args.get('worker', '').strip()
    grade     = request.args.get('grade', '').strip()
    status    = request.args.get('status', '').strip()
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    eq_id     = request.args.get('eq_id', '').strip()

    sql = "SELECT * FROM fault_history WHERE 1=1"
    params = []

    if q:
        sql += " AND (eq_number LIKE ? OR eq_name LIKE ? OR symptom LIKE ? OR worker LIKE ?)"
        params += [f'%{q}%'] * 4
    if sym:
        sql += " AND symptom LIKE ?"
        params.append(f'%{sym}%')
    if worker:
        sql += " AND worker LIKE ?"
        params.append(f'%{worker}%')
    if grade:
        sql += " AND grade=?"
        params.append(grade)
    if status:
        sql += " AND (status=? OR (status IS NULL AND ?='미조치'))"
        params += [status, status]
    if date_from:
        sql += " AND date(occurred_at) >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND date(occurred_at) <= ?"
        params.append(date_to)
    if eq_id:
        sql += " AND equipment_id=?"
        params.append(eq_id)

    PAGE_SIZE = 20
    try:
        page = max(1, int(request.args.get('page', 1)))
    except ValueError:
        page = 1

    db = get_db()
    total_count = db.execute(f"SELECT COUNT(*) c FROM ({sql})", params).fetchone()['c']
    total_pages = max(1, (total_count + PAGE_SIZE - 1) // PAGE_SIZE)
    page = min(page, total_pages)

    sql += " ORDER BY occurred_at DESC LIMIT ? OFFSET ?"
    params_paged = params + [PAGE_SIZE, (page - 1) * PAGE_SIZE]
    faults = db.execute(sql, params_paged).fetchall()
    workers = db.execute(
        "SELECT DISTINCT worker FROM fault_history WHERE worker IS NOT NULL AND worker != '' ORDER BY worker"
    ).fetchall()
    db.close()

    def page_url(p):
        args = request.args.to_dict()
        args['page'] = p
        return '?' + urlencode(args)

    return render_template('fault/list.html',
                           faults=faults, symptoms=SYMPTOMS, workers=workers,
                           grades=GRADES, statuses=STATUSES,
                           q=q, sym=sym, worker=worker, grade=grade,
                           status=status, date_from=date_from, date_to=date_to,
                           page=page, total_pages=total_pages, total_count=total_count,
                           page_url=page_url)


@app.route('/fault/<int:fault_id>')
def fault_detail(fault_id):
    db = get_db()
    fault = db.execute("SELECT * FROM fault_history WHERE id=?", (fault_id,)).fetchone()
    if not fault:
        db.close()
        abort(404)
    parts         = db.execute("SELECT * FROM used_parts WHERE fault_id=?", (fault_id,)).fetchall()
    fault_photos  = db.execute(
        "SELECT * FROM fault_photos WHERE fault_id=? AND photo_type='fault' ORDER BY id", (fault_id,)).fetchall()
    action_photos = db.execute(
        "SELECT * FROM fault_photos WHERE fault_id=? AND photo_type='action' ORDER BY id", (fault_id,)).fetchall()
    db.close()
    return render_template('fault/detail.html', fault=fault, parts=parts,
                           fault_photos=fault_photos, action_photos=action_photos)


# ──────────────────────────────────────────────
# 수정 권한 인증
# ──────────────────────────────────────────────
@app.route('/fault/auth', methods=['POST'])
def fault_auth():
    pw       = request.form.get('password', '').strip()
    next_url = request.form.get('next', '/')
    if check_admin_password(pw):
        session['edit_authorized'] = True
        return redirect(next_url)
    flash('비밀번호가 올바르지 않습니다.', 'danger')
    return redirect(request.referrer or url_for('fault_list'))


# ──────────────────────────────────────────────
# 고장 이력 수정
# ──────────────────────────────────────────────
@app.route('/fault/edit/<int:fault_id>', methods=['GET', 'POST'])
def fault_edit(fault_id):
    if not session.get('edit_authorized'):
        flash('수정 권한이 없습니다. 비밀번호를 입력해 주세요.', 'danger')
        return redirect(url_for('fault_detail', fault_id=fault_id))

    db = get_db()
    fault = db.execute("SELECT * FROM fault_history WHERE id=?", (fault_id,)).fetchone()
    if not fault:
        db.close()
        abort(404)

    if request.method == 'POST':
        symptom_list = request.form.getlist('symptom[]')
        symptom      = ', '.join([s for s in symptom_list if s])
        cause        = request.form.get('cause', '').strip()
        action_detail= request.form.get('action_detail', '').strip()
        worker_sel   = request.form.get('worker_select', '').strip()
        worker_custom= request.form.get('worker_custom', '').strip()
        worker       = worker_custom if worker_sel == '__new__' else worker_sel
        completed_at = request.form.get('completed_at', '').strip() or None
        grade        = request.form.get('grade', '').strip()
        grade_note   = request.form.get('grade_note', '').strip() if grade == 'D' else None
        status       = request.form.get('status', '미조치').strip()
        status_note  = request.form.get('status_note', '').strip() if status == '대기' else None
        occurred_at  = parse_occurred_at(request.form.get('occurred_at', ''))

        if worker:
            register_worker(db, worker)

        # 기존 사진 삭제 요청 처리
        delete_ids = request.form.getlist('delete_photo[]')
        for pid in delete_ids:
            p = db.execute("SELECT filename FROM fault_photos WHERE id=? AND fault_id=?",
                           (pid, fault_id)).fetchone()
            if p:
                fp = os.path.join(PHOTOS_DIR, p['filename'])
                if os.path.exists(fp):
                    os.remove(fp)
                db.execute("DELETE FROM fault_photos WHERE id=?", (pid,))

        db.execute("""
            UPDATE fault_history SET
              occurred_at=?, symptom=?, cause=?, action_detail=?, worker=?,
              completed_at=?, grade=?, grade_note=?, status=?, status_note=?
            WHERE id=?
        """, (occurred_at, symptom, cause, action_detail, worker, completed_at,
              grade, grade_note, status, status_note, fault_id))

        # 새 사진 추가 (고장/완료 각 타입별 최대 5장)
        eq_id_edit = fault['equipment_id']
        photo_save_failed = False
        def save_edit_photos(files_key, labels_key, ptype):
            nonlocal photo_save_failed
            files  = request.files.getlist(files_key)
            labels = request.form.getlist(labels_key)
            exist = db.execute(
                "SELECT COUNT(*) FROM fault_photos WHERE fault_id=? AND photo_type=?",
                (fault_id, ptype)).fetchone()[0]
            for i, pf in enumerate(files):
                if exist >= 5:
                    break
                if pf and pf.filename and allowed_file(pf.filename):
                    label = (labels[i] if i < len(labels) else '').strip()
                    ts = datetime.now().strftime('%Y%m%d%H%M%S%f')[:17]
                    base = os.path.splitext(secure_filename(pf.filename))[0]
                    fn = f"p{eq_id_edit}_{ptype}_{ts}_{i}_{base}.jpg"
                    try:
                        save_and_compress_image(pf, os.path.join(PHOTOS_DIR, fn))
                        db.execute(
                            "INSERT INTO fault_photos (fault_id,filename,label,photo_type) VALUES (?,?,?,?)",
                            (fault_id, fn, label, ptype))
                        exist += 1
                    except Exception:
                        logging.exception('photo save failed: %s', fn)
                        photo_save_failed = True

        save_edit_photos('fault_photos[]', 'fault_photo_labels[]', 'fault')
        save_edit_photos('action_photos[]', 'action_photo_labels[]', 'action')
        db.commit()
        db.close()
        if grade != fault['grade']:
            notify_fault_by_grade(grade, fault_id, fault['eq_number'], fault['eq_name'], symptom, status, worker, action_detail)
        flash('고장 이력이 수정되었습니다.', 'success')
        if photo_save_failed:
            flash('일부 사진 저장에 실패했습니다. 필요하면 다시 첨부해주세요.', 'warning')
        return redirect(url_for('fault_detail', fault_id=fault_id))

    parts         = db.execute("SELECT * FROM used_parts WHERE fault_id=?", (fault_id,)).fetchall()
    fault_photos  = db.execute(
        "SELECT * FROM fault_photos WHERE fault_id=? AND photo_type='fault' ORDER BY id", (fault_id,)).fetchall()
    action_photos = db.execute(
        "SELECT * FROM fault_photos WHERE fault_id=? AND photo_type='action' ORDER BY id", (fault_id,)).fetchall()
    workers = get_workers(db)
    db.close()
    selected_symptoms = [s.strip() for s in (fault['symptom'] or '').split(',') if s.strip()]
    occurred_at_val = (fault['occurred_at'] or '').replace(' ', 'T')[:16]
    return render_template('fault/edit.html',
                           fault=fault, parts=parts,
                           fault_photos=fault_photos, action_photos=action_photos,
                           symptoms=SYMPTOMS, grades=GRADES,
                           parts_list=PARTS, statuses=STATUSES, workers=workers,
                           selected_symptoms=selected_symptoms,
                           occurred_at_val=occurred_at_val,
                           today=date.today().isoformat())


@app.route('/fault/delete/<int:fault_id>', methods=['POST'])
def fault_delete(fault_id):
    if not session.get('edit_authorized'):
        flash('삭제 권한이 없습니다. 비밀번호를 입력해 주세요.', 'danger')
        return redirect(request.referrer or url_for('fault_list'))
    db = get_db()
    fault = db.execute("SELECT * FROM fault_history WHERE id=?", (fault_id,)).fetchone()
    if not fault:
        db.close()
        abort(404)
    eq_id = fault['equipment_id']
    for fname in [fault['photo_fault'], fault['photo_action']]:
        if fname:
            for d in [FAULT_DIR, ACTION_DIR]:
                fp = os.path.join(d, fname)
                if os.path.exists(fp):
                    os.remove(fp)
    # 다중 사진 파일 삭제
    old_photos = db.execute("SELECT filename FROM fault_photos WHERE fault_id=?", (fault_id,)).fetchall()
    for p in old_photos:
        fp = os.path.join(PHOTOS_DIR, p['filename'])
        if os.path.exists(fp):
            os.remove(fp)
    db.execute("DELETE FROM fault_history WHERE id=?", (fault_id,))
    db.commit()
    db.close()
    flash('고장 이력이 삭제되었습니다.', 'warning')
    return redirect(url_for('equipment_detail', eq_id=eq_id))


# ──────────────────────────────────────────────
# 분석
# ──────────────────────────────────────────────
@app.route('/analysis/repeat')
def analysis_repeat():
    db = get_db()
    eq_stat = db.execute("""
        SELECT f.eq_number, f.eq_name, e.section,
               COUNT(*) as total,
               SUM(CASE WHEN f.grade='A' THEN 1 ELSE 0 END) as grade_a,
               SUM(CASE WHEN f.grade='B' THEN 1 ELSE 0 END) as grade_b,
               SUM(CASE WHEN f.grade='C' THEN 1 ELSE 0 END) as grade_c
        FROM fault_history f
        LEFT JOIN equipment e ON f.equipment_id = e.id
        GROUP BY f.equipment_id ORDER BY total DESC
    """).fetchall()

    sym_stat = db.execute("""
        SELECT symptom, COUNT(*) as total FROM fault_history
        GROUP BY symptom ORDER BY total DESC
    """).fetchall()

    parts_stat = db.execute("""
        SELECT part_name, SUM(quantity) as total, COUNT(DISTINCT fault_id) as cases
        FROM used_parts GROUP BY part_name ORDER BY total DESC
    """).fetchall()

    grade_stat = db.execute("""
        SELECT grade, COUNT(*) as total FROM fault_history
        GROUP BY grade ORDER BY grade
    """).fetchall()

    db.close()
    return render_template('analysis/repeat.html',
                           eq_stat=eq_stat, sym_stat=sym_stat,
                           parts_stat=parts_stat, grade_stat=grade_stat)


@app.route('/analysis/trend')
def analysis_trend():
    db = get_db()
    period = request.args.get('period', 'monthly')
    section = request.args.get('section', '')

    base_sql = "FROM fault_history WHERE 1=1"
    params = []
    if section:
        base_sql += " AND equipment_id IN (SELECT id FROM equipment WHERE section=?)"
        params.append(section)

    if period == 'weekly':
        trend = db.execute(f"""
            SELECT strftime('%Y-W%W', occurred_at) as label, COUNT(*) as cnt
            {base_sql}
            GROUP BY label ORDER BY label DESC LIMIT 24
        """, params).fetchall()
    else:
        trend = db.execute(f"""
            SELECT strftime('%Y-%m', occurred_at) as label, COUNT(*) as cnt
            {base_sql}
            GROUP BY label ORDER BY label DESC LIMIT 24
        """, params).fetchall()

    trend = list(reversed(trend))

    label_col = "strftime('%Y-W%W', occurred_at)" if period == 'weekly' else "strftime('%Y-%m', occurred_at)"
    grade_rows = db.execute(f"""
        SELECT {label_col} as label, COALESCE(grade, 'D') as grade, COUNT(*) as cnt
        {base_sql}
        GROUP BY label, grade
    """, params).fetchall()
    grade_map = {}
    for r in grade_rows:
        grade_map.setdefault(r['label'], {})[r['grade']] = r['cnt']
    trend_by_grade = {
        g: [grade_map.get(t['label'], {}).get(g, 0) for t in trend]
        for g in ('A', 'B', 'C', 'D')
    }

    eq_trend = db.execute(f"""
        SELECT eq_number || ' ' || eq_name as label, COUNT(*) as cnt
        {base_sql}
        GROUP BY equipment_id ORDER BY cnt DESC LIMIT 10
    """, params).fetchall()

    db.close()
    return render_template('analysis/trend.html',
                           trend=trend, eq_trend=eq_trend, trend_by_grade=trend_by_grade,
                           period=period, section=section, sections=SECTIONS)


# ──────────────────────────────────────────────
# 엑셀 다운로드
# ──────────────────────────────────────────────
@app.route('/export/excel')
def export_excel():
    date_from = request.args.get('date_from', '').strip()
    date_to   = request.args.get('date_to', '').strip()
    eq_id     = request.args.get('eq_id', '').strip()
    symptom   = request.args.get('symptom', '').strip()

    sql = "SELECT * FROM fault_history WHERE 1=1"
    params = []
    if date_from:
        sql += " AND date(occurred_at) >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND date(occurred_at) <= ?"
        params.append(date_to)
    if eq_id:
        sql += " AND equipment_id=?"
        params.append(eq_id)
    if symptom:
        sql += " AND symptom=?"
        params.append(symptom)
    sql += " ORDER BY occurred_at DESC"

    db = get_db()
    faults = db.execute(sql, params).fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '고장이력'

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ['No', '설비번호', '설비명', '발생일시', '완료일',
               '증상', '원인', '조치내용', '작업자', '고장등급', '사용부품']
    col_widths = [5, 12, 20, 18, 12, 20, 25, 30, 10, 10, 25]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = w

    ws.row_dimensions[1].height = 22

    for ri, fault in enumerate(faults, 2):
        parts_rows = db.execute(
            "SELECT part_name, quantity FROM used_parts WHERE fault_id=?", (fault['id'],)
        ).fetchall()
        parts_str = ', '.join([f"{p['part_name']}x{p['quantity']}" for p in parts_rows])

        row_data = [
            ri - 1,
            fault['eq_number'],
            fault['eq_name'],
            fault['occurred_at'],
            fault['completed_at'] or '',
            fault['symptom'] or '',
            fault['cause'] or '',
            fault['action_detail'] or '',
            fault['worker'] or '',
            fault['grade'] or '',
            parts_str
        ]
        fill_color = 'FFE0E0' if fault['grade'] == 'A' else ('FFFDE0' if fault['grade'] == 'B' else 'E8F5E9')
        row_fill = PatternFill("solid", fgColor=fill_color)

        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = center
            cell.border = border
            cell.fill = row_fill

    db.close()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    fname = f"고장이력_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ──────────────────────────────────────────────
# 백업
# ──────────────────────────────────────────────
BACKUP_DIR  = os.path.join(BASE_DIR, 'equipment_qr_manager_backup')
BACKUP_LOG  = os.path.join(BASE_DIR, 'backup_log.json')
BACKUP_ITEMS = ['app.py', 'database.py', 'requirements.txt', '실행.bat',
                'equipment_qr.db', 'templates', 'static', 'uploads']
BACKUP_EXCLUDE = {'__pycache__', 'backup_log.json', 'equipment_qr_manager_backup'}

def _ignore_patterns(dir, contents):
    return {c for c in contents if c in BACKUP_EXCLUDE}

def _read_backup_log():
    if os.path.exists(BACKUP_LOG):
        try:
            with open(BACKUP_LOG, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {'last': None}

@app.route('/backup/status')
def backup_status():
    return jsonify(_read_backup_log())

@app.route('/backup', methods=['POST'])
def do_backup():
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        for item in BACKUP_ITEMS:
            src = os.path.join(BASE_DIR, item)
            dst = os.path.join(BACKUP_DIR, item)
            if not os.path.exists(src):
                continue
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)
                shutil.copytree(src, dst, ignore=_ignore_patterns)
            else:
                shutil.copy2(src, dst)
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(BACKUP_LOG, 'w', encoding='utf-8') as f:
            json.dump({'last': ts}, f, ensure_ascii=False)
        return jsonify({'ok': True, 'last': ts})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


# ──────────────────────────────────────────────
# 카카오톡 알림 연동
# ──────────────────────────────────────────────
@app.route('/kakao/login')
def kakao_login():
    if not session.get('edit_authorized'):
        flash('카카오 연동은 관리자 인증이 필요합니다. 비밀번호를 입력해 주세요.', 'danger')
        return redirect(request.referrer or url_for('index'))
    cfg = _kakao_load()
    if not cfg.get('rest_api_key'):
        flash('카카오 REST API 키가 설정되지 않았습니다.', 'danger')
        return redirect(url_for('index'))
    return redirect(_kakao_authorize_url())

@app.route('/kakao/callback')
def kakao_callback():
    error = request.args.get('error')
    code  = request.args.get('code')
    if error or not code:
        flash('카카오 인증 실패: ' + (error or '인증 코드 없음'), 'danger')
        return redirect(url_for('index'))
    result, account = _kakao_exchange_code(code)
    logging.info('kakao_exchange result: %s', result)
    if account:
        flash(f'✅ {account["nickname"]}님, 카카오톡 알림 연동이 완료됐습니다.', 'success')
        try:
            ip = get_local_ip()
            _kakao_send_to_account(account, '✅ 설비 QR 이력관리 시스템 카카오톡 알림이 연결되었습니다.', f'http://{ip}:5001/')
        except Exception:
            logging.exception('kakao welcome message failed')
    else:
        flash('카카오 토큰 발급 실패: ' + str(result.get('error_description', result)), 'danger')
    return redirect(url_for('index'))

@app.route('/kakao/remove/<kakao_id>', methods=['POST'])
def kakao_remove(kakao_id):
    if not session.get('edit_authorized'):
        flash('삭제 권한이 없습니다. 비밀번호를 입력해 주세요.', 'danger')
        return redirect(request.referrer or url_for('index'))
    _kakao_remove_account(kakao_id)
    flash('카카오톡 알림 대상에서 제거했습니다.', 'success')
    return redirect(url_for('index'))


# ──────────────────────────────────────────────
# 사진 서빙
# ──────────────────────────────────────────────
@app.route('/uploads/<path:subfolder>/<filename>')
def uploaded_file(subfolder, filename):
    folder = os.path.join(UPLOAD_DIR, subfolder)
    return send_file(os.path.join(folder, filename))


# ──────────────────────────────────────────────
# 에러 페이지
# ──────────────────────────────────────────────
@app.errorhandler(404)
def not_found_error(e):
    return render_template('errors/404.html'), 404

@app.errorhandler(500)
def internal_error(e):
    logging.exception('unhandled server error')
    return render_template('errors/500.html'), 500


if __name__ == '__main__':
    init_db()
    ip = get_local_ip()
    port = 5001
    print(f"\n{'='*50}")
    print(f"  설비 QR 이력관리 시스템 시작")
    print(f"  로컬 접속 : http://127.0.0.1:{port}")
    print(f"  네트워크  : http://{ip}:{port}")
    print(f"{'='*50}\n")
    try:
        from waitress import serve
        serve(app, host='0.0.0.0', port=port, threads=8)
    except ImportError:
        logging.warning('waitress가 설치되어 있지 않아 Flask 개발 서버로 대신 실행합니다. (pip install waitress 권장)')
        app.run(host='0.0.0.0', port=port, debug=False)
