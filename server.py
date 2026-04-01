"""
線上工作回報系統 - Flask 後端
資料儲存：Google Sheets（依日期分頁）
照片儲存：Google Drive（永久保留，15GB 免費）
"""

import os
import json
import io
import uuid
from datetime import datetime, timezone, timedelta
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

# ── Google API ────────────────────────────────────────────────
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SPREADSHEET_NAME = '工作回報系統'
DRIVE_FOLDER_NAME = '工作回報照片'
SCOPES = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
HEADERS = ['ID', '姓名', '日期', '開始時間', '結束時間', '地點', '原因', '解決方法', '照片URLs', '提交時間', '照片預覽']
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'heic'}
MIME_MAP = {
    'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'png': 'image/png',  'gif': 'image/gif',
    'webp': 'image/webp','heic': 'image/heic'
}

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB


# ── 工具函式 ──────────────────────────────────────────────────

def taiwan_now():
    return datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_credentials():
    """從環境變數或檔案讀取 Google 憑證"""
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if creds_json:
        return Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    if os.path.exists('credentials.json'):
        return Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    return None

def get_sheets_client():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return gspread.authorize(creds)
    except Exception as e:
        print(f"[Sheets] 連線失敗：{e}")
        return None

def get_drive_service():
    creds = get_credentials()
    if not creds:
        return None
    try:
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        print(f"[Drive] 連線失敗：{e}")
        return None

def get_or_create_drive_folder(service, folder_name, parent_id=None):
    """取得或建立 Drive 資料夾，回傳 folder_id"""
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service.files().list(q=query, fields='files(id,name)').execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    # 建立新資料夾
    meta = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        meta['parents'] = [parent_id]
    folder = service.files().create(body=meta, fields='id').execute()
    return folder.get('id')

def upload_photo_to_drive(service, file_data, filename, folder_id):
    """上傳照片到 Drive，回傳公開瀏覽 URL"""
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'jpg'
    mime = MIME_MAP.get(ext, 'image/jpeg')
    media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype=mime, resumable=False)
    file_meta = {'name': filename, 'parents': [folder_id]}
    uploaded = service.files().create(
        body=file_meta, media_body=media, fields='id'
    ).execute()
    file_id = uploaded.get('id')
    # 設為任何人可檢視
    service.permissions().create(
        fileId=file_id,
        body={'type': 'anyone', 'role': 'reader'}
    ).execute()
    return f"https://drive.google.com/uc?id={file_id}"

_drive_folder_cache = {}  # 快取避免重複查詢

def ensure_drive_folder(date_str):
    """確保 Drive 內有對應的資料夾，回傳 folder_id"""
    if date_str in _drive_folder_cache:
        return _drive_folder_cache[date_str], get_drive_service()
    service = get_drive_service()
    if not service:
        return None, None
    root_id  = get_or_create_drive_folder(service, DRIVE_FOLDER_NAME)
    month    = date_str[:7]  # 2026-03
    month_id = get_or_create_drive_folder(service, month, root_id)
    day_id   = get_or_create_drive_folder(service, date_str, month_id)
    _drive_folder_cache[date_str] = day_id
    return day_id, service

def get_spreadsheet():
    client = get_sheets_client()
    if not client:
        return None
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        ss = client.create(SPREADSHEET_NAME)
        ss.share(None, perm_type='anyone', role='reader')
        return ss
    except Exception as e:
        print(f"[Sheets] 開啟失敗：{e}")
        return None

def get_or_create_ws(spreadsheet, date_str):
    try:
        return spreadsheet.worksheet(date_str)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=date_str, rows=1000, cols=12)
        ws.append_row(HEADERS)
        ws.freeze(rows=1)
        return ws


# ── 路由 ─────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin')
def admin():
    return render_template('admin.html')


@app.route('/api/submit', methods=['POST'])
def submit_report():
    name       = request.form.get('name',       '').strip()
    date       = request.form.get('date',       '').strip()
    start_time = request.form.get('start_time', '').strip()
    end_time   = request.form.get('end_time',   '').strip()
    location   = request.form.get('location',   '').strip()
    reason     = request.form.get('reason',     '').strip()
    solution   = request.form.get('solution',   '').strip()

    if not all([name, date, start_time, end_time, location, reason, solution]):
        return jsonify({'success': False, 'message': '請填寫所有必填欄位'}), 400
    if start_time >= end_time:
        return jsonify({'success': False, 'message': '結束時間必須晚於開始時間'}), 400

    # 上傳照片到 Google Drive
    photo_urls = []
    files = request.files.getlist('photos')
    if files:
        folder_id, drive_svc = ensure_drive_folder(date)
        if folder_id and drive_svc:
            for photo in files:
                if photo and photo.filename and allowed_file(photo.filename):
                    try:
                        ext = photo.filename.rsplit('.', 1)[-1].lower()
                        fname = f"{uuid.uuid4().hex[:8]}.{ext}"
                        url = upload_photo_to_drive(drive_svc, photo.read(), fname, folder_id)
                        if url:
                            photo_urls.append(url)
                    except Exception as e:
                        print(f"[Drive] 照片上傳失敗：{e}")

    # 寫入 Google Sheets
    spreadsheet = get_spreadsheet()
    if not spreadsheet:
        return jsonify({'success': False, 'message': 'Google Sheets 未連線，請確認設定'}), 500
    try:
        ws = get_or_create_ws(spreadsheet, date)
        # 照片預覽：用 =IMAGE() 公式讓 Sheets 直接顯示縮圖
        photo_formula = f'=IMAGE("{photo_urls[0]}", 1)' if photo_urls else ''
        ws.append_row([
            uuid.uuid4().hex[:8],
            name, date, start_time, end_time,
            location, reason, solution,
            '\n'.join(photo_urls),   # 純文字 URL（後台系統用）
            taiwan_now(),
            photo_formula            # IMAGE 公式（Sheets 顯示縮圖用）
        ], value_input_option='USER_ENTERED')
    except Exception as e:
        return jsonify({'success': False, 'message': f'資料寫入失敗：{e}'}), 500

    return jsonify({'success': True, 'message': '回報已成功送出！'})


@app.route('/api/dates')
def get_dates():
    spreadsheet = get_spreadsheet()
    if not spreadsheet:
        return jsonify({'dates': []})
    try:
        titles = [ws.title for ws in spreadsheet.worksheets()]
        date_titles = sorted(
            [t for t in titles if len(t) == 10 and t[4] == '-' and t[7] == '-'],
            reverse=True
        )
        return jsonify({'dates': date_titles})
    except Exception as e:
        return jsonify({'dates': [], 'error': str(e)})


@app.route('/api/reports/<date_str>')
def get_reports_by_date(date_str):
    spreadsheet = get_spreadsheet()
    if not spreadsheet:
        return jsonify({'date': date_str, 'reports': [], 'count': 0})
    try:
        ws = spreadsheet.worksheet(date_str)
        rows = ws.get_all_records()
        reports = []
        for r in rows:
            photos_raw = r.get('照片URLs', '')
            photos = [p for p in photos_raw.split('\n') if p.strip()] if photos_raw else []
            reports.append({
                'id':         r.get('ID', ''),
                'name':       r.get('姓名', ''),
                'date':       r.get('日期', ''),
                'start_time': r.get('開始時間', ''),
                'end_time':   r.get('結束時間', ''),
                'location':   r.get('地點', ''),
                'reason':     r.get('原因', ''),
                'solution':   r.get('解決方法', ''),
                'photos':     photos,
                'created_at': r.get('提交時間', '')
            })
        return jsonify({'date': date_str, 'reports': reports, 'count': len(reports)})
    except gspread.WorksheetNotFound:
        return jsonify({'date': date_str, 'reports': [], 'count': 0})
    except Exception as e:
        return jsonify({'date': date_str, 'reports': [], 'count': 0, 'error': str(e)})


@app.route('/api/stats')
def get_stats():
    spreadsheet = get_spreadsheet()
    if not spreadsheet:
        return jsonify({'total': 0, 'days': 0, 'latest_date': None})
    try:
        date_sheets = [ws for ws in spreadsheet.worksheets()
                       if len(ws.title) == 10 and ws.title[4] == '-']
        days  = len(date_sheets)
        total = sum(max(ws.row_count - 1, 0) for ws in date_sheets)
        latest = sorted([ws.title for ws in date_sheets], reverse=True)
        return jsonify({'total': total, 'days': days,
                        'latest_date': latest[0] if latest else None})
    except:
        return jsonify({'total': 0, 'days': 0, 'latest_date': None})


@app.route('/api/status')
def status():
    creds = get_credentials()
    sheets_ok = creds is not None
    return jsonify({
        'sheets': {'connected': sheets_ok,
                   'message': '' if sheets_ok else '找不到憑證'},
        'drive':  {'connected': sheets_ok}
    })


# ── 啟動 ─────────────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 50)
    print("  線上工作回報系統 啟動中...")
    print("=" * 50)
    print(f"  前台：http://localhost:{port}")
    print(f"  後台：http://localhost:{port}/admin")
    print("=" * 50)
    app.run(host='0.0.0.0', port=port, debug=False)
