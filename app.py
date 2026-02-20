#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Booth Schedule Generator – Cloud Edition (Render)
Flask + gunicorn + openpyxl
"""
import os, sys, json, random, threading, tempfile, shutil, time, secrets, atexit, traceback
from copy import copy
from collections import defaultdict
from functools import wraps
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

random.seed(42)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB上限
# SECRET_KEYが未設定の場合は固定のフォールバックキーを使用
# （Render再起動でセッションが消えないようにするため）
app.secret_key = os.environ.get('SECRET_KEY', 'booth-schedule-generator-default-key-2026')

# パスワード（環境変数 or デフォルト）
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'booth2026')

# 一時ファイル管理（ディスクベース: gunicornマルチワーカー対応）
UPLOAD_BASE = os.path.join(tempfile.gettempdir(), 'booth_sessions')
os.makedirs(UPLOAD_BASE, exist_ok=True)
SESSION_TIMEOUT = 3600  # 1時間でクリーンアップ

def _session_dir(sid):
    """セッションIDからディレクトリパスを返す"""
    return os.path.join(UPLOAD_BASE, sid)

def _session_meta_path(sid):
    """セッションのメタデータJSONファイルパス"""
    return os.path.join(_session_dir(sid), '_meta.json')

def _load_meta(sid):
    """ディスクからセッションメタデータを読み込む"""
    mp = _session_meta_path(sid)
    if os.path.exists(mp):
        with open(mp, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def _save_meta(sid, meta):
    """セッションメタデータをディスクに保存"""
    mp = _session_meta_path(sid)
    with open(mp, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)

def cleanup_old_sessions():
    """古いセッションの一時ファイルを削除"""
    now = time.time()
    if not os.path.exists(UPLOAD_BASE):
        return
    for name in os.listdir(UPLOAD_BASE):
        sdir = os.path.join(UPLOAD_BASE, name)
        if not os.path.isdir(sdir):
            continue
        meta = _load_meta(name)
        if meta and now - meta.get('last_access', 0) > SESSION_TIMEOUT:
            shutil.rmtree(sdir, ignore_errors=True)

def get_session_data():
    """現在のセッションのデータを取得(なければ作成) - ディスクベース"""
    cleanup_old_sessions()
    sid = session.get('sid')
    sdir = _session_dir(sid) if sid else None
    if not sid or not os.path.exists(_session_meta_path(sid)):
        sid = secrets.token_hex(16)
        session['sid'] = sid
        sdir = _session_dir(sid)
        os.makedirs(sdir, exist_ok=True)
        meta = {'files': {}, 'dir': sdir, 'last_access': time.time()}
        _save_meta(sid, meta)
    meta = _load_meta(sid)
    meta['last_access'] = time.time()
    meta['dir'] = sdir  # 常にパスを保証
    _save_meta(sid, meta)
    # resultはインメモリで保持（大きいため）、ただしfilesパスはディスクから復元
    if not hasattr(get_session_data, '_cache'):
        get_session_data._cache = {}
    if sid not in get_session_data._cache:
        get_session_data._cache[sid] = {'result': {}}
    cached = get_session_data._cache[sid]
    return {**meta, 'result': cached.get('result', {}), '_sid': sid}

def save_session_files(sd):
    """ファイルパス情報をディスクに保存"""
    sid = sd['_sid']
    meta = _load_meta(sid)
    meta['files'] = sd['files']
    meta['last_access'] = time.time()
    _save_meta(sid, meta)

def save_session_result(sd):
    """resultをインメモリキャッシュ + ディスクに保存"""
    sid = sd['_sid']
    if not hasattr(get_session_data, '_cache'):
        get_session_data._cache = {}
    get_session_data._cache[sid] = {'result': sd['result']}
    # ディスクにもJSON保存（サーバー再起動後の復元用）
    _save_result_to_disk(sid, sd['result'])

def _result_json_path(sid):
    return os.path.join(_session_dir(sid), '_result.json')

def _save_result_to_disk(sid, result):
    """スケジュール結果をディスクにJSON保存"""
    rp = _result_json_path(sid)
    try:
        # schedule内のtupleをlistに変換して保存
        saveable = {}
        if 'schedule_json' in result:
            saveable['schedule_json'] = result['schedule_json']
        if 'unplaced' in result:
            saveable['unplaced'] = result['unplaced']
        if 'office_teachers' in result:
            saveable['office_teachers'] = result['office_teachers']
        if 'booth_pref' in result:
            saveable['booth_pref'] = result['booth_pref']
        if 'students' in result:
            # studentsのsetをlistに変換
            stu_save = []
            for s in result['students']:
                sc = dict(s)
                if isinstance(sc.get('avail'), set):
                    sc['avail'] = sorted([list(a) for a in sc['avail']])
                if isinstance(sc.get('backup_avail'), set):
                    sc['backup_avail'] = sorted([list(a) for a in sc['backup_avail']])
                if isinstance(sc.get('ng_dates'), set):
                    sc['ng_dates'] = [list(d) for d in sc['ng_dates']]
                if 'fixed' in sc:
                    sc['fixed'] = [list(f) for f in sc['fixed']]
                stu_save.append(sc)
            saveable['students'] = stu_save
        if 'week_dates' in result:
            saveable['week_dates'] = result['week_dates']
        with open(rp, 'w', encoding='utf-8') as f:
            json.dump(saveable, f, ensure_ascii=False)
    except Exception as e:
        print(f"[save_result] WARNING: ディスク保存失敗: {e}", flush=True)

def _load_result_from_disk(sid):
    """ディスクからスケジュール結果を読み込む"""
    rp = _result_json_path(sid)
    if not os.path.exists(rp):
        return None
    try:
        with open(rp, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None

# ========== 認証 ==========
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('authenticated'):
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'error': '認証が必要です'}), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    error = None
    if request.method == 'POST':
        pw = request.form.get('password', '')
        if pw == APP_PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('index'))
        error = 'パスワードが違います'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    sid = session.get('sid')
    if sid:
        sdir = _session_dir(sid)
        if os.path.exists(sdir):
            shutil.rmtree(sdir, ignore_errors=True)
        if hasattr(get_session_data, '_cache') and sid in get_session_data._cache:
            del get_session_data._cache[sid]
    session.clear()
    return redirect(url_for('login_page'))

ALLOWED_EXT = {'.xlsx'}
def validate_file(f):
    if not f or not f.filename:
        return False, 'ファイルが選択されていません'
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return False, f'許可されていないファイル形式です: {ext}'
    return True, None

RESULT = {}  # 後方互換（セッション内に移行済み）

# ========== 定数 ==========
DAYS = ['月','火','水','木','金','土']
WEEKDAY_TIMES = ['16:00','17:05','18:10','19:15','20:20']
SATURDAY_TIMES = ['14:55','16:00','17:05','18:10']
ALL_TIMES = ['14:55','16:00','17:05','18:10','19:15','20:20']
TIME_SHORT = {'14:55':'14','16:00':'16','17:05':'17','18:10':'18','19:15':'19','20:20':'20'}
MAX_BOOTHS = 6
# メタシート判定キーワード（週シートと区別するため）
# load_teacher_skills の検出キーワードと一致させること
META_KEYWORDS = ['必要コマ', '一覧', 'ブース希望', '指導可能', 'スキル']

NAME_MAP = {}
for full, short in [
    ('寒河江　道也','寒河江T'),('寒河江 道也','寒河江T'),
    ('若林　鈴華','若林T'),('若林 鈴華','若林T'),
    ('石川　隆斗','隆斗T'),('石川 隆斗','隆斗T'),
    ('石川　瑠璃','瑠璃T'),('石川 瑠璃','瑠璃T'),
    ('田村　倫子','田村T'),('田村 倫子','田村T'),
    ('平畑　美優奏','平畑T'),('平畑 美優奏','平畑T'),
    ('粉川　仁','粉川T'),('粉川 仁','粉川T'),
    ('小山　桜','小山T'),('小山 桜','小山T'),
    ('橋本　穂果','橋本T'),('橋本 穂果','橋本T'),
    ('後藤　凜','後藤T'),('後藤 凜','後藤T'),
    ('渡邉　樹希','渡邉T'),('渡邉 樹希','渡邉T'),
    ('越智　三佳','越智T'),('越智 三佳','越智T'),
    ('井上　玲也','井上T'),('井上 玲也','井上T'),
    ('西　泰我','西T'),('西 泰我','西T'),
    ('飯村　定子','飯村T'),('飯村 定子','飯村T'),
]:
    NAME_MAP[full] = short

# デフォルト講師ブース希望（UI から変更可能）
DEFAULT_BOOTH_PREF = {'若林T':1, '粉川T':3, '田村T':4}

SRC_TIME_SLOTS = [
    (6,'14:55',6),(19,'16:00',6),(32,'17:05',6),
    (45,'18:10',6),(58,'19:15',9),(77,'20:20',9),
]
SRC_DAY_COLS = {'月':3,'火':8,'水':13,'木':18,'金':23,'土':28}

SKILL_COL_MAP = {
    3:'小国',4:'小算',5:'小英',6:'小理',7:'小社',
    8:'受国',9:'受算',10:'受理',11:'受社',
    12:'中国',13:'中数',14:'中英',15:'中理',16:'中社',
    17:'高現',18:'高古',
}

LAYOUT = {
    '14:55':(7,6),'16:00':(20,6),'17:05':(33,6),
    '18:10':(46,6),'19:15':(59,6),'20:20':(72,6),
}
DAY_COLS = {
    '月':(3,4,5,6,7),'火':(8,9,10,11,12),'水':(13,14,15,16,17),
    '木':(18,19,20,21,22),'金':(23,24,25,26,27),'土':(28,29,30,31,32),
}
TUTOR_ROWS = [19,32,45,58,71,84]

def to_short(name):
    if not name: return None
    name = str(name).strip()
    if name in NAME_MAP: return NAME_MAP[name]
    parts = name.replace('\u3000',' ').split()
    return parts[0]+'T' if parts else name

# ========== パーサー ==========
def get_skill_keys(grade, subject):
    g = str(grade).upper()
    s = str(subject)
    # 英検は英語スキルでチェック
    if s == '英検':
        s = '英'
    if g.startswith('S'):
        yr = int(g[1:]) if len(g)>1 else 0
        if s == '数': s = '算'  # 小学/受験は「算」
        if yr >= 4:
            # 中学受験には英語がないので小学英語で代替
            if s == '英': return ['小英']
            return ['受'+s]
        return ['小'+s]
    elif g.startswith('C'):
        if s == '算': s = '数'  # 中学は「数」
        return ['中'+s]
    elif g.startswith('K'):
        if s == '算': s = '数'
        if s == '数': return ['高ⅠA','高ⅡB','高Ⅲ','高C']
        return ['高'+s]
    if s == '算': s = '数'
    return ['中'+s]

def can_teach(teacher, grade, subject, skills):
    keys = get_skill_keys(grade, subject)
    if not skills:
        return True  # スキルデータ自体がない場合は全員可
    if teacher not in skills:
        return False  # スキルシートに未登録の講師は配置不可
    return any(k in skills[teacher] for k in keys)

def load_teacher_skills(wb):
    """ブース表xlsx内の講師指導可能科目シートを読み込む"""
    # シート名を自動検出（「一覧表」「指導可能」等を含むシート）
    skill_sheet = None
    for sn in wb.sheetnames:
        if '一覧' in sn or '指導可能' in sn or 'スキル' in sn or 'skill' in sn.lower():
            skill_sheet = sn
            break
    if not skill_sheet:
        return {}  # シートが見つからない場合は空（全講師が全科目可として動作）

    ws = wb[skill_sheet]
    skills = {}
    for r in range(4, ws.max_row+1):
        t = ws.cell(r,2).value
        if not t: break
        t = str(t).strip()
        s = set()
        for c, k in SKILL_COL_MAP.items():
            if ws.cell(r,c).value == '◯': s.add(k)
        for c in range(19, ws.max_column+1):
            v, h = ws.cell(r,c).value, ws.cell(3,c).value
            if v == '◯' and h: s.add('高'+str(h))
        skills[t] = s
    return skills

def load_booth_pref(wb):
    """ブース表xlsx内の講師ブース希望シートを読み込む"""
    for sn in wb.sheetnames:
        if 'ブース希望' in sn:
            ws = wb[sn]
            pref = {}
            for r in range(2, ws.max_row+1):
                t = ws.cell(r, 1).value
                b = ws.cell(r, 2).value
                if t and b:
                    pref[str(t).strip()] = int(b)
            return pref
    return {}

def parse_ng_dates(val, year, month):
    """NG日程を解析して (week_index, day_name) のsetを返す。
    形式例: '2/5', '2/1-2/7', '2/19,2/24,2/25', '12/5'
    """
    if not val: return set()
    import datetime as _dt
    day_names = ['月','火','水','木','金','土','日']
    result = set()

    def add_date(m, d):
        if m != month: return
        try:
            dt = _dt.date(year, m, d)
        except ValueError:
            return
        wd = dt.weekday()
        if wd >= 6: return  # 日曜スキップ
        wi = (d - 1) // 7
        result.add((wi, day_names[wd]))

    def parse_md(s):
        s = s.strip().replace('/', '/')
        if '/' in s:
            parts = s.split('/')
            return int(parts[0]), int(parts[1])
        else:
            # 日のみ → 当月と仮定
            return month, int(s)

    for part in str(val).split(','):
        part = part.strip()
        if not part: continue
        if '-' in part:
            # 範囲: 2/1-2/7
            a, b = part.split('-', 1)
            try:
                m1, d1 = parse_md(a)
                m2, d2 = parse_md(b)
                if m1 == m2:
                    for d in range(d1, d2 + 1):
                        add_date(m1, d)
            except (ValueError, IndexError):
                pass
        else:
            try:
                m, d = parse_md(part)
                add_date(m, d)
            except (ValueError, IndexError):
                pass
    return result

def load_students_from_wb(wb, year=2026, month=2):
    ws = wb['必要コマ数']
    subj_cols = [(5,'英'),(6,'英検'),(7,'数'),(8,'算'),(9,'国'),(10,'理'),
                 (11,'社'),(12,'現'),(13,'古'),(14,'物'),(15,'化'),(16,'生'),
                 (17,'日'),(18,'地'),(19,'政'),(20,'世')]
    students = []
    for r in range(3, 60):
        grade, name = ws.cell(r,2).value, ws.cell(r,4).value
        if not name: break
        needs = {}
        for col, subj in subj_cols:
            v = ws.cell(r,col).value
            if v and isinstance(v,(int,float)) and v>0: needs[subj] = int(v)
        parse_list = lambda v: [t.strip() for t in str(v or '').split(',') if t.strip()]
        students.append({
            'grade':str(grade),'name':str(name),'needs':needs,
            'wish_teachers':parse_list(ws.cell(r,22).value),
            'ng_teachers':parse_list(ws.cell(r,23).value),
            'ng_students':parse_list(ws.cell(r,24).value),
            'avail':parse_avail(ws.cell(r,25).value),
            'backup_avail':parse_avail(ws.cell(r,26).value),
            'ng_dates':parse_ng_dates(ws.cell(r,27).value, year, month),
            'fixed':parse_regular(ws.cell(r,28).value),
            'notes':str(ws.cell(r,29).value or '').strip(),
        })
    return students

def parse_avail(val):
    if not val: return None
    slots = set()
    for p in str(val).split(','):
        p = p.strip()
        if not p: continue
        d, rest = p[0], p[1:]
        if '-' in rest:
            a,b = rest.split('-')
            for t in range(int(a),int(b)+1): slots.add((d,str(t)))
        else:
            slots.add((d,rest))
    return slots

def parse_regular(val):
    if not val: return []
    result = []
    for p in str(val).split(','):
        p = p.strip()
        if ':' not in p: continue
        dt,subj = p.split(':',1)
        result.append((dt[0], dt[1:], subj.strip()))
    return result

def load_weekly_teachers(path):
    """元シートから各週・曜日・時間帯の出勤講師を読み取る（全講師、絞り込み前）"""
    wb = openpyxl.load_workbook(path)
    weeks = []
    
    # シート名でフィルタリング（「ブース表」を含むシートのみ対象）
    target_sheets = []
    for sn in wb.sheetnames:
        # 非表示シートはスキップ
        if wb[sn].sheet_state != 'visible':
            continue
        # 「ブース表」が含まれるシートのみ対象
        if 'ブース表' in sn:
            target_sheets.append(sn)

    for sn in target_sheets:
        ws = wb[sn]
        week = {}
        for day in DAYS:
            col = SRC_DAY_COLS[day]
            dt = {}
            for start, tl, nb in SRC_TIME_SLOTS:
                ts = TIME_SHORT[tl]
                teachers = []
                for b in range(nb):
                    t = to_short(ws.cell(start+b*2, col).value)
                    if t:
                        teachers.append(t)
                dt[ts] = teachers
            week[day] = dt
        
        # 週全体で講師が一人もいない場合はスキップ（空のシートを除外）
        has_teachers = False
        for d_data in week.values():
            for t_list in d_data.values():
                if t_list:
                    has_teachers = True
                    break
            if has_teachers: break
        
        if has_teachers:
            weeks.append(week)
    return weeks

# ========== 元シート集約（講師回答ファイル → 週別出勤データ） ==========
SURVEY_TIME_ROWS = {10: '14:55', 11: '16:00', 12: '17:05', 13: '18:10', 14: '19:15', 15: '20:20'}
WEEKDAY_NORMALIZE = {
    '月曜日':'月','月曜':'月','月':'月','Mon':'月',
    '火曜日':'火','火曜':'火','火':'火','Tue':'火',
    '水曜日':'水','水曜':'水','水':'水','Wed':'水',
    '木曜日':'木','木曜':'木','木':'木','Thu':'木',
    '金曜日':'金','金曜':'金','金':'金','Fri':'金',
    '土曜日':'土','土曜':'土','土':'土','Sat':'土',
    '日曜日':'日','日曜':'日','日':'日','Sun':'日',
}

def _compute_month_week_map(year, month):
    """月の各日を週インデックス（1始まり）にマッピング。月〜土を1週とする。"""
    import datetime as _dt, calendar
    last_day = calendar.monthrange(year, month)[1]
    week_map = {}
    current_week = 0
    first_monday_seen = False
    for d in range(1, last_day + 1):
        dt = _dt.date(year, month, d)
        dow = dt.weekday()
        if dow == 6:  # 日曜はスキップ
            continue
        if dow == 0:  # 月曜で新しい週を開始
            current_week += 1
            first_monday_seen = True
        elif not first_monday_seen and current_week == 0:
            current_week = 1  # 月初が月曜以外の場合も第1週
        week_map[d] = current_week
    return week_map

def _get_merged_cell_value(ws, row, col):
    """結合セルの場合、左上セルの値を返す"""
    val = ws.cell(row, col).value
    if val is not None:
        return val
    from openpyxl.utils import get_column_letter
    coord = f'{get_column_letter(col)}{row}'
    for merge_range in ws.merged_cells.ranges:
        if coord in merge_range:
            return ws.cell(merge_range.min_row, merge_range.min_col).value
    return None

def _excel_serial_to_date(serial):
    """Excel シリアル値を date に変換"""
    import datetime as _dt
    try:
        return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=int(serial))).date()
    except Exception:
        return None

def parse_survey_file(file_path):
    """講師回答xlsxファイルを解析して講師名と出勤可能日時を返す"""
    import datetime as _dt
    import re as _re
    wb = openpyxl.load_workbook(file_path, data_only=True)

    # データシートを探す（「シート」を含むシート名、なければ先頭シート）
    data_sheet = None
    for sn in wb.sheetnames:
        if 'シート' in sn:
            data_sheet = sn
            break
    if not data_sheet:
        data_sheet = wb.sheetnames[0]

    ws = wb[data_sheet]

    # 講師名（row2, col2） — 結合セルにも対応
    raw_name = _get_merged_cell_value(ws, 2, 2)
    # 取得できなければ近傍セル (row1-3, col1-3) も探索
    if not raw_name:
        for r in range(1, 4):
            for c in range(1, 4):
                v = _get_merged_cell_value(ws, r, c)
                if v and isinstance(v, str) and len(v.strip()) >= 2:
                    # 数字だけ・記号だけのセルは除外
                    stripped = v.strip()
                    if not stripped.replace(' ', '').replace('\u3000', '').isdigit():
                        raw_name = v
                        break
            if raw_name:
                break
    teacher_name = to_short(raw_name) if raw_name else None
    if not teacher_name:
        # ファイル名から講師名を推定（例: 飯村　定子_202603シート.xlsx）
        basename = os.path.basename(file_path)
        if basename.startswith('survey_'):
            basename = basename[7:]  # 'survey_' プレフィックスを除去
        name_part = basename.split('_')[0].strip()
        if name_part:
            teacher_name = to_short(name_part)
            if teacher_name:
                raw_name = name_part
                print(f"[survey] ファイル名から講師名を推定: {teacher_name} (from {os.path.basename(file_path)})", flush=True)
    if not teacher_name:
        print(f"[survey] 講師名を検出できません: {file_path}", flush=True)
        return None

    # 日付から年月を取得し、週マップを構築
    year, month = None, None
    for c in range(3, ws.max_column + 1):
        v = _get_merged_cell_value(ws, 6, c)
        if isinstance(v, (_dt.datetime, _dt.date)):
            dt = v if isinstance(v, _dt.date) else v.date()
            year, month = dt.year, dt.month
            break
        elif isinstance(v, (int, float)) and v > 31:
            # Excel シリアル値の可能性
            dt = _excel_serial_to_date(v)
            if dt:
                year, month = dt.year, dt.month
                break

    # ヘッダー行から年月を推定（日付セルから取れなかった場合）
    if year is None or month is None:
        for r in range(1, 6):
            for c in range(1, ws.max_column + 1):
                v = _get_merged_cell_value(ws, r, c)
                if v and isinstance(v, str):
                    m = _re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月', v)
                    if m:
                        year, month = int(m.group(1)), int(m.group(2))
                        break
                    m2 = _re.search(r'(\d{1,2})\s*月', v)
                    if m2 and year is None:
                        month = int(m2.group(1))
                        year = _dt.date.today().year
                        break
                elif isinstance(v, (_dt.datetime, _dt.date)):
                    dt = v if isinstance(v, _dt.date) else v.date()
                    year, month = dt.year, dt.month
                    break
            if year and month:
                break

    week_map = _compute_month_week_map(year, month) if year and month else {}

    # 列ヘッダー（日付・曜日・祝日フラグ）を読み取る
    columns = []
    j = 3
    consecutive_empty = 0
    while consecutive_empty < 3:
        # row 6: 日付, row 7: 曜日, row 9: 祝休日フラグ
        date_val = _get_merged_cell_value(ws, 6, j)
        if date_val is None or str(date_val).strip() == '':
            consecutive_empty += 1
            j += 1
            continue
        consecutive_empty = 0

        weekday_raw = str(_get_merged_cell_value(ws, 7, j) or '').strip()
        weekday = WEEKDAY_NORMALIZE.get(weekday_raw, weekday_raw)

        # 曜日が取れなかった場合はdateオブジェクトから推測
        resolved_date = None
        if isinstance(date_val, (_dt.datetime, _dt.date)):
            resolved_date = date_val if isinstance(date_val, _dt.date) else date_val.date()
        elif isinstance(date_val, (int, float)) and date_val > 31:
            resolved_date = _excel_serial_to_date(date_val)

        if weekday not in DAYS and resolved_date:
            wd_names = ['月','火','水','木','金','土','日']
            weekday = wd_names[resolved_date.weekday()]

        # 日付から週番号を算出（row 8 は曜日出現回数なので使わない）
        day_of_month = None
        if resolved_date:
            day_of_month = resolved_date.day
        elif isinstance(date_val, (int, float)) and 1 <= date_val <= 31:
            day_of_month = int(date_val)
        week_num = week_map.get(day_of_month)

        holiday = _get_merged_cell_value(ws, 9, j)

        columns.append({
            'col': j,
            'weekday': weekday,
            'week_num': week_num,
            'holiday': (holiday == 1 or str(holiday).strip() == '1'),
        })
        j += 1

    # 出勤可能時間帯を読み取る
    availability = []  # list of {weekday, week_num, time_str}
    for row, time_str in SURVEY_TIME_ROWS.items():
        for col_info in columns:
            if col_info['holiday']:
                continue
            if col_info['weekday'] == '日':
                continue
            val = ws.cell(row, col_info['col']).value
            if val == 1 or str(val).strip() == '1':
                availability.append({
                    'weekday': col_info['weekday'],
                    'week_num': col_info['week_num'],
                    'time': time_str,
                })

    print(f"[survey] parsed: {teacher_name} (full: {raw_name}) — {len(availability)}コマ, year={year}, month={month}, cols={len(columns)}", flush=True)

    return {
        'name': teacher_name,
        'full_name': str(raw_name).strip(),
        'availability': availability,
    }

def aggregate_surveys_to_weekly(survey_results):
    """複数の講師回答データを集約して週別出勤講師データを生成する
    Returns: load_weekly_teachers と同じ形式 — weeks[wi][day][ts] = [teacher, ...]
    """
    # 週数を特定
    max_week = 0
    for sr in survey_results:
        for a in sr['availability']:
            wn = a.get('week_num')
            if wn and wn > max_week:
                max_week = wn
    if max_week == 0:
        max_week = 4  # fallback

    weeks = []
    for wi in range(max_week):
        week = {}
        for day in DAYS:
            dt = {}
            for time_str in ALL_TIMES:
                ts = TIME_SHORT[time_str]
                teachers = []
                for sr in survey_results:
                    for a in sr['availability']:
                        if (a.get('week_num') == wi + 1 and
                            a['weekday'] == day and
                            a['time'] == time_str):
                            fn = sr.get('full_name') or sr['name']
                            if fn not in teachers:
                                teachers.append(fn)
                            break
                dt[ts] = teachers
            week[day] = dt
        weeks.append(week)

    return weeks

def generate_src_excel(weekly_teachers, output_path):
    """週別出勤講師データからブース表元シートExcel（1ブック・週別シート）を生成"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for wi, week_data in enumerate(weekly_teachers):
        ws = wb.create_sheet(f'第{wi+1}週')

        for start_row, time_str, num_booths in SRC_TIME_SLOTS:
            ts = TIME_SHORT[time_str]
            for day, col in SRC_DAY_COLS.items():
                teachers = week_data.get(day, {}).get(ts, [])
                for bi in range(min(len(teachers), num_booths)):
                    ws.cell(start_row + bi * 2, col, teachers[bi])

    wb.save(output_path)

def select_teachers_for_day(day, day_data, booth_pref, wish_teachers_set, office_teacher=None):
    """
    1日分の全時間帯データから、ブース⑥まで（最大6名）に講師を絞り込む。
    - 教室業務担当(office_teacher)はブースから除外
    - 早い時間帯から出勤できる講師を優先
    - 希望講師(wish_teachers_set)はブース枠外でも配置可能（例外）
    - ブース希望がある講師はその番号に配置
    """
    times = SATURDAY_TIMES if day == '土' else WEEKDAY_TIMES
    ts_list = [TIME_SHORT[tl] for tl in times]

    # 各講師の最早出勤時間帯を計算（教室業務担当を除外）
    teacher_earliest = {}
    for ts in ts_list:
        for t in day_data.get(ts, []):
            if t == office_teacher:
                continue  # 教室業務担当はブースに入れない
            if t not in teacher_earliest:
                teacher_earliest[t] = ts

    if len(teacher_earliest) <= MAX_BOOTHS:
        selected = set(teacher_earliest.keys())
    else:
        ts_order = {'14':0, '16':1, '17':2, '18':3, '19':4, '20':5}
        ranked = sorted(teacher_earliest.items(), key=lambda x: ts_order.get(x[1], 99))
        selected = set(t for t, _ in ranked[:MAX_BOOTHS])

    # 希望講師は必ず含める（例外）
    for t in teacher_earliest:
        if t in wish_teachers_set:
            selected.add(t)

    # ブース希望を正確に反映した配置を生成
    # 1) 希望ブースがある講師を先にその番号に配置
    # 2) 残りの講師を空きブースに詰める
    def assign_booth_order(available_teachers):
        """available_teachersリストをブース希望に基づいて最大6スロットに配置"""
        slots = [None] * MAX_BOOTHS  # ブース①〜⑥

        # 希望ブースがある講師を先に配置
        remaining = list(available_teachers)
        for t in list(remaining):
            if t in booth_pref:
                bi = booth_pref[t] - 1  # 0-indexed
                if 0 <= bi < MAX_BOOTHS and slots[bi] is None:
                    slots[bi] = t
                    remaining.remove(t)

        # 残りを空きスロットに順番に詰める
        for t in remaining:
            for i in range(MAX_BOOTHS):
                if slots[i] is None:
                    slots[i] = t
                    break

        return [t for t in slots if t is not None]

    # 各講師の出勤範囲（最初〜最後の出勤コマ）を計算
    # 元シートの読み込み制限(nb=6)で中間コマが欠落する場合を補間する
    ts_order = {'14':0, '16':1, '17':2, '18':3, '19':4, '20':5}
    teacher_range = {}  # {teacher: (first_ord, last_ord)}
    for ts in ts_list:
        for t in day_data.get(ts, []):
            if t not in selected or t == office_teacher:
                continue
            o = ts_order.get(ts, 99)
            if t not in teacher_range:
                teacher_range[t] = (o, o)
            else:
                teacher_range[t] = (teacher_range[t][0], max(teacher_range[t][1], o))

    # 1日分のブース配置を1回だけ決定し、全時間帯で同じブース番号を維持する
    # （途中で別講師がそのブースに入らないようにする）
    all_day_teachers = [t for t in selected if t != office_teacher]
    day_booth_order = assign_booth_order(all_day_teachers)

    result = {}
    for ts in ts_list:
        cur_ord = ts_order.get(ts, 99)
        # day_dataに直接含まれる講師 OR 出勤範囲内（first〜last）の講師
        available = set()
        for t in day_booth_order:
            if t in teacher_range:
                first, last = teacher_range[t]
                if first <= cur_ord <= last:
                    available.add(t)
        # 固定ブース位置に基づいてリスト生成（出勤していないコマは空文字）
        booths = []
        for i, t in enumerate(day_booth_order):
            if t in available:
                booths.append(t)
            else:
                booths.append('')  # そのコマは不在だがブース位置を確保
        result[ts] = booths
    return result

def resolve_office_teacher(day, candidates, day_data):
    """教室業務担当を優先順位リストから決定する。
    - 石川Tは出勤チェック不要で即確定
    - それ以外は day_data（その週・その曜日の出勤講師データ）で出勤確認
    - 誰も出勤していなければ None（教室業務なし）
    """
    if isinstance(candidates, str):
        candidates = [candidates]
    for candidate in candidates:
        if candidate == '石川T':
            return candidate
        # day_data: {ts: [teacher, ...]} — いずれかの時間帯に出勤していれば可
        for ts, teachers in day_data.items():
            if candidate in teachers:
                return candidate
    return None

def load_holidays(booth_wb, num_weeks):
    """ブース表の教室業務行(row 5)から休塾日を検出する。
    Returns: [{day: True, ...}, ...] 各週の休塾日マップ
    """
    # 隠しシートとメタデータシートを除外
    week_sheets = []
    for sn in booth_wb.sheetnames:
        if any(k in sn for k in META_KEYWORDS): continue
        if booth_wb[sn].sheet_state != 'visible': continue
        week_sheets.append(sn)

    holidays = []
    for wi in range(min(num_weeks, len(week_sheets))):
        ws = booth_wb[week_sheets[wi]]
        h = {}
        for day, cols in DAY_COLS.items():
            val = ws.cell(5, cols[0]).value
            if val and '休塾' in str(val):
                h[day] = True
        holidays.append(h)
    # 足りない週は空辞書で埋める
    while len(holidays) < num_weeks:
        holidays.append({})
    return holidays

# ========== スケジューラー ==========
def build_schedule(students, weekly_teachers, skills, office_rule, booth_pref, holidays=None):
    remaining = {s['name']: dict(s['needs']) for s in students}
    smap = {s['name']: s for s in students}
    schedule = []
    office_teachers = []
    num_weeks = len(weekly_teachers)

    # 全生徒の希望講師を集約
    wish_teachers_set = set()
    for s in students:
        wish_teachers_set.update(s['wish_teachers'])

    for wi in range(num_weeks):
        ot = {}
        for d in DAYS:
            # 休塾日チェック
            if holidays and wi < len(holidays) and holidays[wi].get(d):
                ot[d] = '休塾日'
            else:
                candidates = office_rule.get(d, ['石川T'])
                d_data = weekly_teachers[wi].get(d, {})
                ot[d] = resolve_office_teacher(d, candidates, d_data)
        office_teachers.append(ot)
        ws = {}
        for day in DAYS:
            # 講師選抜（ブース⑥まで、早い時間帯優先、教室業務担当除外）
            day_data = weekly_teachers[wi].get(day, {})
            ot_teacher = ot.get(day)
            filtered = select_teachers_for_day(day, day_data, booth_pref, wish_teachers_set, ot_teacher)
            ds = {}
            times = SATURDAY_TIMES if day=='土' else WEEKDAY_TIMES
            for tl in times:
                ts = TIME_SHORT[tl]
                tlist = filtered.get(ts, [])
                booths = [{'teacher':t, 'slots':[]} for t in tlist]
                # 常にMAX_BOOTHS(6)ブース分のデータを確保
                while len(booths) < MAX_BOOTHS:
                    booths.append({'teacher':'', 'slots':[]})
                ds[ts] = booths
            ws[day] = ds
        schedule.append(ws)

    def get_placed_days(ws, name, subj):
        days = set()
        for day in DAYS:
            for ts, booths in ws.get(day,{}).items():
                for b in booths:
                    for g,sn,sb in b['slots']:
                        if sn==name and sb==subj: days.add(day)
        return days

    def get_student_slots(ws, name):
        r = set()
        for day in DAYS:
            for ts, booths in ws.get(day,{}).items():
                for b in booths:
                    for g,sn,sb in b['slots']:
                        if sn==name: r.add((day,ts))
        return r

    def get_any_placed_days(ws, name):
        """生徒が任意の科目で既に配置されている曜日の集合を返す"""
        days = set()
        for day in DAYS:
            for ts, booths in ws.get(day,{}).items():
                for b in booths:
                    for g,sn,sb in b['slots']:
                        if sn==name: days.add(day)
        return days

    def get_teacher_booth(ws, day, teacher):
        for ts, booths in ws.get(day,{}).items():
            for bi,b in enumerate(booths):
                if b['teacher']==teacher and b['slots']:
                    return bi
        return None

    def check_booth(booth, bi, s, day, subj, ws):
        t = booth['teacher']
        if not t or len(booth['slots'])>=2: return False
        if t in s['ng_teachers']: return False
        if not can_teach(t, s['grade'], subj, skills): return False
        # 同一ブース内のNG生徒チェック
        for g2,sn2,sb2 in booth['slots']:
            if sn2 in s['ng_students']: return False
            other = smap.get(sn2)
            if other and s['name'] in other.get('ng_students',[]): return False
        # 隣接ブースチェックは廃止（同一ブースのみNGとする要望により）
        
        eb = get_teacher_booth(ws, day, t)
        if eb is not None and eb != bi: return False
        return True

    def place_student(ws, s, day, ts, subj):
        if day not in ws or ts not in ws[day]: return False
        booths = list(enumerate(ws[day][ts]))
        wish = s.get('wish_teachers', [])
        if wish:
            # 希望講師のブースを先に試す
            wish_first = sorted(booths, key=lambda x: (0 if x[1]['teacher'] in wish else 1))
            booths = wish_first
        for bi,b in booths:
            if check_booth(b, bi, s, day, subj, ws):
                b['slots'].append((s['grade'],s['name'],subj))
                return True
        return False

    def find_slot(ws, s, subj, placed_days, existing, wi, any_placed_days):
        cands = []
        checked_avail = False
        reject_full = 0
        reject_ng = 0
        reject_skill = 0
        reject_other = 0
        for day in DAYS:
            if day in placed_days: continue  # 同一科目の同曜日配置を防止
            # NG日程チェック: 配置自体は許可するがペナルティ
            is_ng_date = (wi, day) in s.get('ng_dates', set())
            
            times = SATURDAY_TIMES if day=='土' else WEEKDAY_TIMES
            for tl in times:
                ts = TIME_SHORT[tl]
                is_primary = s['avail'] is None or (day,ts) in s['avail']
                is_backup = (not is_primary) and s.get('backup_avail') and (day,ts) in s['backup_avail']
                if not is_primary and not is_backup: continue
                checked_avail = True
                if (day,ts) in existing: continue
                if ts not in ws.get(day,{}): continue
                for bi,b in enumerate(ws[day][ts]):
                    t = b['teacher']
                    if not t: continue
                    if len(b['slots'])>=2:
                        reject_full += 1
                        continue
                    if t in s['ng_teachers']:
                        reject_ng += 1
                        continue
                    if not can_teach(t, s['grade'], subj, skills):
                        reject_skill += 1
                        continue
                    if not check_booth(b, bi, s, day, subj, ws):
                        reject_other += 1
                        continue
                    sc = 0
                    # NG日程は大きくペナルティ（配置は可能）
                    if is_ng_date: sc -= 5000
                    # 予備時間はペナルティ（希望時間を優先）
                    if is_backup: sc -= 150
                    # 同曜日に既に別科目が配置されている場合
                    # 連続コマを強く推奨（+2000）、飛び石は回避（-200）
                    existing_on_day = [t_ for d_, t_ in existing if d_ == day]
                    if existing_on_day:
                        # 現在の時刻のインデックスを取得
                        try:
                            # timesは '16:00' 等の形式リスト
                            # tl は現在ループ中の時刻文字列 ('16:00')
                            curr_idx = times.index(tl)
                            
                            is_continuous = False
                            for et_short in existing_on_day:
                                # existingは '16' 形式の場合と '16:00' 形式の場合があるため正規化が必要
                                # TIME_SHORTの逆マッピングまたはループで探す
                                # ここでは existing が (day, ts_short) で入っている前提
                                # ts_short ('16') -> tl_long ('16:00')
                                et_long = TSR.get(et_short) if 'TSR' in globals() else None
                                if not et_long:
                                    # TSRがない場合は自力で探す (TIME_SHORTの逆)
                                    for k, v in TIME_SHORT.items():
                                        if v == et_short:
                                            et_long = k
                                            break
                                if et_long in times:
                                    ex_idx = times.index(et_long)
                                    diff = abs(curr_idx - ex_idx)
                                    if diff == 1:
                                        sc += 2000  # 連続コマは最優先
                                        is_continuous = True
                                    elif diff > 1:
                                        sc -= 200   # 飛び石はペナルティ
                        except ValueError:
                            pass
                    
                    if day in any_placed_days:
                        day_count = len(existing_on_day)
                        if day_count < 2:
                            sc += 50   # 2コマ目までは曜日優先（連続ボーナスと累積）
                        else:
                            sc -= 80   # 3コマ目以降は分散推奨
                    if b['teacher'] in s['wish_teachers']: sc += 500
                    if t in booth_pref and booth_pref[t]==bi+1: sc += 10
                    if len(b['slots'])==0: sc += 20
                    cands.append((sc, day, ts, bi))
        if not cands:
            if not checked_avail:
                reason = '希望時間帯なし'
            elif reject_skill:
                reason = '指導可能な講師不在'
            elif reject_ng:
                reason = 'NG講師'
            elif reject_other:
                reason = 'NG生徒/ブース制約'
            elif reject_full:
                reason = 'ブース満席'
            else:
                reason = '空きコマなし'
            return None, reason
        cands.sort(key=lambda x:-x[0])
        best_sc = cands[0][0]
        bests = [c for c in cands if c[0]==best_sc]
        ch = random.choice(bests)
        return (ch[1], ch[2], ch[3]), None

    def distribute(total, weeks):
        t = [total//weeks]*weeks
        for i in range(total%weeks): t[i] += 1
        random.shuffle(t)
        return t

    # Phase1: 固定授業
    for s in students:
        for day, ts_str, subj in s['fixed']:
            for wi in range(num_weeks):
                if (wi, day) in s.get('ng_dates', set()): continue
                if remaining[s['name']].get(subj, 0) <= 0: continue  # 必要コマ数を超えたら配置しない
                if place_student(schedule[wi], s, day, ts_str, subj):
                    remaining[s['name']][subj] -= 1

    # Phase2: 通常配置
    order = sorted(students, key=lambda s: (
        0 if s['wish_teachers'] else 1,
        len(s['avail']) if s['avail'] else 999, sum(s['needs'].values())
    ))
    unplaced_reasons = {}  # (name, subj) -> reason
    for s in order:
        for subj, total in s['needs'].items():
            still = remaining[s['name']].get(subj, 0)
            if still <= 0: continue
            targets = distribute(still, num_weeks)
            for wi in range(num_weeks):
                for _ in range(targets[wi]):
                    if remaining[s['name']].get(subj,0) <= 0: break
                    pd = get_placed_days(schedule[wi], s['name'], subj)
                    ex = get_student_slots(schedule[wi], s['name'])
                    apd = get_any_placed_days(schedule[wi], s['name'])
                    best, reason = find_slot(schedule[wi], s, subj, pd, ex, wi, apd)
                    if best:
                        day, ts, bi = best
                        schedule[wi][day][ts][bi]['slots'].append((s['grade'],s['name'],subj))
                        remaining[s['name']][subj] -= 1
                    elif reason:
                        unplaced_reasons[(s['name'], subj)] = reason

    # Phase3: 未配置リトライ（distribute で割り当てられなかった週にも配置を試行）
    for s in order:
        for subj in s['needs']:
            still = remaining[s['name']].get(subj, 0)
            if still <= 0: continue
            for wi in range(num_weeks):
                if remaining[s['name']].get(subj, 0) <= 0: break
                pd = get_placed_days(schedule[wi], s['name'], subj)
                ex = get_student_slots(schedule[wi], s['name'])
                apd = get_any_placed_days(schedule[wi], s['name'])
                best, reason = find_slot(schedule[wi], s, subj, pd, ex, wi, apd)
                if best:
                    day, ts, bi = best
                    schedule[wi][day][ts][bi]['slots'].append((s['grade'],s['name'],subj))
                    remaining[s['name']][subj] -= 1
                elif reason:
                    unplaced_reasons[(s['name'], subj)] = reason

    unplaced = []
    for s in students:
        for subj, cnt in remaining[s['name']].items():
            if cnt > 0:
                reason = unplaced_reasons.get((s['name'], subj), '')
                unplaced.append({'grade':s['grade'],'name':s['name'],'subject':subj,'count':cnt,'reason':reason})

    return schedule, unplaced, office_teachers

def extract_week_dates(booth_wb, num_weeks):
    """ブース表シート名から各週・各曜日の日付を算出する。
    _compute_month_week_map を使用して正確な週境界で日付をマッピングする。
    Returns: {'year':int, 'month':int, 'weeks':[ {day_name: day_number, ...}, ... ]}
    """
    import datetime as _dt, re
    week_sheets = [sn for sn in booth_wb.sheetnames if not any(k in sn for k in META_KEYWORDS)]

    year, month = None, None
    for sn in week_sheets:
        m = re.search(r'(\d{4})[./](\d{1,2})[./](\d{1,2})', sn)
        if m:
            year, month = int(m.group(1)), int(m.group(2))
            break
    if not year:
        return None

    day_names = ['月','火','水','木','金','土']
    week_map = _compute_month_week_map(year, month)

    # 週番号ごとに日付をグループ化
    by_week = {}
    for day_num, week_num in week_map.items():
        dt = _dt.date(year, month, day_num)
        wd = dt.weekday()  # 0=Mon ... 5=Sat
        if wd < 6:
            if week_num not in by_week:
                by_week[week_num] = {}
            by_week[week_num][day_names[wd]] = day_num

    # 0-indexed リストに変換 (wi=0 → week 1)
    weeks = []
    for wi in range(num_weeks):
        weeks.append(by_week.get(wi + 1, {}))
    return {'year': year, 'month': month, 'weeks': weeks}

# ========== Excel出力 ==========
def write_excel(schedule, unplaced, office_teachers, booth_path, output_path, state_json=None):
    wb = openpyxl.load_workbook(booth_path)
    num_weeks = len(schedule)
    # 週シート以外（必要コマ数、一覧表、ブース希望等）を特定して保持
    meta_sheets = [sn for sn in wb.sheetnames if any(k in sn for k in META_KEYWORDS)]
    week_sheets = [sn for sn in wb.sheetnames if sn not in meta_sheets]
    # 週シート数が足りない場合はそのまま使える分だけ使う
    num_weeks = min(num_weeks, len(week_sheets))

    # 共通書式
    teacher_font = Font(name='MS PGothic', size=8)
    teacher_align = Alignment(textRotation=255, vertical='center', horizontal='center')
    data_font = Font(name='MS PGothic', size=11)
    data_align = Alignment(vertical='center', horizontal='center')

    for wi in range(num_weeks):
        ws = wb[week_sheets[wi]]
        wsched = schedule[wi]

        # クリア（結合は解除しない）
        for tl, (sr, nb) in LAYOUT.items():
            for b in range(nb):
                r1, r2 = sr+b*2, sr+b*2+1
                for day in DAYS:
                    _, lc, gc, sc, sjc = DAY_COLS[day]
                    try: ws.cell(r1, lc).value = None
                    except: pass
                    for c in [gc, sc, sjc]:
                        for r in [r1, r2]:
                            try: ws.cell(r, c).value = None
                            except: pass

        # 書き込み
        for tl, (sr, nb) in LAYOUT.items():
            ts = TIME_SHORT[tl]
            for day in DAYS:
                _, lc, gc, sc, sjc = DAY_COLS[day]
                booths = wsched.get(day,{}).get(ts,[])
                for bi in range(min(nb, len(booths))):
                    r1, r2 = sr+bi*2, sr+bi*2+1
                    b = booths[bi]
                    # 講師名: 結合セルに縦書き + MS PGothic + 中央
                    if b['teacher']:
                        cell = ws.cell(r1, lc)
                        cell.value = b['teacher']
                        cell.font = teacher_font
                        cell.alignment = teacher_align
                    # 生徒1 → 上段
                    if len(b['slots'])>=1:
                        g,sn,subj = b['slots'][0]
                        for c, v in [(gc,g),(sc,sn),(sjc,subj)]:
                            cell = ws.cell(r1,c)
                            cell.value = v
                            cell.font = data_font
                            cell.alignment = data_align
                    # 生徒2 → 下段
                    if len(b['slots'])>=2:
                        g2,sn2,subj2 = b['slots'][1]
                        for c, v in [(gc,g2),(sc,sn2),(sjc,subj2)]:
                            cell = ws.cell(r2,c)
                            cell.value = v
                            cell.font = data_font
                            cell.alignment = data_align

        # 教室業務・チューター
        ot = office_teachers[wi] if wi < len(office_teachers) else {}
        holiday_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        holiday_font = Font(name='MS PGothic', color='FFFFFF', bold=True, size=11)
        for day in DAYS:
            bc = DAY_COLS[day][0]
            t = ot.get(day)
            if t:
                cell = ws.cell(5, bc, t)
                if t == '休塾日':
                    cell.fill = holiday_fill
                    cell.font = holiday_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                for tr in TUTOR_ROWS:
                    try:
                        c = ws.cell(tr, bc, t)
                        if t == '休塾日':
                            c.fill = holiday_fill
                            c.font = holiday_font
                            c.alignment = Alignment(horizontal='center', vertical='center')
                    except: pass

    # 未配置コマシート
    ws_up = wb.create_sheet('未配置コマ')
    for c, h in enumerate(['学年','生徒名','科目','未配置数'], 1):
        ws_up.cell(1, c, h).font = Font(name='MS PGothic', bold=True)
    for i, u in enumerate(unplaced, 2):
        ws_up.cell(i,1,u['grade']); ws_up.cell(i,2,u['name'])
        ws_up.cell(i,3,u['subject']); ws_up.cell(i,4,u['count'])
    ws_up.column_dimensions['A'].width = 6
    ws_up.column_dimensions['B'].width = 12
    ws_up.column_dimensions['C'].width = 6
    ws_up.column_dimensions['D'].width = 10

    # スケジュール状態を隠しシートに保存（再読み込み用）
    if state_json:
        ws_state = wb.create_sheet('_schedule_data')
        ws_state.sheet_state = 'hidden'
        data_str = json.dumps(state_json, ensure_ascii=False)
        # Excelセルの文字数上限(32767)を考慮して分割
        CHUNK = 30000
        for i in range(0, len(data_str), CHUNK):
            ws_state.cell(i // CHUNK + 1, 1, data_str[i:i+CHUNK])

    wb.save(output_path)

# ========== Error Handlers ==========
@app.errorhandler(413)
def request_entity_too_large(error):
    print(f"[error] 413 Request Entity Too Large", flush=True)
    return jsonify({'error': 'ファイルサイズが上限(10MB)を超えています。ファイルを確認してください。'}), 413

# ========== API ==========
@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/api/teachers')
@login_required
def get_teachers():
    """アップロード済みブース表から講師名一覧とブース希望を返す"""
    sd = get_session_data()
    files = sd.get('files', {})
    if 'booth' not in files:
        return jsonify({'error': 'ブース表がアップロードされていません'}), 400
    try:
        wb = openpyxl.load_workbook(files['booth'], data_only=True)
        skills = load_teacher_skills(wb)
        booth_pref = load_booth_pref(wb)
        if not booth_pref:
            booth_pref = dict(DEFAULT_BOOTH_PREF)
        wb.close()
        return jsonify({
            'teachers': sorted(skills.keys()),
            'boothPref': booth_pref,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
@login_required
def upload():
    sd = get_session_data()
    print(f"[upload] sid={sd.get('_sid','?')}, existing_files={list(sd.get('files',{}).keys())}", flush=True)
    saved = {}
    for key in ['src','booth']:
        f = request.files.get(key)
        if f:
            ok, err = validate_file(f)
            if not ok:
                print(f"[upload] ERROR validation failed: key={key}, filename={f.filename}, error={err}", flush=True)
                return jsonify({'error': err}), 400
            path = os.path.join(sd['dir'], key + '_' + f.filename)
            try:
                f.save(path)
            except Exception as e:
                tb = traceback.format_exc()
                print(f"[upload] ERROR saving file: key={key}, filename={f.filename}, path={path}, error={e}\n{tb}", flush=True)
                return jsonify({'error': f'ファイル保存に失敗しました ({key}: {f.filename}): {e}'}), 500
            if not os.path.exists(path):
                print(f"[upload] ERROR file not found after save: key={key}, path={path}", flush=True)
                return jsonify({'error': f'ファイル保存後にファイルが見つかりません ({key}: {f.filename})'}), 500
            size = os.path.getsize(path)
            saved[key] = path
            print(f"[upload] saved {key} -> {path} (size={size}bytes)", flush=True)
    if not saved:
        print(f"[upload] WARNING no files received in request", flush=True)
        return jsonify({'error': 'アップロードするファイルが含まれていません'}), 400
    sd['files'] = {**sd.get('files',{}), **saved}
    save_session_files(sd)
    return jsonify({'ok': True, 'files': {k: os.path.basename(v) for k,v in sd.get('files',{}).items()}})

@app.route('/api/upload_surveys', methods=['POST'])
@login_required
def upload_surveys():
    """講師回答ファイル（複数）をアップロード → 集約 → 元シートを自動生成"""
    sd = get_session_data()
    files = request.files.getlist('surveys')
    if not files or all(not f.filename for f in files):
        return jsonify({'error': '講師回答ファイルが含まれていません'}), 400

    survey_results = []
    errors = []

    for f in files:
        ok, err = validate_file(f)
        if not ok:
            errors.append(f'{f.filename}: {err}')
            continue

        path = os.path.join(sd['dir'], 'survey_' + f.filename)
        try:
            f.save(path)
        except Exception as e:
            errors.append(f'{f.filename}: 保存失敗 - {e}')
            continue

        try:
            result = parse_survey_file(path)
            if result:
                survey_results.append(result)
                print(f"[survey] {f.filename} -> {result['name']} ({len(result['availability'])}コマ)", flush=True)
            else:
                errors.append(f'{f.filename}: 講師情報を読み取れません')
        except Exception as e:
            errors.append(f'{f.filename}: {str(e)}')
            traceback.print_exc()

    if not survey_results:
        return jsonify({'error': '有効な講師回答ファイルがありません', 'details': errors}), 400

    # 集約して元シートExcelを生成
    weekly_teachers = aggregate_surveys_to_weekly(survey_results)
    src_path = os.path.join(sd['dir'], 'generated_src.xlsx')
    generate_src_excel(weekly_teachers, src_path)

    # srcファイルとして登録
    sd['files'] = {**sd.get('files', {}), 'src': src_path}
    save_session_files(sd)

    teacher_names = sorted(set(sr['name'] for sr in survey_results))
    return jsonify({
        'ok': True,
        'teachers': teacher_names,
        'teacherCount': len(teacher_names),
        'weeks': len(weekly_teachers),
        'errors': errors,
        'files': {k: os.path.basename(v) for k, v in sd.get('files', {}).items()},
    })

@app.route('/api/consolidate_booth', methods=['POST'])
@login_required
def consolidate_booth():
    """週別ブース表ファイルとメタデータファイル(必要コマ数等)を1つのブックに統合"""
    sd = get_session_data()
    meta_file = request.files.get('meta')
    week_files = request.files.getlist('weeks')

    if not meta_file or not meta_file.filename:
        return jsonify({'error': 'メタデータファイル（必要コマ数等）を選択してください'}), 400
    if not week_files or all(not f.filename for f in week_files):
        return jsonify({'error': '週別ブース表ファイルを選択してください'}), 400

    # メタデータファイルを保存・読み込み
    ok, err = validate_file(meta_file)
    if not ok:
        return jsonify({'error': f'メタデータファイル: {err}'}), 400
    meta_path = os.path.join(sd['dir'], 'meta_' + meta_file.filename)
    meta_file.save(meta_path)

    try:
        meta_wb = openpyxl.load_workbook(meta_path)
    except Exception as e:
        return jsonify({'error': f'メタデータファイルの読み込みに失敗: {e}'}), 400

    # メタシートを特定（必要コマ数、一覧表、ブース希望、指導可能科目等）
    meta_sheet_names = [sn for sn in meta_wb.sheetnames if any(k in sn for k in META_KEYWORDS)]

    # 古い週シートを削除（メタシート以外）
    old_week_sheets = [sn for sn in meta_wb.sheetnames if sn not in meta_sheet_names]
    for sn in old_week_sheets:
        del meta_wb[sn]
    print(f"[consolidate] メタシート: {meta_sheet_names}, 削除した古い週シート: {old_week_sheets}", flush=True)

    # 週別ファイルを処理
    errors = []
    week_count = 0
    for f in sorted(week_files, key=lambda x: x.filename):
        ok, err = validate_file(f)
        if not ok:
            errors.append(f'{f.filename}: {err}')
            continue

        week_path = os.path.join(sd['dir'], 'week_' + f.filename)
        try:
            f.save(week_path)
            week_wb = openpyxl.load_workbook(week_path)

            for sn in week_wb.sheetnames:
                # 「ブース表」が含まれるシートのみ対象
                if 'ブース表' not in sn:
                    continue
                # 非表示シートはスキップ（念のため）
                if week_wb[sn].sheet_state != 'visible':
                    continue

                src_ws = week_wb[sn]
                
                # シートが空かチェック（簡易チェック: 講師名セルにデータがあるか）
                has_data = False
                for day in DAYS:
                    col = SRC_DAY_COLS[day]
                    for start, tl, nb in SRC_TIME_SLOTS:
                        for b in range(nb):
                            if src_ws.cell(start+b*2, col).value:
                                has_data = True
                                break
                        if has_data: break
                    if has_data: break
                
                if not has_data:
                    print(f"[consolidate] 空の週シートをスキップ: {sn}", flush=True)
                    continue

                week_count += 1

                # シート名の重複を回避
                new_name = sn
                if new_name in meta_wb.sheetnames:
                    new_name = f'第{week_count}週'
                while new_name in meta_wb.sheetnames:
                    new_name = f'週{week_count}_{week_count}'

                dst_ws = meta_wb.create_sheet(new_name)

                # セルのコピー（値 + スタイル）
                for row in src_ws.iter_rows():
                    for cell in row:
                        dst_cell = dst_ws.cell(row=cell.row, column=cell.column)
                        dst_cell.value = cell.value
                        if cell.has_style:
                            dst_cell.font = copy(cell.font)
                            dst_cell.border = copy(cell.border)
                            dst_cell.fill = copy(cell.fill)
                            dst_cell.number_format = cell.number_format
                            dst_cell.protection = copy(cell.protection)
                            dst_cell.alignment = copy(cell.alignment)

                # 結合セルのコピー
                for merged_range in src_ws.merged_cells.ranges:
                    dst_ws.merge_cells(str(merged_range))

                # 列幅のコピー
                for col_letter, dim in src_ws.column_dimensions.items():
                    dst_ws.column_dimensions[col_letter].width = dim.width

                # 行高さのコピー
                for row_num, dim in src_ws.row_dimensions.items():
                    dst_ws.row_dimensions[row_num].height = dim.height

                print(f"[consolidate] {f.filename} -> シート '{new_name}' を追加", flush=True)

        except Exception as e:
            errors.append(f'{f.filename}: {str(e)}')
            traceback.print_exc()

    if week_count == 0:
        return jsonify({'error': '有効な週シートがありません', 'details': errors}), 400

    # 統合ブックを保存
    output_path = os.path.join(sd['dir'], 'consolidated_booth.xlsx')
    meta_wb.save(output_path)

    # boothファイルとして登録
    sd['files'] = {**sd.get('files', {}), 'booth': output_path}
    save_session_files(sd)

    # 統合結果のシート一覧
    final_sheets = meta_wb.sheetnames

    return jsonify({
        'ok': True,
        'weekCount': week_count,
        'metaSheets': meta_sheet_names,
        'removedSheets': old_week_sheets,
        'finalSheets': final_sheets,
        'errors': errors,
        'files': {k: os.path.basename(v) for k, v in sd.get('files', {}).items()},
    })

@app.route('/api/generate', methods=['POST'])
@login_required
def generate():
    sd = get_session_data()
    files = sd.get('files',{})
    print(f"[generate] sid={sd.get('_sid','?')}, files_keys={list(files.keys())}", flush=True)
    if not all(k in files for k in ['src','booth']):
        print(f"[generate] ERROR: ファイル不足 files={files}", flush=True)
        return jsonify({'error': 'ファイルが不足しています。再度アップロードしてください。'}), 400
    # ファイルが実際に存在するか確認
    for k in ['src', 'booth']:
        if not os.path.exists(files[k]):
            print(f"[generate] ERROR: ファイルが見つかりません: {k}={files[k]}", flush=True)
            return jsonify({'error': f'{k}ファイルが見つかりません。再度アップロードしてください。'}), 400

    data = request.get_json() or {}
    office_rule = data.get('officeRule', {
        '月':['石川T'],'火':['石川T'],'水':['西T'],'木':['石川T'],'金':['石川T'],'土':['越智T']
    })
    booth_pref_ui = data.get('boothPref', {})
    booth_pref_ui = {k: int(v) for k, v in booth_pref_ui.items() if v}

    try:
        # ブース表xlsxから全データを読み込み
        booth_wb = openpyxl.load_workbook(files['booth'])
        skills = load_teacher_skills(booth_wb)
        file_booth_pref = load_booth_pref(booth_wb)
        students = load_students_from_wb(booth_wb)

        # ブース希望: UI設定を優先、なければファイルから読んだ値を使用
        booth_pref = {**file_booth_pref, **booth_pref_ui}

        wt = load_weekly_teachers(files['src'])
        total = sum(sum(s['needs'].values()) for s in students)

        # 休塾日検出
        holidays = load_holidays(booth_wb, len(wt))

        schedule, unplaced, office_teachers = build_schedule(
            students, wt, skills, office_rule, booth_pref, holidays=holidays
        )
        placed = sum(len(b['slots']) for w in schedule for d in w.values() for bs in d.values() for b in bs)

        # JSON用にtupleをlistに変換
        schedule_json = []
        for w in schedule:
            wj = {}
            for day, ds in w.items():
                dj = {}
                for ts, booths in ds.items():
                    bj = []
                    for b in booths:
                        bj.append({'teacher': b['teacher'], 'slots': [list(s) for s in b['slots']]})
                    dj[ts] = bj
                wj[day] = dj
            schedule_json.append(wj)

        # 週ごとの日付情報を取得
        week_dates = extract_week_dates(booth_wb, len(schedule))

        sd['result'] = {
            'schedule': schedule,
            'schedule_json': schedule_json,
            'unplaced': unplaced,
            'office_teachers': office_teachers,
            'office_rule': office_rule,
            'booth_pref': booth_pref,
            'students': students,
            'week_dates': week_dates,
        }
        save_session_result(sd)

        # 生徒データJSON化（全情報含む）
        students_json = []
        for s in students:
            avail_list = sorted([list(a) for a in s['avail']]) if s.get('avail') else None
            backup_list = sorted([list(a) for a in s['backup_avail']]) if s.get('backup_avail') else None
            fixed_list = [[d, t, subj] for d, t, subj in s.get('fixed', [])]
            students_json.append({
                'grade': s['grade'], 'name': s['name'],
                'needs': s['needs'],
                'avail': avail_list,
                'backup_avail': backup_list,
                'fixed': fixed_list,
                'notes': s.get('notes', ''),
                'ng_teachers': s['ng_teachers'],
                'wish_teachers': s['wish_teachers'],
                'ng_students': s['ng_students'],
                'ng_dates': [list(d) for d in s.get('ng_dates', set())],
            })

        return jsonify({
            'placed': placed,
            'total': total,
            'schedule': schedule_json,
            'unplaced': unplaced,
            'officeTeachers': office_teachers,
            'boothPref': booth_pref,
            'students': students_json,
            'weekDates': week_dates,
            'weeklyTeachers': wt,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/download')
@login_required
def download():
    sd = get_session_data()
    res = sd.get('result', {})
    if 'schedule' not in res:
        return jsonify({'error': '先にスケジュールを生成してください'}), 400

    output_path = os.path.join(sd['dir'], 'output.xlsx')
    try:
        # スケジュール全状態をJSON化してExcelに埋め込む（再読み込み用）
        state_json = {
            'schedule': res.get('schedule_json', []),
            'unplaced': res.get('unplaced', []),
            'officeTeachers': res.get('office_teachers', []),
            'boothPref': res.get('booth_pref', {}),
            'weekDates': res.get('week_dates'),
            'total': sum(sum(s['needs'].values()) for s in res.get('students', [])),
        }
        # students JSON化
        students_json = []
        for s in res.get('students', []):
            students_json.append({
                'grade': s['grade'], 'name': s['name'],
                'needs': s['needs'],
                'avail': sorted([list(a) for a in s['avail']]) if s.get('avail') else None,
                'backup_avail': sorted([list(a) for a in s['backup_avail']]) if s.get('backup_avail') else None,
                'fixed': [[d, t, subj] for d, t, subj in s.get('fixed', [])],
                'notes': s.get('notes', ''),
                'ng_teachers': s['ng_teachers'],
                'wish_teachers': s['wish_teachers'],
                'ng_students': s['ng_students'],
                'ng_dates': [list(d) for d in s.get('ng_dates', set())],
            })
        state_json['students'] = students_json
        # weeklyTeachers: srcファイルがあれば再取得、なければresultのキャッシュを利用
        wt = None
        if 'src' in sd.get('files', {}):
            try:
                wt = load_weekly_teachers(sd['files']['src'])
            except Exception:
                pass
        if not wt:
            wt = res.get('weekly_teachers')
        if wt:
            state_json['weeklyTeachers'] = wt
        # placed count
        placed = 0
        for w in res.get('schedule', []):
            for d_data in w.values():
                for bs in d_data.values():
                    for b in bs:
                        placed += len(b['slots'])
        state_json['placed'] = placed

        # office_teachers が不足している場合（古いバックアップ等）、デフォルト設定で補完
        _DEFAULT_RULE = {'月':['石川T'],'火':['石川T'],'水':['西T'],'木':['石川T'],'金':['石川T'],'土':['越智T']}
        ot_list = list(res.get('office_teachers', []))
        rule = res.get('office_rule') or _DEFAULT_RULE
        num_sched_weeks = len(res.get('schedule', []))
        while len(ot_list) < num_sched_weeks:
            if ot_list:
                ot_list.append(dict(ot_list[-1]))
            else:
                ot_list.append({d: rule[d][0] for d in DAYS if rule.get(d)})

        write_excel(
            res['schedule'],
            res['unplaced'],
            ot_list,
            sd['files']['booth'],
            output_path,
            state_json=state_json
        )
        return send_file(output_path, as_attachment=True, download_name='時間割_出力.xlsx')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def parse_schedule_from_wb(wb):
    """保存済みExcelの視覚シートからスケジュール配置を再構築する。
    Returns: (schedule, office_teachers) or (None, None) if no valid week sheets found.
    """
    meta_sheets = set(sn for sn in wb.sheetnames if any(k in sn for k in META_KEYWORDS))
    week_sheets = [sn for sn in wb.sheetnames
                   if sn not in meta_sheets and sn != '_schedule_data' and sn != '未配置コマ']
    if not week_sheets:
        return None, None

    schedule = []
    office_teachers = []
    
    # シート名に日付が含まれることが多いので、それを使ってソートしたほうが安全だが
    # write_excelでは week_sheets 順で書き込んでいるため、ここでは単純にリスト順とする

    for sn in week_sheets:
        ws = wb[sn]
        week = {}
        ot = {}
        # 教室業務（行5）
        for day, vals in DAY_COLS.items():
            bc = vals[0]
            val = ws.cell(5, bc).value
            ot[day] = str(val).strip() if val else ''
        
        # 配置データ
        for tl, (sr, nb) in LAYOUT.items():
            ts = TIME_SHORT[tl]
            
            for day, vals in DAY_COLS.items():
                if day not in week:
                    week[day] = {}
                
                lc, gc, sc, sjc = vals[1], vals[2], vals[3], vals[4]
                
                booths = []
                # 表示上のブース数 (nb) まで読むが、データ構造としては MAX_BOOTHS 分確保する
                for bi in range(MAX_BOOTHS):
                    if bi >= nb:
                        # LAYOUTで定義されたブース数を超えた分は空データ
                        booths.append({'teacher': '', 'slots': []})
                        continue

                    r1 = sr + bi * 2
                    r2 = r1 + 1
                    
                    teacher = ws.cell(r1, lc).value
                    teacher = str(teacher).strip() if teacher else ''
                    
                    slots = []
                    # 生徒1
                    g1 = ws.cell(r1, gc).value
                    s1 = ws.cell(r1, sc).value
                    j1 = ws.cell(r1, sjc).value
                    if s1:
                        slots.append([str(g1 or ''), str(s1), str(j1 or '')])
                    
                    # 生徒2
                    g2 = ws.cell(r2, gc).value
                    s2 = ws.cell(r2, sc).value
                    j2 = ws.cell(r2, sjc).value
                    if s2:
                        slots.append([str(g2 or ''), str(s2), str(j2 or '')])
                        
                    booths.append({'teacher': teacher, 'slots': slots})
                
                week[day][ts] = booths
        
        schedule.append(week)
        office_teachers.append(ot)
        
    return schedule, office_teachers

@app.route('/api/load_saved', methods=['POST'])
@login_required
def load_saved():
    """保存済みExcelからスケジュール状態を読み込む"""
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    ok, err = validate_file(f)
    if not ok:
        return jsonify({'error': err}), 400
    sd = get_session_data()
    path = os.path.join(sd['dir'], 'saved_' + f.filename)
    f.save(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        
        state = {'schedule': [], 'officeTeachers': [], 'weekDates': None, 'students': []}
        
        # 1. 隠しJSONシートがあれば読み込む（weekDates, students, boothPref等のため）
        if '_schedule_data' in wb.sheetnames:
            ws = wb['_schedule_data']
            chunks = []
            for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                if row[0] is not None:
                    chunks.append(str(row[0]))
            data_str = ''.join(chunks)
            try:
                state = json.loads(data_str)
            except:
                pass # JSON破損時は無視して視覚シートに頼る

        # 2. 視覚シートからスケジュール配置を上書き/復元（こちらを正とする）
        vis_schedule, vis_ot = parse_schedule_from_wb(wb)
        if vis_schedule:
            state['schedule'] = vis_schedule
            state['officeTeachers'] = vis_ot
        elif not state.get('schedule'):
             # JSONもなく視覚シートも解析できない場合
            wb.close()
            return jsonify({'error': 'このファイルには保存済みスケジュールが含まれていません。\nブース表DLしたファイルを選択してください。'}), 400

        # ======== 生徒データを保存済みExcelのシートから再取得 ========
        # 保存済みExcelは write_excel によりブース表のシートを含んでいるため
        # 別途ブース表アップロードが不要になる。
        if not state.get('weekDates'):
            # JSONがなくてweekDatesが不明な場合は、シート名から推測する
            num_weeks = len(state.get('schedule', []))
            extracted_wd = extract_week_dates(wb, num_weeks)
            if extracted_wd:
                state['weekDates'] = extracted_wd

        wd = state.get('weekDates') or {}
        year = wd.get('year', 2026)
        month = wd.get('month', 3)

        # set→list変換ヘルパー
        def serialize_student(s):
            avail = s.get('avail')
            backup = s.get('backup_avail')
            ng_dates = s.get('ng_dates')
            fixed = s.get('fixed')
            return {
                **s,
                'avail': sorted([list(a) for a in avail]) if isinstance(avail, set) else (avail or []),
                'backup_avail': sorted([list(a) for a in backup]) if isinstance(backup, set) else (backup or []),
                'ng_dates': [list(d) for d in ng_dates] if isinstance(ng_dates, set) else (ng_dates or []),
                'fixed': [list(f) for f in fixed] if fixed and isinstance(next(iter(fixed), None), (list, tuple)) else (fixed or []),
            }

        # 3. 生徒データとブース希望を保存済みExcelから再読み込み（常に最新の状態にする）
        # 元の `state['students']` はJSONパース結果だが、Excelシートの内容を優先する
        try:
            # 年月が必要だが、weekDatesから推測またはデフォルト値
            wd = state.get('weekDates') or {}
            year = wd.get('year', 2026)
            month = wd.get('month', 3)

            # 「必要コマ数」シートがあればそこから生徒情報をフルリロード
            if '必要コマ数' in wb.sheetnames:
                fresh_students = load_students_from_wb(wb, year, month)
                state['students'] = [serialize_student(s) for s in fresh_students]
                print(f"[load_saved] Reloaded {len(fresh_students)} students from '必要コマ数' sheet", flush=True)
            
            # 「講師ブース希望」シートがあればリロード
            fresh_booth_pref = load_booth_pref(wb)
            if fresh_booth_pref:
                 state['boothPref'] = fresh_booth_pref
                 print(f"[load_saved] Reloaded boothPref from '講師ブース希望' sheet", flush=True)

        except Exception as e:
            print(f"[load_saved] Warning: Failed to reload fresh data from sheets: {e}", flush=True)
            # 失敗しても致命的ではない（JSONデータ等を使う）ので続行

        # 必要ならファイルをセッションに登録（次回以降のために）
        # ただし saved_ は一時ファイル扱いなので、src/booth としては登録しないほうが無難
        # ここでは純粋に state を返すことに集中する


        wb.close()

        # ======== スケジュールの時間キー正規化 ========
        # 視覚シートから読んだ場合は既に正規化済みだが、古いJSON由来の場合は必要
        _time_normalize = {
            '14:55': '14', '16:00': '16', '17:05': '17',
            '18:10': '18', '19:15': '19', '20:20': '20',
        }
        raw_schedule = state.get('schedule', [])
        normalized_schedule = []
        for week in raw_schedule:
            nw = {}
            for day, day_data in week.items():
                nd = {}
                for ts_key, booths in day_data.items():
                    short_key = _time_normalize.get(ts_key, ts_key)
                    nd[short_key] = booths
                nw[day] = nd
            normalized_schedule.append(nw)
        state['schedule'] = normalized_schedule

        # ======== placed / total / unplaced を最新データで再計算 ========
        schedule = state['schedule']
        fresh_students_data = state.get('students', [])

        # 実際に配置されているコマ数
        placed = sum(len(b['slots']) for w in schedule for d in w.values() for bs in d.values() for b in bs)

        # 全必要コマ数を fresh_students から計算
        total = sum(sum(s.get('needs', {}).values()) for s in fresh_students_data)
        # total が 0 の場合は placed を使用（表示上の Infinity% を防ぐ）
        if total == 0:
            total = placed

        # 未配置コマを再計算
        # schedule 内に配置されている (name, subject) の数を集計
        placed_count = {}  # {(name, subj): count}
        for week in schedule:
            for day_slots in week.values():
                for booths in day_slots.values():
                    for b in booths:
                        for slot in b.get('slots', []):
                            key = (slot[1], slot[2])  # (name, subject)
                            placed_count[key] = placed_count.get(key, 0) + 1

        unplaced = []
        for s in fresh_students_data:
            name = s['name']
            grade = s.get('grade', '')
            for subj, need in s.get('needs', {}).items():
                done = placed_count.get((name, subj), 0)
                if done < need:
                    unplaced.append({
                        'grade': grade,
                        'name': name,
                        'subject': subj,
                        'count': need - done,
                        'reason': '保存済みファイルから復元'
                    })

        state['placed'] = placed
        state['total'] = total
        state['unplaced'] = unplaced

        # 安全策: weekDatesがない場合のデフォルト
        if not state.get('weekDates'):
            state['weekDates'] = {'year': year, 'month': month, 'weeks': []}

    except json.JSONDecodeError as e:
        return jsonify({'error': f'スケジュールデータの解析に失敗: {e}'}), 500
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

    # 保存済みファイル自体をブース表として登録（再ダウンロード用）
    sd['files'] = {**sd.get('files', {}), 'booth': path}
    save_session_files(sd)

    return jsonify({'ok': True, **state})

# ========== Schedule update API ==========
@app.route('/api/update_schedule', methods=['POST'])
@login_required
def update_schedule():
    sd = get_session_data()
    data = request.get_json()
    if not data or 'schedule' not in data:
        return jsonify({'error': 'Invalid data'}), 400

    sched_json = data['schedule']
    schedule = []
    for w in sched_json:
        wk = {}
        for day, ds in w.items():
            dk = {}
            for ts, booths in ds.items():
                bk = []
                for b in booths:
                    bk.append({'teacher': b['teacher'], 'slots': [tuple(s) for s in b['slots']]})
                dk[ts] = bk
            wk[day] = dk
        schedule.append(wk)

    res = sd.get('result', {})
    res['schedule'] = schedule
    res['schedule_json'] = sched_json
    res['unplaced'] = data.get('unplaced', [])
    sd['result'] = res
    save_session_result(sd)

    placed = sum(len(b['slots']) for w in schedule for d in w.values() for bs in d.values() for b in bs)
    return jsonify({'ok': True, 'placed': placed})

# ========== JSON restore API ==========
@app.route('/api/restore_json', methods=['POST'])
@login_required
def restore_json():
    """JSONバックアップ + 既存ファイルからスケジュール状態を復元する"""
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext != '.json':
        return jsonify({'error': f'JSONファイルを選択してください（選択: {ext}）'}), 400

    sd = get_session_data()
    try:
        raw = f.read().decode('utf-8')
        state = json.loads(raw)
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        return jsonify({'error': f'JSONの解析に失敗しました: {e}'}), 400

    schedule = state.get('schedule') or state.get('schedule_json')
    if not schedule:
        return jsonify({'error': 'スケジュールデータが含まれていません'}), 400

    unplaced = state.get('unplaced', [])
    office_teachers = state.get('officeTeachers') or state.get('office_teachers', [])
    booth_pref = state.get('boothPref') or state.get('booth_pref', {})
    students = state.get('students', [])
    week_dates = state.get('weekDates') or state.get('week_dates')
    weekly_teachers = state.get('weeklyTeachers') or state.get('weekly_teachers')
    placed = state.get('placed', 0)
    total = state.get('total', 0)

    if not placed:
        placed = sum(len(b['slots']) for w in schedule for d in w.values() for bs in d.values() for b in bs)
    if not total and students:
        total = sum(sum(s.get('needs', {}).values()) for s in students)

    # booth / src ファイルが同時にアップロードされた場合はセッションに保存
    files = dict(sd.get('files', {}))
    for key in ['booth', 'src']:
        fx = request.files.get(key)
        if fx and fx.filename:
            ok, err = validate_file(fx)
            if not ok:
                return jsonify({'error': f'{key}ファイル: {err}'}), 400
            path = os.path.join(sd['dir'], key + '_' + fx.filename)
            fx.save(path)
            files[key] = path
    sd['files'] = files
    save_session_files(sd)

    # srcがあればweeklyTeachersを再取得（最新化）
    if 'src' in files:
        try:
            weekly_teachers = load_weekly_teachers(files['src'])
        except Exception:
            pass

    sd['result'] = {
        'schedule_json': schedule,
        'schedule': schedule,
        'unplaced': unplaced,
        'office_teachers': office_teachers,
        'booth_pref': booth_pref,
        'students': students,
        'week_dates': week_dates,
        'weekly_teachers': weekly_teachers,  # ダウンロード時のフォールバック用
    }
    save_session_result(sd)

    resp = {
        'ok': True,
        'placed': placed,
        'total': total,
        'schedule': schedule,
        'unplaced': unplaced,
        'officeTeachers': office_teachers,
        'boothPref': booth_pref,
        'students': students,
        'weekDates': week_dates,
        'hasBooth': 'booth' in files,
    }
    if weekly_teachers:
        resp['weeklyTeachers'] = weekly_teachers
    return jsonify(resp)

# ========== State persistence API ==========
@app.route('/api/state')
@login_required
def get_state():
    """保存済みスケジュール状態を返す（ページリロード時の復元用）"""
    sd = get_session_data()
    res = sd.get('result', {})

    # インメモリキャッシュにあればそれを使う
    if res and 'schedule_json' in res:
        students_json = []
        for s in res.get('students', []):
            avail_list = sorted([list(a) for a in s['avail']]) if isinstance(s.get('avail'), set) else s.get('avail')
            backup_list = sorted([list(a) for a in s['backup_avail']]) if isinstance(s.get('backup_avail'), set) else s.get('backup_avail')
            fixed_list = [list(f) for f in s.get('fixed', [])] if s.get('fixed') else []
            ng_dates_list = [list(d) for d in s.get('ng_dates', set())] if isinstance(s.get('ng_dates'), set) else s.get('ng_dates', [])
            students_json.append({
                'grade': s['grade'], 'name': s['name'],
                'needs': s.get('needs', {}),
                'avail': avail_list,
                'backup_avail': backup_list,
                'fixed': fixed_list,
                'notes': s.get('notes', ''),
                'ng_teachers': s.get('ng_teachers', []),
                'wish_teachers': s.get('wish_teachers', []),
                'ng_students': s.get('ng_students', []),
                'ng_dates': ng_dates_list,
            })
        placed = sum(len(b['slots']) for w in res['schedule_json'] for d in w.values() for bs in d.values() for b in bs)
        total = sum(sum(s.get('needs', {}).values()) for s in res.get('students', []))
        return jsonify({
            'has_state': True,
            'placed': placed,
            'total': total,
            'schedule': res['schedule_json'],
            'unplaced': res.get('unplaced', []),
            'officeTeachers': res.get('office_teachers', []),
            'boothPref': res.get('booth_pref', {}),
            'students': students_json,
            'weekDates': res.get('week_dates') or {'year':2026, 'month':3, 'weeks':[]},
        })

    # ディスクから復元を試みる
    sid = sd.get('_sid')
    if sid:
        disk_result = _load_result_from_disk(sid)
        if disk_result and 'schedule_json' in disk_result:
            students_json = []
            for s in disk_result.get('students', []):
                students_json.append({
                    'grade': s.get('grade', ''), 'name': s.get('name', ''),
                    'needs': s.get('needs', {}),
                    'avail': s.get('avail') if s.get('avail') else [],
                    'backup_avail': s.get('backup_avail') if s.get('backup_avail') else [],
                    'fixed': s.get('fixed', []),
                    'notes': s.get('notes', ''),
                    'ng_teachers': s.get('ng_teachers', []),
                    'wish_teachers': s.get('wish_teachers', []),
                    'ng_students': s.get('ng_students', []),
                    'ng_dates': s.get('ng_dates', []),
                })
            placed = sum(len(b['slots']) for w in disk_result['schedule_json'] for d in w.values() for bs in d.values() for b in bs)
            total = sum(sum(s.get('needs', {}).values()) for s in disk_result.get('students', []))

            # インメモリキャッシュに復元
            sd['result'] = {
                'schedule_json': disk_result['schedule_json'],
                'schedule': disk_result['schedule_json'],  # JSON形式で保持
                'unplaced': disk_result.get('unplaced', []),
                'office_teachers': disk_result.get('office_teachers', []),
                'booth_pref': disk_result.get('booth_pref', {}),
                'students': disk_result.get('students', []),
                'week_dates': disk_result.get('week_dates'),
            }
            save_session_result(sd)

            return jsonify({
                'has_state': True,
                'placed': placed,
                'total': total,
                'schedule': disk_result['schedule_json'],
                'unplaced': disk_result.get('unplaced', []),
                'officeTeachers': disk_result.get('office_teachers', []),
                'boothPref': disk_result.get('booth_pref', {}),
                'students': students_json,
                'weekDates': disk_result.get('week_dates') or {'year':2026, 'month':3, 'weeks':[]},
            })

    return jsonify({'has_state': False})

# ========== 起動 ==========
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"\n  Booth Schedule Generator (Cloud)")
    print(f"  http://localhost:{port}\n")
    app.run(host='0.0.0.0', port=port, debug=False)
