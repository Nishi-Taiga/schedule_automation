#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Booth Schedule Generator – Cloud Edition (Render)
Flask + gunicorn + openpyxl
"""
import os, sys, json, random, threading, tempfile, shutil, time, secrets, atexit, traceback
from collections import defaultdict
from functools import wraps
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import openpyxl
from openpyxl.styles import Font, Alignment

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
    """resultをインメモリキャッシュに保存"""
    sid = sd['_sid']
    if not hasattr(get_session_data, '_cache'):
        get_session_data._cache = {}
    get_session_data._cache[sid] = {'result': sd['result']}

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
    ('西　大地','西T'),('西 大地','西T'),
    ('飯村　','飯村T'),
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
    8:'受国',9:'受算',10:'受英',11:'受理',12:'受社',
    13:'中国',14:'中数',15:'中英',16:'中理',17:'中社',
    18:'高現',19:'高古',
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
    if g.startswith('S'):
        yr = int(g[1:]) if len(g)>1 else 0
        if yr >= 4: return ['受算'] if s=='数' else ['受'+s]
        return ['小'+s]
    elif g.startswith('C'): return ['中'+s]
    elif g.startswith('K'):
        if s=='数': return ['高ⅠA','高ⅡB','高Ⅲ','高C']
        return ['高'+s]
    return ['中'+s]

def can_teach(teacher, grade, subject, skills):
    keys = get_skill_keys(grade, subject)
    if teacher not in skills: return True
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
        for c in range(20, ws.max_column+1):
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
                 (17,'日'),(18,'地'),(19,'政'),(20,'世'),(21,'作')]
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
            'wish_teachers':parse_list(ws.cell(r,23).value),
            'ng_teachers':parse_list(ws.cell(r,24).value),
            'ng_students':parse_list(ws.cell(r,25).value),
            'avail':parse_avail(ws.cell(r,26).value),
            'ng_dates':parse_ng_dates(ws.cell(r,27).value, year, month),
            'fixed':parse_regular(ws.cell(r,28).value),
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
    for wi in range(min(4, len(wb.sheetnames))):
        ws = wb[wb.sheetnames[wi]]
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
        weeks.append(week)
    return weeks

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

    # 各時間帯のブースリストを生成
    result = {}
    for ts in ts_list:
        available = [t for t in day_data.get(ts, []) if t in selected and t != office_teacher]
        result[ts] = assign_booth_order(available)
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

# ========== スケジューラー ==========
def build_schedule(students, weekly_teachers, skills, office_rule, booth_pref):
    remaining = {s['name']: dict(s['needs']) for s in students}
    smap = {s['name']: s for s in students}
    schedule = []
    office_teachers = []

    # 全生徒の希望講師を集約
    wish_teachers_set = set()
    for s in students:
        wish_teachers_set.update(s['wish_teachers'])

    for wi in range(4):
        ot = {}
        for d in DAYS:
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
                ds[ts] = [{'teacher':t, 'slots':[]} for t in tlist]
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
        # 隣接ブース(bi-1, bi+1)のNG生徒チェック
        booths_in_ts = ws.get(day,{}).get(
            next((ts for ts, bs in ws.get(day,{}).items() if any(b is booth for b in bs)), None), [])
        # ts特定のため別アプローチ
        for ts_key, bs in ws.get(day,{}).items():
            if bi < len(bs) and bs[bi] is booth:
                for adj_bi in [bi-1, bi+1]:
                    if 0 <= adj_bi < len(bs):
                        adj_booth = bs[adj_bi]
                        for g2,sn2,sb2 in adj_booth['slots']:
                            if sn2 in s['ng_students']: return False
                            other = smap.get(sn2)
                            if other and s['name'] in other.get('ng_students',[]): return False
                break
        eb = get_teacher_booth(ws, day, t)
        if eb is not None and eb != bi: return False
        return True

    def place_student(ws, s, day, ts, subj):
        if day not in ws or ts not in ws[day]: return False
        for bi,b in enumerate(ws[day][ts]):
            if check_booth(b, bi, s, day, subj, ws):
                b['slots'].append((s['grade'],s['name'],subj))
                return True
        return False

    def find_slot(ws, s, subj, placed_days, existing, wi, any_placed_days):
        cands = []
        for day in DAYS:
            if day in placed_days: continue  # 同一科目の同曜日配置を防止
            # NG日程チェック
            if (wi, day) in s.get('ng_dates', set()): continue
            times = SATURDAY_TIMES if day=='土' else WEEKDAY_TIMES
            for tl in times:
                ts = TIME_SHORT[tl]
                if s['avail'] is not None and (day,ts) not in s['avail']: continue
                if (day,ts) in existing: continue
                if ts not in ws.get(day,{}): continue
                for bi,b in enumerate(ws[day][ts]):
                    if not check_booth(b, bi, s, day, subj, ws): continue
                    sc = 0
                    # 同曜日に既に別科目が配置されている場合を最優先
                    if day in any_placed_days: sc += 150
                    if b['teacher'] in s['wish_teachers']: sc += 100
                    t = b['teacher']
                    if t in booth_pref and booth_pref[t]==bi+1: sc += 10
                    if len(b['slots'])==0: sc += 20
                    cands.append((sc, day, ts, bi))
        if not cands: return None
        cands.sort(key=lambda x:-x[0])
        best_sc = cands[0][0]
        bests = [c for c in cands if c[0]==best_sc]
        ch = random.choice(bests)
        return ch[1], ch[2], ch[3]

    def distribute(total, weeks):
        t = [total//weeks]*weeks
        for i in range(total%weeks): t[i] += 1
        random.shuffle(t)
        return t

    # Phase1: 固定授業
    for s in students:
        for day, ts_str, subj in s['fixed']:
            for wi in range(4):
                if (wi, day) in s.get('ng_dates', set()): continue
                if place_student(schedule[wi], s, day, ts_str, subj):
                    if remaining[s['name']].get(subj,0)>0:
                        remaining[s['name']][subj] -= 1

    # Phase2: 通常配置
    order = sorted(students, key=lambda s: (
        len(s['avail']) if s['avail'] else 999, sum(s['needs'].values())
    ))
    for s in order:
        for subj, total in s['needs'].items():
            still = remaining[s['name']].get(subj, 0)
            if still <= 0: continue
            targets = distribute(still, 4)
            for wi in range(4):
                for _ in range(targets[wi]):
                    if remaining[s['name']].get(subj,0) <= 0: break
                    pd = get_placed_days(schedule[wi], s['name'], subj)
                    ex = get_student_slots(schedule[wi], s['name'])
                    apd = get_any_placed_days(schedule[wi], s['name'])
                    best = find_slot(schedule[wi], s, subj, pd, ex, wi, apd)
                    if best:
                        day, ts, bi = best
                        schedule[wi][day][ts][bi]['slots'].append((s['grade'],s['name'],subj))
                        remaining[s['name']][subj] -= 1

    unplaced = []
    for s in students:
        for subj, cnt in remaining[s['name']].items():
            if cnt > 0:
                unplaced.append({'grade':s['grade'],'name':s['name'],'subject':subj,'count':cnt})

    return schedule, unplaced, office_teachers

# ========== Excel出力 ==========
def write_excel(schedule, unplaced, office_teachers, booth_path, output_path):
    wb = openpyxl.load_workbook(booth_path)
    for sn in wb.sheetnames[4:]:
        del wb[sn]

    # 共通書式
    teacher_font = Font(name='MS PGothic', size=8)
    teacher_align = Alignment(textRotation=255, vertical='center', horizontal='center')
    data_font = Font(name='MS PGothic', size=11)
    data_align = Alignment(vertical='center', horizontal='center')

    for wi in range(4):
        ws = wb[wb.sheetnames[wi]]
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
        ot = office_teachers[wi]
        for day in DAYS:
            bc = DAY_COLS[day][0]
            t = ot.get(day)
            if t:
                ws.cell(5, bc, t)
                for tr in TUTOR_ROWS:
                    try: ws.cell(tr, bc, t)
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
        wb = openpyxl.load_workbook(files['booth'])
        skills = load_teacher_skills(wb)
        booth_pref = load_booth_pref(wb)
        if not booth_pref:
            booth_pref = dict(DEFAULT_BOOTH_PREF)
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

        schedule, unplaced, office_teachers = build_schedule(
            students, wt, skills, office_rule, booth_pref
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

        sd['result'] = {
            'schedule': schedule,
            'schedule_json': schedule_json,
            'unplaced': unplaced,
            'office_teachers': office_teachers,
            'booth_pref': booth_pref,
            'students': students,
        }
        save_session_result(sd)

        # 生徒データJSON化（NG情報含む）
        students_json = []
        for s in students:
            students_json.append({
                'grade': s['grade'], 'name': s['name'],
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
        write_excel(
            res['schedule'],
            res['unplaced'],
            res['office_teachers'],
            sd['files']['booth'],
            output_path
        )
        return send_file(output_path, as_attachment=True, download_name='時間割_出力.xlsx')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

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

# ========== 起動 ==========
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"\n  Booth Schedule Generator (Cloud)")
    print(f"  http://localhost:{port}\n")
    app.run(host='0.0.0.0', port=port, debug=False)
