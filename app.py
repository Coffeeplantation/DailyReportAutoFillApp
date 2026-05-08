
from flask import Flask, render_template, request, send_file, jsonify
from flask_httpauth import HTTPBasicAuth
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import jpholiday
from datetime import date, time
import calendar
import io
import os
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB
auth = HTTPBasicAuth()

USERS = {
    os.getenv('APP_USERNAME', 'root'): generate_password_hash(os.getenv('APP_PASSWORD', 'root')),
}

@auth.verify_password
def verify_password(username, password):
    if username in USERS and check_password_hash(USERS[username], password):
        return username


def parse_time(time_str):
    if not time_str:
        return None
    try:
        parts = time_str.split(':')  # HH:MM または HH:MM:SS に対応
        return time(int(parts[0]), int(parts[1]))
    except Exception:
        return None


def find_column(ws, header_row, label):
    """ヘッダー行からラベルに一致する列を返す（部分一致対応）。見つからなければ None。"""
    label = label.strip()
    if not label:
        return None
    for cell in ws[header_row]:
        if cell.value:
            cell_val = str(cell.value).strip()
            if cell_val == label or label in cell_val or cell_val in label:
                return cell.column_letter
    return None


def find_header_row(ws, labels):
    """ラベルを2つ以上含む行をヘッダー行として検索する。見つからなければ None。"""
    clean = [l.strip() for l in labels if l.strip()]
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        row_vals = [str(c.value).strip() for c in row if c.value]
        hits = sum(
            1 for label in clean
            if any(label in v or v in label for v in row_vals)
        )
        if hits >= 2:
            return row[0].row
    return None


@app.route('/')
@auth.login_required
def index():
    today = date.today()
    return render_template('index.html',
                           current_month=today.month,
                           current_year=today.year)


@app.route('/api/holidays')
@auth.login_required
def get_holidays():
    today = date.today()
    try:
        year  = int(request.args.get('year',  today.year))
        month = int(request.args.get('month', today.month))
        if not (2000 <= year <= 2099) or not (1 <= month <= 12):
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({'error': '年月の値が正しくありません'}), 400
    _, last_day = calendar.monthrange(year, month)
    holidays = {}
    for day in range(1, last_day + 1):
        name = jpholiday.is_holiday_name(date(year, month, day))
        if name:
            holidays[str(day)] = name
    return jsonify(holidays)


@app.route('/api/write', methods=['POST'])
@auth.login_required
def write_excel():
    excel_file = request.files.get('excel_file')
    if not excel_file:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    safe_name = secure_filename(excel_file.filename) or 'report.xlsx'

    # 有給取得日
    paid_leave_str = request.form.get('paid_leave_dates', '')
    paid_leave = set()
    for d in paid_leave_str.split(','):
        d = d.strip()
        if d.isdigit():
            paid_leave.add(int(d))

    # 曜日別勤務時間（0=月〜4=金）
    day_names = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']
    weekday_times = {}
    for i, name in enumerate(day_names):
        start = parse_time(request.form.get(f'{name}_start'))
        end = parse_time(request.form.get(f'{name}_end'))
        brk = parse_time(request.form.get(f'{name}_break')) or time(1, 0)
        if start and end:
            weekday_times[i] = {'start': start, 'end': end, 'break': brk}

    today = date.today()
    try:
        year  = int(request.form.get('year',  today.year))
        month = int(request.form.get('month', today.month))
        if not (2000 <= year <= 2099) or not (1 <= month <= 12):
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({'error': '年月の値が正しくありません（年：2000〜2099、月：1〜12）'}), 400
    _, last_day = calendar.monthrange(year, month)

    # 祝日取得
    holidays = set()
    for day in range(1, last_day + 1):
        if jpholiday.is_holiday(date(year, month, day)):
            holidays.add(day)

    # Excel処理
    try:
        wb = openpyxl.load_workbook(excel_file)
    except Exception:
        return jsonify({'error': 'ファイルを開けませんでした。Excel形式（.xlsx）のファイルを選択してください。'}), 400
    ws = wb.active

    label_start  = request.form.get('label_start',  '開始時間')
    label_end    = request.form.get('label_end',    '終了時間')
    label_break  = request.form.get('label_break',  '休憩時間')
    label_note   = request.form.get('label_note',   '備考')
    note_workday = request.form.get('note_workday', '在宅勤務')

    # ヘッダー行をラベルで動的検索（見つからなければ19行目をフォールバック）
    HEADER_ROW    = find_header_row(ws, [label_start, label_end, label_break, label_note]) or 19
    DATA_START_ROW = HEADER_ROW + 1

    # 例外日の備考（日付→備考のマッピング）
    ex_days  = request.form.getlist('exception_day')
    ex_notes = request.form.getlist('exception_note')
    note_exceptions = {}
    for d, n in zip(ex_days, ex_notes):
        if d.isdigit():
            note_exceptions[int(d)] = n.strip() or None

    # 例外日の勤務時間（日付→時間のマッピング）
    time_ex_days   = request.form.getlist('time_ex_day')
    time_ex_starts = request.form.getlist('time_ex_start')
    time_ex_ends   = request.form.getlist('time_ex_end')
    time_ex_breaks = request.form.getlist('time_ex_break')
    time_exceptions = {}
    for d, s, e, b in zip(time_ex_days, time_ex_starts, time_ex_ends, time_ex_breaks):
        if d.isdigit():
            start = parse_time(s)
            end   = parse_time(e)
            brk   = parse_time(b) or time(1, 0)
            if start and end:
                time_exceptions[int(d)] = {'start': start, 'end': end, 'break': brk}

    col_start = find_column(ws, HEADER_ROW, label_start) or 'F'
    col_end   = find_column(ws, HEADER_ROW, label_end)   or 'I'
    col_break = find_column(ws, HEADER_ROW, label_break) or 'L'
    col_note  = find_column(ws, HEADER_ROW, label_note)  or 'S'

    for day in range(1, last_day + 1):
        row = DATA_START_ROW + (day - 1)
        weekday = date(year, month, day).weekday()  # 0=月, 6=日

        # 対象セルをクリア
        for col in [col_start, col_end, col_break, col_note]:
            ws[f'{col}{row}'].value = None

        if day in time_exceptions:
            # 時間例外が最優先（土日・祝日・有給より上書き可能）
            t = time_exceptions[day]
            ws[f'{col_start}{row}'].value = t['start']
            ws[f'{col_start}{row}'].number_format = 'h:mm'
            ws[f'{col_end}{row}'].value = t['end']
            ws[f'{col_end}{row}'].number_format = 'h:mm'
            ws[f'{col_break}{row}'].value = t['break']
            ws[f'{col_break}{row}'].number_format = 'h:mm'
            ws[f'{col_note}{row}'].value = note_exceptions[day] if day in note_exceptions else note_workday
        elif day in paid_leave:
            ws[f'{col_note}{row}'].value = '私用により、休暇'
        elif weekday >= 5:
            pass  # 土日：空欄
        elif day in holidays:
            ws[f'{col_note}{row}'].value = '祝日'
        elif weekday in weekday_times:
            t = weekday_times[weekday]
            ws[f'{col_start}{row}'].value = t['start']
            ws[f'{col_start}{row}'].number_format = 'h:mm'
            ws[f'{col_end}{row}'].value = t['end']
            ws[f'{col_end}{row}'].number_format = 'h:mm'
            ws[f'{col_break}{row}'].value = t['break']
            ws[f'{col_break}{row}'].number_format = 'h:mm'
            ws[f'{col_note}{row}'].value = note_exceptions[day] if day in note_exceptions else note_workday

    # カーソルをA1に設定
    try:
        ws.sheet_view.selection[0].activeCell = 'A1'
        ws.sheet_view.selection[0].sqref = 'A1'
    except (IndexError, AttributeError):
        pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=safe_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def _build_manual() -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "使い方ガイド"

    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 52
    ws.column_dimensions['D'].width = 28

    C_NAVY  = "1A56DB"; C_ACCENT = "DBEAFE"; C_GREEN = "059669"
    C_ORANGE = "EA580C"; C_GRAY = "F3F4F6"; C_WHITE = "FFFFFF"; C_BORDER = "CBD5E1"

    def _side(c=C_BORDER): return Side(border_style="thin", color=c)
    def _border(): return Border(left=_side(), right=_side(), top=_side(), bottom=_side())
    def _fill(c): return PatternFill("solid", fgColor=c)
    def _font(bold=False, size=11, color="000000", italic=False):
        return Font(name="Meiryo", bold=bold, size=size, color=color, italic=italic)
    def _align(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def _merge(r1, c1, r2, c2): ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    def _rh(row, h): ws.row_dimensions[row].height = h

    # ── タイトル ──
    for r in range(1, 5): _rh(r, 6 if r in (1, 4) else 24)
    for r in range(1, 5):
        for c in range(1, 5): ws.cell(r, c).fill = _fill(C_NAVY)
    _merge(2, 1, 3, 4)
    c = ws.cell(2, 1, "日報自動入力アプリ  使い方ガイド")
    c.font = Font(name="Meiryo", bold=True, size=18, color=C_WHITE)
    c.alignment = _align("center", "center")

    # ── アプリ概要 ──
    _rh(5, 8)
    _merge(6, 1, 6, 4); _rh(6, 28)
    c = ws.cell(6, 1, "■  アプリ概要")
    c.font = _font(bold=True, size=13, color=C_WHITE); c.fill = _fill(C_NAVY)
    c.alignment = _align(); c.border = _border()

    overview = [
        ("目的",     "毎月会社から配布されるExcel作業報告書に、勤務時間・備考を自動で書き込み・保存するWebアプリです。"),
        ("メリット", "月初に最小限の入力をするだけで、1ヶ月分のExcelが自動で完成します。手入力の手間をゼロにします。"),
        ("対象",     "Excel形式（.xlsx）の作業報告書を毎月提出している社員が対象です。"),
        ("動作環境", "ブラウザ（Chrome / Edge / Safari など）で動作します。インストール不要です。"),
    ]
    for i, (label, text) in enumerate(overview):
        r = 7 + i; _rh(r, 36)
        c1 = ws.cell(r, 2, label)
        c1.font = _font(bold=True, size=10, color=C_NAVY); c1.fill = _fill(C_ACCENT)
        c1.alignment = _align("center"); c1.border = _border()
        _merge(r, 3, r, 4)
        c2 = ws.cell(r, 3, text)
        c2.font = _font(size=10); c2.fill = _fill(C_WHITE)
        c2.alignment = _align(wrap=True); c2.border = _border()

    # ── 使い方手順 ──
    _rh(12, 8)
    _merge(13, 1, 13, 4); _rh(13, 28)
    c = ws.cell(13, 1, "■  使い方手順")
    c.font = _font(bold=True, size=13, color=C_WHITE); c.fill = _fill(C_GREEN)
    c.alignment = _align(); c.border = _border()

    steps = [
        ("STEP 1", "アプリを開く",
         "ブラウザでアプリのURLにアクセスします。ユーザー名・パスワードを入力してログインしてください。"),
        ("STEP 2", "対象月を確認する",
         "画面上部に「年・月」が表示されます。自動で当月が設定されます。\n異なる月を処理したい場合は数字を直接変更してください。"),
        ("STEP 3", "曜日別の勤務時間を入力する",
         "月〜金それぞれの開始・終了・休憩時間を入力します。\n設定はブラウザに自動保存され、次回起動時に引き継がれます。"),
        ("STEP 4", "例外日を設定する（任意）",
         "「＋ 例外日を追加」から残業・早退・休日出勤など通常と異なる日を個別設定できます。\n土日・祝日も選択可能で、例外日設定はすべての設定より優先されます。"),
        ("STEP 5", "有給取得日を選択する（任意）",
         "「有給を取得する日がある」にチェックを入れるとカレンダーが表示されます。\n取得する日をクリックして選択してください（土日・祝日も選択可能）。"),
        ("STEP 6", "Excelファイルを選択する",
         "ファイル選択エリアをクリックし、会社から配布された作業報告書（.xlsx）を選びます。\n選択後、✕ ボタンで選択をやり直せます。"),
        ("STEP 7", "入力完了・ダウンロード",
         "「入力完了・ダウンロード」ボタンを押すと、自動入力済みのExcelファイルがダウンロードされます。\nファイルを開いて内容を確認し、所定の場所に保存してください。"),
    ]
    step_colors = [
        ("1E40AF","DBEAFE"),("065F46","D1FAE5"),("92400E","FEF3C7"),
        ("7C3AED","EDE9FE"),("BE185D","FCE7F3"),("1E40AF","DBEAFE"),("065F46","D1FAE5"),
    ]
    cur = 14
    for i, (step, title, desc) in enumerate(steps):
        fc, bc = step_colors[i]; _rh(cur, 14)
        c1 = ws.cell(cur, 2, step)
        c1.font = Font(name="Meiryo", bold=True, size=9, color=C_WHITE)
        c1.fill = _fill(fc); c1.alignment = _align("center"); c1.border = _border()
        ws.merge_cells(start_row=cur, start_column=3, end_row=cur, end_column=4)
        c2 = ws.cell(cur, 3, title)
        c2.font = Font(name="Meiryo", bold=True, size=11, color=fc)
        c2.fill = _fill(bc); c2.alignment = _align(); c2.border = _border()
        ws.cell(cur, 4).fill = _fill(bc); ws.cell(cur, 4).border = _border()
        cur += 1
        lc = desc.count('\n') + 1; _rh(cur, max(32, lc * 28))
        ws.merge_cells(start_row=cur, start_column=2, end_row=cur, end_column=4)
        c3 = ws.cell(cur, 2, desc)
        c3.font = _font(size=10); c3.fill = _fill(C_WHITE)
        c3.alignment = _align(wrap=True); c3.border = _border()
        cur += 1; _rh(cur, 5); cur += 1

    # ── 自動入力ルール ──
    _rh(cur, 8); cur += 1
    _merge(cur, 1, cur, 4); _rh(cur, 28)
    c = ws.cell(cur, 1, "■  自動入力ルール")
    c.font = _font(bold=True, size=13, color=C_WHITE); c.fill = _fill(C_ORANGE)
    c.alignment = _align(); c.border = _border(); cur += 1
    _rh(cur, 22)
    for j, h in enumerate(["日の種類", "開始・終了・休憩", "備考欄"]):
        cx = ws.cell(cur, 2 + j, h)
        cx.font = _font(bold=True, size=10, color=C_WHITE); cx.fill = _fill(C_ORANGE)
        cx.alignment = _align("center"); cx.border = _border()
    cur += 1
    for kind, times, note, bg in [
        ("出勤日（平日）",    "曜日別設定の時間・休憩 1:00", "在宅勤務（変更可）", C_WHITE),
        ("例外日（休日出勤）","例外設定した時間・休憩 1:00", "在宅勤務（変更可）", "FFF9C4"),
        ("有給取得日",        "空欄",                        "私用により、休暇",   "FCE7F3"),
        ("祝日",              "空欄",                        "祝日",               "FFF5F5"),
        ("土日",              "空欄",                        "空欄",               C_GRAY),
    ]:
        _rh(cur, 28)
        for col, val in zip([2, 3, 4], [kind, times, note]):
            cx = ws.cell(cur, col, val)
            cx.font = _font(size=10); cx.fill = _fill(bg)
            cx.alignment = _align("center"); cx.border = _border()
        cur += 1

    # ── よくある疑問 ──
    _rh(cur, 8); cur += 1
    _merge(cur, 1, cur, 4); _rh(cur, 28)
    c = ws.cell(cur, 1, "■  よくある疑問")
    c.font = _font(bold=True, size=13, color=C_WHITE); c.fill = _fill("6D28D9")
    c.alignment = _align(); c.border = _border(); cur += 1
    for q, a in [
        ("設定は毎月入力し直す？",
         "いいえ。曜日別の勤務時間・各種ラベルはブラウザに自動保存されます。次回起動時には前回の設定が自動で反映されます。"),
        ("休日出勤はどう入力する？",
         "「＋ 例外日を追加」から該当日を選択（土日・祝日も選択可）し、出勤時間を入力します。例外日設定は最優先で反映されます。"),
        ("有給を祝日や土日に取得したい場合は？",
         "有給カレンダーで土日・祝日もクリック選択できます。選択した日は「私用により、休暇」が記入されます。"),
        ("ファイルを間違えて選択した場合は？",
         "ファイル名の右に表示される「✕」ボタンを押すと選択をキャンセルできます。再度クリックして正しいファイルを選び直してください。"),
        ("書き込まれる列や行がずれている場合は？",
         "アプリが開始時間・終了時間・休憩時間・備考のヘッダーを自動検索して列を特定します。Excelのヘッダー文字列と設定が一致しているか確認してください。"),
    ]:
        _rh(cur, 18)
        _merge(cur, 2, cur, 4)
        cq = ws.cell(cur, 2, f"Q.  {q}")
        cq.font = Font(name="Meiryo", bold=True, size=10, color="6D28D9")
        cq.fill = _fill("EDE9FE"); cq.alignment = _align(); cq.border = _border(); cur += 1
        _rh(cur, max(32, a.count('\n') * 28 + 28))
        _merge(cur, 2, cur, 4)
        ca = ws.cell(cur, 2, f"A.  {a}")
        ca.font = _font(size=10); ca.fill = _fill(C_WHITE)
        ca.alignment = _align(wrap=True); ca.border = _border(); cur += 1
        _rh(cur, 4); cur += 1

    # ── フッター ──
    _rh(cur, 8); cur += 1
    _merge(cur, 1, cur, 4); _rh(cur, 22)
    cf = ws.cell(cur, 1, "日報自動入力アプリ  操作マニュアル  ／  本ファイルはアプリから自動生成されました")
    cf.font = _font(size=9, color="94A3B8", italic=True); cf.fill = _fill(C_NAVY)
    cf.alignment = _align("center")
    for r in range(1, cur + 1): ws.cell(r, 1).fill = _fill(C_NAVY)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


@app.route('/api/manual')
@auth.login_required
def download_manual():
    buf = _build_manual()
    return send_file(
        buf,
        as_attachment=True,
        download_name='日報アプリ_使い方ガイド.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=False, port=5000)
