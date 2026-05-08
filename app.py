
from flask import Flask, render_template, request, send_file, jsonify
from flask_httpauth import HTTPBasicAuth
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import openpyxl
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
        elif weekday >= 5:
            pass  # 土日：空欄
        elif day in holidays:
            ws[f'{col_note}{row}'].value = '祝日'
        elif day in paid_leave:
            ws[f'{col_note}{row}'].value = '私用により、休暇'
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


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=False, port=5000)
