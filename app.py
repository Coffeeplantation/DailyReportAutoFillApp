from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
import jpholiday
from datetime import date, time
import calendar
import io

app = Flask(__name__)


def parse_time(time_str):
    if not time_str:
        return None
    try:
        h, m = time_str.split(':')
        return time(int(h), int(m))
    except Exception:
        return None


def find_column(ws, header_row, label):
    """ヘッダー行からラベルに一致する列を返す（部分一致対応）。見つからなければ None。"""
    label = label.strip()
    for cell in ws[header_row]:
        if cell.value:
            cell_val = str(cell.value).strip()
            if cell_val == label or label in cell_val or cell_val in label:
                return cell.column_letter
    return None


@app.route('/')
def index():
    today = date.today()
    return render_template('index.html',
                           current_month=today.month,
                           current_year=today.year)


@app.route('/api/holidays')
def get_holidays():
    today = date.today()
    year  = int(request.args.get('year',  today.year))
    month = int(request.args.get('month', today.month))
    _, last_day = calendar.monthrange(year, month)
    holidays = {}
    for day in range(1, last_day + 1):
        name = jpholiday.is_holiday_name(date(year, month, day))
        if name:
            holidays[str(day)] = name
    return jsonify(holidays)


@app.route('/api/write', methods=['POST'])
def write_excel():
    excel_file = request.files.get('excel_file')
    if not excel_file:
        return jsonify({'error': 'ファイルが選択されていません'}), 400

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
    year  = int(request.form.get('year',  today.year))
    month = int(request.form.get('month', today.month))
    _, last_day = calendar.monthrange(year, month)

    # 祝日取得
    holidays = set()
    for day in range(1, last_day + 1):
        if jpholiday.is_holiday(date(year, month, day)):
            holidays.add(day)

    # Excel処理
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    # ヘッダー行からラベルで列を検索（見つからなければデフォルト列を使用）
    HEADER_ROW = 19
    label_start    = request.form.get('label_start',   '開始時間')
    label_end      = request.form.get('label_end',     '終了時間')
    label_break    = request.form.get('label_break',   '休憩時間')
    label_note     = request.form.get('label_note',    '備考')
    note_workday   = request.form.get('note_workday',  '在宅勤務')

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
        row = 19 + day  # 1日目=20行目
        weekday = date(year, month, day).weekday()  # 0=月, 6=日

        # 対象セルをクリア
        for col in [col_start, col_end, col_break, col_note]:
            ws[f'{col}{row}'].value = None

        if weekday >= 5:
            pass  # 土日：空欄
        elif day in holidays:
            ws[f'{col_note}{row}'].value = '祝日'
        elif day in paid_leave:
            ws[f'{col_note}{row}'].value = '私用により、休暇'
        elif day in time_exceptions or weekday in weekday_times:
            t = time_exceptions.get(day) or weekday_times.get(weekday)
            if t:
                ws[f'{col_start}{row}'].value = t['start']
                ws[f'{col_start}{row}'].number_format = 'h:mm'
                ws[f'{col_end}{row}'].value = t['end']
                ws[f'{col_end}{row}'].number_format = 'h:mm'
                ws[f'{col_break}{row}'].value = t['break']
                ws[f'{col_break}{row}'].number_format = 'h:mm'
                ws[f'{col_note}{row}'].value = note_exceptions[day] if day in note_exceptions else note_workday

    # カーソルをA1に設定
    ws.sheet_view.selection[0].activeCell = 'A1'
    ws.sheet_view.selection[0].sqref = 'A1'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=excel_file.filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=5000)
