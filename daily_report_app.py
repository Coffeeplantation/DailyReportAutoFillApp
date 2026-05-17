import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
from datetime import date, time
import calendar as cal_module
import openpyxl
from openpyxl.styles import Font
import jpholiday

WEEKDAY_NAMES = ['月曜日', '火曜日', '水曜日', '木曜日', '金曜日']
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.json')
DEFAULT_TIMES = {'start': '09:00', 'end': '18:00', 'break': '01:00'}


def parse_time(s):
    if not s:
        return None
    try:
        h, m = s.strip().split(':')
        return time(int(h), int(m))
    except Exception:
        return None


class DailyReportApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('日報自動入力アプリ')
        self.root.geometry('560x620')
        self.root.minsize(520, 400)
        self.settings = self._load_settings()
        self._build_ui()
        self._apply_settings()
        self.root.mainloop()

    # ── Settings ──────────────────────────────────────────────────────────

    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_settings(self):
        data = {
            'weekday_times': {
                str(i): {k: self.time_vars[i][k].get() for k in ('start', 'end', 'break')}
                for i in range(5)
            }
        }
        try:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showwarning('警告', f'設定の保存に失敗しました:\n{e}')

    def _apply_settings(self):
        wt = self.settings.get('weekday_times', {})
        for i in range(5):
            t = wt.get(str(i), DEFAULT_TIMES)
            for k in ('start', 'end', 'break'):
                self.time_vars[i][k].set(t.get(k, DEFAULT_TIMES[k]))

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # スクロール可能なメインエリア
        canvas = tk.Canvas(self.root, borderwidth=0, highlightthickness=0)
        sb = ttk.Scrollbar(self.root, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        container = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=container, anchor='nw')
        container.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(-1 * (e.delta // 120), 'units'))

        pad = dict(padx=12, pady=5)

        ttk.Label(container, text='日報自動入力アプリ', font=('', 14, 'bold')).pack(pady=(14, 4))

        self._build_ym(container, pad)
        self._build_weekday(container, pad)
        self._build_paid_leave(container, pad)
        self._build_exceptions(container, pad)
        self._build_file(container, pad)
        ttk.Button(container, text='入力完了・保存', command=self._execute, width=22).pack(pady=14)

    def _build_ym(self, parent, pad):
        f = ttk.LabelFrame(parent, text='対象月')
        f.pack(fill='x', **pad)
        today = date.today()
        self.year_var  = tk.IntVar(value=today.year)
        self.month_var = tk.IntVar(value=today.month)
        inner = ttk.Frame(f)
        inner.pack(padx=10, pady=6)
        ttk.Spinbox(inner, from_=2000, to=2099, textvariable=self.year_var,  width=6).pack(side='left')
        ttk.Label(inner, text=' 年 ').pack(side='left')
        ttk.Spinbox(inner, from_=1,    to=12,   textvariable=self.month_var, width=4).pack(side='left')
        ttk.Label(inner, text(' 月')).pack(side='left')

    def _build_weekday(self, parent, pad):
        f = ttk.LabelFrame(parent, text='曜日別勤務時間  （形式: HH:MM）')
        f.pack(fill='x', **pad)
        for col, txt in enumerate(['', '開始', '終了', '休憩']):
            ttk.Label(f, text=txt, width=8 if col else 7, anchor='center').grid(
                row=0, column=col, padx=3, pady=2)
        self.time_vars = []
        for i, name in enumerate(WEEKDAY_NAMES):
            ttk.Label(f, text=name, width=6, anchor='w').grid(
                row=i + 1, column=0, padx=10, pady=4, sticky='w')
            row_vars = {}
            for col, key in enumerate(('start', 'end', 'break'), 1):
                var = tk.StringVar()
                ttk.Entry(f, textvariable=var, width=8, justify='center').grid(
                    row=i + 1, column=col, padx=3, pady=4)
                row_vars[key] = var
            self.time_vars.append(row_vars)

    def _build_paid_leave(self, parent, pad):
        outer = ttk.LabelFrame(parent, text='有給取得日')
        outer.pack(fill='x', **pad)
        self.pl_var = tk.BooleanVar()
        ttk.Checkbutton(outer, text='有給取得日がある', variable=self.pl_var,
                        command=self._toggle_pl).pack(anchor='w', padx=8, pady=4)
        self.pl_content = ttk.Frame(outer)
        self.day_check_vars = {}

    def _toggle_pl(self):
        for w in self.pl_content.winfo_children():
            w.destroy()
        self.day_check_vars.clear()
        if self.pl_var.get():
            self.pl_content.pack(padx=10, pady=(0, 8))
            self._build_cal_grid(self.year_var.get(), self.month_var.get())
        else:
            self.pl_content.pack_forget()

    def _build_cal_grid(self, year, month):
        _, last = cal_module.monthrange(year, month)
        hdr_names = ['月', '火', '水', '木', '金', '土', '日']
        hdr_fgs   = ['black'] * 5 + ['royalblue', 'red']
        for c, (h, fg) in enumerate(zip(hdr_names, hdr_fgs)):
            tk.Label(self.pl_content, text=h, width=4, fg=fg, font=('', 9, 'bold')).grid(
                row=0, column=c, padx=1)
        first_wd = date(year, month, 1).weekday()
        gr, gc = 1, first_wd
        for day in range(1, last + 1):
            var = tk.BooleanVar()
            self.day_check_vars[day] = var
            fg = 'royalblue' if gc == 5 else ('red' if gc == 6 else 'black')
            tk.Checkbutton(self.pl_content, text=str(day), variable=var, fg=fg, width=3).grid(
                row=gr, column=gc, padx=1, pady=1)
            gc += 1
            if gc > 6:
                gc, gr = 0, gr + 1

    def _build_exceptions(self, parent, pad):
        outer = ttk.LabelFrame(parent, text='例外日（残業・早退・休日出勤など）')
        outer.pack(fill='x', **pad)
        self.ex_var = tk.BooleanVar()
        ttk.Checkbutton(outer, text='例外日がある', variable=self.ex_var,
                        command=self._toggle_ex).pack(anchor='w', padx=8, pady=4)

        self.ex_content = ttk.Frame(outer)

        # ヘッダー（固定）
        hdr = ttk.Frame(self.ex_content)
        hdr.pack(fill='x', padx=2)
        for col, (txt, w) in enumerate([('日付', 5), ('開始', 7), ('終了', 7), ('休憩', 7), ('備考', 14)]):
            ttk.Label(hdr, text=txt, width=w, anchor='center').grid(row=0, column=col, padx=2)

        # 行を追加していくエリア
        self.ex_rows_frame = ttk.Frame(self.ex_content)
        self.ex_rows_frame.pack(fill='x')

        ttk.Button(self.ex_content, text='＋ 例外日を追加', command=self._add_ex_row).pack(pady=5)

        self.exception_rows = []

    def _toggle_ex(self):
        if self.ex_var.get():
            self.ex_content.pack(fill='x', padx=8, pady=(0, 6))
            if not self.exception_rows:
                self._add_ex_row()
        else:
            self.ex_content.pack_forget()
            for row in self.exception_rows:
                row['frame'].destroy()
            self.exception_rows.clear()

    def _add_ex_row(self):
        _, last = cal_module.monthrange(self.year_var.get(), self.month_var.get())
        rf = ttk.Frame(self.ex_rows_frame)
        rf.pack(fill='x', pady=2)

        day_v   = tk.StringVar(value='1')
        start_v = tk.StringVar()
        end_v   = tk.StringVar()
        break_v = tk.StringVar(value='01:00')
        note_v  = tk.StringVar(value='在宅勤務')

        ttk.Spinbox(rf, from_=1, to=last, textvariable=day_v,   width=5).grid(row=0, column=0, padx=2)
        ttk.Entry(rf, textvariable=start_v, width=7, justify='center').grid(row=0, column=1, padx=2)
        ttk.Entry(rf, textvariable=end_v,   width=7, justify='center').grid(row=0, column=2, padx=2)
        ttk.Entry(rf, textvariable=break_v, width=7, justify='center').grid(row=0, column=3, padx=2)
        ttk.Entry(rf, textvariable=note_v,  width=14).grid(row=0, column=4, padx=2)

        row_data = {
            'frame': rf,
            'day': day_v, 'start': start_v, 'end': end_v, 'break': break_v, 'note': note_v,
        }

        def remove(r=row_data):
            r['frame'].destroy()
            self.exception_rows.remove(r)

        ttk.Button(rf, text='✕', width=2, command=remove).grid(row=0, column=5, padx=4)
        self.exception_rows.append(row_data)

    def _build_file(self, parent, pad):
        f = ttk.LabelFrame(parent, text='Excelファイル')
        f.pack(fill='x', **pad)
        inner = ttk.Frame(f)
        inner.pack(fill='x', padx=8, pady=6)
        self.file_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.file_var, width=46).pack(side='left', expand=True, fill='x')
        ttk.Button(inner, text='参照...', command=self._browse).pack(side='left', padx=6)

    def _browse(self):
        p = filedialog.askopenfilename(
            title='Excelファイルを選択してください',
            filetypes=[('Excelファイル', '*.xlsx'), ('すべてのファイル', '*.*')],
        )
        if p:
            self.file_var.set(p)

    # ── Execute ───────────────────────────────────────────────────────────

    def _execute(self):
        path = self.file_var.get().strip()
        if not path:
            messagebox.showerror('エラー', 'Excelファイルを選択してください。')
            return
        if not os.path.exists(path):
            messagebox.showerror('エラー', 'ファイルが見つかりません。\nパスを確認してください。')
            return

        year, month = self.year_var.get(), self.month_var.get()

        # 曜日別勤務時間
        wt = {}
        for i in range(5):
            s = parse_time(self.time_vars[i]['start'].get())
            e = parse_time(self.time_vars[i]['end'].get())
            b = parse_time(self.time_vars[i]['break'].get()) or time(1, 0)
            if s and e:
                wt[i] = {'start': s, 'end': e, 'break': b}
        if not wt:
            messagebox.showerror('エラー', '勤務時間をHH:MM形式で入力してください。\n例: 09:00')
            return

        # 有給取得日
        paid = {d for d, v in self.day_check_vars.items() if v.get()} if self.pl_var.get() else set()

        # 例外日
        tex, nex = {}, {}
        for row in self.exception_rows:
            ds = row['day'].get()
            if not (ds and ds.isdigit()):
                continue
            d = int(ds)
            s = parse_time(row['start'].get())
            e = parse_time(row['end'].get())
            b = parse_time(row['break'].get()) or time(1, 0)
            if s and e:
                tex[d] = {'start': s, 'end': e, 'break': b}
            n = row['note'].get().strip()
            if n:
                nex[d] = n

        # 祝日（オフライン取得）
        _, last = cal_module.monthrange(year, month)
        hols = {d for d in range(1, last + 1) if jpholiday.is_holiday(date(year, month, d))}

        # Excel を開く
        try:
            wb = openpyxl.load_workbook(path)
        except Exception as e:
            messagebox.showerror('エラー', f'ファイルを開けませんでした。\n{e}')
            return
        ws = wb.active

        hr = self._find_hr(ws) or 19
        dr = hr + 1
        cs = self._find_col(ws, hr, '開始時間') or 'F'
        ce = self._find_col(ws, hr, '終了時間') or 'I'
        cb = self._find_col(ws, hr, '休憩時間') or 'L'
        cn = self._find_col(ws, hr, '備考')     or 'S'
        blk = Font(color='000000')

        for day in range(1, last + 1):
            r  = dr + day - 1
            wd = date(year, month, day).weekday()

            for c in (cs, ce, cb, cn):
                ws[f'{c}{r}'].value = None

            if day in tex:
                t = tex[day]
                for c, v in ((cs, t['start']), (ce, t['end']), (cb, t['break'])):
                    ws[f'{c}{r}'].value          = v
                    ws[f'{c}{r}'].number_format  = 'h:mm'
                    ws[f'{c}{r}'].font           = blk
                ws[f'{cn}{r}'].value = nex.get(day, '在宅勤務')
                ws[f'{cn}{r}'].font  = blk
            elif day in paid:
                ws[f'{cn}{r}'].value = '私用により、休暇'
                ws[f'{cn}{r}'].font  = blk
            elif wd >= 5:
                pass  # 土日：空欄
            elif day in hols:
                ws[f'{cn}{r}'].value = '祝日'
                ws[f'{cn}{r}'].font  = blk
            elif wd in wt:
                t = wt[wd]
                for c, v in ((cs, t['start']), (ce, t['end']), (cb, t['break'])):
                    ws[f'{c}{r}'].value          = v
                    ws[f'{c}{r}'].number_format  = 'h:mm'
                    ws[f'{c}{r}'].font           = blk
                ws[f'{cn}{r}'].value = '在宅勤務'
                ws[f'{cn}{r}'].font  = blk

        # カーソルをA1へ
        try:
            ws.sheet_view.selection[0].activeCell = 'A1'
            ws.sheet_view.selection[0].sqref      = 'A1'
        except Exception:
            pass

        # 上書き保存
        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror('エラー', f'保存に失敗しました。\nExcelが開かれていないか確認してください。\n{e}')
            return

        self._save_settings()
        messagebox.showinfo('完了', f'書き込みが完了しました。\n\nファイル: {os.path.basename(path)}')

    # ── Excel helpers ──────────────────────────────────────────────────────

    def _find_col(self, ws, hr, label):
        for cell in ws[hr]:
            if cell.value and (label in str(cell.value) or str(cell.value) in label):
                return cell.column_letter
        return None

    def _find_hr(self, ws):
        labels = ['開始時間', '終了時間', '休憩時間', '備考']
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            vals = [str(c.value).strip() for c in row if c.value]
            if sum(1 for lbl in labels if any(lbl in v or v in lbl for v in vals)) >= 2:
                return row[0].row
        return None


if __name__ == '__main__':
    DailyReportApp()
