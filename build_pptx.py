import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── カラー定義 ──────────────────────────────────────────────
NAVY   = RGBColor(0x1A, 0x56, 0xDB)
NAVY2  = RGBColor(0x1E, 0x40, 0xAF)
ACCENT = RGBColor(0xDB, 0xEA, 0xFE)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GREEN  = RGBColor(0x05, 0x96, 0x69)
GREEN2 = RGBColor(0x06, 0x5F, 0x46)
ORANGE = RGBColor(0xEA, 0x58, 0x0C)
ORANGE2= RGBColor(0x92, 0x40, 0x0E)
GRAY   = RGBColor(0xF3, 0xF4, 0xF6)
DARK   = RGBColor(0x1F, 0x2A, 0x44)
MUTED  = RGBColor(0x94, 0xA3, 0xB8)
PURPLE = RGBColor(0x6D, 0x28, 0xD9)
TEAL   = RGBColor(0x0F, 0x76, 0x6E)
PINK   = RGBColor(0xBE, 0x18, 0x5D)

W = Inches(13.33)
H = Inches(7.5)
HH = Inches(0.72)   # header height
FH = Inches(0.28)   # footer height
CY = HH + Inches(0.25)
CH = H - CY - FH - Inches(0.1)
ML = Inches(0.55)   # margin left
MR = Inches(0.55)   # margin right
CW = W - ML - MR    # content width


# ── 低レベルヘルパー ─────────────────────────────────────────
def _prs():
    p = Presentation()
    p.slide_width  = W
    p.slide_height = H
    return p

def _slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def _box(slide, x, y, w, h, rgb, line_rgb=None, line_w=0.75):
    s = slide.shapes.add_shape(1, int(x), int(y), int(w), int(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb
    if line_rgb:
        s.line.color.rgb = line_rgb
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    return s

def _set_text(shape, text, size=11, bold=False, color=DARK,
              align=PP_ALIGN.LEFT, ml=Inches(0.18), mt=Inches(0.1)):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left   = int(ml)
    tf.margin_top    = int(mt)
    tf.margin_right  = int(Inches(0.1))
    tf.margin_bottom = int(Inches(0.05))
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.name  = "Meiryo"
    r.font.color.rgb = color

def _set_multiline(shape, lines, size=11, bold=False, color=DARK,
                   align=PP_ALIGN.LEFT, ml=Inches(0.18), mt=Inches(0.1)):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left   = int(ml)
    tf.margin_top    = int(mt)
    tf.margin_right  = int(Inches(0.1))
    tf.margin_bottom = int(Inches(0.05))
    for i, line in enumerate(lines):
        p = tf.paragraphs[i] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size  = Pt(size)
        r.font.bold  = bold
        r.font.name  = "Meiryo"
        r.font.color.rgb = color

def _label_box(slide, x, y, w, h, text, bg, tc=WHITE, size=11, bold=True,
               align=PP_ALIGN.CENTER):
    s = _box(slide, x, y, w, h, bg)
    _set_text(s, text, size=size, bold=bold, color=tc, align=align,
              ml=Inches(0.1), mt=Inches(0.08))
    return s

def _tb(slide, x, y, w, h, text, size=11, bold=False, color=DARK,
        align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    _set_text(box, text, size=size, bold=bold, color=color, align=align)
    return box

def _header(slide, title, n, total=10):
    bg = _box(slide, 0, 0, W, HH, NAVY)
    _box(slide, 0, 0, Inches(0.07), HH, RGBColor(0x60, 0xA5, 0xFA))
    _set_text(bg, title, size=18, bold=True, color=WHITE,
              ml=Inches(0.22), mt=Inches(0.17))
    num = _box(slide, W - Inches(1.4), Inches(0.2), Inches(1.2), Inches(0.35), NAVY)
    _set_text(num, f"{n} / {total}", size=10, color=RGBColor(0xBF, 0xD7, 0xFF),
              align=PP_ALIGN.RIGHT, ml=Inches(0.05), mt=Inches(0.05))

def _footer_bar(slide):
    _box(slide, 0, H - FH, W, FH, NAVY)

def _chrome(slide, title, n, total=10):
    _header(slide, title, n, total)
    _footer_bar(slide)


# ── 操作マニュアル ────────────────────────────────────────────
def _build_manual_pptx() -> io.BytesIO:
    prs = _prs()
    TOTAL = 10

    # ── Slide 1: タイトル ──
    sl = _slide(prs)
    _box(sl, 0, 0, W, H, NAVY)
    _box(sl, 0, 0, Inches(0.12), H, RGBColor(0x60, 0xA5, 0xFA))
    _footer_bar(sl)
    title_box = _box(sl, Inches(0.5), Inches(2.0), W - Inches(1.0), Inches(1.5), NAVY)
    _set_text(title_box, "日報自動入力アプリ　操作マニュアル",
              size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.25))
    sub = _box(sl, Inches(0.5), Inches(3.7), W - Inches(1.0), Inches(0.7), NAVY)
    _set_text(sub, "このマニュアルを読めば、月初の作業が 5分で完了します",
              size=15, color=RGBColor(0xBF, 0xD7, 0xFF), align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.1))
    contents = [
        "◆  アプリの使い方（STEP 1〜7）の詳しい手順",
        "◆  自動入力ルールの早見表",
        "◆  よくある疑問と回答",
    ]
    for i, line in enumerate(contents):
        _tb(sl, Inches(4.0), Inches(4.7) + Inches(0.4) * i, Inches(8.0), Inches(0.38),
            line, size=11, color=RGBColor(0xBF, 0xD7, 0xFF))
    num = _box(sl, W - Inches(1.4), H - Inches(0.55), Inches(1.2), Inches(0.3), NAVY)
    _set_text(num, "1 / 10", size=10, color=RGBColor(0xBF, 0xD7, 0xFF),
              align=PP_ALIGN.RIGHT, ml=Inches(0.05), mt=Inches(0.03))

    # ── Slide 2: このアプリについて ──
    sl = _slide(prs)
    _chrome(sl, "このアプリについて", 2, TOTAL)
    features = [
        ("月初 5分以内",      "で操作を完了できます"),
        ("1ヶ月分すべて",     "のExcel行を自動で埋めます"),
        ("土日・祝日・有給",  "を自動で判定・記入します"),
        ("設定は自動保存",    "翌月は変更箇所のみ修正OK"),
    ]
    fw = (CW - Inches(0.3)) / 4
    for i, (kw, rest) in enumerate(features):
        bx = ML + (fw + Inches(0.1)) * i
        card = _box(sl, bx, CY, fw, Inches(1.5), ACCENT)
        _set_text(card, kw, size=13, bold=True, color=NAVY,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.15))
        desc = _box(sl, bx, CY + Inches(1.5), fw, Inches(0.8), GRAY)
        _set_text(desc, rest, size=10, color=DARK, align=PP_ALIGN.CENTER,
                  ml=Inches(0.1), mt=Inches(0.1))

    _tb(sl, ML, CY + Inches(2.5), CW, Inches(0.35),
        "必要なもの（事前確認）", size=13, bold=True, color=NAVY)
    checks = [
        "✔  ブラウザ（Chrome / Edge / Safari）　インストール不要",
        "✔  アプリのURL・ログイン情報　　　　　管理者から入手",
        "✔  会社配布の作業報告書（.xlsx）　　　毎月配布されるファイル",
    ]
    for i, c in enumerate(checks):
        _tb(sl, ML + Inches(0.2), CY + Inches(3.0) + Inches(0.42) * i,
            CW, Inches(0.38), c, size=11, color=DARK)
    notes = [
        "！  インターネット接続が必要です（祝日取得）",
        "！  ダウンロード後、内容を確認してから保存してください",
        "！  元のファイルは上書きされません",
    ]
    note_x = ML + Inches(5.5)
    for i, n in enumerate(notes):
        _tb(sl, note_x, CY + Inches(3.0) + Inches(0.42) * i,
            CW - Inches(5.5), Inches(0.38), n, size=10,
            color=RGBColor(0xB4, 0x5C, 0x09))

    # ── Slide 3: 操作の全体フロー ──
    sl = _slide(prs)
    _chrome(sl, "操作の全体フロー", 3, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.35),
        "7 つのステップで月初の作業を完了させましょう", size=12, color=MUTED)

    steps = [
        ("STEP 1", "アクセス\n・ログイン",    NAVY2, False),
        ("STEP 2", "対象月\nを確認",          GREEN2, False),
        ("STEP 3", "勤務時間\nを設定",         ORANGE2, False),
        ("STEP 4", "例外日\n設定（任意）",     PURPLE, True),
        ("STEP 5", "有給日\n選択（任意）",     PINK, True),
        ("STEP 6", "Excelファイル\n選択",      TEAL, False),
        ("STEP 7", "完了・\nダウンロード",     NAVY2, False),
    ]
    sw = (CW - Inches(0.6)) / 7
    sy = CY + Inches(0.5)
    for i, (label, title, color, optional) in enumerate(steps):
        bx = ML + (sw + Inches(0.1)) * i
        step_box = _box(sl, bx, sy, sw, Inches(0.38), color)
        _set_text(step_box, label, size=9, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.06))
        title_box = _box(sl, bx, sy + Inches(0.38), sw, Inches(0.75),
                         RGBColor(0xFF, 0xED, 0xB5) if optional else ACCENT)
        lines = title.split('\n')
        _set_multiline(title_box, lines, size=9,
                       color=ORANGE2 if optional else NAVY2,
                       align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.08))
        if i < 6:
            _tb(sl, bx + sw, sy + Inches(0.45), Inches(0.1), Inches(0.3),
                "▶", size=10, color=MUTED, align=PP_ALIGN.CENTER)

    _tb(sl, ML, sy + Inches(1.3), CW, Inches(0.3),
        "※ STEP 4・5（オレンジ色）は任意。残業・休日出勤・有給がある月のみ操作してください。",
        size=9, color=MUTED)

    descs = [
        ("必須  STEP 1〜3・6〜7",
         "毎月行う操作です。STEP 3の設定はブラウザに保存されるため、2回目以降は変更がある場合のみ修正してください。"),
        ("任意  STEP 4〜5",
         "残業・早退・休日出勤がある場合はSTEP 4を。有給休暇がある場合はSTEP 5を操作してください。"),
    ]
    desc_y = sy + Inches(1.8)
    desc_w = (CW - Inches(0.2)) / 2
    for i, (head, body) in enumerate(descs):
        bx = ML + (desc_w + Inches(0.2)) * i
        head_bg = NAVY2 if i == 0 else RGBColor(0xFF, 0xED, 0xB5)
        head_tc = WHITE if i == 0 else ORANGE2
        h = _box(sl, bx, desc_y, desc_w, Inches(0.38), head_bg)
        _set_text(h, head, size=11, bold=True, color=head_tc,
                  ml=Inches(0.15), mt=Inches(0.08))
        b = _box(sl, bx, desc_y + Inches(0.38), desc_w, Inches(0.85), GRAY)
        _set_text(b, body, size=10, color=DARK, ml=Inches(0.12), mt=Inches(0.1))

    # ── Slide 4: STEP 1 ──
    sl = _slide(prs)
    _chrome(sl, "STEP 1　アクセス・ログイン", 4, TOTAL)
    step_h = _box(sl, ML, CY, CW, Inches(0.45), NAVY2)
    _set_text(step_h, "1　ブラウザでアプリのURLを開き、ログインします",
              size=14, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.1))
    steps_detail = [
        "①  ブラウザ（Chrome / Edge / Safari）を開きます。",
        "②  アドレスバーにアプリの URL を入力してアクセスします。",
        "③  ユーザー名・パスワードの入力画面が表示されたら、指定の情報を入力して「OK」を押します。",
        "④  アプリのメイン画面が表示されればログイン完了です。",
    ]
    for i, s in enumerate(steps_detail):
        row = _box(sl, ML, CY + Inches(0.55) + Inches(0.65) * i, CW, Inches(0.6),
                   ACCENT if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(row, s, size=12, color=DARK, ml=Inches(0.2), mt=Inches(0.12))

    point = _box(sl, ML, CY + Inches(3.25), CW, Inches(0.85),
                 RGBColor(0xEF, 0xF6, 0xFF), line_rgb=NAVY2)
    _set_multiline(point, [
        "ポイント",
        "・ URL とログイン情報は管理者から入手してください。",
        "・ ログインできない場合は Caps Lock やスペース入力ミスを確認してください。",
    ], size=10, color=NAVY2, ml=Inches(0.2), mt=Inches(0.1))

    # ── Slide 5: STEP 2〜3 ──
    sl = _slide(prs)
    _chrome(sl, "STEP 2〜3　対象月の確認・勤務時間の設定", 5, TOTAL)
    half = (CW - Inches(0.3)) / 2

    left_h = _box(sl, ML, CY, half, Inches(0.45), GREEN2)
    _set_text(left_h, "2　対象月を確認する", size=13, bold=True, color=WHITE,
              ml=Inches(0.15), mt=Inches(0.1))
    left_items = [
        "画面上部に「年・月」が表示されます。",
        "自動で「当月」が設定されています。",
        "別の月を対象にしたい場合は数字を直接変更してください。",
        "",
        "通常は変更不要です。当月が自動で設定されています。",
    ]
    for i, s in enumerate(left_items):
        if s:
            row = _box(sl, ML, CY + Inches(0.55) + Inches(0.58) * i,
                       half, Inches(0.55),
                       RGBColor(0xD1, 0xFA, 0xE5) if i < 3 else RGBColor(0xF0, 0xFD, 0xF4),
                       line_rgb=RGBColor(0x6E, 0xE7, 0xB7))
            _set_text(row, s, size=11, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    right_x = ML + half + Inches(0.3)
    right_h = _box(sl, right_x, CY, half, Inches(0.45), ORANGE2)
    _set_text(right_h, "3　勤務時間を設定する", size=13, bold=True, color=WHITE,
              ml=Inches(0.15), mt=Inches(0.1))
    right_items = [
        "月曜〜金曜それぞれの「開始時間」「終了時間」を入力します。",
        "入力後、下部の月間スケジュールに1日〜末日の予定と合計稼働時間が表示されます。",
        "設定変更時はヘッダーに「✓ 自動保存」と表示され、ブラウザに保存されます。",
        "",
        "設定を保存しておくと翌月以降は変更箇所のみ修正するだけでOKです。",
    ]
    for i, s in enumerate(right_items):
        if s:
            row = _box(sl, right_x, CY + Inches(0.55) + Inches(0.58) * i,
                       half, Inches(0.55),
                       RGBColor(0xFE, 0xF3, 0xC7) if i < 3 else RGBColor(0xFF, 0xFB, 0xEB),
                       line_rgb=RGBColor(0xFB, 0xBF, 0x24))
            _set_text(row, s, size=11, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    # ── Slide 6: STEP 4 ──
    sl = _slide(prs)
    _chrome(sl, "STEP 4　例外日を設定する（任意）", 6, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "残業・早退・休日出勤など、通常と異なる日を個別設定します", size=11, color=MUTED)
    step_h = _box(sl, ML, CY + Inches(0.35), CW, Inches(0.42), PURPLE)
    _set_text(step_h, "4　残業・早退・休日出勤がある日を個別に設定します",
              size=13, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    ex_steps = [
        "①  「＋ 例外日を追加」ボタンをクリックします。",
        "②  カレンダーが表示されるので、対象の日付をクリックして選択します（土日・祝日も選択可）。",
        "③  選択した日の開始時間・終了時間を入力します。備考欄の初期値は「在宅勤務」です（変更可能）。",
        "④  複数日ある場合は必要な分だけ繰り返してください。",
    ]
    for i, s in enumerate(ex_steps):
        row = _box(sl, ML, CY + Inches(0.87) + Inches(0.62) * i, CW, Inches(0.58),
                   RGBColor(0xF5, 0xF3, 0xFF) if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0xC4, 0xB5, 0xFD))
        _set_text(row, s, size=12, color=DARK, ml=Inches(0.2), mt=Inches(0.1))

    warn = _box(sl, ML, CY + Inches(3.4), CW, Inches(0.85),
                RGBColor(0xED, 0xE9, 0xFE), line_rgb=PURPLE)
    _set_multiline(warn, [
        "重要：例外日設定は最優先で適用されます",
        "・ 曜日別設定・有給・祝日・土日の区別に関わらず最優先で反映されます。",
        "・ 休日出勤の記録にも使えます。　・ 例外日がない月はスキップしてください。",
    ], size=10, color=PURPLE, ml=Inches(0.2), mt=Inches(0.08))

    # ── Slide 7: STEP 5 ──
    sl = _slide(prs)
    _chrome(sl, "STEP 5　有給取得日を選択する（任意）", 7, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "有給休暇を取得する日をカレンダーで選択します", size=11, color=MUTED)
    step_h = _box(sl, ML, CY + Inches(0.35), CW, Inches(0.42), PINK)
    _set_text(step_h, "5　有給休暇を取得する日をカレンダーで選択します",
              size=13, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    paid_steps = [
        "①  「有給を取得する日がある」チェックボックスをオンにします。",
        "②  当月のカレンダーが表示されます。",
        "③  有給を取得する日付をクリックして選択します（選択した日はハイライト表示）。",
        "④  複数日ある場合は続けてクリックしてください。もう一度クリックすると解除できます。",
    ]
    for i, s in enumerate(paid_steps):
        row = _box(sl, ML, CY + Inches(0.87) + Inches(0.62) * i, CW, Inches(0.58),
                   RGBColor(0xFF, 0xF0, 0xF7) if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0xFB, 0xCF, 0xE8))
        _set_text(row, s, size=12, color=DARK, ml=Inches(0.2), mt=Inches(0.1))

    point = _box(sl, ML, CY + Inches(3.4), CW, Inches(0.85),
                 RGBColor(0xFF, 0xF0, 0xF7), line_rgb=PINK)
    _set_multiline(point, [
        "ポイント",
        "・ 土曜・日曜・祝日も選択できます（振替休日の有給取得など）。",
        "・ 選択した日は備考欄に「私用により、休暇」と自動入力されます。　・ 有給がない月はスキップしてください。",
    ], size=10, color=PINK, ml=Inches(0.2), mt=Inches(0.08))

    # ── Slide 8: STEP 6〜7 ──
    sl = _slide(prs)
    _chrome(sl, "STEP 6〜7　Excelファイルの選択・ダウンロード", 8, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "最後のステップです。ファイルを選んでボタンを押すだけで完了します", size=11, color=MUTED)
    half = (CW - Inches(0.3)) / 2

    left_h = _box(sl, ML, CY + Inches(0.35), half, Inches(0.42), TEAL)
    _set_text(left_h, "6　Excelファイルを選択する", size=13, bold=True, color=WHITE,
              ml=Inches(0.15), mt=Inches(0.09))
    left_steps = [
        "「Excelファイルを選択」エリアをクリックします。",
        "ファイル選択ダイアログが開くので、配布された作業報告書（.xlsx）を選んで「開く」を押します。",
        "ファイル名が表示されれば選択完了です。",
        "間違えた場合は「✕」ボタンでキャンセルして選び直せます。",
    ]
    for i, s in enumerate(left_steps):
        row = _box(sl, ML, CY + Inches(0.87) + Inches(0.75) * i,
                   half, Inches(0.72),
                   RGBColor(0xCC, 0xFB, 0xF1) if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0x99, 0xF6, 0xE4))
        _set_text(row, s, size=11, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    right_x = ML + half + Inches(0.3)
    right_h = _box(sl, right_x, CY + Inches(0.35), half, Inches(0.42), NAVY2)
    _set_text(right_h, "7　完了・ダウンロード", size=13, bold=True, color=WHITE,
              ml=Inches(0.15), mt=Inches(0.09))
    right_steps = [
        "「入力完了・ダウンロード」ボタンをクリックします。",
        "処理が完了すると、自動入力済みのExcelファイルがダウンロードされます。",
        "ダウンロードしたファイルを開いて内容を確認してください。",
        "問題なければ所定の場所に保存して作業完了です。",
    ]
    for i, s in enumerate(right_steps):
        row = _box(sl, right_x, CY + Inches(0.87) + Inches(0.75) * i,
                   half, Inches(0.72),
                   ACCENT if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(row, s, size=11, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    note = _box(sl, ML, H - Inches(0.9), CW, Inches(0.35), RGBColor(0xEF, 0xF6, 0xFF))
    _set_text(note, "元のExcelファイルは上書きされません。新しいファイルがダウンロードされます。安心して操作してください。",
              size=10, color=NAVY2, ml=Inches(0.2), mt=Inches(0.07))

    # ── Slide 9: 自動入力ルール早見表 ──
    sl = _slide(prs)
    _chrome(sl, "自動入力ルール早見表", 9, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "日の種類に応じて以下のルールでExcelへ自動入力されます", size=11, color=MUTED)

    cols = [Inches(2.8), Inches(4.5), Inches(4.5)]
    headers = ["日の種類", "開始・終了・休憩", "備考欄の記入内容"]
    hx = ML
    hy = CY + Inches(0.4)
    for w, h in zip(cols, headers):
        cell = _box(sl, hx, hy, w, Inches(0.45), NAVY)
        _set_text(cell, h, size=12, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.1))
        hx += w

    rows = [
        ("出勤日（平日）",             "曜日別設定の開始・終了\n休憩：1時間00分", "在宅勤務",          ACCENT,                          RGBColor(0xDB, 0xEA, 0xFE)),
        ("例外日（残業・早退・休日出勤）", "例外日設定の開始・終了\n休憩：1時間00分", "在宅勤務（例外設定が最優先）", RGBColor(0xFE, 0xF3, 0xC7), RGBColor(0xFE, 0xF3, 0xC7)),
        ("有給取得日",                 "空欄",                          "私用により、休暇",    RGBColor(0xFF, 0xF0, 0xF7),      RGBColor(0xFF, 0xF0, 0xF7)),
        ("祝日",                      "空欄",                          "祝日",               RGBColor(0xFF, 0xF5, 0xF5),      RGBColor(0xFF, 0xF5, 0xF5)),
        ("土日",                      "空欄",                          "空欄",               GRAY,                            GRAY),
    ]
    row_h = Inches(0.75)
    for ri, (kind, times, note, bg1, bg2) in enumerate(rows):
        ry = hy + Inches(0.45) + row_h * ri
        rx = ML
        for ci, (val, w, bg) in enumerate([(kind, cols[0], bg1), (times, cols[1], bg2), (note, cols[2], bg2)]):
            cell = _box(sl, rx, ry, w, row_h, bg, line_rgb=RGBColor(0xE5, 0xE7, 0xEB))
            lines = val.split('\n')
            _set_multiline(cell, lines, size=11, color=DARK,
                          align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.12))
            rx += w

    notes_y = hy + Inches(0.45) + row_h * 5 + Inches(0.1)
    for note in [
        "※ 祝日情報は外部API（jpholiday）から自動取得。インターネット未接続時は空欄になる場合があります。",
        "※ 例外日設定は有給・祝日・土日に関わらず最優先で適用されます（休日出勤の記録にも使用可能）。",
    ]:
        _tb(sl, ML, notes_y, CW, Inches(0.3), note, size=9, color=MUTED)
        notes_y += Inches(0.32)

    # ── Slide 10: Q&A ──
    sl = _slide(prs)
    _chrome(sl, "よくある疑問（Q&A）", 10, TOTAL)
    qas = [
        ("Q1　設定は毎月入力し直す必要がありますか？",
         "不要です。曜日別の勤務時間・設定はブラウザに自動保存されます。次回起動時に前回の設定が自動で読み込まれます。"),
        ("Q2　休日出勤はどうやって入力しますか？",
         "「＋ 例外日を追加」から対象日（土日・祝日も可）を選択し、出勤時間を入力します。例外日設定が最優先で反映されます。"),
        ("Q3　有給を土日・祝日に取得する場合は？",
         "STEP 5のカレンダーで土日・祝日も選択できます。選択した日の備考欄に「私用により、休暇」が記入されます。"),
        ("Q4　ファイルを間違えて選択した場合は？",
         "ファイル名の右側の「✕」ボタンで選択をキャンセルできます。再度クリックして正しいファイルを選び直してください。"),
        ("Q5　Excelの列の位置が変わっても使えますか？",
         "はい。アプリがヘッダー文字列を自動検索して対象列を特定するため、列の位置が変わっても正常に動作します。"),
        ("Q6　ダウンロードしたファイルはどこに保存される？",
         "ブラウザの設定に従い、通常はダウンロードフォルダに保存されます。確認後、所定の場所へ移動してください。"),
    ]
    qa_h = (CH - Inches(0.1)) / len(qas)
    for i, (q, a) in enumerate(qas):
        qy = CY + qa_h * i
        qrow = _box(sl, ML, qy, CW, qa_h * 0.46, RGBColor(0xED, 0xE9, 0xFE),
                    line_rgb=RGBColor(0xC4, 0xB5, 0xFD))
        _set_text(qrow, q, size=11, bold=True, color=PURPLE,
                  ml=Inches(0.2), mt=Inches(0.06))
        arow = _box(sl, ML, qy + qa_h * 0.46, CW, qa_h * 0.5, WHITE,
                    line_rgb=RGBColor(0xE5, 0xE7, 0xEB))
        _set_text(arow, a, size=10, color=DARK, ml=Inches(0.25), mt=Inches(0.07))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ── 導入説明資料 ─────────────────────────────────────────────
def _build_report_pptx() -> io.BytesIO:
    from datetime import date as _date
    TODAY = f"{_date.today().year}年{_date.today().month}月"
    prs = _prs()
    TOTAL = 10

    # ── Slide 1: タイトル ──
    sl = _slide(prs)
    _box(sl, 0, 0, W, H, NAVY)
    _box(sl, 0, 0, Inches(0.12), H, RGBColor(0x60, 0xA5, 0xFA))
    _footer_bar(sl)
    t = _box(sl, Inches(0.5), Inches(1.8), W - Inches(1.0), Inches(1.6), NAVY)
    _set_text(t, "日報自動入力アプリ", size=34, bold=True, color=WHITE,
              align=PP_ALIGN.CENTER, ml=Inches(0.2), mt=Inches(0.28))
    sub = _box(sl, Inches(0.5), Inches(3.5), W - Inches(1.0), Inches(0.65), NAVY)
    _set_text(sub, "導入説明資料", size=20, bold=True,
              color=RGBColor(0xBF, 0xD7, 0xFF), align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.1))
    cap = _box(sl, Inches(0.5), Inches(4.3), W - Inches(1.0), Inches(0.5), NAVY)
    _set_text(cap, "― 月次作業報告書の自動化による業務効率化 ―", size=12,
              color=RGBColor(0xBF, 0xD7, 0xFF), align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.1))
    for label, val in [("作成日", TODAY), ("作成者", ""), ("提出先", "")]:
        pass
    info_y = Inches(5.1)
    for label, val in [("作成日", TODAY), ("作成者", ""), ("提出先", "")]:
        lb = _box(sl, Inches(3.5), info_y, Inches(1.5), Inches(0.4), NAVY2)
        _set_text(lb, label, size=11, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.08))
        vb = _box(sl, Inches(5.0), info_y, Inches(4.5), Inches(0.4),
                  RGBColor(0x1E, 0x3A, 0x8A))
        _set_text(vb, val, size=11, color=WHITE, ml=Inches(0.15), mt=Inches(0.08))
        info_y += Inches(0.5)
    num = _box(sl, W - Inches(1.4), H - Inches(0.55), Inches(1.2), Inches(0.3), NAVY)
    _set_text(num, "1 / 10", size=10, color=RGBColor(0xBF, 0xD7, 0xFF),
              align=PP_ALIGN.RIGHT, ml=Inches(0.05), mt=Inches(0.03))

    # ── Slide 2: 目次 ──
    sl = _slide(prs)
    _chrome(sl, "目次　CONTENTS", 2, TOTAL)
    items = [
        ("1", "現状の課題"),
        ("2", "アプリ概要・解決策"),
        ("3", "期待される効果"),
        ("4", "動作環境・対象ユーザー"),
        ("5", "操作手順（全体フロー）"),
        ("6", "自動入力ルール"),
        ("7", "Q&A"),
        ("8", "まとめ・導入のお願い"),
    ]
    iw = (CW - Inches(0.2)) / 2
    for i, (num_s, title) in enumerate(items):
        col = i % 2
        row = i // 2
        ix = ML + col * (iw + Inches(0.2))
        iy = CY + Inches(0.2) + Inches(1.1) * row
        num_b = _box(sl, ix, iy, Inches(0.55), Inches(0.75), NAVY)
        _set_text(num_b, num_s, size=22, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.1))
        title_b = _box(sl, ix + Inches(0.55), iy, iw - Inches(0.55), Inches(0.75), ACCENT)
        _set_text(title_b, title, size=13, bold=True, color=NAVY,
                  ml=Inches(0.2), mt=Inches(0.18))

    # ── Slide 3: 現状の課題 ──
    sl = _slide(prs)
    _chrome(sl, "1.  現状の課題", 3, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "毎月発生している手作業の実態と、本アプリによる解決", size=11, color=MUTED)

    half = (CW - Inches(0.3)) / 2
    before_h = _box(sl, ML, CY + Inches(0.35), half, Inches(0.45),
                    RGBColor(0xEF, 0x44, 0x44))
    _set_text(before_h, "✗  現状の課題", size=13, bold=True, color=WHITE,
              ml=Inches(0.2), mt=Inches(0.1))
    before_items = [
        "毎月末〜月初に1ヶ月分の勤務時間をExcelへ手作業で入力",
        "開始・終了・休憩・備考を全日付分繰り返す単調な入力作業",
        "手入力ミス・記入漏れのリスクが常に存在する",
        "月あたり約15〜30分の入力工数が継続的に発生",
    ]
    for i, s in enumerate(before_items):
        row = _box(sl, ML, CY + Inches(0.9) + Inches(0.72) * i, half, Inches(0.68),
                   RGBColor(0xFF, 0xF5, 0xF5), line_rgb=RGBColor(0xFE, 0xCA, 0xCA))
        _set_text(row, s, size=11, color=DARK, ml=Inches(0.2), mt=Inches(0.12))

    right_x = ML + half + Inches(0.3)
    after_h = _box(sl, right_x, CY + Inches(0.35), half, Inches(0.45), GREEN)
    _set_text(after_h, "✔  導入後の姿", size=13, bold=True, color=WHITE,
              ml=Inches(0.2), mt=Inches(0.1))
    after_items = [
        "月初に数分の確認だけで1ヶ月分の作業がすべて完結",
        "ボタン1つで全日程を自動入力・Excelファイル生成",
        "ルールベース処理でミス・漏れを根本から排除",
        "翌月以降は設定の確認・変更のみで対応可能",
    ]
    for i, s in enumerate(after_items):
        row = _box(sl, right_x, CY + Inches(0.9) + Inches(0.72) * i, half, Inches(0.68),
                   RGBColor(0xF0, 0xFD, 0xF4), line_rgb=RGBColor(0x6E, 0xE7, 0xB7))
        _set_text(row, s, size=11, color=DARK, ml=Inches(0.2), mt=Inches(0.12))

    arrow = _tb(sl, ML + half + Inches(0.03), CY + Inches(1.6), Inches(0.27), Inches(0.5),
                "▶▶", size=14, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)

    # ── Slide 4: アプリ概要 ──
    sl = _slide(prs)
    _chrome(sl, "2.  アプリ概要・解決策", 4, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.35),
        "ブラウザだけで完結する Excel 自動入力 Web アプリ", size=13, bold=True, color=NAVY)
    _tb(sl, ML, CY + Inches(0.38), CW, Inches(0.38),
        "会社から毎月配布されるExcel作業報告書への記入を自動化するWebアプリケーションです。インストール不要・ブラウザのみで動作。",
        size=10, color=MUTED)
    features = [
        ("📅", "曜日別 勤務時間設定",    "月〜金それぞれの開始・終了時間を設定。設定変更時はヘッダーに「自動保存」を表示。リセットボタンで初期値に戻せます。"),
        ("🗓", "例外日・有給 個別設定",   "残業・早退・休日出勤・有給取得日を個別設定。例外日は祝日・土日を上書き（休日出勤対応）。"),
        ("📊", "月間スケジュール プレビュー", "1日〜末日の予定を自動計算して一覧表示。就業時間合計もリアルタイムで確認できます。"),
        ("💾", "Excelへ 自動書込み",     "配布ファイルを選ぶだけ。自動入力済みファイルをダウンロード。"),
    ]
    fw = (CW - Inches(0.3)) / 4
    fy = CY + Inches(0.9)
    for i, (icon, title, desc) in enumerate(features):
        fx = ML + (fw + Inches(0.1)) * i
        icon_b = _box(sl, fx, fy, fw, Inches(0.5), NAVY)
        _set_text(icon_b, f"{icon} {title}", size=10, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.1))
        desc_b = _box(sl, fx, fy + Inches(0.5), fw, Inches(1.8), ACCENT,
                      line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(desc_b, desc, size=9, color=DARK, ml=Inches(0.1), mt=Inches(0.1))

    # ── Slide 5: 期待される効果 ──
    sl = _slide(prs)
    _chrome(sl, "3.  期待される効果", 5, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "導入によって得られる定量・定性効果", size=11, color=MUTED)
    effects = [
        ("⏱",  "工数削減",   "約15〜30分/月", "月次の手入力をほぼゼロに。\n年間で3〜6時間分の業務時間を創出。",   NAVY),
        ("✔",  "入力ミス防止", "精度 100%",   "手動入力によるミス・漏れを排除。\n正確な報告書を自動生成。",         GREEN),
        ("🖱", "操作の簡便性", "月初 5分以内", "ブラウザのみ、インストール不要。\nプログラミング知識不要で完結。",    ORANGE),
        ("💾", "設定の引き継ぎ", "翌月は確認のみ", "設定がブラウザに自動保存。\n翌月以降は変更箇所のみ修正。",    PURPLE),
    ]
    ew = (CW - Inches(0.3)) / 4
    ey = CY + Inches(0.4)
    for i, (icon, title, metric, desc, color) in enumerate(effects):
        ex = ML + (ew + Inches(0.1)) * i
        head = _box(sl, ex, ey, ew, Inches(0.45), color)
        _set_text(head, f"{icon}  {title}", size=12, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.1))
        metric_b = _box(sl, ex, ey + Inches(0.45), ew, Inches(0.65), ACCENT)
        _set_text(metric_b, metric, size=16, bold=True, color=color,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.1))
        desc_b = _box(sl, ex, ey + Inches(1.1), ew, Inches(1.3), GRAY,
                      line_rgb=RGBColor(0xE5, 0xE7, 0xEB))
        lines = desc.split('\n')
        _set_multiline(desc_b, lines, size=10, color=DARK,
                       align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.15))

    summary = _box(sl, ML, ey + Inches(2.55), CW, Inches(0.55), NAVY)
    _set_text(summary,
              "月初5分以内で1ヶ月分が自動完成。手入力ゼロ・ミスゼロを実現します。",
              size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.12))

    # ── Slide 6: 動作環境・対象ユーザー ──
    sl = _slide(prs)
    _chrome(sl, "4.  動作環境・対象ユーザー", 6, TOTAL)
    env_rows = [
        ("対象ユーザー", "Excel形式（.xlsx）の作業報告書を毎月提出している社員"),
        ("対象ファイル", "会社から毎月配布される Excel 作業報告書（.xlsx）"),
        ("動作環境",    "ブラウザ（Chrome / Edge / Safari など）　※ インストール不要"),
        ("アクセス方法", "社内ネットワーク上のURL、またはローカル環境からアクセス"),
        ("認証方式",    "ユーザー名・パスワードによる Basic 認証（不正アクセス防止）"),
        ("ファイル保存", "自動入力済みファイルはブラウザ経由でダウンロード。サーバーへの保存なし"),
    ]
    lw = Inches(2.8)
    rw = CW - lw
    rh = Inches(0.65)
    for i, (label, val) in enumerate(env_rows):
        ry = CY + Inches(0.1) + rh * i
        lb = _box(sl, ML, ry, lw, rh, NAVY2, line_rgb=RGBColor(0x3B, 0x82, 0xF6))
        _set_text(lb, label, size=12, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.15))
        vb = _box(sl, ML + lw, ry, rw, rh,
                  ACCENT if i % 2 == 0 else WHITE,
                  line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(vb, val, size=11, color=DARK, ml=Inches(0.2), mt=Inches(0.15))

    foot = _box(sl, ML, CY + rh * 6 + Inches(0.15), CW, Inches(0.45), NAVY)
    _set_text(foot,
              "インストール不要・ブラウザだけで完結。既存の PC 環境をそのままご利用いただけます。",
              size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
              ml=Inches(0.2), mt=Inches(0.1))

    # ── Slide 7: 操作手順 ──
    sl = _slide(prs)
    _chrome(sl, "5.  操作手順（全体フロー）", 7, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "月初の操作は 7 ステップで完了します（目安：5分以内）", size=11, color=MUTED)

    steps7 = [
        ("STEP 1", "アクセス\n・ログイン",    NAVY2,  False),
        ("STEP 2", "対象月\nを確認",          GREEN2, False),
        ("STEP 3", "勤務時間\n設定",           ORANGE2,False),
        ("STEP 4", "例外日\n設定（任意）",     PURPLE, True),
        ("STEP 5", "有給日\n選択（任意）",     PINK,   True),
        ("STEP 6", "Excelファイル\n選択",      TEAL,   False),
        ("STEP 7", "完了・\nダウンロード",     NAVY2,  False),
    ]
    sw = (CW - Inches(0.6)) / 7
    sy = CY + Inches(0.4)
    for i, (label, title, color, optional) in enumerate(steps7):
        bx = ML + (sw + Inches(0.1)) * i
        sh = _box(sl, bx, sy, sw, Inches(0.38), color)
        _set_text(sh, label, size=9, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.06))
        th = _box(sl, bx, sy + Inches(0.38), sw, Inches(0.72),
                  RGBColor(0xFF, 0xED, 0xB5) if optional else ACCENT)
        lines = title.split('\n')
        _set_multiline(th, lines, size=9,
                       color=ORANGE2 if optional else NAVY2,
                       align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.06))
        if i < 6:
            _tb(sl, bx + sw, sy + Inches(0.4), Inches(0.1), Inches(0.3),
                "▶", size=10, color=MUTED, align=PP_ALIGN.CENTER)

    _tb(sl, ML, sy + Inches(1.2), CW, Inches(0.28),
        "※ STEP 4・5（オレンジ色）は任意。残業・有給がある月のみ操作してください。",
        size=9, color=MUTED)

    step_descs = [
        ("STEP 1", "ブラウザでアプリのURLを開き、ユーザー名・パスワードを入力してログイン。"),
        ("STEP 2", "画面上部の「年・月」が自動で当月に設定。別の月は数字を直接変更。"),
        ("STEP 3", "月〜金の開始・終了時間を入力。月間スケジュールで全日程・合計稼働時間を確認。"),
        ("STEP 4", "「＋ 例外日」から残業・休日出勤など個別設定（土日・祝日も上書き可）。"),
        ("STEP 5", "「有給を取得する日がある」にチェックしカレンダーで日付を選択。"),
        ("STEP 6", "Excelファイル選択エリアをクリックし、配布された報告書（.xlsx）を選択。"),
        ("STEP 7", "「入力完了・ダウンロード」を押すと自動入力済みファイルを取得。"),
    ]
    desc_w = (CW - Inches(0.6)) / 2
    desc_y = sy + Inches(1.6)
    for i, (step, desc) in enumerate(step_descs):
        col = i % 2
        row = i // 2
        if i == 6:
            dx = ML
            dy = desc_y + Inches(1.0) * 3
            dw = CW
        else:
            dx = ML + col * (desc_w + Inches(0.6))
            dy = desc_y + Inches(1.0) * row
            dw = desc_w
        step_b = _box(sl, dx, dy, Inches(0.8), Inches(0.38), NAVY2)
        _set_text(step_b, step, size=9, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.05), mt=Inches(0.06))
        desc_b = _box(sl, dx + Inches(0.8), dy, dw - Inches(0.8), Inches(0.38),
                      ACCENT, line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(desc_b, desc, size=9, color=DARK, ml=Inches(0.1), mt=Inches(0.07))

    # ── Slide 8: 自動入力ルール ──
    sl = _slide(prs)
    _chrome(sl, "6.  自動入力ルール", 8, TOTAL)
    _tb(sl, ML, CY, CW, Inches(0.3),
        "日の種類に応じて以下のルールで各セルへ自動入力されます", size=11, color=MUTED)

    cols8 = [Inches(2.8), Inches(4.5), Inches(4.5)]
    headers8 = ["日の種類", "開始・終了・休憩時間", "備考欄の記入内容"]
    hx = ML
    hy8 = CY + Inches(0.35)
    for w, h in zip(cols8, headers8):
        cell = _box(sl, hx, hy8, w, Inches(0.42), NAVY)
        _set_text(cell, h, size=12, bold=True, color=WHITE,
                  align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.09))
        hx += w

    rows8 = [
        ("出勤日（平日）",              "曜日別設定の開始・終了時間\n休憩時間：1時間00分", "在宅勤務",             ACCENT,                          ACCENT),
        ("例外日（残業・早退・休日出勤）","例外日に設定した開始・終了時間\n休憩時間：1時間00分", "在宅勤務（例外設定が最優先）", RGBColor(0xFE, 0xF3, 0xC7), RGBColor(0xFE, 0xF3, 0xC7)),
        ("有給取得日",                  "空欄",                           "私用により、休暇",    RGBColor(0xFF, 0xF0, 0xF7),      RGBColor(0xFF, 0xF0, 0xF7)),
        ("祝日",                       "空欄",                           "祝日",               RGBColor(0xFF, 0xF5, 0xF5),      RGBColor(0xFF, 0xF5, 0xF5)),
        ("土日",                       "空欄",                           "空欄",               GRAY,                            GRAY),
    ]
    row_h8 = Inches(0.72)
    for ri, (kind, times, note, bg1, bg2) in enumerate(rows8):
        ry = hy8 + Inches(0.42) + row_h8 * ri
        rx = ML
        for ci, (val, w, bg) in enumerate([(kind, cols8[0], bg1), (times, cols8[1], bg2), (note, cols8[2], bg2)]):
            cell = _box(sl, rx, ry, w, row_h8, bg, line_rgb=RGBColor(0xE5, 0xE7, 0xEB))
            lines = val.split('\n')
            _set_multiline(cell, lines, size=11, color=DARK,
                          align=PP_ALIGN.CENTER, ml=Inches(0.1), mt=Inches(0.1))
            rx += w

    notes8_y = hy8 + Inches(0.42) + row_h8 * 5 + Inches(0.1)
    for n in [
        "※ 祝日情報は外部API（jpholiday）から自動取得します。インターネット未接続時は祝日が空欄になる場合があります。",
        "※ 例外日の設定は有給・祝日・土日に関わらず最優先で適用されます（休日出勤の記録にも使用可能）。",
    ]:
        _tb(sl, ML, notes8_y, CW, Inches(0.28), n, size=9, color=MUTED)
        notes8_y += Inches(0.3)

    # ── Slide 9: Q&A ──
    sl = _slide(prs)
    _chrome(sl, "7.  Q&A（よくある質問）", 9, TOTAL)
    qas9 = [
        ("Q1　設定は毎月入力し直す必要がありますか？",
         "不要です。曜日別の勤務時間・各種設定はブラウザに自動保存されます。次回起動時に前回の設定が自動で読み込まれます。"),
        ("Q2　休日出勤が発生した場合はどうすれば？",
         "「＋ 例外日を追加」から対象日（土日・祝日も可）を選択し、出勤時間を入力します。例外日設定が最優先で反映されます。"),
        ("Q3　有給を土日・祝日に取得した場合は？",
         "有給取得日のカレンダーで土日・祝日も選択できます。選択した日の備考欄に「私用により、休暇」が記入されます。"),
        ("Q4　Excelの列構成が違っても使えますか？",
         "はい。アプリがヘッダー文字列を自動検索して対象列を特定するため、列の位置が変わっても正常に動作します。"),
    ]
    qa_h9 = (CH - Inches(0.2)) / len(qas9)
    for i, (q, a) in enumerate(qas9):
        qy = CY + Inches(0.1) + qa_h9 * i
        qrow = _box(sl, ML, qy, CW, qa_h9 * 0.42, ACCENT,
                    line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
        _set_text(qrow, q, size=12, bold=True, color=NAVY,
                  ml=Inches(0.2), mt=Inches(0.1))
        arow = _box(sl, ML, qy + qa_h9 * 0.42, CW, qa_h9 * 0.55, WHITE,
                    line_rgb=RGBColor(0xE5, 0xE7, 0xEB))
        _set_text(arow, a, size=11, color=DARK, ml=Inches(0.25), mt=Inches(0.1))

    # ── Slide 10: まとめ・導入のお願い ──
    sl = _slide(prs)
    _chrome(sl, "8.  まとめ・導入のお願い", 10, TOTAL)

    can_head = _box(sl, ML, CY, CW, Inches(0.42), GREEN)
    _set_text(can_head, "本アプリで実現できること", size=13, bold=True, color=WHITE,
              ml=Inches(0.2), mt=Inches(0.09))
    can_items = [
        "● 月初5分以内の操作で1ヶ月分のExcel作業報告書が自動完成",
        "● 手入力ゼロ・入力ミスゼロを実現",
        "● ブラウザだけで動作。インストール不要・既存PC環境をそのまま利用可能",
        "● 設定が自動保存されるため、翌月以降は変更のある日だけ修正するだけ",
    ]
    for i, s in enumerate(can_items):
        row = _box(sl, ML, CY + Inches(0.52) + Inches(0.58) * i, CW, Inches(0.55),
                   RGBColor(0xF0, 0xFD, 0xF4) if i % 2 == 0 else WHITE,
                   line_rgb=RGBColor(0x6E, 0xE7, 0xB7))
        _set_text(row, s, size=12, color=DARK, ml=Inches(0.2), mt=Inches(0.1))

    req_head = _box(sl, ML, CY + Inches(2.9), CW, Inches(0.42), NAVY)
    _set_text(req_head, "導入にあたってのお願い", size=13, bold=True, color=WHITE,
              ml=Inches(0.2), mt=Inches(0.09))
    req_body = _box(sl, ML, CY + Inches(3.32), CW, Inches(1.5), ACCENT,
                    line_rgb=RGBColor(0xBF, 0xD7, 0xFF))
    _set_multiline(req_body, [
        "本アプリを社内に導入・展開するにあたり、ご承認をお願いいたします。",
        "現在は主に1名での使用を想定していますが、同一フォーマットを使用する他の社員への展開も容易に行えます。",
        "ご不明な点やご要望がございましたら、お気軽にお申し付けください。",
        "",
        "以上",
    ], size=11, color=DARK, ml=Inches(0.25), mt=Inches(0.15))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
