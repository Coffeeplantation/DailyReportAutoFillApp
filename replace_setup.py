"""build_pptx.py の _build_setup_guide_pptx を差し替えるスクリプト"""
import re

NEW_FUNC = r'''

# ════════════════════════════════════════════════════
#  開発環境セットアップガイド
#  Google アカウント → GitHub（Google 連携）→ Claude Code（GitHub 連携）→ Codespace 開発
# ════════════════════════════════════════════════════
def _build_setup_guide_pptx():
    prs = _prs()
    TOTAL = 12

    G_BLUE   = RGBColor(0x42, 0x85, 0xF4)   # Google blue
    GH_BLACK = RGBColor(0x24, 0x29, 0x2F)   # GitHub dark
    GH_GRAY  = RGBColor(0xF6, 0xF8, 0xFA)   # GitHub light bg
    ANT_RUST = RGBColor(0xCC, 0x5C, 0x36)   # Anthropic brand
    CS_VIOLT = RGBColor(0x6E, 0x40, 0xC9)   # Codespace purple
    CODE_BG  = RGBColor(0x0D, 0x11, 0x17)   # terminal dark
    CODE_FG  = RGBColor(0x7E, 0xE7, 0x87)   # terminal green
    WARN_BG  = RGBColor(0xFF, 0xFB, 0xEB)

    # ── ローカルヘルパー ──────────────────────────────────────
    def _code_block(slide, x, y, w, h, text):
        bg = _rbox(slide, x, y, w, h, CODE_BG)
        _set_tf(bg, text, size=9, color=CODE_FG, ml=Inches(0.18), mt=Inches(0.1))
        return bg

    def _url_chip(slide, x, y, w, url_text, color=NAVY):
        chip = _rbox(slide, x, y, w, Inches(0.36), ACCENT2, color, line_w=1.0)
        _set_tf(chip, "\U0001f517  " + url_text, size=10, color=color, ml=Inches(0.12), mt=Inches(0.06))
        return chip

    def _step_hdr(slide, num_str, title, color):
        _num_badge(slide, ML, CY + Inches(0.05), Inches(0.58), num_str, color)
        _tb(slide, ML + Inches(0.68), CY + Inches(0.08), CW - Inches(0.68), Inches(0.44),
            title, size=13, bold=True, color=color)

    def _rows(slide, items, start_y, rh=Inches(0.65), lcolor=NAVY_D):
        for i, item in enumerate(items):
            if len(item) == 2:
                icon_s, text = item
                sub = None
            else:
                icon_s, text, sub = item
            bg = ACCENT2 if i % 2 == 0 else WHITE
            h_row = rh + (Inches(0.32) if sub else Inches(0))
            row = _rbox(slide, ML, start_y + rh * i, CW, h_row - Inches(0.04), bg, GRAY2)
            badge = _oval(slide, ML + Inches(0.1), start_y + rh * i + Inches(0.16),
                          Inches(0.3), Inches(0.3), lcolor)
            _set_tf(badge, icon_s, size=7, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER, ml=Inches(0.02), mt=Inches(0.04))
            _tb(slide, ML + Inches(0.5), start_y + rh * i + Inches(0.1),
                CW - Inches(0.55), rh - Inches(0.16), text, size=11, color=DARK)
            if sub:
                _code_block(slide, ML + Inches(0.5), start_y + rh * i + rh - Inches(0.08),
                            CW - Inches(0.55), Inches(0.3), sub)

    def _tip(slide, text, color=NAVY_L, bg=None):
        if bg is None:
            bg = ACCENT2
        tip = _rbox(slide, ML, H - FH - Inches(0.52), CW, Inches(0.4), bg, color, line_w=1.0)
        _set_tf(tip, "\U0001f4a1  " + text, size=10, color=DARK, ml=Inches(0.2), mt=Inches(0.08))

    def _warn(slide, text):
        w = _rbox(slide, ML, H - FH - Inches(0.52), CW, Inches(0.4), WARN_BG, ORANGE, line_w=1.0)
        _set_tf(w, "⚠  " + text, size=10, color=ORANGE_D, ml=Inches(0.2), mt=Inches(0.08))

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 1: タイトル
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _box(sl, 0, 0, W, H, NAVY_D)
    _oval(sl, W - Inches(5.5), H - Inches(5.5), Inches(7.5), Inches(7.5), RGBColor(0x1E, 0x40, 0xAF))
    _oval(sl, W - Inches(3.8), H - Inches(3.8), Inches(5.5), Inches(5.5), RGBColor(0x1D, 0x4E, 0xD8))
    _oval(sl, -Inches(1.5), -Inches(1.2), Inches(4.5), Inches(4.5), RGBColor(0x1E, 0x40, 0xAF))
    _box(sl, 0, Inches(1.8), Inches(0.18), Inches(2.2), NAVY_L)
    t = _box(sl, Inches(0.38), Inches(1.85), W - Inches(4.8), Inches(1.5), NAVY_D)
    _set_tf(t, "開発環境セットアップガイド", size=30, bold=True, color=WHITE, ml=Inches(0.25), mt=Inches(0.22))
    sub = _box(sl, Inches(0.38), Inches(3.45), W - Inches(4.8), Inches(0.55), NAVY_D)
    _set_tf(sub, "Google アカウント取得からアプリ開発開始まで  ／  所要時間：30〜60 分",
            size=13, color=RGBColor(0xBF, 0xD7, 0xFF), ml=Inches(0.25), mt=Inches(0.06))
    for i, (c, s, tl) in enumerate([
        (G_BLUE,   "STEP 1", "Google アカウントの作成"),
        (GH_BLACK, "STEP 2", "GitHub アカウントの登録（Google アカウント連携）"),
        (ANT_RUST, "STEP 3", "Claude Code の登録（GitHub アカウント連携）"),
        (CS_VIOLT, "STEP 4", "リポジトリ作成と Codespace 起動"),
        (GREEN,    "STEP 5", "Claude Code でアプリ開発開始"),
    ]):
        _oval(sl, Inches(0.45), Inches(4.35) + Inches(0.46) * i, Inches(0.26), Inches(0.26), c)
        _tb(sl, Inches(0.83), Inches(4.32) + Inches(0.46) * i, Inches(6.0), Inches(0.38),
            f"{s}  {tl}", size=10, color=RGBColor(0xBF, 0xD7, 0xFF))
    _box(sl, 0, H - Inches(0.32), W, Inches(0.32), NAVY_D)
    _tb(sl, W - Inches(1.5), H - Inches(0.28), Inches(1.3), Inches(0.25),
        "1  /  12", size=9, color=MUTED_L, align=PP_ALIGN.RIGHT)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 2: 全体の流れ
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "全体の流れ", 2, TOTAL)
    _tb(sl, ML, CY + Inches(0.04), CW, Inches(0.28),
        "5 ステップで開発環境を整えます。本ガイドは Codespace（クラウド開発環境）を使った方法の一例です。インストール不要でブラウザだけで完結します。",
        size=10, color=MUTED, italic=True)
    _step_flow(sl, [
        ("STEP 1", "Google\nアカウント作成",          G_BLUE,   False),
        ("STEP 2", "GitHub\n登録\n(Google連携)",       GH_BLACK, False),
        ("STEP 3", "Claude Code\n登録\n(GitHub連携)",  ANT_RUST, False),
        ("STEP 4", "リポジトリ作成\nCodespace起動",   CS_VIOLT, False),
        ("STEP 5", "Claude Code\nで開発開始",          GREEN,    False),
    ], ML, CY + Inches(0.38), CW, Inches(1.2))
    for i, (color, step, head, body) in enumerate([
        (G_BLUE,   "STEP 1", "Google アカウントの作成",
         "Gmail アドレスを取得。STEP 2・3 の登録にもこのアカウントを使います。"),
        (GH_BLACK, "STEP 2", "GitHub アカウントの登録（Google アカウント連携）",
         "Gmail アドレスで GitHub に登録し、設定から Google アカウントを OAuth 連携します。"),
        (ANT_RUST, "STEP 3", "Claude Code の登録（GitHub アカウント連携）",
         "claude.ai に GitHub アカウントで登録。Pro プランで Claude Code がフル活用できます。"),
        (CS_VIOLT, "STEP 4", "リポジトリ作成と Codespace 起動",
         "GitHub にリポジトリを作成し、ブラウザ上の開発環境「Codespace」を起動します。"),
        (GREEN,    "STEP 5", "Claude Code でアプリ開発開始",
         "Codespace のターミナルで Claude Code をインストールし、日本語で指示するだけで開発開始。"),
    ]):
        col = i % 2; row = i // 2
        if i == 4:
            dx, dy, dw = ML, CY + Inches(1.75) + Inches(0.85) * 2, CW
        else:
            dw = (CW - Inches(0.3)) / 2
            dx = ML + col * (dw + Inches(0.3)); dy = CY + Inches(1.75) + Inches(0.85) * row
        s = _rbox(sl, dx, dy, Inches(1.1), Inches(0.38), color)
        _set_tf(s, step, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.05), mt=Inches(0.08))
        h2 = _rbox(sl, dx + Inches(1.1), dy, dw - Inches(1.1), Inches(0.38), ACCENT2, GRAY2)
        _box(sl, dx + Inches(1.1), dy, Inches(0.05), Inches(0.38), color)
        _set_tf(h2, f"{head}  /  {body}", size=9, color=DARK, ml=Inches(0.15), mt=Inches(0.08))

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 3: STEP 1 Google アカウント作成
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 1  Google アカウントの作成", 3, TOTAL)
    _step_hdr(sl, "1", "Gmail アドレスを取得します（GitHub・Claude Code の登録にも使います）", G_BLUE)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(5.5), "https://accounts.google.com/signup", G_BLUE)
    _rows(sl, [
        ("①", "上記 URL をブラウザで開き、「アカウントを作成」→「個人用」をクリックします。"),
        ("②", "姓・名（ニックネームでも可）を入力して「次へ」を押します。"),
        ("③", "生年月日と性別を入力して「次へ」を押します。"),
        ("④", "ユーザー名（例：yamada.taro2024）を入力します。これが Gmail アドレスの @ 前になります。"),
        ("⑤", "パスワードを設定します（8文字以上・英字大小・数字・記号を混ぜると安全）。"),
        ("⑥", "電話番号を入力し、SMS で届いた確認コード（6桁）を入力します。"),
        ("⑦", "利用規約を確認して「同意する」をクリック → 作成完了です。"),
    ], CY + Inches(1.02), rh=Inches(0.625), lcolor=G_BLUE)
    _tip(sl, "すでに Google アカウントをお持ちの場合は STEP 2 へ進んでください。")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 4: STEP 2 GitHub アカウント登録
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 2  GitHub アカウントの登録", 4, TOTAL)
    _step_hdr(sl, "2", "Gmail アドレスで GitHub に登録します", GH_BLACK)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(4.0), "https://github.com/signup", GH_BLACK)
    _rows(sl, [
        ("①", "上記 URL を開き「メールアドレス」欄に STEP 1 の Gmail アドレスを入力して「Continue」。"),
        ("②", "パスワードを設定します（GitHub 専用パスワードを新しく決めてください）。「Continue」を押します。"),
        ("③", "ユーザー名を入力します（例：yamada-taro）。英字・数字・ハイフンのみ使用可。後から変更できます。"),
        ("④", "「Would you like to receive product updates?」→ 「n」で「Continue」。"),
        ("⑤", "「Verify your account」でパズル認証を完了して「Create account」を押します。"),
        ("⑥", "Gmail に届いた件名「Your GitHub launch code」のメールを開き、6桁のコードを入力します。"),
        ("⑦", "「How many team members...」などのアンケートはすべて「Skip personalization」で OK です。"),
    ], CY + Inches(1.02), rh=Inches(0.625), lcolor=GH_BLACK)
    _warn(sl, "ユーザー名は公開されます。本名ではなくニックネームでも構いません。後から変更可能です。")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 5: STEP 2続 GitHub に Google アカウントを連携 & 2FA
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 2  GitHub 初期設定：Google アカウント連携と 2FA", 5, TOTAL)
    half5 = (CW - Inches(0.3)) / 2
    # 左：Google 連携
    gl = _rbox(sl, ML, CY + Inches(0.05), half5, Inches(0.44), G_BLUE)
    _set_tf(gl, "Google アカウントの連携（ソーシャルログイン）", size=11, bold=True,
            color=WHITE, ml=Inches(0.15), mt=Inches(0.1))
    for i, (icon_s, text) in enumerate([
        ("①", "GitHub にログイン後、右上のアイコン → 「Settings」を開きます。"),
        ("②", "左メニューの「Password and authentication」をクリックします。"),
        ("③", "ページ下部「Social accounts」→「Link Google account」をクリック。"),
        ("④", "Google のログイン画面が開くので、STEP 1 のアカウントを選択します。"),
        ("⑤", "「Connected accounts」に Google が追加されれば連携完了です。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        rr = _rbox(sl, ML, CY + Inches(0.56) + Inches(0.65) * i, half5, Inches(0.61), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), CY + Inches(0.56) + Inches(0.65) * i + Inches(0.15),
                      Inches(0.28), Inches(0.28), G_BLUE)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.46), CY + Inches(0.56) + Inches(0.65) * i + Inches(0.1),
            half5 - Inches(0.5), Inches(0.48), text, size=10, color=DARK)
    _tip(sl, "Google 連携すると「次回から Google でログイン」が使えて便利です。")
    # 右：2FA
    rx5 = ML + half5 + Inches(0.3)
    gr = _rbox(sl, rx5, CY + Inches(0.05), half5, Inches(0.44), GH_BLACK)
    _set_tf(gr, "2段階認証（2FA）の設定（強く推奨）", size=11, bold=True,
            color=WHITE, ml=Inches(0.15), mt=Inches(0.1))
    for i, (icon_s, text) in enumerate([
        ("①", "「Password and authentication」ページ内「Two-factor authentication」→「Enable」。"),
        ("②", "スマホに「Google Authenticator」アプリをインストールします（App Store / Google Play）。"),
        ("③", "「Authenticator app」を選択し、表示された QR コードをアプリでスキャンします。"),
        ("④", "アプリに表示された 6 桁のコードを GitHub 画面に入力して「Continue」を押します。"),
        ("⑤", "バックアップコードが表示されます。必ずダウンロード・印刷して保管してください。"),
    ]):
        bg = GH_GRAY if i % 2 == 0 else WHITE
        rr = _rbox(sl, rx5, CY + Inches(0.56) + Inches(0.65) * i, half5, Inches(0.61), bg, GRAY2)
        badge = _oval(sl, rx5 + Inches(0.1), CY + Inches(0.56) + Inches(0.65) * i + Inches(0.15),
                      Inches(0.28), Inches(0.28), GH_BLACK)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, rx5 + Inches(0.46), CY + Inches(0.56) + Inches(0.65) * i + Inches(0.1),
            half5 - Inches(0.5), Inches(0.48), text, size=10, color=DARK)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 6: STEP 3 claude.ai アカウント作成（GitHub 連携）
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 3  Claude Code の登録（GitHub アカウント連携）", 6, TOTAL)
    _step_hdr(sl, "3", "GitHub アカウントで claude.ai に登録し、Claude Code を有効化します", ANT_RUST)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(3.8), "https://claude.ai", ANT_RUST)
    _rows(sl, [
        ("①", "上記 URL をブラウザで開き、「Sign up」をクリックします。"),
        ("②", "「Continue with GitHub」をクリックします（GitHub アカウントで連携登録）。"),
        ("③", "GitHub のログイン画面が表示されたら、STEP 2 で作成したアカウントでログインします。"),
        ("④", "「Authorize Anthropic」の確認画面が表示されたら「Authorize」をクリックします。"),
        ("⑤", "電話番号認証が求められる場合は、SMS で届いた 6 桁のコードを入力します。"),
        ("⑥", "利用規約に同意して「Continue」→ アカウント作成完了です。"),
    ], CY + Inches(1.02), rh=Inches(0.65), lcolor=ANT_RUST)
    # プランの説明
    plan_y = CY + Inches(1.02) + Inches(0.65) * 6 + Inches(0.12)
    ph = _rbox(sl, ML, plan_y, CW, Inches(0.38), NAVY_D)
    _set_tf(ph, "プランの選択  ─  Claude Code をフル活用するには Pro プランが必要です", size=11, bold=True,
            color=WHITE, ml=Inches(0.2), mt=Inches(0.08))
    pw = (CW - Inches(0.2)) / 2
    for i, (plan, price, desc, bg2, is_rec) in enumerate([
        ("Free プラン", "無料",   "基本的な会話が可能。Claude Code の使用回数に厳しい制限があります。", GRAY, False),
        ("Pro プラン",  "$20/月", "Claude Code 使い放題。本格的なアプリ開発・自動コーディングに対応。",  ACCENT2, True),
    ]):
        px = ML + (pw + Inches(0.2)) * i
        pp = _rbox(sl, px, plan_y + Inches(0.38), pw, Inches(0.4), NAVY if is_rec else GRAY2)
        _set_tf(pp, f"{plan}  /  {price}" + ("  ★推奨" if is_rec else ""),
                size=11, bold=True, color=WHITE if is_rec else DARK, ml=Inches(0.15), mt=Inches(0.08))
        pb = _rbox(sl, px, plan_y + Inches(0.78), pw, Inches(0.42), WHITE, GRAY2)
        _set_tf(pb, desc, size=10, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 7: STEP 3続 Pro プランへのアップグレード
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 3  Pro プランへのアップグレード", 7, TOTAL)
    _tb(sl, ML, CY + Inches(0.04), CW, Inches(0.3),
        "Claude Code をフル活用するために Pro プラン（$20/月）にアップグレードします",
        size=11, color=MUTED, italic=True)
    up_head = _rbox(sl, ML, CY + Inches(0.38), CW, Inches(0.44), ANT_RUST)
    _set_tf(up_head, "Pro プランへのアップグレード手順", size=13, bold=True,
            color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    for i, (icon_s, text) in enumerate([
        ("①", "claude.ai にログイン後、左下または画面上部の「Upgrade to Pro」をクリックします。"),
        ("②", "プラン選択画面で「Pro  $20/month」の「Subscribe」ボタンをクリックします。"),
        ("③", "クレジットカード情報を入力して「Subscribe to Pro」をクリックします（Visa / Mastercard 可）。"),
        ("④", "登録完了メールが届き、ダッシュボードに「Pro」バッジが表示されれば完了です。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        row = _rbox(sl, ML, CY + Inches(0.9) + Inches(0.7) * i, CW, Inches(0.66), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), CY + Inches(0.9) + Inches(0.7) * i + Inches(0.18),
                      Inches(0.3), Inches(0.3), ANT_RUST)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.5), CY + Inches(0.9) + Inches(0.7) * i + Inches(0.12),
            CW - Inches(0.55), Inches(0.5), text, size=11, color=DARK)
    # Claude Code 確認
    cc_y = CY + Inches(0.9) + Inches(0.7) * 4 + Inches(0.12)
    cc_head = _rbox(sl, ML, cc_y, CW, Inches(0.44), NAVY_D)
    _set_tf(cc_head, "Claude Code の利用確認（Web 版）", size=13, bold=True,
            color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    for i, (icon_s, text) in enumerate([
        ("①", "claude.ai にアクセスし、チャット画面が開いていることを確認します。"),
        ("②", "右上のメニューから「Claude Code」または「Use Claude Code」を選択します。"),
        ("③", "コマンド入力欄が表示されれば Pro プランで Claude Code が有効になっています。"),
    ]):
        bg = PURPLE_L if i % 2 == 0 else WHITE
        row = _rbox(sl, ML, cc_y + Inches(0.44) + Inches(0.56) * i, CW, Inches(0.52), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), cc_y + Inches(0.44) + Inches(0.56) * i + Inches(0.12),
                      Inches(0.26), Inches(0.26), PURPLE_D)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.03))
        _tb(sl, ML + Inches(0.44), cc_y + Inches(0.44) + Inches(0.56) * i + Inches(0.08),
            CW - Inches(0.5), Inches(0.42), text, size=11, color=DARK)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 8: STEP 4 リポジトリ作成
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 4  GitHub リポジトリを作成する", 8, TOTAL)
    _step_hdr(sl, "4", "コードを保存・管理する「リポジトリ」を GitHub に作成します", GH_BLACK)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(4.5), "https://github.com/new", GH_BLACK)
    _rows(sl, [
        ("①", "上記 URL を開きます（または GitHub ホーム右上の「＋」→「New repository」をクリック）。"),
        ("②", "「Repository name」に英字でプロジェクト名を入力します（例：my-first-app）。スペース不可。"),
        ("③", "説明（Description）は任意です。スキップして「Public」または「Private」を選択します。"),
        ("④", "「Add a README file」にチェックを入れます（Codespace 起動に必要）。"),
        ("⑤", "「Add .gitignore」→「Python」を選択します（不要ファイルを除外）。"),
        ("⑥", "「Create repository」ボタンをクリックします。"),
    ], CY + Inches(1.02), rh=Inches(0.65), lcolor=GH_BLACK)
    tip8a = _rbox(sl, ML, H - FH - Inches(0.78), CW, Inches(0.25), GH_GRAY, GRAY2)
    _set_tf(tip8a, "Private にしておけば自分だけが閲覧できます。後から Public に変更することも可能です。",
            size=10, color=MUTED, ml=Inches(0.15), mt=Inches(0.04))
    _warn(sl, "README ファイルがないとリポジトリが空の状態になり、Codespace を起動できません。必ずチェックを入れてください。")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 9: STEP 4続 Codespace 起動
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 4  GitHub Codespace を起動する", 9, TOTAL)
    _tb(sl, ML, CY + Inches(0.04), CW, Inches(0.3),
        "本ガイドでは Codespace（ブラウザ上の VS Code 環境）を使う方法を紹介します。Codespace はあくまで一例であり、VS Code のローカルインストールなど他の方法でも問題ありません。",
        size=11, color=MUTED, italic=True)
    # What is Codespace
    cs_info_y = CY + Inches(0.38)
    ci = _rbox(sl, ML, cs_info_y, CW, Inches(0.44), CS_VIOLT)
    _set_tf(ci, "Codespace を使うメリット", size=12, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    feat_w = (CW - Inches(0.3)) / 3
    for i, (icon_s, head, body) in enumerate([
        ("\U0001f4bb", "ブラウザだけで動く", "VS Code がそのままブラウザで使えます。インストール不要。"),
        ("\U000026a1", "環境が最初から整っている", "Node.js・Python・Git がプリインストール済みです。"),
        ("\U00002601", "クラウド保存",         "作業内容はクラウドに保存。どのPCからでもアクセス可。"),
    ]):
        fx = ML + (feat_w + Inches(0.15)) * i
        fb = _rbox(sl, fx, cs_info_y + Inches(0.44), feat_w, Inches(0.9), ACCENT2, GRAY2)
        _tb(sl, fx + Inches(0.05), cs_info_y + Inches(0.5), feat_w - Inches(0.1), Inches(0.3),
            icon_s + "  " + head, size=11, bold=True, color=CS_VIOLT)
        _tb(sl, fx + Inches(0.1), cs_info_y + Inches(0.82), feat_w - Inches(0.15), Inches(0.5),
            body, size=10, color=DARK)
    # 起動手順
    launch_y = cs_info_y + Inches(1.42)
    lh = _rbox(sl, ML, launch_y, CW, Inches(0.44), GH_BLACK)
    _set_tf(lh, "Codespace の起動手順", size=12, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.09))
    for i, (icon_s, text) in enumerate([
        ("①", "作成したリポジトリのページを開きます（github.com → Your repositories → リポジトリ名）。"),
        ("②", "緑色の「<> Code」ボタンをクリックし、「Codespaces」タブを選択します。"),
        ("③", "「Create codespace on main」をクリックします。"),
        ("④", "ブラウザ内に VS Code が起動します（初回は 1〜2 分かかります）。"),
        ("⑤", "画面下部の「Terminal」タブをクリックしてターミナルを開きます。表示されない場合は メニュー「Terminal」→「New Terminal」。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        rr = _rbox(sl, ML, launch_y + Inches(0.44) + Inches(0.62) * i, CW, Inches(0.58), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), launch_y + Inches(0.44) + Inches(0.62) * i + Inches(0.14),
                      Inches(0.28), Inches(0.28), GH_BLACK)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.46), launch_y + Inches(0.44) + Inches(0.62) * i + Inches(0.08),
            CW - Inches(0.5), Inches(0.46), text, size=10, color=DARK)
    _tip(sl, "無料プランでは月 120 時間・ストレージ 15 GB が使用できます。Pro プランは月 180 時間まで無料。")

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 10: STEP 4続 Claude Code インストール & 認証
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 4  Codespace 内で Claude Code をインストールする", 10, TOTAL)
    _tb(sl, ML, CY + Inches(0.04), CW, Inches(0.28),
        "Codespace のターミナルに以下のコマンドを順番に入力してください",
        size=11, color=MUTED, italic=True)
    cmd_y = CY + Inches(0.38)
    half10 = (CW - Inches(0.3)) / 2
    # 左列：コマンド手順
    for i, (head, cmd, note) in enumerate([
        ("① Node.js のバージョンを確認する",
         "node --version\n# 出力例: v20.x.x または v22.x.x",
         "v18 以上であれば問題ありません。Codespace には最初から入っています。"),
        ("② Claude Code をインストールする",
         "npm install -g @anthropic-ai/claude-code",
         "「claude」コマンドが使えるようになります。数十秒かかります。"),
        ("③ インストールを確認する",
         "claude --version\n# 出力例: 1.x.x",
         "バージョン番号が表示されればインストール成功です。"),
        ("④ 初回認証を行う",
         "claude\n# → ブラウザに認証URLが表示されます",
         "表示された URL をクリックし、claude.ai の Pro アカウントでログインします。"),
    ]):
        hy = cmd_y + Inches(1.4) * i
        hb = _rbox(sl, ML, hy, half10, Inches(0.35), GH_BLACK)
        _set_tf(hb, head, size=10, bold=True, color=WHITE, ml=Inches(0.12), mt=Inches(0.07))
        _code_block(sl, ML, hy + Inches(0.35), half10, Inches(0.6), cmd)
        nb = _rbox(sl, ML, hy + Inches(0.95), half10, Inches(0.38), ACCENT2, GRAY2)
        _set_tf(nb, note, size=9, color=DARK, ml=Inches(0.12), mt=Inches(0.07))
    # 右列：認証フロー説明
    rx10 = ML + half10 + Inches(0.3)
    ah = _rbox(sl, rx10, cmd_y, half10, Inches(0.44), ANT_RUST)
    _set_tf(ah, "初回認証の流れ", size=12, bold=True, color=WHITE, ml=Inches(0.18), mt=Inches(0.09))
    for i, (icon_s, text) in enumerate([
        ("\U0001f4bb", "ターミナルで「claude」と入力して Enter を押します。"),
        ("\U0001f517", "「Please open the following URL...」というメッセージとURLが表示されます。"),
        ("\U0001f310", "そのURLをクリック（または Ctrl クリック）でブラウザが開きます。"),
        ("\U0001f511", "claude.ai の Pro アカウントでログインし「Authorize」をクリックします。"),
        ("✅",     "ターミナルに「Authenticated successfully」と表示されれば認証完了です。"),
        ("\U0001f4ac", "チャットが開始します。日本語でメッセージを入力してみましょう！"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        ar = _rbox(sl, rx10, cmd_y + Inches(0.44) + Inches(0.58) * i, half10, Inches(0.54), bg, GRAY2)
        _tb(sl, rx10 + Inches(0.1), cmd_y + Inches(0.44) + Inches(0.58) * i + Inches(0.06),
            Inches(0.3), Inches(0.38), icon_s, size=14, align=PP_ALIGN.CENTER)
        _tb(sl, rx10 + Inches(0.48), cmd_y + Inches(0.44) + Inches(0.58) * i + Inches(0.1),
            half10 - Inches(0.52), Inches(0.38), text, size=10, color=DARK)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 11: STEP 5 開発開始とコード保存
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "STEP 5  Claude Code でアプリ開発を開始する", 11, TOTAL)
    _tb(sl, ML, CY + Inches(0.04), CW, Inches(0.3),
        "あとは Claude Code に日本語で指示するだけです！（Codespace のターミナルで実行）",
        size=13, bold=True, color=GREEN_D)
    half11 = (CW - Inches(0.3)) / 2
    # 左：指示の例
    eh = _rbox(sl, ML, CY + Inches(0.42), half11, Inches(0.44), GREEN_D)
    _set_tf(eh, "Claude Code への指示の例", size=12, bold=True, color=WHITE,
            ml=Inches(0.2), mt=Inches(0.09))
    for i, ex in enumerate([
        "「Flask を使ったシンプルな Todo アプリを作ってください」",
        "「このコードにコメントを日本語で追加してください」",
        "「バグを見つけて直してください」",
        "「README.md を日本語で書いてください」",
        "「ログイン機能を追加してください」",
    ]):
        bg = GREEN_L if i % 2 == 0 else WHITE
        er = _rbox(sl, ML, CY + Inches(0.94) + Inches(0.55) * i, half11, Inches(0.51), bg, GRAY2)
        _box(sl, ML, CY + Inches(0.94) + Inches(0.55) * i, Inches(0.05), Inches(0.51), GREEN)
        _set_tf(er, "\U0001f4ac  " + ex, size=10, color=DARK, ml=Inches(0.2), mt=Inches(0.12))
    # 右：git コマンド
    rx11 = ML + half11 + Inches(0.3)
    gh = _rbox(sl, rx11, CY + Inches(0.42), half11, Inches(0.44), GH_BLACK)
    _set_tf(gh, "作業後は GitHub に保存（push）", size=12, bold=True, color=WHITE,
            ml=Inches(0.18), mt=Inches(0.09))
    for i, (cmd, note) in enumerate([
        ("git status",                          "変更されたファイルを確認します。"),
        ('git add .',                           "すべての変更をステージングします。"),
        ('git commit -m "変更内容の説明"',      "変更をコミット（記録）します。"),
        ("git push",                            "GitHub に変更をアップロードします。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        cr = _rbox(sl, rx11, CY + Inches(0.94) + Inches(0.65) * i, half11, Inches(0.61), bg, GRAY2)
        _code_block(sl, rx11 + Inches(0.1), CY + Inches(0.94) + Inches(0.65) * i + Inches(0.06),
                    half11 - Inches(0.2), Inches(0.3), cmd)
        _tb(sl, rx11 + Inches(0.1), CY + Inches(0.94) + Inches(0.65) * i + Inches(0.36),
            half11 - Inches(0.15), Inches(0.22), note, size=9, color=MUTED)
    _tip(sl, "Codespace 内の変更は自動保存されます。git push するまでは GitHub 側には反映されません。", NAVY_L, ACCENT2)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    #  Slide 12: Q&A
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    sl = _slide(prs)
    _chrome(sl, "よくある質問（Q&A）", 12, TOTAL)
    qa_h = CH / 4
    for i, (qn_s, q, a, color) in enumerate([
        ("Q1", "Codespace は無料で使えますか？",
         "個人アカウントは月120時間・15GBまで無料です。Pro プランユーザーは月180時間まで無料。超過後は課金が発生しますが、通常の使用では無料枠内に収まります。",
         G_BLUE),
        ("Q2", "Claude Code の Pro プランは必須ですか？",
         "Free プランでも試用できますが、使用回数に厳しい制限があります。本格的な開発には Pro プラン（$20/月）を推奨します。Claude Code で自動生成できるコード量が大幅に増えます。",
         ANT_RUST),
        ("Q3", "Codespace を閉じてもデータは消えませんか？",
         "ブラウザを閉じても Codespace は保持されます。GitHub の「Codespaces」ページから再開できます。ただし非アクティブ状態が 30 日続くと自動削除されます。",
         CS_VIOLT),
        ("Q4", "GitHub・claude.ai のパスワードを忘れた場合は？",
         "GitHub：ログイン画面の「Forgot password?」からメール（Gmail）でリセットできます。claude.ai：「Forgot password?」または「Continue with GitHub」を使って再ログインできます。",
         GH_BLACK),
    ]):
        qy = CY + qa_h * i
        badge = _oval(sl, ML, qy + qa_h * 0.1, Inches(0.55), Inches(0.55), color)
        _set_tf(badge, qn_s, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.03), mt=Inches(0.09))
        qrow = _rbox(sl, ML + Inches(0.65), qy + qa_h * 0.08,
                     CW - Inches(0.65), qa_h * 0.38, ACCENT2, GRAY2)
        _box(sl, ML + Inches(0.65), qy + qa_h * 0.08, Inches(0.06), qa_h * 0.38, color)
        _set_tf(qrow, q, size=11, bold=True, color=NAVY_D, ml=Inches(0.22), mt=Inches(0.08))
        arow = _rbox(sl, ML + Inches(0.65), qy + qa_h * 0.46,
                     CW - Inches(0.65), qa_h * 0.5, WHITE, GRAY2)
        _box(sl, ML + Inches(0.65), qy + qa_h * 0.46, Inches(0.06), qa_h * 0.5, MUTED_L)
        _set_tf(arow, a, size=10, color=DARK, ml=Inches(0.22), mt=Inches(0.07))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
'''

with open('build_pptx.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 古い setup guide 関数を新しいものに差し替え
marker = 'def _build_setup_guide_pptx():'
idx = content.find(marker)
if idx == -1:
    print('ERROR: marker not found')
else:
    # コメントブロックも含めて置き換えるため、直前の空行まで遡る
    pre = content[:idx]
    last_blank = pre.rfind('\n\n')
    trim_idx = last_blank  # \n\n の位置
    new_content = content[:trim_idx] + NEW_FUNC
    with open('build_pptx.py', 'w', encoding='utf-8') as f:
        f.write(new_content)
    print(f'Done. Replaced from index {idx}')
