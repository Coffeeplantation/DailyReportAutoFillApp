"""build_pptx.py にセットアップガイド関数を追記するスクリプト"""
code = r'''

# ════════════════════════════════════════════════════
#  開発環境セットアップガイド
#  Google アカウント取得 → GitHub → Claude Code → 開発開始
# ════════════════════════════════════════════════════
def _build_setup_guide_pptx():
    prs = _prs()
    TOTAL = 12

    G_BLUE   = RGBColor(0x42, 0x85, 0xF4)
    GH_BLACK = RGBColor(0x24, 0x29, 0x2F)
    GH_GRAY  = RGBColor(0xF6, 0xF8, 0xFA)
    ANT_RUST = RGBColor(0xCC, 0x5C, 0x36)
    CODE_BG  = RGBColor(0x1E, 0x1E, 0x2E)
    CODE_FG  = RGBColor(0xA6, 0xE3, 0xA1)

    def _code_block(slide, x, y, w, h, text):
        bg = _rbox(slide, x, y, w, h, CODE_BG)
        _set_tf(bg, text, size=10, color=CODE_FG, ml=Inches(0.18), mt=Inches(0.1))
        return bg

    def _url_chip(slide, x, y, w, url_text, color=NAVY):
        chip = _rbox(slide, x, y, w, Inches(0.36), ACCENT2, color, line_w=1.0)
        _set_tf(chip, "\U0001f517  " + url_text, size=10, color=color, ml=Inches(0.12), mt=Inches(0.06))
        return chip

    def _step_hdr(slide, num_str, title, color):
        _num_badge(slide, ML, CY + Inches(0.05), Inches(0.58), num_str, color)
        _tb(slide, ML + Inches(0.68), CY + Inches(0.1), CW - Inches(0.68), Inches(0.4),
            title, size=14, bold=True, color=color)

    def _rows(slide, items, start_y, rh=Inches(0.66), lcolor=NAVY_D):
        for i, (icon_s, text) in enumerate(items):
            bg = ACCENT2 if i % 2 == 0 else WHITE
            row = _rbox(slide, ML, start_y + rh * i, CW, rh - Inches(0.04), bg, GRAY2)
            badge = _oval(slide, ML + Inches(0.1), start_y + rh * i + rh * 0.22,
                          Inches(0.3), Inches(0.3), lcolor)
            _set_tf(badge, icon_s, size=7, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER, ml=Inches(0.02), mt=Inches(0.04))
            _tb(slide, ML + Inches(0.5), start_y + rh * i + Inches(0.11),
                CW - Inches(0.55), rh - Inches(0.15), text, size=11, color=DARK)

    # ─────────── Slide 1: タイトル ───────────
    sl = _slide(prs)
    _box(sl, 0, 0, W, H, NAVY_D)
    _oval(sl, W - Inches(5.5), H - Inches(5.5), Inches(7.5), Inches(7.5), RGBColor(0x1E, 0x40, 0xAF))
    _oval(sl, W - Inches(3.8), H - Inches(3.8), Inches(5.5), Inches(5.5), RGBColor(0x1D, 0x4E, 0xD8))
    _oval(sl, -Inches(1.5), -Inches(1.2), Inches(4.5), Inches(4.5), RGBColor(0x1E, 0x40, 0xAF))
    _box(sl, 0, Inches(1.8), Inches(0.18), Inches(2.5), NAVY_L)
    t = _box(sl, Inches(0.38), Inches(1.85), W - Inches(5.0), Inches(1.5), NAVY_D)
    _set_tf(t, "開発環境セットアップガイド", size=30, bold=True, color=WHITE, ml=Inches(0.25), mt=Inches(0.22))
    sub = _box(sl, Inches(0.38), Inches(3.5), W - Inches(5.0), Inches(0.6), NAVY_D)
    _set_tf(sub, "Google アカウント取得からアプリ開発開始まで",
            size=14, color=RGBColor(0xBF, 0xD7, 0xFF), ml=Inches(0.25), mt=Inches(0.08))
    for i, (c, s, tl) in enumerate([
        (G_BLUE,   "STEP 1", "Google アカウントの作成"),
        (GH_BLACK, "STEP 2", "GitHub アカウントの作成と設定"),
        (ANT_RUST, "STEP 3", "Claude.ai アカウントの作成"),
        (NAVY,     "STEP 4", "Git・Claude Code のインストール"),
        (GREEN,    "STEP 5", "リポジトリ作成とアプリ開発開始"),
    ]):
        _oval(sl, Inches(0.45), Inches(4.6) + Inches(0.48) * i, Inches(0.28), Inches(0.28), c)
        _tb(sl, Inches(0.85), Inches(4.58) + Inches(0.48) * i, Inches(5.5), Inches(0.38),
            f"{s}  {tl}", size=11, color=RGBColor(0xBF, 0xD7, 0xFF))
    _box(sl, 0, H - Inches(0.32), W, Inches(0.32), NAVY_D)
    _tb(sl, W - Inches(1.5), H - Inches(0.28), Inches(1.3), Inches(0.25),
        "1  /  12", size=9, color=MUTED_L, align=PP_ALIGN.RIGHT)

    # ─────────── Slide 2: 全体フロー ───────────
    sl = _slide(prs)
    _chrome(sl, "全体の流れ", 2, TOTAL)
    _tb(sl, ML, CY + Inches(0.05), CW, Inches(0.28),
        "以下の 5 ステップで開発環境を整えます。所要時間の目安：30〜60 分",
        size=10, color=MUTED, italic=True)
    _step_flow(sl, [
        (G_BLUE,   "STEP 1", "Google\nアカウント作成",       False),
        (GH_BLACK, "STEP 2", "GitHub\nアカウント作成",       False),
        (ANT_RUST, "STEP 3", "Claude.ai\n登録",              False),
        (NAVY,     "STEP 4", "Git・Claude Code\nインストール", False),
        (GREEN,    "STEP 5", "リポジトリ作成\n開発開始",       False),
    ], ML, CY + Inches(0.42), CW, Inches(1.1))
    dw2 = (CW - Inches(0.4)) / 2
    dy2 = CY + Inches(1.7)
    for i, (color, step, head, body) in enumerate([
        (G_BLUE,   "STEP 1", "Google アカウントの作成",
         "Gmail アドレスを取得します。GitHub・Claude.ai の登録でも使用します。"),
        (GH_BLACK, "STEP 2", "GitHub アカウントの作成",
         "ソースコード管理サービス。作成したアプリを保存・公開できます。"),
        (ANT_RUST, "STEP 3", "Claude.ai アカウントの作成",
         "Google アカウントで簡単に登録できます。Pro プランで Claude Code が使えます。"),
        (NAVY,     "STEP 4", "Git・Claude Code のインストール",
         "Git はバージョン管理、Claude Code は AI コーディングアシスタントです。"),
        (GREEN,    "STEP 5", "リポジトリ作成とアプリ開発開始",
         "GitHub にリポジトリを作り、Claude Code に指示するだけで開発が始まります。"),
    ]):
        col = i % 2; row = i // 2
        if i == 4:
            dx2, dy2_i, dw2_i = ML, dy2 + Inches(0.9) * 2, CW
        else:
            dx2 = ML + col * (dw2 + Inches(0.4)); dy2_i = dy2 + Inches(0.9) * row; dw2_i = dw2
        s = _rbox(sl, dx2, dy2_i, Inches(0.85), Inches(0.38), color)
        _set_tf(s, step, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.05), mt=Inches(0.07))
        h2 = _rbox(sl, dx2 + Inches(0.85), dy2_i, dw2_i - Inches(0.85), Inches(0.38), ACCENT2, GRAY2)
        _box(sl, dx2 + Inches(0.85), dy2_i, Inches(0.05), Inches(0.38), color)
        _set_tf(h2, f"{head}  {body}", size=9, color=DARK, ml=Inches(0.15), mt=Inches(0.08))

    # ─────────── Slide 3: STEP 1 Google アカウント ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 1  Google アカウントの作成", 3, TOTAL)
    _step_hdr(sl, "1", "Gmail アドレスを取得します（GitHub・Claude.ai の登録にも使います）", G_BLUE)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(5.5), "https://accounts.google.com", G_BLUE)
    _rows(sl, [
        ("①", "ブラウザで上記URLを開き、「アカウントを作成」→「個人用」をクリックします。"),
        ("②", "姓・名を入力して「次へ」を押します。"),
        ("③", "生年月日と性別を入力して「次へ」を押します。"),
        ("④", "ユーザー名（メールアドレスの @ より前の部分）を決めて入力します。"),
        ("⑤", "パスワードを設定します（8文字以上、英字・数字・記号の組み合わせ推奨）。"),
        ("⑥", "電話番号を入力して確認コードを受信し、コードを入力します。"),
        ("⑦", "利用規約を確認して「同意する」をクリックすれば作成完了です。"),
    ], CY + Inches(1.02), rh=Inches(0.63), lcolor=G_BLUE)
    tip = _rbox(sl, ML, H - FH - Inches(0.52), CW, Inches(0.4), ACCENT2, NAVY_L, line_w=1.0)
    _set_tf(tip, "\U0001f4a1  すでに Google アカウントをお持ちの場合は STEP 2 へ進んでください。",
            size=10, color=NAVY_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 4: STEP 2 GitHub アカウント ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 2  GitHub アカウントの作成", 4, TOTAL)
    _step_hdr(sl, "2", "ソースコード管理サービス GitHub にアカウントを作成します", GH_BLACK)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(4.0), "https://github.com", GH_BLACK)
    _rows(sl, [
        ("①", "上記URLを開き、右上の「Sign up」をクリックします。"),
        ("②", "メールアドレスに STEP 1 で作成した Gmail アドレスを入力します。"),
        ("③", "パスワードを設定します（GitHub 独自のパスワードを設定してください）。"),
        ("④", "ユーザー名を決めて入力します（英数字・ハイフンのみ、後から変更可能）。"),
        ("⑤", "メール通知の設定（スキップ可）のあと、パズル認証を通過します。"),
        ("⑥", "GitHub から届いた確認メール（Gmail）を開き、認証コードを入力します。"),
        ("⑦", "プラン選択画面が表示されたら「Continue for free」（無料プラン）を選択します。"),
    ], CY + Inches(1.02), rh=Inches(0.63), lcolor=GH_BLACK)
    warn = _rbox(sl, ML, H - FH - Inches(0.52), CW, Inches(0.4),
                 RGBColor(0xFF, 0xFB, 0xEB), ORANGE, line_w=1.0)
    _set_tf(warn, "⚠  ユーザー名は GitHub 上で公開されます。本名ではなく任意のニックネームでも構いません。",
            size=10, color=ORANGE_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 5: STEP 2続 2FA ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 2  GitHub 初期設定（2段階認証）", 5, TOTAL)
    _tb(sl, ML, CY + Inches(0.05), CW, Inches(0.28),
        "2段階認証（2FA）を設定してアカウントを保護します（強く推奨）", size=11, color=MUTED, italic=True)
    half5 = (CW - Inches(0.3)) / 2
    h5l = _rbox(sl, ML, CY + Inches(0.4), half5, Inches(0.42), GH_BLACK)
    _set_tf(h5l, "2段階認証（2FA）の設定手順", size=12, bold=True, color=WHITE, ml=Inches(0.18), mt=Inches(0.1))
    rh5 = Inches(0.55)
    for i, (icon_s, text) in enumerate([
        ("①", "右上のアイコン → 「Settings」を開きます。"),
        ("②", "左メニューの「Password and authentication」をクリックします。"),
        ("③", "「Two-factor authentication」の「Enable」をクリックします。"),
        ("④", "スマートフォンに「Google Authenticator」アプリをインストールします。"),
        ("⑤", "画面に表示されたQRコードをアプリでスキャンします。"),
        ("⑥", "アプリに表示された6桁のコードを入力して「Enable」をクリックします。"),
        ("⑦", "バックアップコードを安全な場所に保存します。"),
    ]):
        bg = GH_GRAY if i % 2 == 0 else WHITE
        row = _rbox(sl, ML, CY + Inches(0.9) + rh5 * i, half5, rh5 - Inches(0.03), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), CY + Inches(0.9) + rh5 * i + rh5 * 0.22,
                      Inches(0.28), Inches(0.28), GH_BLACK)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.48), CY + Inches(0.9) + rh5 * i + Inches(0.1),
            half5 - Inches(0.5), rh5 - Inches(0.12), text, size=10, color=DARK)
    rx5 = ML + half5 + Inches(0.3)
    h5r = _rbox(sl, rx5, CY + Inches(0.4), half5, Inches(0.42), G_BLUE)
    _set_tf(h5r, "Google Authenticator のインストール", size=12, bold=True, color=WHITE, ml=Inches(0.18), mt=Inches(0.1))
    for i, (icon_s, text) in enumerate([
        ("\U0001f4f1", "スマートフォンの App Store または Google Play を開きます。"),
        ("\U0001f50d", "「Google Authenticator」で検索します。"),
        ("⬇",     "Google LLC が提供するアプリをインストールします。"),
        ("✅",     "アプリを起動し「コードを追加」からQRスキャンを行います。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        row = _rbox(sl, rx5, CY + Inches(0.9) + Inches(0.9) * i, half5, Inches(0.82), bg, GRAY2)
        _tb(sl, rx5 + Inches(0.1), CY + Inches(0.9) + Inches(0.9) * i + Inches(0.05),
            Inches(0.35), Inches(0.38), icon_s, size=16, align=PP_ALIGN.CENTER)
        _tb(sl, rx5 + Inches(0.5), CY + Inches(0.9) + Inches(0.9) * i + Inches(0.2),
            half5 - Inches(0.55), Inches(0.45), text, size=10, color=DARK)
    note5 = _rbox(sl, ML, H - FH - Inches(0.5), CW, Inches(0.38), ACCENT2, NAVY_L, line_w=0.8)
    _set_tf(note5, "\U0001f4a1  2FA を設定しておくと、不正ログインを防止できます。強く推奨します。",
            size=10, color=NAVY_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 6: STEP 3 Claude.ai ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 3  Claude.ai アカウントの作成", 6, TOTAL)
    _step_hdr(sl, "3", "Google アカウントで Claude.ai に登録します", ANT_RUST)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(4.0), "https://claude.ai", ANT_RUST)
    _rows(sl, [
        ("①", "上記URLをブラウザで開き、「Sign up」をクリックします。"),
        ("②", "「Continue with Google」を選択します。"),
        ("③", "STEP 1 で作成した Google アカウントを選択します（自動でログイン）。"),
        ("④", "電話番号認証が求められる場合は、SMS でコードを受け取り入力します。"),
        ("⑤", "利用規約に同意して「Continue」をクリックします。"),
    ], CY + Inches(1.02), rh=Inches(0.68), lcolor=ANT_RUST)
    plan_y = CY + Inches(1.02) + Inches(0.68) * 5 + Inches(0.15)
    plan_head = _rbox(sl, ML, plan_y, CW, Inches(0.4), NAVY_D)
    _set_tf(plan_head, "プランの選択", size=12, bold=True, color=WHITE, ml=Inches(0.2), mt=Inches(0.08))
    pw = (CW - Inches(0.2)) / 2
    for i, (plan, price, desc, bg2, is_rec) in enumerate([
        ("Free プラン", "無料",   "基本的な Claude との会話が可能。Claude Code の利用には制限があります。", GRAY, False),
        ("Pro プラン",  "$20/月", "Claude Code がフル機能で使えます。本格的なアプリ開発に推奨です。",    ACCENT2, True),
    ]):
        px = ML + (pw + Inches(0.2)) * i
        ph = _rbox(sl, px, plan_y + Inches(0.4), pw, Inches(0.42), NAVY if is_rec else bg2)
        _set_tf(ph, f"{plan}  /  {price}" + ("  ★推奨" if is_rec else ""),
                size=12, bold=True, color=WHITE if is_rec else DARK, ml=Inches(0.15), mt=Inches(0.09))
        pb = _rbox(sl, px, plan_y + Inches(0.82), pw, Inches(0.58), WHITE, GRAY2)
        _set_tf(pb, desc, size=10, color=DARK, ml=Inches(0.15), mt=Inches(0.1))

    # ─────────── Slide 7: STEP 4 Git ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 4  Git のインストール", 7, TOTAL)
    _step_hdr(sl, "4", "バージョン管理ツール Git を Windows にインストールします", TEAL)
    _url_chip(sl, ML, CY + Inches(0.58), Inches(5.5), "https://git-scm.com/download/win", TEAL)
    _rows(sl, [
        ("①", "上記URLを開き、「Click here to download」をクリックしてインストーラーをダウンロードします。"),
        ("②", "ダウンロードした「Git-*-64-bit.exe」を実行します。"),
        ("③", "インストーラーの設定はすべてデフォルト（Next を押し続ける）で問題ありません。"),
        ("④", "インストール完了後、「Git Bash」を起動して以下のコマンドを入力します。"),
    ], CY + Inches(1.02), rh=Inches(0.7), lcolor=TEAL)
    _code_block(sl, ML, CY + Inches(1.02) + Inches(0.7) * 4 + Inches(0.08), CW, Inches(0.78),
                'git config --global user.name  "あなたの名前"\ngit config --global user.email "あなたのGmailアドレス"')
    tip7 = _rbox(sl, ML, H - FH - Inches(0.52), CW, Inches(0.4), ACCENT2, NAVY_L, line_w=0.8)
    _set_tf(tip7, "\U0001f4a1  ダブルクォート内を自分の名前・Gmail アドレスに書き換えて実行してください。",
            size=10, color=NAVY_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 8: STEP 4続 Claude Code ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 4  Claude Code のインストール", 8, TOTAL)
    _tb(sl, ML, CY + Inches(0.05), CW, Inches(0.28),
        "AI コーディングアシスタント Claude Code を使えるようにします", size=11, color=MUTED, italic=True)
    half8 = (CW - Inches(0.3)) / 2
    ha = _rbox(sl, ML, CY + Inches(0.38), half8, Inches(0.45), NAVY)
    _set_tf(ha, "方法 A  VS Code 拡張（推奨・簡単）", size=12, bold=True, color=WHITE, ml=Inches(0.18), mt=Inches(0.09))
    _url_chip(sl, ML, CY + Inches(0.9), half8, "https://code.visualstudio.com", NAVY)
    for i, (icon_s, text) in enumerate([
        ("①", "上記URLから VS Code をダウンロード・インストールします。"),
        ("②", "VS Code を起動し、左サイドバーの拡張機能アイコンをクリックします。"),
        ("③", "検索欄に「Claude Code」と入力し、Anthropic 製の拡張機能をインストールします。"),
        ("④", "サインインボタンから claude.ai アカウントで認証します。"),
    ]):
        bg = ACCENT2 if i % 2 == 0 else WHITE
        row = _rbox(sl, ML, CY + Inches(1.32) + Inches(0.62) * i, half8, Inches(0.58), bg, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), CY + Inches(1.32) + Inches(0.62) * i + Inches(0.14),
                      Inches(0.28), Inches(0.28), NAVY)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.03))
        _tb(sl, ML + Inches(0.48), CY + Inches(1.32) + Inches(0.62) * i + Inches(0.1),
            half8 - Inches(0.52), Inches(0.46), text, size=10, color=DARK)
    rx8 = ML + half8 + Inches(0.3)
    hb = _rbox(sl, rx8, CY + Inches(0.38), half8, Inches(0.45), TEAL)
    _set_tf(hb, "方法 B  CLI（上級者向け）", size=12, bold=True, color=WHITE, ml=Inches(0.18), mt=Inches(0.09))
    _url_chip(sl, rx8, CY + Inches(0.9), half8, "https://nodejs.org  （LTS 版）", TEAL)
    for i, (icon_s, text) in enumerate([
        ("①", "上記URLから Node.js（LTS版）をダウンロード・インストールします。"),
        ("②", "コマンドプロンプト（管理者権限）を開きます。"),
        ("③", "以下のコマンドを実行します。"),
        ("④", "「claude」と入力してブラウザで認証します。"),
    ]):
        bg = TEAL_L if i % 2 == 0 else WHITE
        row = _rbox(sl, rx8, CY + Inches(1.32) + Inches(0.62) * i, half8, Inches(0.58), bg, GRAY2)
        badge = _oval(sl, rx8 + Inches(0.1), CY + Inches(1.32) + Inches(0.62) * i + Inches(0.14),
                      Inches(0.28), Inches(0.28), TEAL)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.03))
        _tb(sl, rx8 + Inches(0.48), CY + Inches(1.32) + Inches(0.62) * i + Inches(0.1),
            half8 - Inches(0.52), Inches(0.46), text, size=10, color=DARK)
    _code_block(sl, rx8, CY + Inches(1.32) + Inches(0.62) * 3 + Inches(0.02), half8, Inches(0.36),
                "npm install -g @anthropic-ai/claude-code")

    # ─────────── Slide 9: STEP 5 リポジトリ作成 ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 5  最初のリポジトリを作成する", 9, TOTAL)
    _step_hdr(sl, "5", "GitHub にコードを保存する「リポジトリ」を作成します", GREEN_D)
    _rows(sl, [
        ("①", "github.com にログインし、右上の「＋」→「New repository」をクリックします。"),
        ("②", "「Repository name」に任意の名前を入力します（例：MyFirstApp）。"),
        ("③", "「Public」（公開）または「Private」（非公開）を選択します。"),
        ("④", "「Add a README file」にチェックを入れます。"),
        ("⑤", "「Create repository」ボタンをクリックします。"),
        ("⑥", "リポジトリが作成されたら、緑色の「Code」ボタンをクリックします。"),
        ("⑦", "「HTTPS」タブのURLをコピーします（https://github.com/ユーザー名/リポジトリ名.git）。"),
    ], CY + Inches(1.02), rh=Inches(0.63), lcolor=GREEN_D)
    tip9 = _rbox(sl, ML, H - FH - Inches(0.52), CW, Inches(0.4), GREEN_L, GREEN, line_w=1.0)
    _set_tf(tip9, "\U0001f4a1  Private にしておけば自分だけが見られます。後から Public に変更することもできます。",
            size=10, color=GREEN_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 10: STEP 5続 クローン ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 5  リポジトリをローカルに取得する（クローン）", 10, TOTAL)
    _tb(sl, ML, CY + Inches(0.05), CW, Inches(0.28),
        "GitHub 上のリポジトリを自分のパソコンにダウンロードします", size=11, color=MUTED, italic=True)
    rh10 = Inches(0.68)
    for i, (icon_s, text, bg3) in enumerate([
        ("①", "デスクトップなど作業したいフォルダを決めます。", ACCENT2),
        ("②", "エクスプローラーで対象フォルダを開き、アドレスバーに「cmd」と入力して Enter を押します。", WHITE),
        ("③", "コマンドプロンプトが開いたら、以下のコマンドを入力して Enter を押します。", ACCENT2),
    ]):
        row = _rbox(sl, ML, CY + Inches(0.38) + rh10 * i, CW, rh10 - Inches(0.03), bg3, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), CY + Inches(0.38) + rh10 * i + rh10 * 0.22,
                      Inches(0.3), Inches(0.3), NAVY_D)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.5), CY + Inches(0.38) + rh10 * i + Inches(0.12),
            CW - Inches(0.55), rh10 - Inches(0.15), text, size=11, color=DARK)
    _code_block(sl, ML, CY + Inches(0.38) + rh10 * 3 + Inches(0.06), CW, Inches(0.42),
                "git clone https://github.com/ユーザー名/リポジトリ名.git")
    for i, (icon_s, text) in enumerate([
        ("④", "クローンが完了すると、フォルダ内にリポジトリ名のフォルダが作成されます。"),
        ("⑤", "VS Code を開き、「ファイル」→「フォルダーを開く」からクローンしたフォルダを選択します。"),
        ("⑥", "VS Code 左下の Claude Code アイコンをクリックすると、チャット画面が開きます。"),
    ]):
        iy = CY + Inches(0.38) + rh10 * 3 + Inches(0.56) + Inches(0.7) * i
        bg4 = GREEN_L if i % 2 == 0 else WHITE
        row = _rbox(sl, ML, iy, CW, Inches(0.66), bg4, GRAY2)
        badge = _oval(sl, ML + Inches(0.1), iy + Inches(0.18), Inches(0.3), Inches(0.3), GREEN_D)
        _set_tf(badge, icon_s, size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.02), mt=Inches(0.04))
        _tb(sl, ML + Inches(0.5), iy + Inches(0.12), CW - Inches(0.55), Inches(0.5),
            text, size=11, color=DARK)

    # ─────────── Slide 11: STEP 5続 開発開始 ───────────
    sl = _slide(prs)
    _chrome(sl, "STEP 5  Claude Code でアプリ開発を開始する", 11, TOTAL)
    _tb(sl, ML, CY + Inches(0.05), CW, Inches(0.3),
        "あとは Claude Code に日本語で指示するだけです！", size=13, bold=True, color=GREEN_D)
    ex_head = _rbox(sl, ML, CY + Inches(0.42), CW, Inches(0.42), GREEN_D)
    _set_tf(ex_head, "Claude Code への指示の例", size=12, bold=True, color=WHITE,
            ml=Inches(0.2), mt=Inches(0.09))
    for i, ex in enumerate([
        "「Flask を使ったシンプルな Todo アプリを作ってください」",
        "「このコードのバグを直してください」",
        "「README.md を日本語で書いてください」",
        "「ログイン機能を追加してください」",
    ]):
        bg5 = GREEN_L if i % 2 == 0 else WHITE
        ex_row = _rbox(sl, ML, CY + Inches(0.92) + Inches(0.58) * i, CW, Inches(0.55), bg5, GRAY2)
        _box(sl, ML, CY + Inches(0.92) + Inches(0.58) * i, Inches(0.06), Inches(0.55), GREEN)
        _set_tf(ex_row, "\U0001f4ac  " + ex, size=11, color=DARK, ml=Inches(0.22), mt=Inches(0.13))
    push_head = _rbox(sl, ML, CY + Inches(3.28), CW, Inches(0.42), NAVY_D)
    _set_tf(push_head, "作業後は GitHub にプッシュして保存しましょう", size=12, bold=True, color=WHITE,
            ml=Inches(0.2), mt=Inches(0.09))
    _code_block(sl, ML, CY + Inches(3.7), CW, Inches(0.85),
                'git add .\ngit commit -m "変更内容の説明"\ngit push')
    tip11 = _rbox(sl, ML, H - FH - Inches(0.52), CW, Inches(0.38), ACCENT2, NAVY_L, line_w=0.8)
    _set_tf(tip11,
            "\U0001f4a1  Claude Code は変更履歴（git diff）も読んでいます。「直前の変更を元に戻して」などと言うだけでOKです。",
            size=10, color=NAVY_D, ml=Inches(0.2), mt=Inches(0.08))

    # ─────────── Slide 12: Q&A ───────────
    sl = _slide(prs)
    _chrome(sl, "よくある質問（Q&A）", 12, TOTAL)
    qa_h12 = CH / 4
    for i, (qn_s, q, a) in enumerate([
        ("Q1", "GitHub と Claude Code は必ずセットで使う必要がありますか？",
         "いいえ。Claude Code はローカルフォルダ内でも動作します。ただし GitHub と組み合わせることで変更履歴の管理や他端末との共有が容易になります。"),
        ("Q2", "Claude Code の Pro プランは必須ですか？",
         "Free プランでも基本的な会話は可能ですが、Claude Code でのコーディングサポートをフル活用するには Pro プラン（$20/月）を推奨します。"),
        ("Q3", "VS Code 以外のエディタでも使えますか？",
         "VS Code と JetBrains IDE（IntelliJ など）に公式拡張があります。また CLI 版はターミナルがあれば任意の環境で使用できます。"),
        ("Q4", "コマンドプロンプトに慣れていないのですが大丈夫ですか？",
         "VS Code 拡張版であれば、コマンド入力はほぼ不要です。Git の操作も VS Code の「ソース管理」パネルから GUI で行えます。"),
    ]):
        qy12 = CY + qa_h12 * i
        badge12 = _oval(sl, ML, qy12 + qa_h12 * 0.07, Inches(0.55), Inches(0.55), NAVY)
        _set_tf(badge12, qn_s, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                ml=Inches(0.03), mt=Inches(0.09))
        qrow12 = _rbox(sl, ML + Inches(0.65), qy12 + qa_h12 * 0.06,
                       CW - Inches(0.65), qa_h12 * 0.4, ACCENT2, GRAY2)
        _set_tf(qrow12, q, size=11, bold=True, color=NAVY_D, ml=Inches(0.2), mt=Inches(0.08))
        arow12 = _rbox(sl, ML + Inches(0.65), qy12 + qa_h12 * 0.46,
                       CW - Inches(0.65), qa_h12 * 0.5, WHITE, GRAY2)
        _box(sl, ML + Inches(0.65), qy12 + qa_h12 * 0.46, Inches(0.06), qa_h12 * 0.5, NAVY_L)
        _set_tf(arow12, a, size=10, color=DARK, ml=Inches(0.22), mt=Inches(0.08))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
'''

with open('build_pptx.py', 'a', encoding='utf-8') as f:
    f.write(code)
print('Done')
