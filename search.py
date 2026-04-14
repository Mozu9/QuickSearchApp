import tkinter as tk
from tkinter import colorchooser, messagebox
from urllib.parse import quote
import subprocess
import configparser
import os
import winreg
import sys
import threading
import pystray
from PIL import Image

# ===== パス設定 =====
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

config_file = os.path.join(current_dir, 'config.ini')
config = configparser.ConfigParser()
config.optionxform = str

APP_NAME = "QuickSearchApp"
EXE_PATH = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
ICO_PATH = os.path.join(current_dir, 'icon.ico')

def _get_ico_path():
    """exe化時は一時展開フォルダ、通常実行時はスクリプトと同じフォルダを参照する"""
    if getattr(sys, 'frozen', False):
        # PyInstallerが--add-dataで埋め込んだファイルはここに展開される
        return os.path.join(sys._MEIPASS, 'icon.ico')
    return ICO_PATH

# ===== デフォルト値 =====
DEFAULT_SETTINGS = {
    'topmost':          'True',
    'opacity':          '0.9',
    'background_color': '#1e1e1e',
    'font_color':       '#00ddff',
    'font_size':        '9',
    'window_width':     '250',
    'window_height':    '400',
    'last_x':           '100',
    'last_y':           '100',
    'auto_start':       'False',
    'save_geometry':    'True',
}
DEFAULT_SITES = {}

# ===== 安全な型変換ヘルパー =====
def safe_int(value, default, min_val=None, max_val=None):
    """変換失敗・範囲外は default を返す"""
    try:
        v = int(float(str(value)))
        if min_val is not None:
            v = max(min_val, v)
        if max_val is not None:
            v = min(max_val, v)
        return v
    except (ValueError, TypeError):
        return default

def safe_float(value, default, min_val=None, max_val=None):
    """変換失敗・範囲外は default を返す"""
    try:
        v = float(str(value))
        if min_val is not None:
            v = max(min_val, v)
        if max_val is not None:
            v = min(max_val, v)
        return v
    except (ValueError, TypeError):
        return default

def safe_color(value, default):
    """#rrggbb 形式でなければ default を返す"""
    v = str(value).strip()
    if len(v) == 7 and v.startswith('#'):
        try:
            int(v[1:], 16)
            return v
        except ValueError:
            pass
    return default

def safe_bool(value, default):
    """True/False 文字列以外は default を返す"""
    v = str(value).strip().lower()
    if v == 'true':
        return True
    if v == 'false':
        return False
    return default

# ===== スタートアップ登録 =====
def set_startup(enabled):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{EXE_PATH}"')
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Startup error: {e}")

# ===== 設定の読み込み =====
def load_settings():
    """
    1. メモリにデフォルト値をセット
    2. config.ini を読み込んで上書き
      - ファイル構造が壊れていたら .bak に退避してデフォルトで起動
      - 個別の値が壊れていても safe_* 関数が後でフォールバックする
    """
    config['SETTINGS']     = DEFAULT_SETTINGS.copy()
    config['SEARCH_SITES'] = DEFAULT_SITES.copy()

    if not os.path.exists(config_file):
        # 初回起動：デフォルト値でファイルを新規生成
        _write_config()
        return

    try:
        config.read(config_file, encoding='utf-8')
    except configparser.Error as e:
        # ファイル構造レベルで壊れている場合
        backup = config_file + ".bak"
        try:
            os.replace(config_file, backup)
            backup_msg = f"\n壊れたファイルは {backup} に保存しました。"
        except OSError:
            backup_msg = ""

        messagebox.showwarning(
            "設定ファイルエラー",
            f"config.ini の読み込みに失敗しました。\n"
            f"デフォルト設定で起動します。{backup_msg}\n\n"
            f"詳細: {e}"
        )
        # デフォルト値でメモリをリセットして新規ファイル生成
        config['SETTINGS']     = DEFAULT_SETTINGS.copy()
        config['SEARCH_SITES'] = DEFAULT_SITES.copy()
        _write_config()
        return

    # SETTINGS セクションが丸ごと消えていた場合の補完
    if 'SETTINGS' not in config:
        config['SETTINGS'] = DEFAULT_SETTINGS.copy()

    # SEARCH_SITES セクションが丸ごと消えていた場合の補完
    if 'SEARCH_SITES' not in config:
        config['SEARCH_SITES'] = DEFAULT_SITES.copy()

def save_settings():
    if config.getboolean('SETTINGS', 'save_geometry', fallback=True):
        try:
            config['SETTINGS']['last_x']        = str(root.winfo_x())
            config['SETTINGS']['last_y']         = str(root.winfo_y())
            config['SETTINGS']['window_width']   = str(root.winfo_width())
            config['SETTINGS']['window_height']  = str(root.winfo_height())
        except Exception:
            pass
    _write_config()

def _write_config():
    try:
        with open(config_file, 'w', encoding='utf-8') as f:
            config.write(f)
    except OSError as e:
        messagebox.showerror("保存エラー", f"config.ini の書き込みに失敗しました。\n{e}")

# ===== マルチモニター対応：座標検証 =====
def is_on_any_monitor(x, y, w, h):
    """ウィンドウ矩形がいずれかのモニターに少しでも重なっているか確認する"""
    try:
        import ctypes
        import ctypes.wintypes
        MONITOR_DEFAULTTONULL = 0
        rect = ctypes.wintypes.RECT(x, y, x + w, y + h)
        monitor = ctypes.windll.user32.MonitorFromRect(
            ctypes.byref(rect), MONITOR_DEFAULTTONULL
        )
        return monitor != 0
    except Exception:
        return True  # 取得失敗時はそのまま使う

# ===== ウィンドウ初期化 =====
root = tk.Tk()
root.overrideredirect(True)
load_settings()  # ← ここで壊れたファイルを検出してフォールバック済み

# safe_* で個別の値をすべて検証してから使う
s = config['SETTINGS']

bg_color  = safe_color(s.get('background_color'), '#1e1e1e')
fg_color  = safe_color(s.get('font_color'),        '#00ddff')
font_size = safe_int  (s.get('font_size'),    9,  min_val=6,   max_val=20)
opacity   = safe_float(s.get('opacity'),      0.9, min_val=0.1, max_val=1.0)
topmost   = safe_bool (s.get('topmost'),      True)

# メインモニターのサイズ（ウィンドウサイズの上限とフォールバック座標に使用）
screen_w = root.winfo_screenwidth()
screen_h = root.winfo_screenheight()

win_w = safe_int(s.get('window_width'),  250, min_val=150, max_val=screen_w)
win_h = safe_int(s.get('window_height'), 400, min_val=100, max_val=screen_h)

# 座標はクランプせず、どのモニターにも収まらない場合のみメイン画面中央にリセット
last_x = safe_int(s.get('last_x'), 100)
last_y = safe_int(s.get('last_y'), 100)
if not is_on_any_monitor(last_x, last_y, win_w, win_h):
    last_x = (screen_w - win_w) // 2
    last_y = (screen_h - win_h) // 2

root.geometry(f"{win_w}x{win_h}+{last_x}+{last_y}")
root.configure(bg=bg_color)
root.attributes("-alpha",   opacity)
root.attributes("-topmost", topmost)

# 検証済みの値を config に書き戻す（次回保存時に整合性を保つため）
config['SETTINGS']['background_color'] = bg_color
config['SETTINGS']['font_color']       = fg_color
config['SETTINGS']['font_size']        = str(font_size)
config['SETTINGS']['opacity']          = str(opacity)
config['SETTINGS']['topmost']          = str(topmost)
config['SETTINGS']['window_width']     = str(win_w)
config['SETTINGS']['window_height']    = str(win_h)
config['SETTINGS']['last_x']           = str(last_x)
config['SETTINGS']['last_y']           = str(last_y)

# ===== リサイズ =====
def start_resize(event):
    root._resize_start_x = event.x_root
    root._resize_start_y = event.y_root
    root._start_w = root.winfo_width()
    root._start_h = root.winfo_height()
    return "break"

def do_resize(event):
    dw = event.x_root - root._resize_start_x
    dh = event.y_root - root._resize_start_y
    new_w = max(150, root._start_w + dw)
    new_h = max(100, root._start_h + dh)
    root.geometry(f"{new_w}x{new_h}")
    return "break"

# ===== ウィジェット一覧 =====
labels  = []
entries = []

def get_bg():        return config['SETTINGS'].get('background_color', '#1e1e1e')
def get_fg():        return config['SETTINGS'].get('font_color',        '#00ddff')
def get_font_size(): return safe_int(config['SETTINGS'].get('font_size'), 9, 6, 20)

# ===== 貼り付けヘルパー =====
def paste_to_entry(widget):
    try:
        widget.insert(tk.INSERT, root.clipboard_get())
    except tk.TclError:
        pass

# ===== タスクトレイ =====
_tray_icon = None

def _load_tray_image():
    """アイコン画像を読み込む。icon.icoがなければ動的生成する"""
    ico = _get_ico_path()
    if os.path.exists(ico):
        return Image.open(ico).resize((64, 64))
    # フォールバック：シンプルな塗りつぶし円を生成
    img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    draw.ellipse([4, 4, 60, 60], fill='#00ddff')
    return img

def hide_to_tray():
    """ウィンドウを隠してタスクトレイに収納する"""
    global _tray_icon
    root.withdraw()  # ウィンドウを非表示

    tray_menu = pystray.Menu(
        pystray.MenuItem('表示に戻す', lambda icon, item: show_from_tray(icon)),
        pystray.MenuItem('設定を開く', lambda icon, item: root.after(0, open_settings)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('終了',       lambda icon, item: quit_app(icon)),
    )

    _tray_icon = pystray.Icon(
        APP_NAME,
        _load_tray_image(),
        APP_NAME,
        tray_menu
    )

    # pystray はブロッキングなので別スレッドで動かす
    threading.Thread(target=_tray_icon.run, daemon=True).start()

def show_from_tray(icon=None):
    """タスクトレイからウィンドウを復元する"""
    if icon is not None:
        icon.stop()
    root.after(0, root.deiconify)  # メインスレッドで復元

def quit_app(icon=None):
    """トレイアイコンから終了する"""
    if icon is not None:
        icon.stop()
    root.after(0, lambda: (save_settings(), root.destroy()))

# ===== タイトルバー（最小化ボタン付き） =====
def build_titlebar():
    bg = get_bg()
    fg = get_fg()

    bar = tk.Frame(root, bg=bg, height=22)
    bar.pack(fill="x", side="bottom")
    bar.pack_propagate(False)

    # アプリ名（ドラッグ移動の起点にもなる）
    title_lbl = tk.Label(bar, text=APP_NAME, bg=bg, fg=fg, font=("Arial", 7))
    title_lbl.pack(side="left", padx=6)

    # 最小化ボタン（トレイに収納）
    min_btn = tk.Button(
        bar, text="↓",
        bg=bg, fg=fg,
        relief="flat", cursor="hand2",
        font=("Arial", 10, "bold"),
        padx=6, pady=0,
        command=hide_to_tray
    )
    min_btn.pack(side="right", padx=4)

    # タイトルバー自体もドラッグで移動できるように
    bar.bind("<Button-1>",  on_press)
    bar.bind("<B1-Motion>", on_drag)
    title_lbl.bind("<Button-1>",  on_press)
    title_lbl.bind("<B1-Motion>", on_drag)


def rebuild_ui():
    global labels, entries

    for widget in container.winfo_children():
        widget.destroy()
    labels.clear()
    entries.clear()

    bg    = get_bg()
    fg    = get_fg()
    fs    = get_font_size()
    sites = dict(config['SEARCH_SITES'])

    # サイト未登録時のガイド
    if not sites:
        guide = tk.Label(
            container,
            text="右クリック →「設定を表示」\nからサイトを追加してください",
            bg=bg, fg=fg,
            font=("Arial", fs),
            justify="center"
        )
        guide.pack(expand=True)
        labels.append(guide)
        grip.lift()
        return

    for site_name, url in sites.items():
        lbl = tk.Label(
            container, text=site_name,
            bg=bg, fg=fg,
            font=("Arial", fs, "bold")
        )
        lbl.pack(pady=(8, 0), padx=12, anchor="w")
        labels.append(lbl)

        # 検索窓と→ボタンを横並び
        row = tk.Frame(container, bg=bg)
        row.pack(padx=12, pady=(2, 6), fill="x")

        e = tk.Entry(
            row, bg="#2d2d2d", fg="white",
            insertbackground="white", relief="flat",
            font=("Arial", fs)
        )
        e.pack(side="left", fill="x", expand=True)
        entries.append(e)

        # 検索窓の右クリック：貼り付けのみ
        entry_menu = tk.Menu(root, tearoff=0)
        entry_menu.add_command(label="貼り付け", command=lambda w=e: paste_to_entry(w))
        e.bind("<Button-3>", lambda event, m=entry_menu: m.post(event.x_root, event.y_root))

        def make_search(entry_widget, base_url):
            def do_search(event=None):
                query = entry_widget.get().strip()
                if not query:
                    return
                subprocess.Popen(f'start "" "{base_url + quote(query)}"', shell=True)
                entry_widget.delete(0, tk.END)
                entry_widget.focus_set()
            return do_search

        search_fn = make_search(e, url)
        e.bind("<Return>", search_fn)

        btn = tk.Button(
            row, text="→",
            bg="#2d2d2d", fg=fg,
            relief="flat", cursor="hand2",
            padx=4, pady=0,
            font=("Arial", fs),
            command=search_fn
        )
        btn.pack(side="left", padx=(2, 0))

    grip.lift()

# ===== コンテナ＆グリップ =====
container = tk.Frame(root, bg=get_bg())
container.pack(fill="both", expand=True)

grip = tk.Frame(root, bg=get_bg(), cursor="size_nw_se")
grip.place(relx=1.0, rely=1.0, anchor="se", width=16, height=16)
grip.bind("<Button-1>", start_resize)
grip.bind("<B1-Motion>", do_resize)

# ===== ウィンドウ移動 =====
def on_press(event):
    root._drag_x = event.x
    root._drag_y = event.y

def on_drag(event):
    dx = event.x - root._drag_x
    dy = event.y - root._drag_y
    root.geometry(f"+{root.winfo_x() + dx}+{root.winfo_y() + dy}")

root.bind("<Button-1>", on_press)
root.bind("<B1-Motion>", on_drag)

# ===== 設定画面 =====
_settings_win = None

def open_settings():
    global _settings_win
    if _settings_win is not None and _settings_win.winfo_exists():
        _settings_win.lift()
        _settings_win.focus_force()
        return

    _settings_win = tk.Toplevel(root)
    _settings_win.title("アプリ設定")
    _settings_win.geometry("340x680")
    _settings_win.attributes("-topmost", True)
    _settings_win.grab_set()

    pad = {"padx": 20, "pady": 4}

    tk.Label(_settings_win, text="透過率").pack(pady=(12, 0))
    s_opacity = tk.Scale(
        _settings_win, from_=0.1, to=1.0, resolution=0.05, orient="horizontal",
        command=lambda v: (root.attributes("-alpha", float(v)),
                           config.set('SETTINGS', 'opacity', str(v)))
    )
    s_opacity.set(safe_float(config['SETTINGS'].get('opacity'), 0.9, 0.1, 1.0))
    s_opacity.pack(fill="x", **pad)

    tk.Label(_settings_win, text="文字サイズ").pack(pady=(10, 0))
    s_font = tk.Scale(
        _settings_win, from_=6, to=20, orient="horizontal",
        command=lambda v: apply_font_size(int(float(v)))
    )
    s_font.set(get_font_size())
    s_font.pack(fill="x", **pad)

    tk.Label(_settings_win, text="横幅").pack(pady=(10, 0))
    s_width = tk.Scale(
        _settings_win, from_=150, to=1000, orient="horizontal",
        command=lambda v: (root.geometry(f"{int(float(v))}x{root.winfo_height()}"),
                           config.set('SETTINGS', 'window_width', str(int(float(v)))))
    )
    s_width.set(root.winfo_width())
    s_width.pack(fill="x", **pad)

    tk.Label(_settings_win, text="高さ").pack(pady=(10, 0))
    s_height = tk.Scale(
        _settings_win, from_=100, to=1200, orient="horizontal",
        command=lambda v: (root.geometry(f"{root.winfo_width()}x{int(float(v))}"),
                           config.set('SETTINGS', 'window_height', str(int(float(v)))))
    )
    s_height.set(root.winfo_height())
    s_height.pack(fill="x", **pad)

    tk.Button(_settings_win, text="背景色を選択", command=lambda: pick_color("bg")).pack(fill="x", **pad)
    tk.Button(_settings_win, text="文字色を選択", command=lambda: pick_color("fg")).pack(fill="x", **pad)

    top_var = tk.BooleanVar(value=safe_bool(config['SETTINGS'].get('topmost'), True))
    tk.Checkbutton(
        _settings_win, text="常に最前面に表示", variable=top_var,
        command=lambda: (root.attributes("-topmost", top_var.get()),
                         config.set('SETTINGS', 'topmost', str(top_var.get())))
    ).pack(anchor="w", padx=20, pady=2)

    geo_var = tk.BooleanVar(value=safe_bool(config['SETTINGS'].get('save_geometry'), True))
    tk.Checkbutton(
        _settings_win, text="終了時の位置・サイズを記憶", variable=geo_var,
        command=lambda: config.set('SETTINGS', 'save_geometry', str(geo_var.get()))
    ).pack(anchor="w", padx=20, pady=2)

    start_var = tk.BooleanVar(value=safe_bool(config['SETTINGS'].get('auto_start'), False))
    tk.Checkbutton(
        _settings_win, text="Windows起動時に自動実行", variable=start_var,
        command=lambda: (set_startup(start_var.get()),
                         config.set('SETTINGS', 'auto_start', str(start_var.get())))
    ).pack(anchor="w", padx=20, pady=2)

    # サイト管理
    tk.Frame(_settings_win, height=1, bg="#aaaaaa").pack(fill="x", padx=20, pady=10)
    tk.Label(_settings_win, text="検索サイト管理", font=("Arial", 9, "bold")).pack()

    btn_frame = tk.Frame(_settings_win)
    btn_frame.pack(fill="x", padx=20, pady=4)
    tk.Button(btn_frame, text="＋ サイトを追加", command=add_site_dialog).pack(side="left", expand=True, fill="x", padx=(0, 4))
    tk.Button(btn_frame, text="－ サイトを削除", command=remove_site_dialog).pack(side="left", expand=True, fill="x")

    tk.Button(
        _settings_win, text="設定を保存して閉じる",
        bg="#00ddff", fg="#000000",
        command=lambda: (save_settings(), _settings_win.destroy())
    ).pack(pady=14, fill="x", padx=20)

    # 投げ銭ボタン（設定画面最下部）
    tk.Frame(_settings_win, height=1, bg="#aaaaaa").pack(fill="x", padx=20, pady=(0, 8))
    tk.Button(
        _settings_win, text="☕ 作者に投げ銭する（OFUSEへ）",
        bg="#f5a623", fg="#ffffff",
        relief="flat", cursor="hand2",
        command=open_donate
    ).pack(pady=(0, 16), fill="x", padx=20)

def apply_font_size(size):
    config.set('SETTINGS', 'font_size', str(size))
    for lbl in labels:
        lbl.configure(font=("Arial", size, "bold"))
    for e in entries:
        e.configure(font=("Arial", size))

def pick_color(target):
    color = colorchooser.askcolor(title="色を選択")[1]
    if not color:
        return
    if target == "bg":
        config.set('SETTINGS', 'background_color', color)
        root.configure(bg=color)
        container.configure(bg=color)
        grip.configure(bg=color)
        for lbl in labels:
            lbl.configure(bg=color)
    else:
        config.set('SETTINGS', 'font_color', color)
        for lbl in labels:
            lbl.configure(fg=color)

DONATE_URL = "https://ofuse.me/d4dfbffb"

def open_donate():
    import webbrowser
    webbrowser.open(DONATE_URL)

def open_about():
    win = tk.Toplevel(root)
    win.title("このアプリについて")
    win.geometry("300x220")
    win.attributes("-topmost", True)
    win.grab_set()
    win.resizable(False, False)

    tk.Label(win, text="QuickSearchApp", font=("Arial", 13, "bold")).pack(pady=(20, 4))
    tk.Label(win, text="常駐型クイック検索ランチャー", font=("Arial", 9)).pack()
    tk.Label(win, text="右クリックで設定・サイト管理ができます。\nconfig.ini を編集してカスタマイズも可能です。",
             font=("Arial", 8), justify="center", fg="#555555").pack(pady=(8, 0))

    tk.Frame(win, height=1, bg="#cccccc").pack(fill="x", padx=20, pady=12)

    tk.Label(win, text="気に入ったら投げ銭していただけると\n作者の励みになります 🙏",
             font=("Arial", 9), justify="center").pack()

    tk.Button(
        win, text="☕ 投げ銭する（OFUSEへ）",
        bg="#f5a623", fg="#ffffff",
        relief="flat", cursor="hand2",
        command=open_donate
    ).pack(pady=10, fill="x", padx=30)

# ===== サイト追加 =====
def add_site_dialog():
    win = tk.Toplevel(root)
    win.title("サイトを追加")
    win.geometry("320x180")
    win.attributes("-topmost", True)
    win.grab_set()

    tk.Label(win, text="表示名").pack(pady=(12, 0))
    name_entry = tk.Entry(win)
    name_entry.pack(padx=20, fill="x")
    name_entry.focus_set()

    tk.Label(win, text="検索URL（末尾に検索語が付きます）").pack(pady=(8, 0))
    url_entry = tk.Entry(win)
    url_entry.insert(0, "https://example.com/search?q=")
    url_entry.pack(padx=20, fill="x")

    def on_add():
        name = name_entry.get().strip()
        url  = url_entry.get().strip()
        if not name or not url:
            messagebox.showwarning("入力エラー", "名前とURLを両方入力してください", parent=win)
            return
        if name in config['SEARCH_SITES']:
            messagebox.showwarning("重複", f"「{name}」はすでに存在します", parent=win)
            return
        config['SEARCH_SITES'][name] = url
        rebuild_ui()
        win.destroy()

    tk.Button(win, text="追加", bg="#00ddff", fg="#000000", command=on_add).pack(pady=12)
    win.bind("<Return>", lambda e: on_add())

# ===== サイト削除 =====
def remove_site_dialog():
    sites = list(config['SEARCH_SITES'].keys())
    if not sites:
        messagebox.showinfo("情報", "削除できるサイトがありません")
        return

    win = tk.Toplevel(root)
    win.title("サイトを削除")
    win.geometry("280x300")
    win.attributes("-topmost", True)
    win.grab_set()

    tk.Label(win, text="削除するサイトを選択").pack(pady=(12, 4))

    listbox = tk.Listbox(win, selectmode="single")
    listbox.pack(fill="both", expand=True, padx=20, pady=4)
    for s in sites:
        listbox.insert(tk.END, s)

    def on_delete():
        sel = listbox.curselection()
        if not sel:
            return
        name = sites[sel[0]]
        if messagebox.askyesno("確認", f"「{name}」を削除しますか？", parent=win):
            del config['SEARCH_SITES'][name]
            rebuild_ui()
            win.destroy()

    tk.Button(win, text="削除", bg="#ff4444", fg="white", command=on_delete).pack(pady=8)

# ===== 右クリックメニュー（アプリ本体） =====
menu = tk.Menu(root, tearoff=0)
menu.add_command(label="設定を表示",     command=open_settings)
menu.add_command(label="トレイに収納",   command=hide_to_tray)
menu.add_command(label="config.ini を開く", command=lambda: subprocess.Popen(f'notepad "{config_file}"', shell=True))
menu.add_separator()
menu.add_command(label="このアプリについて", command=open_about)
menu.add_separator()
menu.add_command(label="終了", command=lambda: (save_settings(), root.destroy()))

root.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))

# ===== 初回描画 =====
build_titlebar()
rebuild_ui()

root.mainloop()
