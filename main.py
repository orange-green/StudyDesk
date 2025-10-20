#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ...existing code...
import multiprocessing
import tkinter as tk
from tkinter import font, colorchooser
import json
import threading
import time
import tempfile
import os
import random
import sys
from functools import partial
from tkinter import messagebox
from PIL import Image, ImageDraw
from pynput import mouse, keyboard
import pystray
from pystray import Menu as menu, MenuItem as item
from datetime import datetime

from dictionary import all_dicts

APP_TITLE = 'StudyDesk'

shuffle_mode = {'value': False}
topmost_timer = {'enabled': True}
pronunciation_type = {'value': 1}  # 1: 英音，2: 美音

# 可视化设置（颜色和字号），会保存到配置文件
visual_settings = {
    'word_color': 'blue',
    'meaning_color': 'black',
    'word_size': 24,
    'meaning_size': 12
}

# 热键默认设置（用户可在设置界面修改）
DEFAULT_HOTKEYS = {
    'toggle': '<ctrl>+<alt>+s',
    'settings': '<ctrl>+<alt>+o',
    'next': '<ctrl>+<alt>+right',
    'prev': '<ctrl>+<alt>+left',
    'known': '<ctrl>+<alt>+k',
    'forgot': '<ctrl>+<alt>+f'
}
user_hotkeys = DEFAULT_HOTKEYS.copy()

# 复习数据文件（按词条统计熟记/忘记次数）
REVIEW_DATA_PATH = os.path.join(tempfile.gettempdir(), 'study_desk_review.json')
review_data = {}

audio_cache_dir = os.path.join(tempfile.gettempdir(), "study_desk_cache")
os.makedirs(audio_cache_dir, exist_ok=True)

CONFIG_FILE_PATH = os.path.join(tempfile.gettempdir(), 'study_desk_config.json')
current_dict_path = {'value': ''}

# 全局托盘图标句柄（可能在单独线程中创建）
icon = None

# 全局 hotkey listener handle (pynput.GlobalHotKeys)
gh_listener = None

def save_review_data():
    try:
        with open(REVIEW_DATA_PATH, 'w', encoding='utf-8') as f:
            json.dump(review_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存复习数据失败: {e}")

def load_review_data():
    global review_data
    if not os.path.exists(REVIEW_DATA_PATH):
        review_data = {}
        return
    try:
        with open(REVIEW_DATA_PATH, 'r', encoding='utf-8') as f:
            review_data = json.load(f) or {}
    except Exception as e:
        print(f"加载复习数据失败: {e}")
        review_data = {}

def save_config():
    config = {
        'pronunciation_type': pronunciation_type['value'],
        'shuffle_mode': shuffle_mode['value'],
        'topmost_enabled': topmost_timer['enabled'],
        'current_dict': current_dict_path.get('value', ''),
        'current_index': app.index if 'app' in globals() else 0,
        'visual_settings': visual_settings,
        'hotkeys': user_hotkeys
    }
    try:
        with open(CONFIG_FILE_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"配置保存失败: {e}")

def load_config():
    if not os.path.exists(CONFIG_FILE_PATH):
        return {}
    try:
        with open(CONFIG_FILE_PATH, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
            # merge visual settings if present
            if 'visual_settings' in cfg:
                visual_settings.update(cfg.get('visual_settings') or {})
            # load user hotkeys if present
            if 'hotkeys' in cfg:
                user_hotkeys.update(cfg.get('hotkeys') or {})
            return cfg
    except Exception as e:
        print(f"配置加载失败: {e}")
        return {}

def play_pronunciation(word):
    if not word:
        return
    from playsound import playsound
    import requests
    import re

    def clean_text(input_str):
        if re.match(r'^\s*[a-zA-Z_]\w*\s*\(', input_str):
            func_name = re.sub(r'(\w+)\s*\(.*\)', r'\1', input_str)
            return func_name.strip()
        else:
            cleaned_str = re.sub(r'[^\w\s\u4e00-\u9fff]', ' ', input_str)
            cleaned_str = re.sub(r'\s+', ' ', cleaned_str).strip()
            return cleaned_str

    word = clean_text(word)
    type_id = pronunciation_type['value']
    filename = f"{word}_{type_id}.mp3"
    file_path = os.path.join(audio_cache_dir, filename)

    if not os.path.exists(file_path):
        url = f"https://dict.youdao.com/dictvoice?audio={word}&type={type_id}"
        try:
            response = requests.get(url, timeout=5)
            response.raise_for_status()
            with open(file_path, 'wb') as f:
                f.write(response.content)
        except Exception as e:
            print(f"下载音频失败: {e}")
            return

    try:
        # block=False 如果 playsound 后台播放在某些系统不稳定，可改为 True
        playsound(os.path.abspath(file_path), block=False)
    except Exception as e:
        print(f"播放出错：{e}")

def mark_word_known(word_name):
    if not word_name:
        return
    entry = review_data.get(word_name, {'known': 0, 'forgot': 0, 'last': None})
    entry['known'] = entry.get('known', 0) + 1
    entry['last'] = datetime.utcnow().isoformat()
    review_data[word_name] = entry
    save_review_data()

def mark_word_forgot(word_name):
    if not word_name:
        return
    entry = review_data.get(word_name, {'known': 0, 'forgot': 0, 'last': None})
    entry['forgot'] = entry.get('forgot', 0) + 1
    entry['last'] = datetime.utcnow().isoformat()
    review_data[word_name] = entry
    save_review_data()

class TransparentWordWindow:
    def __init__(self, root, words, start_index=0):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.overrideredirect(True)
        self.root.wm_attributes("-topmost", True)
        # 使白色透明（Windows/pyinstaller 下通常可用）
        try:
            self.root.wm_attributes("-transparentcolor", "white")
        except Exception:
            pass
        self.root.configure(bg='white')

        self.canvas = tk.Canvas(root, bg='white', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 使用 visual_settings 初始化字体与颜色
        self.font_word = font.Font(family='Helvetica', size=visual_settings.get('word_size', 24), weight='bold')
        self.font_phone = font.Font(family='Helvetica', size=visual_settings.get('meaning_size', 12), slant='italic')
        self.font_trans = font.Font(family='Helvetica', size=visual_settings.get('meaning_size', 12), weight='bold')

        self.word_text = self.canvas.create_text(10, 10, anchor='nw', text='', font=self.font_word, fill=visual_settings.get('word_color', 'blue'))
        self.phone_text = self.canvas.create_text(10, 50, anchor='nw', text='', font=self.font_phone, fill='gray')
        self.trans_text = self.canvas.create_text(10, 80, anchor='nw', text='', font=self.font_trans, fill=visual_settings.get('meaning_color', 'black'))
        self.status_text = self.canvas.create_text(10, 110, anchor='nw', text='', font=font.Font(size=10), fill='green')

        self.words = words or []
        self.index = start_index if 0 <= start_index < len(self.words) else 0

        self.update_display()

        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<ButtonPress-1>", self.start_move)
        self.canvas.bind("<B1-Motion>", self.do_move)
        self.canvas.tag_bind(self.word_text, "<Button-1>", self.on_word_click)

        start_scroll_listener(self.show_prev_word, self.show_next_word, interval=0.2)
        set_window_topmost(self.root)

    def toggle_visible(self):
        if self.root.state() == 'normal':
            self.hide()
        else:
            self.show()

    def on_word_click(self, event):
        current_word = self.words[self.index].get('name', '') if self.words else ''
        play_pronunciation(current_word)

    def update_display(self):
        if not self.words:
            self.canvas.itemconfig(self.word_text, text='（词典为空）')
            self.canvas.itemconfig(self.phone_text, text='')
            self.canvas.itemconfig(self.trans_text, text='')
            self.canvas.itemconfig(self.status_text, text='')
            return

        word = self.words[self.index]
        name = word.get('name', '')
        self.canvas.itemconfig(self.word_text, text=name)
        self.canvas.itemconfig(
            self.phone_text,
            text=f"US: /{word.get('usphone', '')}/   UK: /{word.get('ukphone', '')}/"
        )
        self.canvas.itemconfig(
            self.trans_text,
            text='；'.join(word.get('trans', []))
        )

        # 显示复习标记
        entry = review_data.get(name)
        if entry:
            status = f"熟记:{entry.get('known',0)} 忘记:{entry.get('forgot',0)}"
            self.canvas.itemconfig(self.status_text, text=status, fill='green')
        else:
            self.canvas.itemconfig(self.status_text, text='未标记', fill='gray')

        # 异步播放发音以避免阻塞 UI
        threading.Thread(target=play_pronunciation, args=(name,), daemon=True).start()

    def on_mousewheel(self, event):
        if not self.words:
            return
        self.index = (self.index - 1) % len(self.words) if event.delta > 0 else (self.index + 1) % len(self.words)
        self.update_display()
        save_config()

    def update_word_list(self, new_list, dict_path=''):
        self.words = new_list or []
        self.index = 0
        current_dict_path['value'] = dict_path
        self.update_display()
        save_config()

    def show(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes('-topmost', True)
        except Exception:
            pass

    def hide(self):
        try:
            self.root.withdraw()
        except Exception:
            pass

    def quit(self):
        save_config()
        try:
            if 'icon' in globals() and icon:
                icon.stop()
        except Exception:
            pass
        self.root.destroy()

    def start_move(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def do_move(self, event):
        x = self.root.winfo_pointerx() - self._drag_start_x
        y = self.root.winfo_pointery() - self._drag_start_y
        self.root.geometry(f'+{x}+{y}')

    def show_next_word(self):
        if not self.words:
            return
        if shuffle_mode['value'] and len(self.words) > 1:
            new_index = self.index
            while new_index == self.index:
                new_index = random.randint(0, len(self.words) - 1)
            self.index = new_index
        else:
            self.index = (self.index + 1) % len(self.words)
        self.update_display()
        save_config()

    def show_prev_word(self):
        if not self.words:
            return
        if shuffle_mode['value'] and len(self.words) > 1:
            new_index = self.index
            while new_index == self.index:
                new_index = random.randint(0, len(self.words) - 1)
            self.index = new_index
        else:
            self.index = (self.index - 1) % len(self.words)
        self.update_display()
        save_config()

    def apply_visual_settings(self):
        # 更新字体与颜色并刷新显示
        self.font_word.configure(size=visual_settings.get('word_size', 24))
        self.font_phone.configure(size=visual_settings.get('meaning_size', 12))
        self.font_trans.configure(size=visual_settings.get('meaning_size', 12))
        self.canvas.itemconfig(self.word_text, fill=visual_settings.get('word_color', 'blue'), font=self.font_word)
        self.canvas.itemconfig(self.phone_text, font=self.font_phone)
        self.canvas.itemconfig(self.trans_text, fill=visual_settings.get('meaning_color', 'black'), font=self.font_trans)
        self.update_display()

    def mark_current_known(self):
        if not self.words:
            return
        name = self.words[self.index].get('name', '')
        mark_word_known(name)
        self.update_display()

    def mark_current_forgot(self):
        if not self.words:
            return
        name = self.words[self.index].get('name', '')
        mark_word_forgot(name)
        self.update_display()

    def open_settings(self):
        # 在主线程打开设置窗口（可能从托盘线程调用）
        def _open():
            dlg = tk.Toplevel(self.root)
            dlg.title("设置")
            dlg.transient(self.root)
            dlg.grab_set()
            dlg.resizable(False, False)

            # 视觉设置区域（保留原有）
            frm_visual = tk.LabelFrame(dlg, text="视觉设置")
            frm_visual.grid(row=0, column=0, padx=6, pady=6, sticky='ew')
            tk.Label(frm_visual, text="单词颜色:").grid(row=0, column=0, padx=6, pady=6, sticky='e')
            word_color_var = tk.StringVar(value=visual_settings.get('word_color'))
            entry_word_color = tk.Entry(frm_visual, textvariable=word_color_var, width=12)
            entry_word_color.grid(row=0, column=1, padx=6, pady=6)
            def pick_word_color():
                c = colorchooser.askcolor(title="选择单词颜色", initialcolor=word_color_var.get())
                if c and c[1]:
                    word_color_var.set(c[1])
            tk.Button(frm_visual, text="选择", command=pick_word_color).grid(row=0, column=2, padx=6)

            tk.Label(frm_visual, text="意思颜色:").grid(row=1, column=0, padx=6, pady=6, sticky='e')
            meaning_color_var = tk.StringVar(value=visual_settings.get('meaning_color'))
            entry_meaning_color = tk.Entry(frm_visual, textvariable=meaning_color_var, width=12)
            entry_meaning_color.grid(row=1, column=1, padx=6, pady=6)
            def pick_meaning_color():
                c = colorchooser.askcolor(title="选择意思颜色", initialcolor=meaning_color_var.get())
                if c and c[1]:
                    meaning_color_var.set(c[1])
            tk.Button(frm_visual, text="选择", command=pick_meaning_color).grid(row=1, column=2, padx=6)

            tk.Label(frm_visual, text="单词字号:").grid(row=2, column=0, padx=6, pady=6, sticky='e')
            word_size_var = tk.IntVar(value=visual_settings.get('word_size'))
            tk.Spinbox(frm_visual, from_=8, to=96, textvariable=word_size_var, width=6).grid(row=2, column=1, padx=6, pady=6, sticky='w')

            tk.Label(frm_visual, text="意思字号:").grid(row=3, column=0, padx=6, pady=6, sticky='e')
            meaning_size_var = tk.IntVar(value=visual_settings.get('meaning_size'))
            tk.Spinbox(frm_visual, from_=8, to=96, textvariable=meaning_size_var, width=6).grid(row=3, column=1, padx=6, pady=6, sticky='w')

            # 热键设置区域
            frm_hotkeys = tk.LabelFrame(dlg, text="快捷键设置（示例格式: ctrl+alt+s 或 right/left/up/down）")
            frm_hotkeys.grid(row=1, column=0, padx=6, pady=6, sticky='ew')

            def add_hotkey_row(row, label_text, key_var):
                tk.Label(frm_hotkeys, text=label_text).grid(row=row, column=0, padx=6, pady=4, sticky='e')
                ent = tk.Entry(frm_hotkeys, textvariable=key_var, width=22)
                ent.grid(row=row, column=1, padx=6, pady=4)
                def capture():
                    ent.delete(0, tk.END)
                    ent.insert(0, "按组合键...")
                    def _capture():
                        captured = capture_hotkey_blocking()
                        if captured:
                            key_var.set(captured)
                    threading.Thread(target=_capture, daemon=True).start()
                tk.Button(frm_hotkeys, text="捕获", command=capture, width=6).grid(row=row, column=2, padx=4)

            hk_toggle = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('toggle', DEFAULT_HOTKEYS['toggle'])))
            hk_settings = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('settings', DEFAULT_HOTKEYS['settings'])))
            hk_next = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('next', DEFAULT_HOTKEYS['next'])))
            hk_prev = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('prev', DEFAULT_HOTKEYS['prev'])))
            hk_known = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('known', DEFAULT_HOTKEYS['known'])))
            hk_forgot = tk.StringVar(value=humanize_hotkey(user_hotkeys.get('forgot', DEFAULT_HOTKEYS['forgot'])))

            add_hotkey_row(0, "切换 显示/隐藏:", hk_toggle)
            add_hotkey_row(1, "打开 设置页:", hk_settings)
            add_hotkey_row(2, "下一个 词:", hk_next)
            add_hotkey_row(3, "上一个 词:", hk_prev)
            add_hotkey_row(4, "标记 熟记:", hk_known)
            add_hotkey_row(5, "标记 忘记:", hk_forgot)

            # 复习操作按钮
            frm_review = tk.LabelFrame(dlg, text="当前词操作")
            frm_review.grid(row=2, column=0, padx=6, pady=6, sticky='ew')
            tk.Button(frm_review, text="熟记", command=lambda: [self.mark_current_known(), messagebox.showinfo("已标记", "已标记为熟记")]).grid(row=0, column=0, padx=6, pady=6)
            tk.Button(frm_review, text="忘记", command=lambda: [self.mark_current_forgot(), messagebox.showinfo("已标记", "已标记为忘记")]).grid(row=0, column=1, padx=6, pady=6)

            def on_save():
                visual_settings['word_color'] = word_color_var.get()
                visual_settings['meaning_color'] = meaning_color_var.get()
                visual_settings['word_size'] = int(word_size_var.get())
                visual_settings['meaning_size'] = int(meaning_size_var.get())

                # 保存热键 (normalize from human form to pynput form)
                try:
                    user_hotkeys['toggle'] = normalize_hotkey_string(hk_toggle.get())
                    user_hotkeys['settings'] = normalize_hotkey_string(hk_settings.get())
                    user_hotkeys['next'] = normalize_hotkey_string(hk_next.get())
                    user_hotkeys['prev'] = normalize_hotkey_string(hk_prev.get())
                    user_hotkeys['known'] = normalize_hotkey_string(hk_known.get())
                    user_hotkeys['forgot'] = normalize_hotkey_string(hk_forgot.get())
                except Exception as e:
                    print(f"热键格式化失败: {e}")

                self.apply_visual_settings()
                save_config()
                restart_hotkeys_listener(self)
                dlg.destroy()

            tk.Button(dlg, text="保存", command=on_save).grid(row=3, column=0, padx=6, pady=10, sticky='w')
            tk.Button(dlg, text="取消", command=dlg.destroy).grid(row=3, column=1, padx=6, pady=10, sticky='e')
            dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

        # 确保在主线程创建对话框
        self.root.after(0, _open)

def humanize_hotkey(pynput_style):
    # 将 '<ctrl>+<alt>+s' -> 'ctrl+alt+s' 便于显示/编辑
    try:
        parts = pynput_style.split('+')
        parts = [p.strip().strip('<>').lower() for p in parts if p]
        return '+'.join(parts)
    except Exception:
        return pynput_style

def normalize_hotkey_string(user_input):
    # 将 'ctrl+alt+s' 或 'ctrl+alt+right' 等 -> '<ctrl>+<alt>+s'
    if not user_input:
        return ''
    parts = [p.strip().lower() for p in user_input.replace(' ', '').split('+') if p.strip()]
    res = []
    modifiers = {'ctrl', 'alt', 'shift'}
    special_keys = {'right','left','up','down','enter','space','tab','esc','delete','home','end','pageup','pagedown'}
    for p in parts:
        if p in modifiers or p in special_keys:
            res.append(f"<{p}>")
        else:
            # single character or word -> use as-is (letters should be lower)
            res.append(p)
    return '+'.join(res)

def capture_hotkey_blocking(timeout=8):
    # 在单独线程中调用：阻塞等待用户按下一次组合键并返回 humanized form 'ctrl+alt+s'
    pressed = set()
    captured_result = {'value': None}
    finished = threading.Event()

    def on_press(key):
        try:
            if hasattr(key, 'char') and key.char:
                pressed.add(key.char.lower())
            else:
                pressed.add(str(key).replace('Key.', '').lower())
        except Exception:
            pass

    def on_release(key):
        # 构造组合（优先 ctrl/alt/shift，然后主键）
        mods = []
        main = None
        for p in pressed:
            if p in ('ctrl', 'ctrl_l', 'ctrl_r', 'lctrl', 'rctrl'):
                mods.append('ctrl')
            elif p in ('alt', 'alt_l', 'alt_r'):
                mods.append('alt')
            elif p in ('shift', 'shift_l', 'shift_r'):
                mods.append('shift')
            elif p in ('left', 'right', 'up', 'down', 'enter', 'space', 'tab', 'esc'):
                main = p
            else:
                # letters or others -> choose as main if not yet set
                if not main:
                    main = p
        if not main and mods:
            main = mods.pop()  # fallback
        combo = '+'.join(mods + ([main] if main else []))
        if combo:
            captured_result['value'] = combo
        finished.set()
        return False

    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        finished.wait(timeout=timeout)
        try:
            listener.stop()
        except Exception:
            pass
    return captured_result['value']

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_dict_from_file(path):
    try:
        with open(resource_path(path), 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"词典加载失败: {path} - {e}")
        return []

def create_image():
    img = Image.new('RGB', (64, 64), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([16, 16, 48, 48], fill='black')
    return img

def create_tray_icon(app_window):
    def on_show(icon_obj, item): app_window.show()
    def on_hide(icon_obj, item): app_window.hide()
    def on_quit(icon_obj, item): app_window.quit()

    def set_pronunciation(value):
        pronunciation_type['value'] = value
        save_config()

    def set_shuffle_mode(mode):
        shuffle_mode['value'] = mode
        save_config()

    def load_and_update_dict(path, icon_obj=None, item=None):
        new_words = load_dict_from_file(path)
        app_window.update_word_list(new_words, path)

    def toggle_remove_topmost(icon_obj, item):
        topmost_timer['enabled'] = not topmost_timer['enabled']
        save_config()

    def open_about_page(icon_obj=None, item=None):
        import webbrowser
        webbrowser.open("https://github.com/muieay")

    def load_custom_icon():
        icon_path = resource_path("logo.png")
        if os.path.exists(icon_path):
            try:
                return Image.open(icon_path).resize((64, 64))
            except Exception:
                pass
        return create_image()

    def build_dict_menu():
        def make_item(entry):
            return item(entry["name"], partial(load_and_update_dict, entry["url"]))
        return [item(category, menu(*[make_item(e) for e in entries])) for category, entries in all_dicts.items()]

    menu_structure = menu(
        item('显示', on_show),
        item('隐藏', on_hide),
        item('关于', open_about_page),
        item('设置', lambda icon_obj, it: app_window.root.after(0, app_window.open_settings)),
        item('置顶', toggle_remove_topmost, checked=lambda item: topmost_timer['enabled']),
        item('词典', menu(*build_dict_menu())),
        item('发音', menu(
            item('英音', lambda icon_obj, it: set_pronunciation(1)),
            item('美音', lambda icon_obj, it: set_pronunciation(2))
        )),
        item('模式', menu(
            item('顺序', lambda icon_obj, it: set_shuffle_mode(False)),
            item('随机', lambda icon_obj, it: set_shuffle_mode(True))
        )),
        item('退出', on_quit)
    )

    global icon
    try:
        icon = pystray.Icon("WordDisplay", load_custom_icon(), "屏幕单词", menu_structure)
        icon.run()
    except Exception as e:
        print(f"托盘图标启动失败: {e}")

def start_scroll_listener(callback_scroll_up, callback_scroll_down, interval=0.2):
    last_time = [0]
    def on_scroll(x, y, dx, dy):
        now = time.time()
        if now - last_time[0] < interval:
            return
        last_time[0] = now
        if dy > 0:
            callback_scroll_up()
        elif dy < 0:
            callback_scroll_down()
    listener = mouse.Listener(on_scroll=on_scroll)
    listener.daemon = True
    listener.start()

def get_top_windows():
    try:
        import win32gui, win32con
    except Exception:
        return []
    top_windows = []
    def enum_windows_callback(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                if style & win32con.WS_EX_TOPMOST:
                    top_windows.append(hwnd)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except Exception:
        pass
    return top_windows

def unset_topmost(hwnd):
    try:
        import win32gui, win32con
        title = win32gui.GetWindowText(hwnd)
        if APP_TITLE not in title:
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                   win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    except Exception:
        pass

def set_window_topmost(root_window):
    def check_topmost():
        while getattr(root_window, '_keep_topmost', False):
            try:
                if not root_window.attributes('-topmost'):
                    root_window.attributes('-topmost', True)
                if topmost_timer['enabled']:
                    close_topmost_windows()
            except Exception:
                pass
            time.sleep(5)
    root_window._keep_topmost = True
    try:
        root_window.deiconify()
        root_window.attributes('-topmost', True)
    except Exception:
        pass
    threading.Thread(target=check_topmost, daemon=True).start()

def close_topmost_windows():
    for hwnd in get_top_windows():
        unset_topmost(hwnd)

def set_window_to_bottom_right(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = screen_width - width - 10
    y = screen_height - height - 10
    window.geometry(f'{width}x{height}+{x}+{y}')
# ...existing code...
def hotkey_action_factory(app_window):
    """
    构造供 pynput.GlobalHotKeys 使用的热键 -> 回调 映射。
    对每个热键做 normalize_hotkey_string 规范化，并且仅在非空时注册。
    """
    mapping = {}
    try:
        actions = {
            'toggle': app_window.toggle_visible,
            'settings': app_window.open_settings,
            'next': app_window.show_next_word,
            'prev': app_window.show_prev_word,
            'known': app_window.mark_current_known,
            'forgot': app_window.mark_current_forgot
        }
        for name, handler in actions.items():
            raw = user_hotkeys.get(name, DEFAULT_HOTKEYS.get(name, ''))
            key = normalize_hotkey_string(raw)
            if key:
                # 绑定当前 handler（使用默认参数避免闭包问题）
                mapping[key] = (lambda h=handler: app_window.root.after(0, h))
    except Exception as e:
        print(f"构造热键映射失败: {e}")
    return mapping
# ...existing code...

def start_hotkeys_listener(app_window):
    global gh_listener
    # stop existing
    try:
        if gh_listener:
            gh_listener.stop()
    except Exception:
        pass
    mapping = hotkey_action_factory(app_window)
    try:
        gh_listener = keyboard.GlobalHotKeys(mapping)
        gh_listener.daemon = True
        gh_listener.start()
    except Exception as e:
        print(f"全局热键监听启动失败: {e}")

def restart_hotkeys_listener(app_window):
    # small delay to ensure old listener stopped
    threading.Thread(target=lambda: (time.sleep(0.1), start_hotkeys_listener(app_window)), daemon=True).start()

if __name__ == "__main__":
    # 支持 Windows 上的可执行文件冻结
    try:
        multiprocessing.freeze_support()
    except Exception:
        pass

    load_review_data()
    config = load_config()
    pronunciation_type['value'] = config.get('pronunciation_type', 1)
    shuffle_mode['value'] = config.get('shuffle_mode', False)
    topmost_timer['enabled'] = config.get('topmost_enabled', True)

    # hotkeys 已在 load_config 中合并到 user_hotkeys

    dict_path = config.get('current_dict')
    if not dict_path or not os.path.exists(resource_path(dict_path)):
        # 选取第一个内置字典
        try:
            dict_path = next(iter(all_dicts.values()))[0]['url']
        except Exception:
            dict_path = ''
    current_dict_path['value'] = dict_path

    initial_words = load_dict_from_file(dict_path) if dict_path else []
    start_index = config.get('current_index', 0)

    root = tk.Tk()
    # 兼容打包环境，隐藏主窗口图标
    try:
        root.iconbitmap(default=resource_path("logo.ico"))
    except Exception:
        pass

    app = TransparentWordWindow(root, initial_words, start_index=start_index)

    # 启动全局热键监听（守护线程）
    start_hotkeys_listener(app)

    # 托盘在独立线程启动，避免阻塞主线程（注意：在某些平台 pystray 要求主线程创建图标）
    threading.Thread(target=create_tray_icon, args=(app,), daemon=True).start()

    window_width, window_height = 700, 150
    set_window_to_bottom_right(root, window_width, window_height)
    root.protocol("WM_DELETE_WINDOW", lambda: [save_config(), save_review_data(), root.destroy()])
    # 将 visual 设置应用到初始窗口
    app.apply_visual_settings()
    root.mainloop()