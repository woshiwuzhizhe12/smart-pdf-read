"""
Excel/WPS ↔ PDF 联动高亮工具 + Excel vs Excel 双文件对比 + PDF差异对比
- 支持 Excel 和 WPS 表格
- 四种模式：
  ① 跟随点击：点哪个单元格高亮哪个关键词（PDF联动）
  ② 全量扫描：扫描指定列所有内容，同时高亮表格和PDF中能互相匹配的部分
  ③ Excel对比：对比两份Excel不同列的关键词，双向高亮差异（匹配/仅左表/仅右表）
  ④ PDF差异对比：对比两份PDF的指定页，生成差异图像
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog
from PIL import Image, ImageTk, ImageDraw
from io import BytesIO
from collections import OrderedDict
import fitz
import numpy as np
import os
import sys
import subprocess
import re
import datetime
import ctypes
import threading
import queue
from ctypes import wintypes

try:
    import win32com.client as win32
except ImportError:
    win32 = None

# ════════════════════════════════════════════════════════════════
# Everything SDK 封装
# ════════════════════════════════════════════════════════════════



# ════════════════════════════════════════════════════════════════
# AnyTXT Searcher 内容检索封装
# ════════════════════════════════════════════════════════════════

class AnyTxtSearch:
    """封装 AnyTXT Searcher JSON-RPC API（本机 127.0.0.1:9920，进程 ATGUI.exe）。

    用于在 Everything 文件名搜索结果基础上做跨盘内容检索：GetResult 返回
    命中文件的 [fid, 修改时间, 大小, 完整路径]，content_filter 与给定文件
    列表求交集，返回命中的完整路径。
    """
    URL = "http://127.0.0.1:9920"

    @classmethod
    def _rpc(cls, method, input_, timeout=15):
        import json, urllib.request
        body = json.dumps({"id": 1, "jsonrpc": "2.0", "method": method,
                           "params": {"input": input_}}).encode("utf-8")
        req = urllib.request.Request(cls.URL, data=body, headers={
            "Content-Type": "application/json", "Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    @classmethod
    def get_result_page(cls, keyword, folder, limit=500, offset=0):
        """单页查询，返回 [(fid, path), ...]。"""
        inp = {"pattern": keyword, "filterDir": folder, "filterExt": "*",
               "lastModifyBegin": 0, "lastModifyEnd": 2147483647,
               "limit": limit, "offset": offset, "order": 0}
        out = cls._rpc("ATRpcServer.Searcher.V1.GetResult", inp)["result"]["data"]["output"]
        return [(f[0], f[3]) for f in out["files"]]

    @classmethod
    def get_all_matches(cls, keyword, folder, limit=500):
        """翻页拉全 folder 下所有命中，返回 {path: fid}。"""
        path2fid, offset = {}, 0
        while True:
            page = cls.get_result_page(keyword, folder, limit=limit, offset=offset)
            if not page:
                break
            for fid, path in page:
                path2fid[path] = fid
            offset += len(page)
        return path2fid

    @classmethod
    def content_filter(cls, paths, keyword):
        """在给定文件列表里按内容过滤（跨盘）。返回命中路径列表；不可用/失败返回 None。"""
        if not paths or not keyword or not keyword.strip():
            return None
        try:
            drives = sorted({os.path.splitdrive(p)[0] for p in paths if p})
            path2fid = {}
            for d in drives:
                path2fid.update(cls.get_all_matches(keyword, d + "\\"))
        except Exception:
            return None
        norm_all = {os.path.normcase(p) for p in paths}
        return [p for p in path2fid if os.path.normcase(p) in norm_all]

    @classmethod
    def get_raw_text(cls, fid):
        """按 fid 取文件全文（供后续页内高亮，暂未接入）。"""
        out = cls._rpc("ATRpcServer.Searcher.V1.GetRawTextByFID",
                       {"fid": str(fid)})["result"]["data"]["output"]
        return out["text"]


# ── 颜色常量（深藏青 + 紫罗兰强调色 + 青绿点缀，科技感暗色主题）──────────
BG      = "#0f1117"
PANEL   = "#171a26"
CARD    = "#1e2230"
ENTRY   = "#262b3d"
BORDER  = "#171a26"  # = PANEL，隐形分隔（VS Code 风格：用间距替代线条）
ACCENT  = "#7c6cf0"
GREEN   = "#2dd4a8"
TEXT    = "#e9eaf5"
DIM     = "#8d92ab"
CANVAS  = "#0a0b12"

# ── 派生二级色（随主题明暗切换的弹窗/浮层背景）──────────

# ── 语义色（阅读/状态反馈常用色，集中定义便于统一换色）──────────
HIGHLIGHT   = "#ffc832"   # PDF 文本高亮（搜索命中底色，黄）
ANNOTATE    = "#ffdc00"   # 手绘标注 / EPUB 链接高亮（黄，比高亮更亮）
DANGER      = "#ff4d6d"   # 危险/重要/错误（红）


# ════════════════════════════════════════════════════════════════
# 全局快捷键（隐藏/恢复主窗口）相关常量
# ════════════════════════════════════════════════════════════════
MAGNIFIER_TRANSPARENT_COLOR = "#ff00fe"  # 放大镜悬浮窗的"透明色"，配合-transparentcolor做出圆形效果
BUTTON_CORNER_RADIUS = 8  # 按钮圆角半径（像素）：在原来方角矩形基础上稍微带一点弧度，
                           # 不是全圆的胶囊形状。改这一个数字就能统一调整全部按钮的弧度大小。
MATCH_CHIP_COLOR_GROUPS = [
    "#7c6cf0", "#2dd4a8", "#e0a83e", "#e0607a",
    "#4ea8e0", "#b87ce0", "#6cd47c", "#e08a4e",
]  # 匹配页码列表每10个一组循环使用的颜色，方便长列表里快速分辨大致位置

# ── 圆盘菜单：全部可选操作池（key → 显示文案），顺序即默认顺序 ──
# key 同时作为 action 的唯一标识，自定义时只增删改顺序，不新增未知 key。
RADIAL_ACTION_POOL = [
    ("magnifier", "🔍 放大镜"),
    ("crop", "✂ 智能裁边"),
    ("body_crop", "⛶ 框选正文"),
    ("annotate", "✏ 标注"),
    ("popup_mode", "📋 弹窗模式"),
    ("copy", "📄 复制文本"),
    ("bookmark", "🔖 加书签"),
    ("fit_width", "🔲 适应宽度"),
]
DEFAULT_RADIAL_ACTIONS = [key for key, _ in RADIAL_ACTION_POOL]



# 显示器矩形结构体（用于放大镜多屏定位：按「光标所在显示器」clamp，而非主屏尺寸）
class _MONITOR_POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class _MONITOR_RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _MONITOR_RECT),
                ("rcWork", _MONITOR_RECT), ("dwFlags", ctypes.c_ulong)]

# 按键名 → 虚拟键码（VK_*），用于把"Ctrl+Alt+E"这种字符串解析成RegisterHotKey需要的参数
VK_MAP = {chr(c): c for c in range(0x30, 0x3A)}      # '0'-'9'
VK_MAP.update({chr(c): c for c in range(0x41, 0x5B)})  # 'A'-'Z'（大写ASCII正好等于VK码）
VK_MAP.update({f"F{i}": 0x70 + i - 1 for i in range(1, 13)})  # F1-F12
VK_MAP.update({
    "SPACE": 0x20, "TAB": 0x09, "ESC": 0x1B, "ESCAPE": 0x1B,
    "ENTER": 0x0D, "RETURN": 0x0D, "BACKSPACE": 0x08,
    "UP": 0x26, "DOWN": 0x28, "LEFT": 0x25, "RIGHT": 0x27,
    "HOME": 0x24, "END": 0x23, "INSERT": 0x2D, "DELETE": 0x2E,
    "PAGEUP": 0x21, "PAGEDOWN": 0x22,
    "`": 0xC0, "-": 0xBD, "=": 0xBB, "[": 0xDB, "]": 0xDD,
    ";": 0xBA, "'": 0xDE, ",": 0xBC, ".": 0xBE, "/": 0xBF, "\\": 0xDC,
})

# ════════════════════════════════════════════════════════════════
# 文件阅读器 - 代码语法高亮（VSCode Dark+ 配色方案）
# ════════════════════════════════════════════════════════════════
CODE_BG = "#1e1e1e"
CODE_FG = "#d4d4d4"
CODE_GUTTER_BG = "#1e1e1e"
CODE_GUTTER_FG = "#6e7681"
CODE_SELECT_BG = "#264f78"


# 每种"语言"对应一组 (分组名, 正则, 是否需要MULTILINE) -> 统一拼成一个大正则，按顺序优先匹配


# 扩展名 → 语言规则集 的映射；值为 None 表示走代码阅读器但不做语法高亮（仅深色+行号+等宽字体）
EXT_TO_LANG = {
    ".py": "python",
    ".c": "c", ".h": "c",
    ".cpp": "c", ".cc": "c", ".cxx": "c", ".hpp": "c", ".hxx": "c",
    ".java": "c", ".cs": "c",
    ".html": "html", ".htm": "html", ".xml": "html",
    ".css": "css",
    ".js": "js", ".ts": "js", ".jsx": "js", ".tsx": "js",
    ".json": "json",
    ".sh": None, ".bat": None, ".ps1": None,
    ".yaml": None, ".yml": None, ".ini": None, ".cfg": None,
    ".txt": None, ".log": None, ".sql": None,
}
# 走"代码阅读器"（深色VSCode风格 + 行号）的全部扩展名（含上面没有高亮规则的纯文本类）
CODE_VIEW_EXTS = set(EXT_TO_LANG.keys())
# 走"浅色纯文本预览"（旧有逻辑：Markdown原文/PPT文本/EPUB文本）的扩展名
TEXT_VIEW_EXTS = {".md", ".ppt", ".pptx", ".epub"}
# 图片扩展名（走图片预览，用 PIL 打开显示在主阅读区 Canvas 上）
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tiff", ".tif"}
# Office 二进制格式（无内置预览，双击用系统默认程序打开）
OFFICE_EXTS = {".doc", ".docx", ".xls", ".xlsx"}
# 文件阅读器支持打开的全部扩展名（PDF + 代码/文本类 + Markdown/PPT/EPUB + 图片 + Office）
READER_SUPPORTED_EXTS = (".pdf",)










def _hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 3:
        hex_color = "".join(c * 2 for c in hex_color)
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    r, g, b = (max(0, min(255, int(round(c)))) for c in rgb)
    return f"#{r:02x}{g:02x}{b:02x}"


def _lighten_color(hex_color, factor=0.18):
    """把颜色往白色方向混合，用于按钮hover状态。factor越大越亮。"""
    try:
        r, g, b = _hex_to_rgb(hex_color)
    except Exception:
        return hex_color
    r = r + (255 - r) * factor
    g = g + (255 - g) * factor
    b = b + (255 - b) * factor
    return _rgb_to_hex((r, g, b))




# ── 小米汽车车漆色主题（车漆色 → 深色主题，base 做强调色）──────────
# 官方只公布色名、未公布 HEX，下表色值为按实车观感的估算值，可逐项微调。
# 顺序按车型分组：SU7 → SU7 Ultra → YU7 → YU7 GT，菜单里作为「小米汽车色」分类追加在末尾。
_XIAOMI_COLORS = [
    # ── SU7（9 色）──
    ("mi_gulf_blue",       "海湾蓝", "#0E6BA8"),
    ("mi_elegant_gray",    "雅灰",   "#A7ACB1"),
    ("mi_olive_green",     "橄榄绿", "#6B7A3C"),
    ("mi_meteor_blue",     "流星蓝", "#6FA8DC"),
    ("mi_pearl_white",     "珍珠白", "#F0F1EC"),
    ("mi_lava_orange",     "熔岩橙", "#F06A0A"),
    ("mi_cambrian_gray",   "寒武岩灰", "#6F7983"),
    ("mi_aurora_purple",   "霞光紫", "#7B5CC4"),
    ("mi_diamond_black",   "钻石黑", "#1A1C20"),
    # ── SU7 Ultra（新增 4 色）──
    ("mi_lightning_yellow", "闪电黄", "#F5C800"),
    ("mi_parrot_green",     "鹦鹉绿", "#14A04F"),
    ("mi_space_silver",     "太空银", "#C1C7CC"),
    ("mi_obsidian_black",   "曜石黑", "#0E0F12"),
    # ── YU7（新增 8 色）──
    ("mi_gem_green",       "宝石绿", "#17A673"),
    ("mi_titanium",        "钛金属", "#A6ABB0"),
    ("mi_shadow_cyan",     "影青",   "#55B2AC"),
    ("mi_gold_pink",       "流金粉", "#E6B1A5"),
    ("mi_danxia_purple",   "丹霞紫", "#8D5FC2"),
    ("mi_deep_sea_blue",   "深海蓝", "#12395F"),
    ("mi_cherry_red",      "车厘子红", "#A60F2C"),
    ("mi_volcanic_gray",   "火山灰", "#565A60"),
    # ── YU7 GT 高端定制漆（3 款哑光/液态）──
    ("mi_warrior_black_matte", "武士黑哑光", "#232426"),
    ("mi_titanium_matte",      "钛金属哑光", "#8E9499"),
    ("mi_liquid_silver",       "液态银",     "#C9CDD1"),
]


def _theme_from_car_color(name, base):
    """由车漆色生成一套深色主题：base 做 ACCENT 强调色，背景向 base 去饱和偏色。

    车漆色多为高饱和色，直接铺满背景会刺眼，所以背景统一「去饱和 + 压暗」，
    只在强调色处保留车漆原色，保证主题一眼能认出车色、又耐看。
    """
    r, g, b = _hex_to_rgb(base)
    luma = 0.299 * r + 0.587 * g + 0.114 * b

    def tone(sat, light):
        """sat: 保留饱和度的比例(0~1)；light: 明度缩放(0~1，越小越暗)。"""
        rr = luma + (r - luma) * sat
        gg = luma + (g - luma) * sat
        bb = luma + (b - luma) * sat
        return _rgb_to_hex((rr * light, gg * light, bb * light))

    return {
        "name": name,
        "BG": tone(0.30, 0.20),
        "PANEL": tone(0.28, 0.28),
        "CARD": tone(0.26, 0.36),
        "ENTRY": tone(0.24, 0.46),
        "BORDER": tone(0.28, 0.52),
        "ACCENT": base,
        "GREEN": "#2dd4a8",
        "TEXT": "#e8eaf0",
        "DIM": "#8b92a0",
        "CANVAS": tone(0.34, 0.12),
    }


class RoundedButton(tk.Canvas):
    """
    VS Code 风格按钮：默认透明背景 + 彩色文字/图标，hover 时显示微妙底色。
    原 bg 参数自动转为文字颜色（亮色=彩色文字，暗色=普通文字）。
    """

    def __init__(self, parent, text="", command=None, bg=None, fg="white",
                 font=None, padx=12, pady=6, cursor="hand2",
                 activebackground=None, activeforeground=None,
                 width=None, height=None, wraplength=None, state="normal",
                 bd=0, relief=None, **kwargs):
        try:
            parent_bg = kwargs.pop("bg_canvas", None) or parent["bg"]
        except Exception:
            parent_bg = BG
        super().__init__(parent, bg=parent_bg, highlightthickness=0, bd=0)

        self._text = text
        self._command = command
        self._parent_bg = parent_bg
        # ── VS Code 风格：bg 参数转为文字颜色 ──
        raw_bg = bg or TEXT
        self._text_color = self._map_to_text_color(raw_bg)
        self._fg = self._text_color  # 默认文字颜色
        self._font = font or ("Microsoft YaHei UI", 9)
        self._padx = padx
        self._pady = pady
        # hover/pressed 背景色（VS Code 微妙层次）
        self._hover_bg = _lighten_color(parent_bg, 0.10)
        self._press_bg = _lighten_color(parent_bg, 0.18)
        self._state = state
        self._wraplength = wraplength
        self._fixed_w = width
        self._fixed_h = height

        self._hover = False
        self._pressed = False

        try:
            super().config(cursor=cursor)
        except Exception:
            pass

        self.bind("<Configure>", self._redraw)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

        self._apply_size()
        self.after_idle(self._redraw)

    @staticmethod
    def _map_to_text_color(raw_bg):
        """将原 bg 参数映射为 VS Code 风格的文字颜色。"""
        # 已知的深色UI背景 → 转为普通文字颜色
        DARK_UI = {"#454b63", "#343a52", "#262b3d", "#1e2230", "#3a3f56",
                   ENTRY, CARD, PANEL, BG, CANVAS, BORDER}
        if raw_bg in DARK_UI:
            return TEXT
        # 暗色 → 普通文字
        try:
            r, g, b = _hex_to_rgb(raw_bg)
            brightness = 0.299 * r + 0.587 * g + 0.114 * b
            if brightness < 90:
                return TEXT
        except Exception:
            pass
        # 亮色/彩色 → 用原色作为文字颜色
        return raw_bg

    # ── 尺寸计算：按文字+内边距自动估算，使其行为接近原生Button ──
    def _apply_size(self):
        import tkinter.font as tkfont
        try:
            f = tkfont.Font(font=self._font)
            text_w = max(f.measure(self._text), 1)
            text_h = f.metrics("linespace")
        except Exception:
            text_w, text_h = max(10 * len(self._text), 20), 16

        if self._fixed_w:
            # tk.Button的width参数是"字符数"，这里粗略按字符平均宽度估算
            try:
                w = max(text_w, int(self._fixed_w) * 8) + self._padx * 2
            except Exception:
                w = text_w + self._padx * 2
        else:
            w = text_w + self._padx * 2
        h = (self._fixed_h * 18 if self._fixed_h else text_h) + self._pady * 2
        self._natural_w = max(int(w), 24)
        self._natural_h = max(int(h), 20)
        self.config(width=self._natural_w, height=self._natural_h)

    # ── 颜色状态（VS Code 风格：透明 → hover 微妙底色 → pressed 更深）──
    def _current_colors(self):
        if self._state == "disabled":
            return self._parent_bg, DIM
        if self._pressed:
            return self._press_bg, self._text_color
        if self._hover:
            return self._hover_bg, self._text_color
        return self._parent_bg, self._text_color

    def _refresh_theme(self, color_map=None, new_palette=None):
        """主题切换后重算自身颜色（父背景 + 文字色），让按钮跟随明暗主题。"""
        try:
            self._parent_bg = self.master["bg"]
        except Exception:
            self._parent_bg = (new_palette or {}).get("PANEL", BG)
        old_text = str(self._text_color).lower()
        if color_map and old_text in color_map:
            self._text_color = color_map[old_text]
        self._hover_bg = _lighten_color(self._parent_bg, 0.10)
        self._press_bg = _lighten_color(self._parent_bg, 0.18)
        try:
            super().config(bg=self._parent_bg)
        except Exception:
            pass
        self._redraw()

    def _draw_pill(self, x0, y0, x1, y1, fill):
        """
        画一个带圆角的矩形（不是全圆的"胶囊"两端半圆，而是在原本方角矩形
        基础上，四个角用小圆弧稍微"削"一下，弧度可以通过模块级常量
        BUTTON_CORNER_RADIUS 统一调整，数值越大弧度越明显，0就是纯方角）。
        """
        w, h = x1 - x0, y1 - y0
        r = max(0, min(BUTTON_CORNER_RADIUS, w / 2, h / 2))
        if r < 1:
            self.create_rectangle(x0, y0, x1, y1, fill=fill, outline=fill)
            return
        # 四个角各画一个90度扇形
        self.create_arc(x0, y0, x0 + 2 * r, y0 + 2 * r,
                         start=90, extent=90, fill=fill, outline=fill)
        self.create_arc(x1 - 2 * r, y0, x1, y0 + 2 * r,
                         start=0, extent=90, fill=fill, outline=fill)
        self.create_arc(x0, y1 - 2 * r, x0 + 2 * r, y1,
                         start=180, extent=90, fill=fill, outline=fill)
        self.create_arc(x1 - 2 * r, y1 - 2 * r, x1, y1,
                         start=270, extent=90, fill=fill, outline=fill)
        # 中间用两条矩形条补满（横向一条+纵向一条，覆盖四角之间的全部区域）
        self.create_rectangle(x0 + r, y0, x1 - r, y1, fill=fill, outline=fill)
        self.create_rectangle(x0, y0 + r, x1, y1 - r, fill=fill, outline=fill)

    def _redraw(self, event=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1 or h <= 1:
            w, h = self._natural_w, self._natural_h
        fill, fg = self._current_colors()
        self._draw_pill(1, 1, w - 1, h - 1, fill)
        self.create_text(
            w / 2, h / 2, text=self._text, fill=fg, font=self._font,
            width=self._wraplength if self._wraplength else 0,
            justify=tk.CENTER)

    # ── 鼠标事件 ──
    def _on_enter(self, _e=None):
        if self._state != "disabled":
            self._hover = True
            self._redraw()

    def _on_leave(self, _e=None):
        self._hover = False
        self._pressed = False
        self._redraw()

    def _on_press(self, _e=None):
        if self._state == "disabled":
            return
        self._pressed = True
        self._redraw()

    def _on_release(self, event=None):
        if self._state == "disabled":
            return
        was_pressed = self._pressed
        self._pressed = False
        self._redraw()
        if was_pressed and event is not None:
            w, h = self.winfo_width(), self.winfo_height()
            if 0 <= event.x <= w and 0 <= event.y <= h and self._command:
                self._command()

    # ── tk.Button兼容的 config()/configure()（VS Code 风格）──
    def config(self, **kwargs):
        if not kwargs:
            return
        need_resize = False
        if "text" in kwargs:
            self._text = kwargs.pop("text")
            need_resize = True
        if "bg" in kwargs:
            raw_bg = kwargs.pop("bg")
            self._text_color = self._map_to_text_color(raw_bg)
        if "fg" in kwargs:
            v = kwargs.pop("fg")
            if v is not None:
                self._text_color = v
        if "font" in kwargs:
            self._font = kwargs.pop("font")
            need_resize = True
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self._state = kwargs.pop("state")
        if "cursor" in kwargs:
            try:
                super().config(cursor=kwargs.pop("cursor"))
            except Exception:
                kwargs.pop("cursor", None)
        if kwargs:
            try:
                super().config(**kwargs)
            except Exception:
                pass
        if need_resize:
            self._apply_size()
        self._redraw()

    configure = config

    def cget(self, key):
        mapping = {"text": self._text, "bg": self._text_color, "fg": self._text_color,
                   "state": self._state, "font": self._font,
                   "command": self._command}
        if key in mapping:
            return mapping[key]
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)


class SlimScrollbar(tk.Canvas):
    """超薄滚动条：默认 6px，鼠标悬停临时加宽，颜色随全局主题。

    对外兼容 ttk.Scrollbar 的关键接口：
      - set(first, last)：画布 yscrollcommand 回调，更新滑块位置/大小
      - 构造时传入 command：拖动/点击时回调 command("moveto"/"scroll", ...)
    这样可以直接替换原来的 ttk.Scrollbar，画布配置不用改。
    """
    def __init__(self, parent, command=None, orient="vertical",
                 width=6, hover_width=14):
        self._orient = orient
        self._command = command
        self._width = width
        self._hover_width = hover_width
        self._first = 0.0
        self._last = 1.0
        self._dragging = False
        self._bg = CANVAS
        self._thumb = DIM
        self._thumb_active = ACCENT
        if orient == "vertical":
            super().__init__(parent, width=width, bg=CANVAS,
                             highlightthickness=0, bd=0)
        else:
            super().__init__(parent, height=width, bg=CANVAS,
                             highlightthickness=0, bd=0)
        self.bind("<Configure>", lambda e: self._redraw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self._redraw()

    def _refresh_theme(self, color_map=None, new_palette=None):
        """主题切换后重读新配色（recolor 会先于 globals().update 调用，
        所以不能读模块全局，要用传入的 new_palette）。"""
        if new_palette:
            self._bg = new_palette.get("CANVAS", self._bg)
            self._thumb = new_palette.get("DIM", self._thumb)
            self._thumb_active = new_palette.get("ACCENT", self._thumb_active)
        self.configure(bg=self._bg)
        self._redraw()

    def set(self, first, last):
        """画布 yscrollcommand 回调：更新滑块位置/大小。"""
        try:
            first = max(0.0, min(1.0, float(first)))
            last = max(first, min(1.0, float(last)))
        except Exception:
            return
        if abs(first - self._first) < 1e-6 and abs(last - self._last) < 1e-6:
            return
        self._first = first
        self._last = last
        self._redraw()

    # ── 几何计算 ──
    def _span(self):
        """返回 (轨道长度, 轨道厚度)，长度沿滚动方向。"""
        if self._orient == "vertical":
            return max(1, self.winfo_height() - 1), max(1, self.winfo_width())
        return max(1, self.winfo_width() - 1), max(1, self.winfo_height())

    def _thumb_rect(self):
        length, thick = self._span()
        inner = max(1, length - 2)
        frac = max(0.01, self._last - self._first)
        thumb_len = max(12, int(inner * frac))
        max_off = max(0, inner - thumb_len)
        off = max(0, min(int(self._first * inner), max_off))
        if self._orient == "vertical":
            return (0, off + 1, thick, off + 1 + thumb_len)
        return (off + 1, 0, off + 1 + thumb_len, thick)

    def _redraw(self):
        self.delete("all")
        x0, y0, x1, y1 = self._thumb_rect()
        color = self._thumb_active if self._dragging else self._thumb
        self.create_rectangle(x0, y0, x1, y1, fill=color, outline="",
                              width=0, tags="thumb")

    def _on_enter(self, _e):
        if self._orient == "vertical":
            self.configure(width=self._hover_width)
        else:
            self.configure(height=self._hover_width)
        self._redraw()

    def _on_leave(self, _e):
        if self._dragging:
            return
        if self._orient == "vertical":
            self.configure(width=self._width)
        else:
            self.configure(height=self._width)
        self._redraw()

    def _frac_from_event(self, event):
        length, _ = self._span()
        inner = max(1, length - 2)
        x0, y0, x1, y1 = self._thumb_rect()
        thumb_len = (y1 - y0) if self._orient == "vertical" else (x1 - x0)
        pos = event.y if self._orient == "vertical" else event.x
        # 让滑块中心对齐鼠标，返回滑块顶部对应的比例
        return max(0.0, min(1.0, (pos - thumb_len / 2.0) / inner))

    def _on_press(self, event):
        x0, y0, x1, y1 = self._thumb_rect()
        hit = (x0 <= event.x <= x1 and y0 <= event.y <= y1)
        if hit:
            self._dragging = True
            self._redraw()
            self._drag_move(event)
        else:
            # 点击轨道：直接跳到对应位置
            frac = self._frac_from_event(event)
            self._invoke("moveto", frac)

    def _drag_move(self, event):
        self._invoke("moveto", self._frac_from_event(event))

    def _on_drag(self, event):
        if self._dragging:
            self._drag_move(event)

    def _on_release(self, event):
        if self._dragging:
            self._dragging = False
            self._redraw()
            # 拖动期间 Leave 被忽略，松开时若指针已移出控件则收回加宽
            if (event.x < 0 or event.x > self.winfo_width()
                    or event.y < 0 or event.y > self.winfo_height()):
                self._on_leave(event)

    def _invoke(self, *args):
        if self._command:
            try:
                self._command(*args)
            except Exception:
                pass


class ExcelPdfSync(tk.Tk):
    POLL_MS  = 600
    PDF_DPI  = 150

    def _apply_app_icon(self):
        """设置标题栏 + 任务栏图标（杠杆+北极星，替换 Tk 默认红羽毛）。"""
        try:
            base = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else "."
            ico = os.path.join(base, "app_icon.ico")
            png = os.path.join(base, "app_icon.png")
            if os.path.isfile(ico):
                self.iconbitmap(ico)
            if os.path.isfile(png):
                self._app_icon_photo = tk.PhotoImage(file=png)
                self.iconphoto(True, self._app_icon_photo)
        except Exception:
            pass  # 图标缺失/加载失败不影响启动

    def __init__(self):
        super().__init__()
        self.title("卓有成效的人生")
        self._apply_app_icon()
        self.geometry("1500x900")
        self.configure(bg=BG)
        # 全局去除 Entry/Text/Listbox 边框（VS Code 风格无边框）
        self.option_add("*Entry.highlightthickness", 0)
        self.option_add("*Text.highlightthickness", 0)
        self.option_add("*Listbox.highlightthickness", 0)
        self.option_add("*Entry.borderwidth", 0)
        self.option_add("*Entry.relief", "flat")
        self.option_add("*Text.borderwidth", 0)
        self.option_add("*Text.relief", "flat")

        self.pdf_doc   = None
        self.cur_page  = 0
        self.scale     = 1.0
        self._tk_img   = None
        self._poll_job = None
        self._last_val = None
        self._highlighted_cells = []
        self._mode     = tk.StringVar(value="follow")
        # 高亮颜色（RGB tuple，默认黄色）
        self._hl_color = (255, 220, 0)
        # 搜索高亮颜色（RGB tuple，默认红色）
        self._search_hl_color = (255, 80, 80)

        # Excel对比模式专用状态
        self._xl_app1 = None      # 左表COM对象
        self._xl_app2 = None      # 右表COM对象
        self._xl_wb1 = None       # 左侧实际选中的工作簿
        self._xl_wb2 = None       # 右侧实际选中的工作簿
        self._xl_name1 = ""       # 左表名称
        self._xl_name2 = ""       # 右表名称
        self._pdf_word_cache = OrderedDict()  # PDF文字缓存，避免每次搜索重复解析

        # ── 双PDF联动模式（PDF联动页签内）──
        self._pdf_doc_left = None
        self._pdf_doc_right = None
        self._cur_page_left = 0
        self._cur_page_right = 0
        self._scale_left = 1.0
        self._scale_right = 1.0
        self._tk_img_left = None
        self._tk_img_right = None

        # PDF阅读器专用状态
        self._reader_pdf = None
        self._reader_page = 0
        self._reader_scale = 1.0
        self._reader_fit_width = True  # 是否自动适应窗口宽度
        self._reader_scroll_lines = 3  # 滚轮每次滚动行数（可自定义）
        self._reader_tk_img = None
        self._reader_tk_imgs = []   # 当前渲染窗口的图片引用
        # 文本缓存：加速搜索性能
        self._reader_text_cache = None   # [page_words, ...] 每页的 (x0,y0,x1,y1,word) 列表
        self._reader_word_index = None   # {word_lower: [page_idx, ...]} 词→页面索引
        # 页面 Y 偏移缓存：加速点击定位
        self._reader_page_yoffsets = []  # [y0, y1, y2, ...] 每页顶部的Y坐标（渲染后）
        self._reader_dir = ""
        # ── 网络 PDF 搜索（Bing filetype:pdf）状态 ──
        self._reader_net_save_dir = ""      # 默认保存目录（持久化到 config.json）
        self._reader_net_results = []       # [(title, url, domain), ...] 搜索结果
        self._reader_net_downloaded = None  # 当前已下载到本地的临时 PDF 路径
        self._reader_net_downloaded_name = ""  # 下载文件的建议名（含 .pdf）
        self._reader_net_searching = False  # 是否正在搜索/下载中（防重入）
        self._reader_net_downloading = False  # 是否正在下载 PDF（防重入）
        self._reader_net_search_id = 0      # 搜索任务 ID（丢弃过期结果）
        # PPT渲染素材缓存：{full_path: (mtime, slide_w, slide_h, slides_data)}
        # 缓存的是"导出的图片字节+文字位置"这些素材，不是最终fitz.Document对象，
        # 因为切换文件时旧文档会被close()，缓存活对象会被废掉；
        # 命中缓存时用 _reader_build_ppt_doc 重新拼一个全新对象即可，很快。
        self._reader_ppt_data_cache = OrderedDict()
        self._reader_ppt_cache_max = 4  # 最多缓存4个PPT的渲染素材，控制内存占用
        self._reader_html_path = ""     # 当前打开的HTML文件完整路径（用于浏览器打开/内嵌预览）
        self._reader_active_file_path = ""  # 当前打开的文件路径（任意类型都会更新，常用文件功能用这个）
        self._reader_html_mode = "code"  # "code" 或 "preview"：HTML文件当前显示的是源码还是网页预览
        self._reader_pdf_files = []      # 当前显示的PDF列表（可能已被过滤）
        self._reader_pdf_files_all = []  # 原始完整PDF列表（用于恢复）
        self._reader_sort_mode = "name"  # 排序方式："name"/"date"/"size"/"type"
        self._reader_sort_reverse = False  # 是否降序
        self._open_counts = {}   # 文件打开次数统计 {full_path: int}，持久化到config.json，用于"次数"排序
        self._reader_resume = {}  # 续读记录 {full_path: {"page","scale","fit_width"}}，持久化到config.json
        self._reader_content_search_text = ""   # 当前内容搜索关键词
        self._reader_content_match_pages = {}  # 内容搜索匹配页码 {full_path: [page_idx, ...]}
        self._reader_content_searching = False  # 是否正在搜索中
        self._reader_search_id = 0              # 搜索任务ID（用于取消旧搜索）
        self._reader_content_text_cache = OrderedDict()  # 内容搜索全文缓存 {full_path: (mtime_ns, size, (kind,payload))}
        self._reader_content_cache_max = 10000  # 内容搜索缓存最多保留的文件数（LRU淘汰，控制内存）
        self._reader_highlight_keywords = []    # 当前高亮关键词列表（预览搜索）
        self._reader_highlight_exact = False    # 点击文字时只高亮完整同词
        self._reader_thumb_images = []   # 缩略图图片引用
        self._reader_thumb_positions = []  # 缩略图位置
        # ── 优化新增：PDF阅读器性能优化状态 ──
        self._reader_opening = False      # 防止快速重复打开导致闪退
        self._reader_select_job = None    # 单击延迟任务（防双击冲突）
        self._reader_zoom_job = None      # 缩放防抖定时器
        self._reader_zoom_pending_scale = None  # 待处理缩放值
        self._reader_zoom_anchor = None   # Ctrl+滚轮缩放锚点（PDF坐标），缩放后回卷保持光标内容不动
        self._reader_zoom_anchor_view = None  # 锚点对应的画布视口坐标 (x, y)
        self._reader_sync_scroll = False  # 同步滚动/缩放：滚一个阅读区，其它阅读区按页码/缩放联动
        self._reader_sync_active = False  # 同步传播中的重入保护标志（防止 A→B→A 回传）
        self._reader_pil_cache = OrderedDict()  # 预渲染缓存 (id(doc), page_idx, zoom) -> PIL 原图（主题/裁边无关）
        self._reader_pil_cap = 16          # 缓存上限，超出按 LRU 淘汰
        self._reader_pil_lock = threading.Lock()
        self._reader_prefetch_queue = None   # 预取任务队列（首次预取时惰性创建）
        self._reader_prefetch_thread = None  # 预取后台线程
        self._reader_render_job = None    # 渲染防抖定时器
        self._reader_lazy_job = None      # 懒加载任务
        self._reader_doc_generation = 0   # 打开新文件时使旧的延迟任务失效
        self._reader_render_start = 0     # 当前画布已渲染页范围起点
        self._reader_render_upto = 0      # 当前画布已渲染页范围终点（不含）
        self._reader_occurrence_window = None
        self._reader_occurrence_generation = 0
        self._reader_occurrence_geometry = None
        self._reader_occurrence_size_locked = True
        self._reader_occurrence_locked_size = None
        self._reader_thumb_virtual = True  # 缩略图使用虚拟滚动模式
        self._reader_thumb_visible_start = 0  # 可见缩略图起始索引
        self._reader_thumb_visible_end = 0    # 可见缩略图结束索引
        self._reader_page_cache = OrderedDict()  # 页面渲染缓存 LRU {(page, zoom, theme, crop模式): (zoom, tk_img, 裁剪偏移)}
        self._reader_popup_page_cache = OrderedDict()  # 点击文本弹窗缩略图缓存 LRU {(doc, page_idx, theme): PIL.Image(已应用主题)}
        self._reader_yoffset_estimated = False  # Y偏移是否已精确计算
        self._reader_scroll_job = None    # 滚动检测防抖
        self._reader_scroll_anim_job = None
        self._reader_scroll_velocity = 0.0
        self._reader_file_type = ""       # 当前打开的文件类型
        self._reader_current_path = ""
        self._reader_text_frame = None    # 文本预览框架
        self._reader_text_preview = None  # 文本预览Text控件
        self._reader_left_visible = True  # 左侧面板是否可见
        self._reader_left_frame = None    # 左侧面板引用
        self._reader_left_toggle_btn = None  # 折叠/展开按钮引用
        self._reader_edge_frame = None    # Edge浏览器面板引用
        self._reader_edge_visible = False # Edge浏览器面板是否可见
        self._reader_ai_frame = None      # AI阅读助手面板引用
        self._reader_ai_visible = False   # AI阅读助手面板是否可见
        self._reader_ai_task_id = 0       # 取消过期AI任务
        self._reader_ai_online_enabled = tk.BooleanVar(value=False)
        self._reader_ai_base_url = tk.StringVar(value="https://api.openai.com/v1")
        self._reader_ai_model = tk.StringVar(value="gpt-4o-mini")
        self._reader_ai_api_key = tk.StringVar(value="")
        self._reader_ai_max_chars = 18000
        self._reader_highlight_color = (255, 200, 50)
        self._reader_highlight_hex = tk.StringVar(value=HIGHLIGHT)
        self._reader_click_popup_mode = tk.StringVar(value="hidden")  # full/minimal/hidden
        self._reader_code_font_size = tk.IntVar(value=10)  # 代码阅读字体大小，可调节并会保存
        self._reader_popup_pad = 45.0   # 点击弹窗：以命中文字为中心的裁剪范围（PDF点），Ctrl+滚轮缩放
        self._reader_popup_zoom_job = None  # Ctrl+滚轮缩放后的重渲染防抖任务

        # ── PDF阅读器 - 放大镜（不放大整页的情况下查看局部高清细节）──
        self._reader_magnifier_enabled = False
        self._reader_magnifier_ratio = 3.0  # 放大倍率，可通过－/＋调节
        self._reader_magnifier_ratio_text = tk.StringVar(value="3.0x")
        self._reader_magnifier_window = None   # 跟随鼠标的悬浮放大镜窗口
        self._reader_magnifier_label = None    # 放大镜窗口里显示图像的Label
        self._reader_magnifier_job = None      # 节流任务：约每16ms执行一次
        self._reader_magnifier_pending = None  # 节流期间记录的最新鼠标位置 (slot, x, y, root_x, root_y)
        self._reader_magnifier_size = 260      # 放大镜窗口边长（像素）
        self._reader_match_preview_window = None  # 匹配页码悬停预览的悬浮窗
        self._reader_match_preview_label = None
        self._reader_match_preview_job = None      # 悬停预览防抖任务
        self._reader_match_preview_pending = None  # 防抖期间记录的最新悬停目标
        self._reader_match_preview_hide_job = None  # 延迟隐藏的防抖任务

        # ── PDF阅读器 - 简单标注（手绘半透明高亮框，按文件+页码保存）──
        self._reader_page_marks = {}  # {文件路径: {"页码字符串": [{"rect":[x0,y0,x1,y1],"color":"#.."}]}}
        self._reader_annotate_mode = False
        self._reader_annotate_color = ANNOTATE
        self._reader_annotate_start_xy = None
        self._reader_annotate_rect_id = None
        # ── 页面添加文本（程序内覆盖层，不动原PDF）──
        self._reader_text_marks = {}  # {文件路径: {页码字符串: [{"x","y","text","color"}]}}
        self._reader_text_mode = False
        self._reader_text_color = "#ffdc00"
        # ── 高亮原文（点选文字 → 持久高亮 + 存原文文本，供全局标记搜索）──
        self._reader_text_hl = {}     # {文件路径: {页码字符串: [{"rect","text","color"}]}}
        self._reader_text_hl_mode = False
        self._reader_text_hl_color = "#ff5050"

        # ── PDF阅读器 - 智能裁边（只渲染内容区域，裁掉上下左右空白边距）──
        self._reader_content_bbox_cache = {}   # {(文件路径,页码): fitz.Rect或None}
        self._reader_crop_margins_mode = False
        self._reader_page_crop_origins = []    # 与_reader_page_yoffsets平行的数组：
                                                 # 每页裁剪后图片左上角对应的PDF点坐标(x0,y0)

        # ── PDF阅读器 - 手动框选正文区（切表头/外框，只预览正文）──
        self._reader_body_crop_regions = {}     # {文件路径: (x0,y0,x1,y1) 归一化0~1相对页宽高}
        self._reader_body_crop_mode = False     # 框选模式开关（拖拽在画布上画框）
        self._reader_body_crop_start = None     # 框选起点 (canvas_x, canvas_y)
        self._reader_body_crop_rect_id = None   # 框选过程中的临时矩形 id
        self._reader_body_crop_slot = None      # 框选发生在哪个阅读区

        # ── PDF阅读器 - 圆盘菜单（右键/双击弹出，操作可自定义）──
        self._reader_radial_actions = list(DEFAULT_RADIAL_ACTIONS)
        self._reader_radial_win = None          # 圆盘菜单窗口引用
        self._reader_pending_click = None       # 单击延迟判定：(slot, event)
        self._reader_click_job = None           # 单击延迟任务的 after id
        self._reader_press_moved = False        # 按下后是否拖拽移动（判定单击/拖拽）
        self._reader_press_xy = None            # 按下时的画布坐标，用于移动阈值判断
        self._reader_last_clicked_word = ""     # 最近点选的文字（供圆盘"复制文本"）

        # ── UI 位置标注（给程序各区域起名，浮层显示，方便描述优化位置）──
        self._ui_region_labels = []             # 已显示的浮层标签 Toplevel 引用

        # ── 文本搜索页签（txt/xlsx/c/cpp/py 关键词全文搜索 + 内容预览）──
        self._txt_search_save_path = os.getcwd()      # 保存路径
        self._txt_search_search_path = os.getcwd()    # 搜索路径
        self._txt_search_recent_paths = []            # 常用路径（持久化到 config）
        self._txt_search_current_file = None          # 当前预览的文件完整路径
        self._txt_search_matches = []                 # 当前文件内的关键词命中位置
        self._txt_search_match_index = -1             # 当前命中索引
        self._txt_search_file_type = ".txt"           # 当前文件类型
        self._txt_search_font_size = tk.IntVar(value=10)  # 预览字体大小（可调+持久化）

        # ── Excel 内容搜索页签（xlsx/xls 结构化关键词搜索 + 层级展示，复刻 PowerShell）──
        self._excel_search_tab_built = False
        self._excel_search_search_path = os.getcwd()   # 搜索路径
        self._excel_search_recent_paths = []           # 常用路径（持久化到 config）
        self._excel_search_recursive = tk.BooleanVar(value=True)  # 是否递归子目录
        self._excel_search_searching = False           # 防重入标志
        self._excel_search_id = 0                      # 任务ID（丢弃过期结果）
        self._excel_search_results = []                # 结构化结果 [{file, sheets:[{sheet,header,rows}]}]
        self._excel_search_keywords = []               # 当前清理后的关键词（用于状态栏显示）
        self._excel_search_logic = "OR"                # 当前搜索逻辑（AND/OR）
        self._excel_search_kw_color = {}               # 关键词 -> 高亮颜色
        self._excel_search_kw_idx = {}                 # 关键词 -> 索引（对应 kw{i} tag）

        # ── 联动搜索 ──
        self._reader_linked_search = tk.BooleanVar(value=True)  # 联动搜索开关

        self._reader_magnifier_is_visible = False  # 悬浮窗当前是否已经显示（避免重复deiconify）
        self._reader_magnifier_last_pos = None      # 上一次实际移动窗口到的位置，移动太小就跳过
        self._magnifier_mask_cache = {}  # {(w,h): (圆形遮罩, RGB背景色)}，避免每帧重新画遮罩
        self._reader_magnifier_cache = None  # 放大镜 overscan 缓存：{"key","oc","img"}
        self._reader_magnifier_smooth = None  # 位置平滑状态 (slot,page_idx,pdf_x,pdf_y)，抗高倍率"飘"
        self._reader_magnifier_thumb_job = None    # 点击弹窗缩略图悬停放大镜的节流任务
        self._reader_magnifier_thumb_pending = None  # 缩略图悬停放大镜的最新待渲染位置
        self._reader_magnifier_pinned = False  # 放大镜是否已锁定位置（锁定后用方向键精确微调）
        self._reader_magnifier_pin = None      # 锁定位置 (slot, page_idx, pdf_x, pdf_y, root_x, root_y)
        self._reader_anytxt_hits = {}  # {slot: [路径...]}：搜索栏跨文件命中结果，供点击弹窗跳转
        self._reader_anytxt_seq = 0    # 跨文件提示的自增序号，用于丢弃过期结果

        # ── 全局快捷键（隐藏/恢复主窗口）状态 ──
        self._hotkey_combo = tk.StringVar(value="Ctrl+Alt+E")  # 快捷键组合字符串，可在设置里修改
        self._hotkey_thread = None       # 后台监听线程
        self._hotkey_stop_flag = False   # 用于通知线程退出
        self._hotkey_registered = False  # 是否成功注册
        self._was_zoomed = False          # 隐藏前是否最大化（恢复时还原）

        # ── 常用文件（收藏链接）状态 ──
        self._favorite_files = []  # [{"name": "...", "path": "..."}, ...]，持久化到config.json
        self._reader_workspaces = []  # 工作区快照：[{"name","app_theme","slots",...}]
        self._reader_workspace_folders = []  # 文件夹树：[{"id","name","parent_id","created_at"}]
        self._reader_current_workspace_name = None  # 当前正在使用的工作区名（"更新当前工作区"用）
        self._reader_workspace_sort_mode = "name"  # 工作区/文件夹排序："name"/"time"/"usage"
        self._reader_ws_sort_var = tk.StringVar(value="name")  # 排序菜单 radiobutton 用
        self._reader_manage_refresh_cb = None  # 管理工作区对话框列表刷新回调
        self._reader_ws_dirty = False  # 本实例是否改过工作区/文件夹（多实例防覆盖用）
        self._favorite_sort_mode = "name"    # 常用文件排序方式："name"/"date"/"type"
        self._favorite_sort_reverse = False  # 是否降序
        self._favorite_search_var = tk.StringVar(value="")  # 常用文件搜索过滤关键字
        self._favorite_filtered = []  # 过滤后显示的常用文件子集（下拉框索引与全量列表不同）

        # ── PDF标记/笔记（按文件路径区分，持久化到config.json）──
        self._reader_annotations = {}  # {文件路径: [{"page": 页码, "label": "标签", "note": "笔记"}, ...]}
        self._reader_toc_mode = "toc"  # "toc" 或 "bookmark"：侧边面板当前显示目录还是标记列表
        # ── 标记搜索模式：mark(标记搜索) / tag(标签搜索) ──
        self._reader_mark_mode = "mark"
        self._reader_mark_tree = None       # 标记结果树（ttk.Treeview：mark 两级(文件→位置) / tag 三级树）
        self._reader_mark_tree_data = {}   # {iid: 节点数据}（mark 两级树用）
        self._reader_tag_tree_data = {}     # {iid: 节点数据 dict}
        self._reader_tag_hover_iid = None   # 当前 hover 的树节点 iid（供 tooltip 定位）

        # ── 顶部细条：激励语 + 番茄时钟 ──
        self._motivation_text = tk.StringVar(value="今天也要高效拿下任务！💪")
        self._motivation_color = TEXT  # 激励语文字颜色，默认使用主题文字色
        self._pomodoro_running = False
        self._pomodoro_mode = "work"  # "work" 或 "break"
        self._pomodoro_work_minutes = 25
        self._pomodoro_break_minutes = 5
        self._pomodoro_remaining = self._pomodoro_work_minutes * 60
        self._pomodoro_style = "ring"  # 视觉风格："ring"环形进度/"glow"数字荧光/"pulse"脉冲呼吸/"classic"经典文字
        self._pomodoro_image_path = ""  # 倒计时结束时弹出的自定义图片路径（空=不显示图片）

        # ── EPUB阅读样式（PDF1/PDF2/PDF3三个面板共用同一套设置）──
        self._epub_bg_color = "#1e2230"
        self._epub_font_size = 12
        self._epub_font_family = "Microsoft YaHei UI"
        self._epub_line_spacing = 6
        self._epub_chapters_cache = {}  # {文件路径: [(章节标题, 正文), ...]}，避免重复解析同一个文件

        # ── EPUB标注（按"章节序号+正文内字符偏移量"定位，不依赖字号/
        #    行间距/字体——这些样式怎么变，只要原文不变，标注位置都稳定）──
        self._epub_marks = {}  # {文件路径: [{"chapter":章节序号,"start":int,"end":int,"color":"#.."}]}
        self._epub_annotate_color = ANNOTATE
        self._pomodoro_blink_on = True  # 脉冲风格里冒号闪烁的开关状态
        self._pomodoro_job = None

        # ── PDF差异对比页签 - 笔记功能 ──
        self._notes_folder = ""  # 笔记保存位置，未设置时在_build_notes_panel里给默认值
        self._notes_window = None  # 独立的"笔记/四象限任务"合并窗口（不在任何页签里，单独一个浮动窗口）

        # ── BJYZ 页签：内嵌浏览器状态 ──
        self._bjyz_html_widget = None   # tkinterweb.HtmlFrame（加载完页签后赋值）
        self._bjyz_url_var = None       # URL输入栏StringVar
        self._bjyz_tree = None          # BJYZ左侧文件树（ttk.Treeview）
        self._bjyz_tree_paths = {}      # BJYZ文件树 item_id -> 绝对路径
        self._bjyz_notes_text = None    # BJYZ右侧笔记编辑器
        self._bjyz_note_link_map = {}   # BJYZ链接tag->url映射

        # ── 四象限任务（重要/紧急矩阵）：和笔记共用同一个窗口（顶部标签切换），
        #    数据直接存进config.json ──
        # 结构：{"important_urgent": [{"id":..,"text":..,"done":bool}, ...], ...}
        self._quadrant_tasks = {
            "important_urgent": [], "important_not_urgent": [],
            "not_important_urgent": [], "not_important_not_urgent": [],
        }
        self._quadrant_listboxes = {}  # {象限key: tk.Listbox}，渲染/刷新时用
        self._quadrant_next_id = 1  # 任务条目自增ID，用于精确定位某一条（避免用列表下标，删除/排序后容易错位）
        self._quadrant_reminder_running = False  # 提醒检查器运行标记
        self._quadrant_reminder_stop = False     # 停止标记

        # ── 每日一看：1/2/3/4 分屏布局，每格可固定图片/网址/PDF页 ──
        self._daily_view_grid = 1  # 1, 2, 3, 4
        self._daily_view_cells = [  # 始终4个配置位
            {"type": None, "path": "", "page": 0, "url": "", "label": ""},
            {"type": None, "path": "", "page": 0, "url": "", "label": ""},
            {"type": None, "path": "", "page": 0, "url": "", "label": ""},
            {"type": None, "path": "", "page": 0, "url": "", "label": ""},
        ]
        self._daily_view_tk_imgs = [None, None, None, None]  # PhotoImage引用
        self._daily_view_zooms = [1.0, 1.0, 1.0, 1.0]  # 每格独立缩放比例
        self._daily_view_active_cell = 0  # 当前选中的格子（搜索固定目标）
        self._daily_view_drag_start = {}  # {idx: (x, y)} 拖拽起点
        self._daily_view_magnifier_active = [False, False, False, False]  # 每格放大镜开关
        self._daily_view_magnifier_win = None  # 放大镜悬浮窗引用
        self._daily_view_magnifier_img = None  # 放大镜当前图像引用
        self._daily_view_magnifier_ratio = 3.0  # 放大镜倍率（−/＋调节）
        self._daily_view_magnifier_ratio_text = tk.StringVar(value="3.0x")

        # ── PDF阅读器右侧双开视图 ──
        self._reader_pdf2 = None
        self._reader_path2 = ""
        self._reader_page2 = 0
        self._reader_scale2 = 1.0
        self._reader_fit_width2 = True
        self._reader_tk_img2 = None
        self._reader_dual_visible = False
        self._cross_send_enabled = tk.BooleanVar(value=False)  # 跨程序发送(已停用，保留状态避免残留引用报错)
        self._reader2_resize_job = None
        self._reader_aux = {}
        self._reader_active_slot = 1
        self._reader_merged_occurrence_window = None
        self._reader_merged_occurrence_geometry = None
        self._reader_merged_occurrence_position = None
        self._reader_merged_geometry_job = None
        self._reader_merged_occurrence_locked = True
        self._reader_merged_occurrence_size = None
        self._reader_merged_occurrence_generation = 0
        self._reader_merged_occurrence_results = {}
        # 列级增量：弹窗内每栏骨架/widget 引用，按 slot 索引，切换点击文本时
        # 只清空并重渲染被点击的那一栏，其它栏（PDF1/PDF3）原样保留。
        self._reader_merged_columns = {}
        self._reader_merged_summary_label = None
        self._reader_merged_lock_btn = None
        self._reader_merged_columns_frame = None
        self._reader_merged_active_slots = []

        # ── 滚轮横滚模式（PDF1-4 左右滑动切换）──
        self._reader_hscroll_on = False       # 横滚模式开关
        self._reader_hscroll_visible = 2      # 固定显示的窗口数(1~4)
        self._reader_hscroll_pin1 = True      # 钉死PDF1：滚轮只滑 PDF2/3/4
        self._reader_hscroll_win = 0          # 横滚窗口起始（滑动位置）
        self._reader_hscroll_pdf1_visible = True
        self._reader_pdf1_pane = None         # PDF1 在 PanedWindow 里的直接子窗格引用

        # ── 程序全局外观主题 ──
        self._app_theme = tk.StringVar(value="daylight")
        self._app_themes = {
            "graphite": {
                "name": "石墨深色", "BG": "#1e2128", "PANEL": "#282c34",
                "CARD": "#2f343f", "ENTRY": "#363b46", "BORDER": "#3a3f4b",
                "ACCENT": "#4f8cff", "GREEN": "#3ec97c",
                "TEXT": "#e6e9ef", "DIM": "#9aa0ac", "CANVAS": "#15171c"},
            "cyber": {
                "name": "科技蓝", "BG": "#07111f", "PANEL": "#0d1b2a",
                "CARD": "#10263d", "ENTRY": "#15334d", "BORDER": "#1f6f8b",
                "ACCENT": "#00bfe8", "GREEN": "#32f5a9",
                "TEXT": "#dff8ff", "DIM": "#7fb3c8", "CANVAS": "#030912"},
            "neon": {
                "name": "霓虹紫", "BG": "#120b24", "PANEL": "#1b1233",
                "CARD": "#241a42", "ENTRY": "#302453", "BORDER": "#59407d",
                "ACCENT": "#a855f7", "GREEN": "#22d3ee",
                "TEXT": "#f4eeff", "DIM": "#ad9bc8", "CANVAS": "#090512"},
            "aurora": {
                "name": "极光青", "BG": "#071a1b", "PANEL": "#0d292a",
                "CARD": "#123637", "ENTRY": "#194647", "BORDER": "#267375",
                "ACCENT": "#19c5b4", "GREEN": "#8ee66b",
                "TEXT": "#e2fffb", "DIM": "#8ebfba", "CANVAS": "#031011"},
            "titanium": {
                "name": "钛金浅色", "BG": "#e9eef5", "PANEL": "#f7f9fc",
                "CARD": "#ffffff", "ENTRY": "#e7edf5", "BORDER": "#b8c5d6",
                "ACCENT": "#2563eb", "GREEN": "#059669",
                "TEXT": "#172033", "DIM": "#64748b", "CANVAS": "#cfd8e5"},
            "pure_white": {
                "name": "纯白", "BG": "#ffffff", "PANEL": "#f8f8f8",
                "CARD": "#f0f0f0", "ENTRY": "#e8e8e8", "BORDER": "#d0d0d0",
                "ACCENT": "#1a6ef5", "GREEN": "#16a34a",
                "TEXT": "#111111", "DIM": "#888888", "CANVAS": "#ececec"},
            "dark": {
                "name": "暗黑", "BG": "#121212", "PANEL": "#1a1a1a",
                "CARD": "#242424", "ENTRY": "#2d2d2d", "BORDER": "#3a3a3a",
                "ACCENT": "#4f8cff", "GREEN": "#3ec97c",
                "TEXT": "#e6e6e6", "DIM": "#9a9a9a", "CANVAS": "#0d0d0d"},
            "daylight": {
                "name": "日间模式", "BG": "#f5f6f8", "PANEL": "#ffffff",
                "CARD": "#fafbfc", "ENTRY": "#f0f1f3", "BORDER": "#d0d3d8",
                "ACCENT": "#2563eb", "GREEN": "#16a34a",
                "TEXT": "#1a1a1a", "DIM": "#6b7280", "CANVAS": "#f6f7f9"},
            "warm_eye": {
                "name": "暖橙护眼", "BG": "#fdf6e3", "PANEL": "#fef8eb",
                "CARD": "#fffdf5", "ENTRY": "#f5edce", "BORDER": "#e0c97a",
                "ACCENT": "#d97706", "GREEN": "#2e7d32",
                "TEXT": "#3b2a0e", "DIM": "#8a6c3a", "CANVAS": "#f0e6c8"},
            "rose": {
                "name": "玫瑰粉", "BG": "#fff0f3", "PANEL": "#fff5f7",
                "CARD": "#ffffff", "ENTRY": "#ffe4ea", "BORDER": "#ffb3c1",
                "ACCENT": "#e11d48", "GREEN": "#059669",
                "TEXT": "#3d0014", "DIM": "#9f4057", "CANVAS": "#ffd6de"},
            "forest": {
                "name": "森林绿", "BG": "#f0fff4", "PANEL": "#f5fff8",
                "CARD": "#ffffff", "ENTRY": "#dcfce7", "BORDER": "#86efac",
                "ACCENT": "#15803d", "GREEN": "#0ea5e9",
                "TEXT": "#052e16", "DIM": "#4a7c59", "CANVAS": "#c6f6d5"},
            # ── 经典开发工具配色（VSCode/Atom/Sublime/GitHub 等）──
            "vscode": {
                "name": "VSCode 深色", "BG": "#1e1e1e", "PANEL": "#252526",
                "CARD": "#2d2d30", "ENTRY": "#3c3c3c", "BORDER": "#3e3e42",
                "ACCENT": "#007acc", "GREEN": "#4ec9b0",
                "TEXT": "#d4d4d4", "DIM": "#858585", "CANVAS": "#171717"},
            "one_dark": {
                "name": "One Dark", "BG": "#282c34", "PANEL": "#21252b",
                "CARD": "#2c313a", "ENTRY": "#3b4048", "BORDER": "#3e4451",
                "ACCENT": "#61afef", "GREEN": "#98c379",
                "TEXT": "#abb2bf", "DIM": "#5c6370", "CANVAS": "#1b1e24"},
            "dracula": {
                "name": "Dracula", "BG": "#282a36", "PANEL": "#21222c",
                "CARD": "#343746", "ENTRY": "#44475a", "BORDER": "#6272a4",
                "ACCENT": "#bd93f9", "GREEN": "#50fa7b",
                "TEXT": "#f8f8f2", "DIM": "#6b7394", "CANVAS": "#191a21"},
            "nord": {
                "name": "Nord", "BG": "#2e3440", "PANEL": "#3b4252",
                "CARD": "#434c5e", "ENTRY": "#4c566a", "BORDER": "#4f5b70",
                "ACCENT": "#88c0d0", "GREEN": "#a3be8c",
                "TEXT": "#eceff4", "DIM": "#d8dee9", "CANVAS": "#242933"},
            "tokyo_night": {
                "name": "Tokyo Night", "BG": "#1a1b26", "PANEL": "#16161e",
                "CARD": "#24283b", "ENTRY": "#2f334d", "BORDER": "#3b4261",
                "ACCENT": "#7aa2f7", "GREEN": "#9ece6a",
                "TEXT": "#c0caf5", "DIM": "#565f89", "CANVAS": "#13141a"},
            "github_dark": {
                "name": "GitHub 暗色", "BG": "#0d1117", "PANEL": "#161b22",
                "CARD": "#21262d", "ENTRY": "#30363d", "BORDER": "#333a42",
                "ACCENT": "#58a6ff", "GREEN": "#3fb950",
                "TEXT": "#c9d1d9", "DIM": "#8b949e", "CANVAS": "#010409"},
            "monokai": {
                "name": "Monokai", "BG": "#272822", "PANEL": "#1e1f1c",
                "CARD": "#3e3d32", "ENTRY": "#49483e", "BORDER": "#75715e",
                "ACCENT": "#fd971f", "GREEN": "#a6e22e",
                "TEXT": "#f8f8f2", "DIM": "#6e6a58", "CANVAS": "#1b1c19"},
            # ── 纯黑 + 高饱和鲜艳色（拉开色相差异，OLED/活力色系）──
            "pure_black": {
                "name": "纯黑", "BG": "#000000", "PANEL": "#0a0a0a",
                "CARD": "#141414", "ENTRY": "#1c1c1c", "BORDER": "#2e2e2e",
                "ACCENT": "#00d4ff", "GREEN": "#00e676",
                "TEXT": "#f5f5f5", "DIM": "#8a8a8a", "CANVAS": "#050505"},
            "vivid_orange": {
                "name": "活力橙", "BG": "#1a0c00", "PANEL": "#261100",
                "CARD": "#351a00", "ENTRY": "#452300", "BORDER": "#7c4000",
                "ACCENT": "#ff8c00", "GREEN": "#ffd600",
                "TEXT": "#fff4e8", "DIM": "#d09a66", "CANVAS": "#120800"},
            "vivid_cyan": {
                "name": "荧光青", "BG": "#001410", "PANEL": "#001f18",
                "CARD": "#002c22", "ENTRY": "#00392c", "BORDER": "#007356",
                "ACCENT": "#00ffc8", "GREEN": "#aaff00",
                "TEXT": "#e6fff8", "DIM": "#66c4aa", "CANVAS": "#000d09"},
            "vivid_magenta": {
                "name": "洋红", "BG": "#1a0010", "PANEL": "#260016",
                "CARD": "#350020", "ENTRY": "#450029", "BORDER": "#7c0050",
                "ACCENT": "#ff2e88", "GREEN": "#00e5ff",
                "TEXT": "#ffeaf5", "DIM": "#c466a0", "CANVAS": "#12000a"},
        }
        # ── 小米汽车色：追加车漆色主题，并记录 key 用于菜单分类分隔 ──
        self._xiaomi_theme_keys = []
        for _key, _name, _base in _XIAOMI_COLORS:
            self._app_themes[_key] = _theme_from_car_color(_name, _base)
            self._xiaomi_theme_keys.append(_key)
        # 初始 active_palette 必须等于「构建控件时实际使用的全局变量值」，
        # 而不是某个目标主题。否则首次 _apply_app_theme 时 old==new（恒等映射），
        # 控件身上那套旧全局色（#171a26 等）不在 color_map 里 → 永远不 remap。
        self._active_app_palette = {
            "BG": BG, "PANEL": PANEL, "CARD": CARD, "ENTRY": ENTRY,
            "BORDER": BORDER, "ACCENT": ACCENT, "GREEN": GREEN,
            "TEXT": TEXT, "DIM": DIM, "CANVAS": CANVAS,
        }

        # ── PDF阅读器颜色主题 ──
        self._reader_theme = tk.StringVar(value="normal")  # normal, dark, green, sepia, black_white
        self._reader_themes = {
            "normal": {"name": "正常", "bg": "#FFFFFF", "fg": "#000000", "canvas_bg": CANVAS},
            "dark": {"name": "暗黑", "bg": "#1a1a1a", "fg": "#e0e0e0", "canvas_bg": "#1a1a1a"},
            "green": {"name": "护眼绿", "bg": "#c7edcc", "fg": "#333333", "canvas_bg": "#2d3a2e"},
            "sepia": {"name": " sepia", "bg": "#f4ecd8", "fg": "#5c4b37", "canvas_bg": "#3e3226"},
            "black_white": {"name": "黑白", "bg": "#000000", "fg": "#ffffff", "canvas_bg": "#000000"},
            "blue": {"name": "护眼蓝", "bg": "#e8f0f8", "fg": "#1a1a2e", "canvas_bg": "#1a1a2e"},
        }

        self._pdf_doc1 = None     # 左侧PDF
        self._pdf_doc2 = None     # 右侧PDF
        self._cur_page1 = 0
        self._cur_page2 = 0
        self._scale1 = 1.0
        self._scale2 = 1.0
        self._tk_img1 = None
        self._tk_img2 = None
        self._cross_index = {}    # {word: {1: [(page, [rect]), ...], 2: [(page, [rect]), ...]}}
        self._cmp_last_click = None  # (side, word) 上次点击的信息
        self._cmp_hl_color = (255, 80, 80)  # 对比高亮颜色，默认红色
        self._cmp_active_hl = {1: set(), 2: set()}  # 每侧当前应持续显示的高亮词集合（缩放/翻页时不丢失）
        self._hovered_card = None   # 当前悬停的缩略图卡片
        self._hovered_page = None   # 当前悬停的页码
        self._fs_popup = None       # 全屏弹窗引用
        self._ctrl_held = False     # Ctrl键是否按下
        self._shift_held = False    # Shift键是否按下（卡片缩放）
        self._word_popups = []       # 悬停时多个出现位置弹窗列表
        self._word_popup_close_id = None  # 延迟关闭弹窗的 after id
        self._popup_enabled = True    # 悬停多弹窗开关（默认开启）
        self._popup_pad = 80          # 裁剪显示范围（PDF点数，关键词周围留白）
        self._card_zoom_popup = None  # 卡片悬停放大弹窗
        self._card_zoom_geometry = None  # 记住用户拖动后的窗口位置和大小
        self._hz_card = None         # 统合弹窗中当前悬停的卡片 frame
        self._hz_img = None          # 悬停卡片的 PIL 原图（250DPI）
        self._hz_page = None         # 悬停卡片所属的 fitz Page
        self._hz_rect = None         # 悬停卡片的目标词 fitz Rect
        self._hz_idx = 0             # 悬停卡片序号
        self._hz_pad = None          # 悬停卡片 pad 值
        self._link_result_rows = {"left": [], "right": []}

        # ── PDF对比页签：内嵌Excel差异对比（独立于「Excel对比」页签的连接）──
        self._cmpxl_app1 = None      # 左侧PDF对应的Excel COM对象
        self._cmpxl_wb1 = None       # 左侧实际选中的工作簿
        self._cmpxl_name1 = ""
        self._cmpxl_app2 = None      # 右侧PDF对应的Excel COM对象
        self._cmpxl_wb2 = None
        self._cmpxl_name2 = ""
        self._cmp_sidebar_mode = tk.StringVar(value="keyword")  # keyword | xldiff
        self._xldiff_records = []     # [{"token","status","addrs"}, ...]
        self._xldiff_listed_tokens = []  # 与 _xldiff_lb 当前显示行一一对应的 token 列表

        # ── PDF差异对比页签状态 ──
        self._diff_file1 = ""
        self._diff_file2 = ""
        self._diff_dir = ""
        self._diff_images = []
        self._diff_page_idx = 0
        self._diff_scale = 1.0
        self._diff_tk_img = None
        self._diff_action_btns = []
        self._diff_paths1 = []
        self._diff_paths2 = []
        self._diff_popup = None
        self._diff_auto_align = tk.BooleanVar(value=False)   # 自动对齐开关（按文本/位号消除整体偏移）
        self._diff_align_info = tk.StringVar(value="")        # 自动对齐结果提示

        self._build_ui()
        # 阅读窗口最大化状态
        self._reader_maximized = False       # 阅读窗口是否最大化
        self._reader_max_restore_geo = ""    # 最大化前的窗口几何信息

    # ════════════════════════════════════════════════════════════════
    # UI
    # ════════════════════════════════════════════════════════════════

    def _build_ui(self):
        # ── 顶部工具栏（外观主题切换）──
        bar = tk.Frame(self, bg=PANEL, height=52)
        bar.pack(fill=tk.X)
        bar.pack_propagate(False)
        self._app_theme_btn = RoundedButton(
            bar, text="✦ 外观", bg=TEXT,
            font=("Microsoft YaHei UI", 9, "bold"), padx=9, pady=5,
            cursor="hand2", command=self._show_app_theme_menu)
        self._app_theme_btn.pack(side=tk.RIGHT, padx=(4, 8), pady=10)

        # ── 主体区（Notebook分页）──
        self._body_frame = tk.Frame(self, bg=BG)
        self._body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        self._notebook = ttk.Notebook(self._body_frame)
        self._notebook.pack(fill=tk.BOTH, expand=True)

        # === 页签1: 文件阅读 ===
        self._tab_reader = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_reader, text="📂 文件阅读")
        self._build_pdf_reader_tab()

        # === 页签2: PDF差异对比 ===
        self._tab_pdf_diff = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_pdf_diff, text="🔍 PDF差异对比")
        self._build_pdf_diff_tab()

        # === 页签3: Excel 搜索 ===
        self._tab_excel_search = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_excel_search, text="📊 Excel 搜索")
        self._build_excel_search_tab(self._tab_excel_search)

        # ── 页签文字配色 ──
        tab_colors = {
            self._tab_reader:       "#e0a83e",  # 琥珀金 — 文件阅读
            self._tab_pdf_diff:     "#e0607a",  # 玫红色 — PDF差异对比
            self._tab_excel_search: "#4cc38a",  # 翠绿色 — Excel 搜索
        }
        for tab_widget, color in tab_colors.items():
            try:
                tab_id = str(tab_widget)
                self._notebook.tk.call(
                    self._notebook._w, "tab", tab_id, "-foreground", color)
            except Exception:
                pass

        # === 配置持久化 ===
        self._shortcut_bindings = {}
        self._config_file = os.path.join(os.path.dirname(__file__), "config.json")
        self._load_config()
        # 默认选中「文件阅读」页签
        self.after_idle(lambda: self._notebook.select(self._tab_reader))

        # ── 同步已加载的配置到按钮状态 ──
        self._apply_reader_settings_to_ui()
        self._reader_refresh_bookmark_list()

        # === 应用内快捷键（翻页/缩放/聚焦搜索）===
        self._setup_keyboard_shortcuts()

    def _toggle_reader_maximize(self):
        """切换阅读窗口最大化：隐藏工具栏和其他页签，只保留阅读区域。"""
        if not self._reader_maximized:
            # ── 进入最大化 ──
            self._reader_max_restore_geo = self.geometry()
            self._reader_maximized = True

            # 隐藏顶部工具栏
            if self._top_bar_frame.winfo_ismapped():
                self._top_bar_frame.pack_forget()

            # 隐藏工具栏手柄（折叠按钮、番茄钟、激励语）
            if self._toolbar_handle_frame.winfo_ismapped():
                self._toolbar_handle_frame.pack_forget()

            # 去掉 body 边距，让阅读区域填满窗口
            self._body_frame.pack_forget()
            self._body_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

            # 隐藏除"文件阅读"外的所有页签
            reader_tab_id = str(self._tab_reader)
            for tab_id in self._notebook.tabs():
                if tab_id != reader_tab_id:
                    try:
                        self._notebook.tab(tab_id, state="hidden")
                    except Exception:
                        pass

            # 切换到阅读页签
            self._notebook.select(self._tab_reader)

            # 最大化窗口
            self.state('zoomed')

            # 更新按钮样式
            self._reader_max_btn.config(text="⛶ 退出最大化", bg=ACCENT)
        else:
            # ── 退出最大化 ──
            self._reader_maximized = False

            # 先恢复窗口状态（在布局恢复之前）
            self.state('normal')
            if self._reader_max_restore_geo:
                self.geometry(self._reader_max_restore_geo)

            # 恢复 body 边距
            self._body_frame.pack_forget()
            self._body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

            # 恢复工具栏手柄（必须先pack，top_bar_frame 依赖它作为 before 锚点）
            if not self._toolbar_handle_frame.winfo_ismapped():
                self._toolbar_handle_frame.pack(fill=tk.X, before=self._body_frame)

            # 恢复顶部工具栏
            if not self._top_bar_frame.winfo_ismapped():
                self._top_bar_frame.pack(fill=tk.X, before=self._toolbar_handle_frame)

            # 恢复所有页签
            for tab_id in self._notebook.tabs():
                try:
                    self._notebook.tab(tab_id, state="normal")
                except Exception:
                    pass

            # 更新按钮样式
            self._reader_max_btn.config(text="⛶ 最大化阅读", bg="#454b63")

    # ════════════════════════════════════════════════════════════════
    # 顶部细条 - 激励语（自由编辑）
    # ════════════════════════════════════════════════════════════════




    # ════════════════════════════════════════════════════════════════
    # 顶部细条 - 番茄时钟
    # ════════════════════════════════════════════════════════════════















    # ── 小工具启动 ──
    _TOOLS = [
        ("🔁 dBm/mV 换算器", "Bode-100", "Bode-100-dbm-mV.py"),
        ("🖱 鼠标定位助手", "mouse-location", "mouse_location_try_3.py"),
    ]






    # ════════════════════════════════════════════════════════════════


















# Excel对比页签（新增）
    # ════════════════════════════════════════════════════════════════






    # ════════════════════════════════════════════════════════════════
    # PDF
    # ════════════════════════════════════════════════════════════════





    def _clean_word(self, word: str) -> str:
        """清理PDF提取的词，去除常见标点和不必要字符。"""
        return word.strip().strip('.,;:\'"()[]{}').strip()














    # ════════════════════════════════════════════════════════════════
    # Excel / WPS COM（通用）
    # ════════════════════════════════════════════════════════════════






    # ════════════════════════════════════════════════════════════════
    # Excel对比模式 - 连接管理
    # ════════════════════════════════════════════════════════════════



    # ════════════════════════════════════════════════════════════════
    # Excel对比模式 - 核心逻辑
    # ════════════════════════════════════════════════════════════════









    # ════════════════════════════════════════════════════════════════
    # 跟随模式轮询
    # ════════════════════════════════════════════════════════════════



    # ════════════════════════════════════════════════════════════════
    # 全量扫描模式
    # ════════════════════════════════════════════════════════════════


    # ════════════════════════════════════════════════════════════════
    # 跟随模式辅助
    # ════════════════════════════════════════════════════════════════





    # ════════════════════════════════════════════════════════════════
    # 清除高亮
    # ════════════════════════════════════════════════════════════════





    # ════════════════════════════════════════════════════════════════
    # PDF双文件对比页签
    # ════════════════════════════════════════════════════════════════





    # ── 文件打开 ──




    # ── 渲染 ──




    # ── 翻页 & 缩放 ──




    # ── 点击同步 ──




    # ── 交叉索引构建 ──









    # ════════════════════════════════════════════════════════════════
    # PDF对比页签 - 内嵌Excel差异对比（独立连接，仅用于驱动PDF高亮，不修改Excel）
    # ════════════════════════════════════════════════════════════════












    # ════════════════════════════════════════════════════════════════
    # 多页展开视图
    # ════════════════════════════════════════════════════════════════






    # ════════════════════════════════════════════════════════════════
    # 全屏弹窗放大（Ctrl+悬停）
    # ════════════════════════════════════════════════════════════════








    # ════════════════════════════════════════════════════════════════
    # 悬停弹窗 — 单个滚动窗口内网格展示所有出现位置
    # ════════════════════════════════════════════════════════════════




    # ── 卡片悬停放大 ──






    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 配置持久化
    # ════════════════════════════════════════════════════════════════

    def _apply_reader_settings_to_ui(self):
        """将已加载的阅读器设置同步到按钮/UI 状态。"""
        # 智能裁边按钮
        if hasattr(self, "_reader_crop_btn"):
            if self._reader_crop_margins_mode:
                self._reader_crop_btn.config(text="✂ 智能裁边(开)", bg=ACCENT)
            else:
                self._reader_crop_btn.config(text="✂ 智能裁边", bg="#454b63")
        # 适应宽度按钮
        if hasattr(self, "_reader_fit_btn"):
            if self._reader_fit_width:
                self._reader_fit_btn.config(text="🔲 适应宽度", bg=ACCENT)
            else:
                self._reader_fit_btn.config(text="🔲 自由缩放", bg="#454b63")
        # 滚动速度预设下拉同步
        if hasattr(self, "_reader_speed_combo_var"):
            self._reader_scroll_lbl_var.set(str(self._reader_scroll_lines))
            self._reader_sync_speed_combo()
        # 阅读器画布背景色
        theme = self._reader_theme.get()
        if theme in self._reader_themes and hasattr(self, "_reader_canvas"):
            bg_color = self._reader_themes[theme].get("canvas_bg", CANVAS)
            self._reader_canvas.config(bg=bg_color)
        # 联动搜索 / 弹窗模式 / 匹配块 按钮状态同步（需求1：默认状态）
        if hasattr(self, "_reader_linked_btn"):
            ll_colors = {True: GREEN, False: "#454b63"}
            ll_labels = {True: "🔗 联动搜索:开", False: "🔗 联动搜索:关"}
            self._reader_linked_btn.config(
                text=ll_labels[self._reader_linked_search.get()],
                bg=ll_colors[self._reader_linked_search.get()])
        if hasattr(self, "_reader_click_popup_btn"):
            pop_labels = {"full": "📋 弹窗:完整", "minimal": "📋 弹窗:简略", "hidden": "📋 弹窗:关闭"}
            pop_colors = {"full": GREEN, "minimal": ACCENT, "hidden": "#454b63"}
            self._reader_click_popup_btn.config(
                text=pop_labels[self._reader_click_popup_mode.get()],
                bg=pop_colors[self._reader_click_popup_mode.get()])
        if hasattr(self, "_reader_match_nav_btn"):
            self._reader_match_nav_btn.config(
                bg=ACCENT if self._reader_match_visible else "#454b63")

    def _load_config(self):
        import json
        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            saved_shortcuts = data.get("shortcut_bindings", {})
            if isinstance(saved_shortcuts, dict):
                for sid, bl in saved_shortcuts.items():
                    if isinstance(bl, list):
                        seqs = [str(b) for b in bl if isinstance(b, str) and b]
                        if seqs:
                            self._shortcut_bindings[str(sid)] = seqs
            self._reader_dir = data.get("pdf_folder", "")
            saved_highlight = str(data.get("reader_highlight_color", HIGHLIGHT)).lower()
            if re.fullmatch(r"#[0-9a-f]{6}", saved_highlight):
                self._reader_highlight_hex.set(saved_highlight)
                self._reader_highlight_color = tuple(
                    int(saved_highlight[i:i + 2], 16) for i in (1, 3, 5))
                if hasattr(self, "_reader_highlight_btn"):
                    self._reader_highlight_btn.config(bg=saved_highlight, fg="#111111")
            open_counts = data.get("reader_open_counts", {})
            if isinstance(open_counts, dict):
                cleaned_counts = {}
                for k, v in open_counts.items():
                    try:
                        n = int(v)
                        if n > 0:
                            cleaned_counts[str(k)] = n
                    except (TypeError, ValueError):
                        continue
                self._open_counts = cleaned_counts
            resume = data.get("reader_resume", {})
            if isinstance(resume, dict):
                self._reader_resume = {
                    str(k): dict(v) for k, v in resume.items()
                    if isinstance(v, dict)
                }
            annotations = data.get("reader_annotations", {})
            if isinstance(annotations, dict):
                cleaned = {}
                for path, items in annotations.items():
                    if not isinstance(items, list):
                        continue
                    valid_items = []
                    for it in items:
                        if not isinstance(it, dict):
                            continue
                        try:
                            page = int(it.get("page", 0))
                        except (TypeError, ValueError):
                            continue
                        valid_items.append({
                            "page": page,
                            "label": str(it.get("label", "")),
                            "note": str(it.get("note", "")),
                        })
                    if valid_items:
                        cleaned[str(path)] = valid_items
                self._reader_annotations = cleaned

            marks = data.get("reader_page_marks", {})
            if isinstance(marks, dict):
                cleaned_marks = {}
                for path, by_page in marks.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, rect_list in by_page.items():
                        if not isinstance(rect_list, list):
                            continue
                        valid_rects = []
                        for it in rect_list:
                            if not isinstance(it, dict):
                                continue
                            rect = it.get("rect")
                            if (isinstance(rect, list) and len(rect) == 4 and
                                    all(isinstance(v, (int, float)) for v in rect)):
                                valid_rects.append({
                                    "rect": [float(v) for v in rect],
                                    "color": str(it.get("color", "#ffdc00")),
                                })
                        if valid_rects:
                            cleaned_pages[str(page_key)] = valid_rects
                    if cleaned_pages:
                        cleaned_marks[str(path)] = cleaned_pages
                self._reader_page_marks = cleaned_marks

            text_marks = data.get("reader_text_marks", {})
            if isinstance(text_marks, dict):
                cleaned_tm = {}
                for path, by_page in text_marks.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, tm_list in by_page.items():
                        if not isinstance(tm_list, list):
                            continue
                        valid = []
                        for it in tm_list:
                            if not isinstance(it, dict):
                                continue
                            x = it.get("x"); y = it.get("y")
                            if not (isinstance(x, (int, float)) and isinstance(y, (int, float))):
                                continue
                            valid.append({
                                "x": float(x), "y": float(y),
                                "text": str(it.get("text", "")),
                                "color": str(it.get("color", "#ffdc00")),
                            })
                        if valid:
                            cleaned_pages[str(page_key)] = valid
                    if cleaned_pages:
                        cleaned_tm[str(path)] = cleaned_pages
                self._reader_text_marks = cleaned_tm

            text_hl = data.get("reader_text_hl", {})
            if isinstance(text_hl, dict):
                cleaned_hl = {}
                for path, by_page in text_hl.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, hl_list in by_page.items():
                        if not isinstance(hl_list, list):
                            continue
                        valid = []
                        for it in hl_list:
                            if not isinstance(it, dict):
                                continue
                            rect = it.get("rect")
                            if not (isinstance(rect, list) and len(rect) == 4 and
                                    all(isinstance(v, (int, float)) for v in rect)):
                                continue
                            valid.append({
                                "rect": [float(v) for v in rect],
                                "text": str(it.get("text", "")),
                                "color": str(it.get("color", "#ff5050")),
                            })
                        if valid:
                            cleaned_pages[str(page_key)] = valid
                    if cleaned_pages:
                        cleaned_hl[str(path)] = cleaned_pages
                self._reader_text_hl = cleaned_hl
            # ── 阅读器设置 ──
            saved_reader_theme = data.get("reader_theme", "normal")
            if saved_reader_theme in ("normal", "dark", "green", "sepia", "black_white"):
                self._reader_theme.set(saved_reader_theme)
                theme_bg = self._reader_themes[saved_reader_theme].get("canvas_bg", CANVAS)
                if hasattr(self, "_reader_canvas"):
                    self._reader_canvas.config(bg=theme_bg)
            saved_crop = data.get("reader_crop_margins", False)
            if isinstance(saved_crop, bool):
                self._reader_crop_margins_mode = saved_crop
            saved_fit = data.get("reader_fit_width", True)
            if isinstance(saved_fit, bool):
                self._reader_fit_width = saved_fit
            saved_scroll = data.get("reader_scroll_lines", 3)
            if isinstance(saved_scroll, int) and 1 <= saved_scroll <= 20:
                self._reader_scroll_lines = saved_scroll
                if hasattr(self, "_reader_speed_combo_var"):
                    self._reader_sync_speed_combo()
            # 手动框选正文区
            saved_body = data.get("reader_body_crop_regions", {})
            if isinstance(saved_body, dict):
                self._reader_body_crop_regions = {}
                for k, v in saved_body.items():
                    if isinstance(v, (list, tuple)) and len(v) == 4:
                        try:
                            self._reader_body_crop_regions[str(k)] = tuple(float(x) for x in v)
                        except (TypeError, ValueError):
                            pass
            # Excel 搜索页签
            _esr = data.get("excel_search_search_path")
            if isinstance(_esr, str) and _esr:
                self._excel_search_search_path = _esr
            _esrp = data.get("excel_search_recent_paths")
            if isinstance(_esrp, list):
                self._excel_search_recent_paths = [p for p in _esrp if isinstance(p, str)][:10]
            if isinstance(data.get("excel_search_recursive"), bool):
                self._excel_search_recursive.set(data["excel_search_recursive"])
            # 外观主题
            saved_theme = data.get("app_theme", "daylight")
            if saved_theme in self._app_themes:
                self._app_theme.set(saved_theme)
                self.after_idle(lambda: self._apply_app_theme(persist=False))
            # 搜索历史（Everything 等搜索栏）
            search_histories = data.get("search_histories", {})
            if isinstance(search_histories, dict):
                cleaned_hist = {}
                for bar_key, inner in search_histories.items():
                    if not isinstance(inner, dict):
                        continue
                    cleaned_inner = {}
                    for kw, cnt in inner.items():
                        try:
                            n = int(cnt)
                            if n > 0:
                                cleaned_inner[str(kw)] = n
                        except (TypeError, ValueError):
                            continue
                    if cleaned_inner:
                        cleaned_hist[str(bar_key)] = cleaned_inner
                self._search_histories = cleaned_hist
        except Exception:
            self._reader_dir = ""
    def _save_config(self):
        import json
        try:
            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump({
                    "pdf_folder": self._reader_dir,
                    "reader_highlight_color": self._reader_highlight_hex.get(),
                    "reader_open_counts": self._open_counts,
                    "reader_resume": self._reader_resume,
                    "reader_annotations": self._reader_annotations,
                    "reader_page_marks": self._reader_page_marks,
                    "reader_text_marks": self._reader_text_marks,
                    "reader_text_hl": self._reader_text_hl,
                    "reader_theme": self._reader_theme.get(),
                    "reader_crop_margins": self._reader_crop_margins_mode,
                    "reader_fit_width": self._reader_fit_width,
                    "reader_scroll_lines": self._reader_scroll_lines,
                    "reader_body_crop_regions": {
                        k: list(v) for k, v in self._reader_body_crop_regions.items()},
                    "excel_search_search_path": self._excel_search_search_path,
                    "excel_search_recent_paths": self._excel_search_recent_paths,
                    "excel_search_recursive": self._excel_search_recursive.get(),
                    "app_theme": self._app_theme.get(),
                    "search_histories": getattr(self, "_search_histories", {}),
                    "shortcut_bindings": getattr(self, "_shortcut_bindings", {}),
                }, f, indent=2)
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════════
    # PDF差异对比页签 - 核心逻辑与UI
    # 集成自 pdf_diff_optimized.py
    # ════════════════════════════════════════════════════════════════

    # Anaglyph 矩阵定义（预转为 numpy float32）
    _zero_vec = [0, 0, 0]
    _ident_mat = [[1, 0, 0], [0, 1, 0], [0, 0, 1]]
    _anaglyph = (
        [[0, 0, 1.0], _zero_vec, _zero_vec],   # 旧版 → 纯蓝色
        [_zero_vec, _ident_mat[1], _ident_mat[2]]  # 新版 → 绿+蓝通道
    )
    _MAT_OLD = np.array(_anaglyph[0], dtype=np.float32)
    _MAT_NEW = np.array(_anaglyph[1], dtype=np.float32)

    def _diff_parse_pages(self, pages_str: str) -> list:
        """解析页码字符串，支持逗号分隔和连字符范围。"""
        pages = []
        for part in pages_str.strip().split(','):
            part = part.strip()
            if '-' in part:
                start, end = part.split('-', 1)
                pages.extend(range(int(start.strip()), int(end.strip()) + 1))
            else:
                pages.append(int(part))
        return pages

    def _diff_pdf2img(self, doc, page_number: int, target_dpi: int = 600) -> Image.Image:
        """直接将PDF指定页渲染为PIL Image，不写临时文件（复用已打开的doc）。"""
        page = doc.load_page(page_number - 1)
        zoom = target_dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def _diff_shift_image(self, img: Image.Image, sx: int, sy: int, fill: int = 255):
        """平移图像 (sx,sy) 像素，空出区域填 fill（默认白，适配原理图白底）。

        正值表示内容向右/下移动。用 numpy 切片做零拷贝式平移，返回新 PIL 图。
        """
        if not sx and not sy:
            return img
        arr = np.asarray(img, dtype=np.uint8)
        h, w = arr.shape[:2]
        out = np.full_like(arr, fill)
        src_x0 = max(0, -sx); src_x1 = min(w, w - sx)
        src_y0 = max(0, -sy); src_y1 = min(h, h - sy)
        dst_x0 = max(0, sx); dst_x1 = min(w, w + sx)
        dst_y0 = max(0, sy); dst_y1 = min(h, h + sy)
        sw = src_x1 - src_x0; sh = src_y1 - src_y0
        if sw > 0 and sh > 0:
            out[dst_y0:dst_y0 + sh, dst_x0:dst_x0 + sw] = \
                arr[src_y0:src_y1, src_x0:src_x1]
        img.close()
        return Image.fromarray(out, "RGB")

    def _diff_auto_align_offset(self, doc1, page1: int, doc2, page2: int):
        """基于文本/位号匹配，估计 page2 相对 page1 的整体平移 (dx, dy)（PDF点）。

        优先用两侧都出现的位号（R1/C1/U3…，唯一性强、最可靠），否则退化为
        任意公共文本；对多组配对取中位数，抗个别错位号/错词。无公共文本返回 None。
        """
        try:
            p1 = doc1[page1 - 1]
            p2 = doc2[page2 - 1]
        except Exception:
            return None

        def _centers(page):
            d = {}
            try:
                words = page.get_text("words")
            except Exception:
                return d
            for w in words:
                try:
                    x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], w[4]
                except (IndexError, TypeError, ValueError):
                    continue
                key = re.sub(r"[^0-9a-zA-Z]", "", txt).lower()
                if not key:
                    continue
                d.setdefault(key, []).append(((x0 + x1) / 2.0, (y0 + y1) / 2.0))
            return d

        c1 = _centers(p1)
        c2 = _centers(p2)
        common = set(c1) & set(c2)
        if not common:
            return None
        refdes = [k for k in common if re.fullmatch(r"[a-z]+\d+", k)]
        pool = refdes if refdes else list(common)
        dxs, dys = [], []
        for k in pool:
            dxs.append(c2[k][0][0] - c1[k][0][0])
            dys.append(c2[k][0][1] - c1[k][0][1])
        if not dxs:
            return None
        dxs.sort(); dys.sort()
        return (dxs[len(dxs) // 2], dys[len(dys) // 2])

    def _diff_process_images(self, page1: int, page2: int, doc1, doc2,
                              target_dpi: int = 600, offset=None):
        """对比两个PDF的指定页，返回 Anaglyph 差异图像（复用已打开的文档句柄）。

        offset: 可选 (dx, dy) PDF点偏移，表示 page2 内容相对 page1 的整体平移。
        传入后先把 page2 反向平移对齐到 page1，再叠加，用于解决"同一电路只是
        页面坐标整体偏移"导致的整页误判为差异的问题。为 None 时保持原行为。

        合成等价于旧版 float32 矩阵乘：旧版蓝→红、新版绿→绿、新版蓝→蓝，
        这里直接用 uint8 通道赋值，内存约为原来的 1/4，无需 clip。
        """
        img1 = self._diff_pdf2img(doc1, page1, target_dpi=target_dpi)
        img2 = self._diff_pdf2img(doc2, page2, target_dpi=target_dpi)

        if offset is not None and (offset[0] or offset[1]):
            zoom = target_dpi / 72.0
            img2 = self._diff_shift_image(
                img2, int(round(-offset[0] * zoom)), int(round(-offset[1] * zoom)))

        target_w = max(img1.width, img2.width)
        target_h = max(img1.height, img2.height)
        if img1.size != (target_w, target_h):
            img1 = img1.resize((target_w, target_h), Image.LANCZOS)
        if img2.size != (target_w, target_h):
            img2 = img2.resize((target_w, target_h), Image.LANCZOS)

        arr1 = np.asarray(img1, dtype=np.uint8)
        img1.close(); del img1
        arr2 = np.asarray(img2, dtype=np.uint8)
        img2.close(); del img2

        result = np.empty(arr1.shape, dtype=np.uint8)
        result[..., 0] = arr1[..., 2]   # 旧版蓝 → 红
        result[..., 1] = arr2[..., 1]   # 新版绿 → 绿
        result[..., 2] = arr2[..., 2]   # 新版蓝 → 蓝
        del arr1, arr2

        diff_img = Image.fromarray(result, "RGB")
        del result
        return diff_img

    def _build_pdf_diff_tab(self):
        """构建PDF差异对比页签"""
        parent = self._tab_pdf_diff
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # ── 顶部工具栏卡片 ──
        card = tk.Frame(parent, bg=CARD, bd=0,
                        highlightbackground=BORDER, highlightthickness=1)
        card.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 6))

        bf = tk.Frame(card, bg=CARD)
        bf.grid(row=0, column=0, sticky="w", padx=14, pady=12)

        # 第1行：旧PDF
        RoundedButton(bf, text="📂 选择旧PDF", command=self._diff_select_doc1,
                  bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9, "bold"), padx=12, pady=7,
                  cursor="hand2").grid(row=0, column=0, sticky=tk.W, pady=4)
        b_list1 = RoundedButton(bf, text="▼", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                            font=("Microsoft YaHei UI", 9), width=2)
        b_list1.grid(row=0, column=1, sticky=tk.W, padx=(2, 0))
        b_list1.bind("<Enter>", lambda e: self._diff_show_popup(b_list1, self._diff_paths1, doc=1))

        tk.Label(bf, text="页码", bg=CARD, fg=DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=0, column=2, padx=(14, 4))
        self._diff_pages1 = tk.Entry(bf, width=14, bg=ENTRY, fg=TEXT,
                                      insertbackground=TEXT, relief=tk.FLAT,
                                      font=("Microsoft YaHei UI", 10))
        self._diff_pages1.grid(row=0, column=3, sticky=tk.W)
        tk.Label(bf, text="例 1,3,5-7", bg=CARD, fg=DIM,
                 font=("Microsoft YaHei UI", 8)).grid(row=0, column=4, padx=8)
        RoundedButton(bf, text="📁 扫描文件夹", command=lambda: self._diff_scan_folder(self._diff_paths1, doc=1),
                  bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9), padx=10, pady=6,
                  cursor="hand2").grid(row=0, column=5, sticky=tk.W, padx=4)

        # 第2行：新PDF
        RoundedButton(bf, text="📂 选择新PDF", command=self._diff_select_doc2,
                  bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9, "bold"), padx=12, pady=7,
                  cursor="hand2").grid(row=1, column=0, sticky=tk.W, pady=4)
        b_list2 = RoundedButton(bf, text="▼", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                            font=("Microsoft YaHei UI", 9), width=2)
        b_list2.grid(row=1, column=1, sticky=tk.W, padx=(2, 0))
        b_list2.bind("<Enter>", lambda e: self._diff_show_popup(b_list2, self._diff_paths2, doc=2))

        tk.Label(bf, text="页码", bg=CARD, fg=DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=1, column=2, padx=(14, 4))
        self._diff_pages2 = tk.Entry(bf, width=14, bg=ENTRY, fg=TEXT,
                                      insertbackground=TEXT, relief=tk.FLAT,
                                      font=("Microsoft YaHei UI", 10))
        self._diff_pages2.grid(row=1, column=3, sticky=tk.W)
        tk.Label(bf, text="DPI", bg=CARD, fg=DIM,
                 font=("Microsoft YaHei UI", 9)).grid(row=1, column=4, padx=(14, 4), sticky=tk.E)
        self._diff_dpi = tk.Entry(bf, width=6, bg=ENTRY, fg=TEXT,
                                   insertbackground=TEXT, relief=tk.FLAT,
                                   font=("Microsoft YaHei UI", 10))
        self._diff_dpi.insert(0, "600")
        self._diff_dpi.grid(row=1, column=5, sticky=tk.W)

        # 分隔线
        tk.Frame(card, bg=BORDER, height=1).grid(
            row=1, column=0, sticky="ew", padx=14, pady=2)

        # 第3行：操作按钮
        af = tk.Frame(card, bg=CARD)
        af.grid(row=2, column=0, sticky="w", padx=14, pady=(8, 12))

        b_preview = RoundedButton(af, text="👁  预览", command=self._diff_preview,
                              bg=ACCENT, fg="white", relief=tk.FLAT,
                              font=("Microsoft YaHei UI", 9, "bold"), padx=12, pady=7,
                              cursor="hand2")
        b_preview.grid(row=0, column=0, padx=(0, 6))
        b_save = RoundedButton(af, text="💾  对比 & 保存", command=self._diff_compare_and_save,
                           bg=GREEN, fg="white", relief=tk.FLAT,
                           font=("Microsoft YaHei UI", 9, "bold"), padx=12, pady=7,
                           cursor="hand2")
        b_save.grid(row=0, column=1, padx=6)
        b_prev = RoundedButton(af, text="◀  上一页", command=self._diff_previous_page,
                           bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                           font=("Microsoft YaHei UI", 9), padx=10, pady=6,
                           cursor="hand2")
        b_prev.grid(row=0, column=2, padx=(20, 6))
        b_next = RoundedButton(af, text="下一页  ▶", command=self._diff_next_page,
                           bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                           font=("Microsoft YaHei UI", 9), padx=10, pady=6,
                           cursor="hand2")
        b_next.grid(row=0, column=3, padx=6)
        self._diff_action_btns = [b_preview, b_save, b_prev, b_next]

        # 进度标签
        self._diff_progress = tk.StringVar(value="就绪")
        tk.Label(af, textvariable=self._diff_progress, bg=CARD,
                 fg=ACCENT, font=("Microsoft YaHei UI", 9)).grid(
                 row=0, column=4, padx=20, sticky=tk.W)

        # ── 自动对齐（消除"同一电路、页面坐标整体偏移"导致的整页误判差异）──
        align_f = tk.Frame(card, bg=CARD)
        align_f.grid(row=3, column=0, sticky="w", padx=14, pady=(0, 12))
        tk.Checkbutton(align_f, text="🔗 自动对齐（按文本/位号消除整体偏移）",
                       variable=self._diff_auto_align, bg=CARD, fg=TEXT,
                       selectcolor=ENTRY, activebackground=CARD, activeforeground=TEXT,
                       font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
        tk.Label(align_f, textvariable=self._diff_align_info, bg=CARD, fg=DIM,
                 font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT, padx=10)

        # ── 画布区 ──
        canvas_frame = tk.Frame(parent, bg=CARD,
                                highlightbackground=BORDER, highlightthickness=1)
        canvas_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(6, 14))
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self._diff_canvas = tk.Canvas(canvas_frame, bg=CANVAS,
                                       width=800, height=900, highlightthickness=0, bd=0)
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self._diff_canvas.xview)
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self._diff_canvas.yview)
        self._diff_canvas.config(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        self._diff_canvas.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        self._diff_canvas.bind("<MouseWheel>", self._diff_zoom_image)
        self._diff_canvas.bind("<ButtonPress-1>", self._diff_start_move)
        self._diff_canvas.bind("<B1-Motion>", self._diff_move_image)

    # ════════════════════════════════════════════════════════════════
    # BJYZ 页签：左侧只读笔记浏览 + 右侧内嵌浏览器
    # ════════════════════════════════════════════════════════════════





    # ── BJYZ笔记：独立的CRUD和序列化（不依赖_notes_text等悬浮窗变量）──








































    # ════════════════════════════════════════════════════════════════
    # 笔记功能（PDF差异对比页签右侧）：文本+图片+链接，自动保存到指定文件夹
    # ════════════════════════════════════════════════════════════════






































    # ── PDF差异对比 - 文件选择 ──

    def _diff_fill_pages(self, pdf_path: str, entry: tk.Entry):
        """读取PDF总页数，自动填入页码输入框。"""
        doc = fitz.open(pdf_path)
        n = doc.page_count
        doc.close()
        entry.delete(0, tk.END)
        entry.insert(0, f"1-{n}" if n > 1 else "1")

    def _diff_select_doc1(self):
        path = filedialog.askopenfilename(
            title="选择旧PDF原理图",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if path:
            self._diff_file1 = path
            self._diff_fill_pages(path, self._diff_pages1)

    def _diff_select_doc2(self):
        path = filedialog.askopenfilename(
            title="选择新PDF原理图",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if path:
            self._diff_file2 = path
            self._diff_fill_pages(path, self._diff_pages2)

    def _diff_scan_folder(self, paths: list, doc: int):
        """扫描文件夹，结果存入paths，刷新对应弹窗。"""
        folder = filedialog.askdirectory(title="选择要扫描的文件夹")
        if not folder:
            return
        found = []
        for root, _, files in os.walk(folder):
            for f in sorted(files):
                if f.lower().endswith(".pdf"):
                    found.append(os.path.join(root, f))
        paths.clear()
        paths.extend(found)
        paths.insert(0, folder)
        if not found:
            messagebox.showinfo("提示", "该文件夹下未找到PDF文件")

    def _diff_show_popup(self, trigger_btn, paths: list, doc: int):
        """鼠标悬停在▼按钮上时，弹出PDF列表悬浮窗。"""
        self._diff_close_popup()
        if len(paths) < 2:
            return
        folder = paths[0]
        pdf_paths = paths[1:]

        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg=ACCENT)

        x = trigger_btn.winfo_rootx()
        y = trigger_btn.winfo_rooty() + trigger_btn.winfo_height()
        popup.geometry(f"360x340+{x}+{y}")

        frame = tk.Frame(popup, bg=CARD, bd=0)
        frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        title = "旧PDF — 点击选择" if doc == 1 else "新PDF — 点击选择"
        tk.Label(frame, text=title, bg=PANEL, fg=TEXT,
                 font=("Microsoft YaHei UI", 9, "bold"), anchor="w",
                 padx=10, pady=6).pack(fill=tk.X)

        body = tk.Frame(frame, bg=CARD)
        body.pack(fill=tk.BOTH, expand=True)

        lb = tk.Listbox(body, font=("Microsoft YaHei UI", 9),
                        bg=CARD, fg=TEXT,
                        selectbackground=ACCENT, selectforeground="white",
                        highlightthickness=0, bd=0, activestyle="none",
                        relief=tk.FLAT)
        sb = ttk.Scrollbar(body, command=lb.yview)
        lb.config(yscrollcommand=sb.set)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0), pady=4)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        for p in pdf_paths:
            lb.insert(tk.END, os.path.relpath(p, folder))

        def on_select(e):
            sel = lb.curselection()
            if not sel:
                return
            path = pdf_paths[sel[0]]
            if doc == 1:
                self._diff_file1 = path
                self._diff_fill_pages(path, self._diff_pages1)
            else:
                self._diff_file2 = path
                self._diff_fill_pages(path, self._diff_pages2)
            self._diff_close_popup()

        lb.bind("<ButtonRelease-1>", on_select)

        def schedule_close(e=None):
            self._diff_close_job = self.after(300, self._diff_close_popup)

        def cancel_close(e=None):
            if getattr(self, "_diff_close_job", None):
                self.after_cancel(self._diff_close_job)
                self._diff_close_job = None

        for w in (popup, frame, body, lb, sb):
            w.bind("<Enter>", cancel_close)
            w.bind("<Leave>", schedule_close)
        trigger_btn.bind("<Leave>", schedule_close)

        self._diff_popup = popup

    def _diff_close_popup(self):
        if getattr(self, "_diff_close_job", None):
            try:
                self.after_cancel(self._diff_close_job)
            except Exception:
                pass
            self._diff_close_job = None
        if self._diff_popup:
            try:
                self._diff_popup.destroy()
            except Exception:
                pass
            self._diff_popup = None

    # ── PDF差异对比 - 对比入口 ──

    def _diff_preview(self):
        self._diff_start_comparison(preview=True)

    def _diff_compare_and_save(self):
        diff_dir = filedialog.askdirectory(title="选择保存对比结果的目录")
        if not diff_dir:
            return
        self._diff_dir = diff_dir
        self._diff_start_comparison(preview=False)

    def _diff_validate_inputs(self) -> tuple:
        """校验输入，返回 (pages1, pages2) 或 None。"""
        if not self._diff_file1 or not self._diff_file2:
            messagebox.showerror("错误", "请先选择两份PDF文件")
            return None
        pages_str1 = self._diff_pages1.get().strip()
        pages_str2 = self._diff_pages2.get().strip()
        if not pages_str1 or not pages_str2:
            messagebox.showerror("错误", "请输入页码范围\n示例：1,3,5-7")
            return None
        try:
            pages1 = self._diff_parse_pages(pages_str1)
            pages2 = self._diff_parse_pages(pages_str2)
        except ValueError:
            messagebox.showerror("格式错误", "页码格式不正确\n示例：1,3,5-7 或 2-10")
            return None
        if len(pages1) != len(pages2):
            messagebox.showerror("页数不匹配",
                                 f"旧PDF选了 {len(pages1)} 页，新PDF选了 {len(pages2)} 页\n两边页数必须相同")
            return None
        return pages1, pages2

    def _diff_set_buttons(self, state):
        for b in self._diff_action_btns:
            b.config(state=state)

    def _diff_start_comparison(self, preview: bool):
        result = self._diff_validate_inputs()
        if result is None:
            return
        pages1, pages2 = result
        try:
            target_dpi = int(self._diff_dpi.get().strip())
            if target_dpi < 72 or target_dpi > 2400:
                raise ValueError
        except ValueError:
            messagebox.showerror("DPI错误", "渲染DPI请填72~2400之间的整数")
            return

        self._diff_set_buttons(tk.DISABLED)
        self._diff_progress.set(f"正在对比 {len(pages1)} 页（{target_dpi} DPI），请稍候...")
        self._diff_align_info.set("")
        self.update()
        auto_align = self._diff_auto_align.get()

        def worker():
            doc1 = doc2 = None
            try:
                doc1 = fitz.open(self._diff_file1)
                doc2 = fitz.open(self._diff_file2)
                images = []
                total = len(pages1)
                for i, (p1, p2) in enumerate(zip(pages1, pages2)):
                    self.after(0, lambda i=i: self._diff_progress.set(
                        f"正在对比 {i+1}/{total} 页（{target_dpi} DPI）..."))
                    offset = None
                    if auto_align:
                        offset = self._diff_auto_align_offset(doc1, p1, doc2, p2)
                        if offset is not None:
                            self.after(0, lambda o=offset, i=i: self._diff_align_info.set(
                                f"第{i+1}页已对齐 dx={o[0]:.1f}, dy={o[1]:.1f} pt"))
                        else:
                            self.after(0, lambda i=i: self._diff_align_info.set(
                                f"第{i+1}页未找到可对齐文本，按原样对比"))
                    img = self._diff_process_images(p1, p2, doc1, doc2,
                                                     target_dpi=target_dpi, offset=offset)
                    if img is not None:
                        images.append(img)
                self.after(0, lambda: self._diff_on_done(images, preview))
            except Exception as e:
                self.after(0, lambda: self._diff_on_error(str(e)))
            finally:
                if doc1 is not None:
                    doc1.close()
                if doc2 is not None:
                    doc2.close()

        threading.Thread(target=worker, daemon=True).start()

    def _diff_on_done(self, images: list, preview: bool):
        self._diff_set_buttons(tk.NORMAL)
        self._diff_progress.set("就绪")
        self._diff_images = images
        self._diff_page_idx = 0

        if not images:
            messagebox.showwarning("无结果", "未生成任何对比图，请检查页码是否有效")
            return

        if not preview:
            import io
            pdf_path = os.path.join(self._diff_dir, "diff_result.pdf")
            first = images[0].convert("RGB")
            rest = [img.convert("RGB") for img in images[1:]]
            first.save(pdf_path, save_all=True, append_images=rest, resolution=100.0)
            del first, rest
            messagebox.showinfo("完成", f"已保存 {len(images)} 页对比结果到：\n{pdf_path}")

        self._diff_show_page(0)

    def _diff_on_error(self, msg: str):
        self._diff_set_buttons(tk.NORMAL)
        self._diff_progress.set("就绪")
        messagebox.showerror("处理出错", f"对比过程中发生错误：\n{msg}")

    # ── 翻页 & 显示 ──

    def _diff_show_page(self, index: int):
        if 0 <= index < len(self._diff_images):
            self._diff_page_idx = index
            self._diff_original_image = self._diff_images[index]
            total = len(self._diff_images)
            self._diff_progress.set(f"第 {index + 1} / {total} 页")

            canvas_w = self._diff_canvas.winfo_width() or 800
            img_w = self._diff_original_image.width
            if img_w > 0:
                self._diff_scale = max(0.05, min(canvas_w / img_w, 1.0))

            self._diff_display_image()

    def _diff_previous_page(self):
        if self._diff_page_idx > 0:
            self._diff_show_page(self._diff_page_idx - 1)

    def _diff_next_page(self):
        if self._diff_page_idx < len(self._diff_images) - 1:
            self._diff_show_page(self._diff_page_idx + 1)

    def _diff_display_image(self):
        w = int(self._diff_original_image.width * self._diff_scale)
        h = int(self._diff_original_image.height * self._diff_scale)
        interp = Image.LANCZOS if self._diff_scale < 1.0 else Image.NEAREST
        scaled = self._diff_original_image.resize((w, h), interp)
        self._diff_tk_img = ImageTk.PhotoImage(scaled)

        self._diff_canvas.delete("all")
        self._diff_canvas.create_image(0, 0, anchor="nw", image=self._diff_tk_img)
        self._diff_canvas.config(scrollregion=self._diff_canvas.bbox("all"))

    # ── 缩放 & 拖拽 ──

    def _diff_zoom_image(self, event):
        factor = 1.1 if event.delta > 0 else 0.9
        self._diff_scale = max(0.05, min(self._diff_scale * factor, 10.0))
        self._diff_display_image()

    def _diff_start_move(self, event):
        self._diff_canvas.scan_mark(event.x, event.y)

    def _diff_move_image(self, event):
        self._diff_canvas.scan_dragto(event.x, event.y, gain=1)
























    # ════════════════════════════════════════════════════════════════
    # 四象限任务（重要/紧急矩阵）
    # ════════════════════════════════════════════════════════════════

    _QUADRANT_DEFS = [
        ("important_urgent", "🔴 重要 · 紧急", DANGER),
        ("important_not_urgent", "🟢 重要 · 不紧急", "#2dd4a8"),
        ("not_important_urgent", "🟡 不重要 · 紧急", "#ffb84d"),
        ("not_important_not_urgent", "⚪ 不重要 · 不紧急", "#8d92ab"),
    ]












    def _build_pdf_reader_tab(self):
        parent = self._tab_reader
        parent.columnconfigure(0, weight=0)   # 左侧面板
        parent.columnconfigure(1, weight=1)   # 中间：PDF预览 + Edge面板（PanedWindow）
        parent.rowconfigure(0, weight=1)

        # ── 水平PanedWindow：PDF预览区（左） + Edge面板（右） ──
        paned = tk.PanedWindow(parent, orient=tk.HORIZONTAL, bg=BORDER,
                                 sashwidth=3, sashrelief=tk.FLAT,
                                 sashpad=0, handlesize=6,
                                 borderwidth=0)
        paned.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self._reader_paned = paned

        # ── 左侧：文件夹+PDF列表（整栏可上下滚动）──
        left = tk.Frame(parent, bg=PANEL, width=280,
                        highlightbackground=BORDER, highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.pack_propagate(False)
        self._reader_left_frame = left

        # 折叠/展开按钮（顶部独立条，始终可见，不随滚动滚走）
        _left_top = tk.Frame(left, bg=PANEL, height=26)
        _left_top.pack(fill=tk.X)
        _left_top.pack_propagate(False)
        self._reader_left_toggle_btn = RoundedButton(
            _left_top, text="◀", bg=PANEL, fg=DIM, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 9, "bold"), padx=4, cursor="hand2",
            bd=0, command=self._toggle_reader_left_panel)
        self._reader_left_toggle_btn.pack(side=tk.RIGHT, padx=4, pady=2)

        # 可滚动视口：所有搜索区块放进 inner，内容超高时整栏上下滚动
        _left_canvas = tk.Canvas(left, bg=PANEL, highlightthickness=0, bd=0,
                                 yscrollincrement=20)
        _left_sb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=_left_canvas.yview)
        _left_canvas.configure(yscrollcommand=_left_sb.set)
        _left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        _left_sb.pack(side=tk.RIGHT, fill=tk.Y)
        _left_inner = tk.Frame(_left_canvas, bg=PANEL)
        _left_inner_id = _left_canvas.create_window((0, 0), window=_left_inner, anchor="nw")
        def _left_fit_inner(event):
            _left_canvas.itemconfigure(_left_inner_id, width=event.width)
        def _left_update_scrollregion(event):
            _left_canvas.configure(scrollregion=_left_canvas.bbox("all"))
        _left_canvas.bind("<Configure>", _left_fit_inner)
        _left_inner.bind("<Configure>", _left_update_scrollregion)
        def _left_wheel(event):
            # 只在鼠标位于左侧栏滚动区时滚动。搜索框/按钮/列表会吞掉滚轮事件，
            # 用 master 链判断鼠标下的 widget 是否属于 _left_inner，避免 bind_all
            # 误伤其它区域（与下方导航栏 _nav_wheel 同理）。
            w = event.widget
            while w is not None:
                if w is _left_inner:
                    _left_canvas.yview_scroll(int(-event.delta / 120) * 3, "units")
                    return "break"
                w = w.master
            return None
        self.bind_all("<MouseWheel>", _left_wheel, add="+")
        left = _left_inner   # 后续所有区块都建在内层滚动帧上

        # 文件夹选择
        hdr = tk.Frame(left, bg=PANEL)
        hdr.pack(fill=tk.X, padx=8, pady=(8, 0))
        RoundedButton(hdr, text="\U0001f4c1 选择文件夹", command=self._reader_select_dir,
                  bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9, "bold"),
                  cursor="hand2", padx=10, pady=6).pack(fill=tk.X)

        self._reader_dir_var = tk.StringVar(value=self._reader_dir or "(未选择)")
        tk.Label(left, textvariable=self._reader_dir_var, bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8), anchor="w", padx=10,
                 wraplength=260, justify=tk.LEFT).pack(fill=tk.X, pady=2)

        tk.Frame(left, bg=BORDER, height=1).pack(fill=tk.X, padx=8, pady=2)

        # ── Everything 全局搜索 ──
        search_hdr = tk.Frame(left, bg=PANEL)
        search_hdr.pack(fill=tk.X, padx=8, pady=(2, 0))
        tk.Label(search_hdr, text="\U0001f50d Everything 搜索", bg=PANEL, fg=TEXT,
                 font=("Microsoft YaHei UI", 9, "bold"), anchor="w").pack(fill=tk.X)

        self._everything_query = tk.StringVar(value="")
        eq_f = tk.Frame(search_hdr, bg=PANEL)
        eq_f.pack(fill=tk.X, pady=(2, 0))
        eq_entry = tk.Entry(eq_f, textvariable=self._everything_query,
                            bg=ENTRY, fg=TEXT, insertbackground=TEXT,
                            relief=tk.FLAT, font=("Microsoft YaHei UI", 9))
        eq_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self._bind_reader_text_input(eq_entry)
        eq_entry.bind("<Return>", lambda e: self._everything_search())
        self._bind_search_history(
            eq_entry, "everything",
            lambda kw: (self._everything_query.set(kw), self._everything_search()))
        RoundedButton(eq_f, text="搜索", bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 8), padx=8, cursor="hand2",
                  command=self._everything_search).pack(side=tk.RIGHT, padx=(4, 0))

        eq_tip = tk.Label(search_hdr, text="空格分隔多关键词，回车执行搜索",
                          bg=PANEL, fg=DIM, font=("Microsoft YaHei UI", 7), anchor="w")
        eq_tip.pack(fill=tk.X, pady=(0, 2))

        # 常用文件格式快捷筛选
        eq_quick_f = tk.Frame(search_hdr, bg=PANEL)
        eq_quick_f.pack(fill=tk.X, pady=(0, 2))
        for ext_label in (".epub", ".pdf", ".ppt", ".jpg", ".png"):
            RoundedButton(
                eq_quick_f, text=ext_label, bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                font=("Microsoft YaHei UI", 7), padx=6, cursor="hand2",
                command=lambda e=ext_label: self._everything_quick_filter(e)
            ).pack(side=tk.LEFT, padx=(0, 4))

        # Everything 搜索状态 + 清除按钮
        eq_status_f = tk.Frame(left, bg=PANEL)
        eq_status_f.pack(fill=tk.X, padx=8, pady=(0, 2))
        self._everything_info = tk.StringVar(value="Everything 就绪")
        tk.Label(eq_status_f, textvariable=self._everything_info, bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8), anchor="w").pack(side=tk.LEFT)
        RoundedButton(eq_status_f, text="返回文件夹", bg="#454b63", fg="white",
                  relief=tk.FLAT, font=("Microsoft YaHei UI", 8), padx=6,
                  cursor="hand2", command=self._everything_clear).pack(side=tk.RIGHT)

        tk.Frame(left, bg=BORDER, height=1).pack(fill=tk.X, padx=8, pady=2)

        # 搜索过滤
        fhdr = tk.Frame(left, bg=PANEL)
        fhdr.pack(fill=tk.X, padx=8, pady=(2, 0))
        self._reader_filter = tk.StringVar(value="")
        self._reader_filter.trace_add("write", self._on_reader_filter)
        fe = tk.Entry(fhdr, textvariable=self._reader_filter,
                      bg=ENTRY, fg=TEXT, insertbackground=TEXT,
                      relief=tk.FLAT, font=("Microsoft YaHei UI", 9))
        fe.pack(fill=tk.X, ipady=3)
        self._bind_reader_text_input(fe)
        tk.Label(fhdr, text="\U0001f50d 过滤文件名", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8)).pack(anchor="w")

        # ── 排序控件 ──
        sort_frame = tk.Frame(left, bg=PANEL)
        sort_frame.pack(fill=tk.X, padx=8, pady=(2, 0))
        tk.Label(sort_frame, text="排序:", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT)
        sort_modes = [
            ("名称", "name"), ("日期", "date"), ("大小", "size"), ("类型", "type"),
            ("次数", "count"),
        ]
        self._reader_sort_btns = {}
        for label, mode in sort_modes:
            btn = RoundedButton(
                sort_frame, text=label, bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                font=("Microsoft YaHei UI", 7), padx=5, pady=2, cursor="hand2",
                command=lambda m=mode: self._reader_set_sort(m))
            btn.pack(side=tk.LEFT, padx=(0, 3))
            self._reader_sort_btns[mode] = btn
        # 升降序切换
        self._reader_sort_dir_btn = RoundedButton(
            sort_frame, text="↑", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 7, "bold"), padx=5, pady=2, cursor="hand2",
            command=self._reader_toggle_sort_direction)
        self._reader_sort_dir_btn.pack(side=tk.LEFT, padx=(2, 0))
        self._reader_update_sort_btn_styles()

        # PDF文件列表
        self._reader_info = tk.StringVar(value="0 个PDF文件")
        tk.Label(left, textvariable=self._reader_info, bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8), anchor="w", padx=10).pack(fill=tk.X)

        lb_f = tk.Frame(left, bg=PANEL)
        lb_f.pack(fill=tk.BOTH, expand=True, padx=8, pady=(2, 8))
        self._reader_lb = tk.Listbox(lb_f, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                                      font=("Consolas", 9), relief=tk.FLAT, height=14,
                                      bd=0, highlightthickness=0, activestyle="none")
        vsb = ttk.Scrollbar(lb_f, command=self._reader_lb.yview)
        self._reader_lb.config(yscrollcommand=vsb.set)
        self._reader_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._reader_lb.bind("<<ListboxSelect>>", self._on_reader_select)
        self._reader_lb.bind("<Double-Button-1>", self._on_reader_open)
        # 右键菜单：打开原文档 / 重命名 / 打开文件位置
        self._reader_lb.bind("<Button-3>", self._reader_show_context_menu)
        self._reader_ctx_menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT,
                                        activebackground=ACCENT, activeforeground="white",
                                        font=("Microsoft YaHei UI", 9))
        self._reader_ctx_menu.add_command(label="📂 打开原文档", command=self._reader_ctx_open)
        self._reader_ctx_menu.add_command(
            label="▥ 在阅读区 2 打开",
            command=lambda: self._reader_dual_open_selected(2))
        self._reader_ctx_menu.add_command(label="✏️ 重命名", command=self._reader_ctx_rename)
        self._reader_ctx_menu.add_command(label="📁 打开文件位置", command=self._reader_ctx_open_folder)
        self._reader_ctx_menu.add_separator()

        # ── 右侧：PDF预览（PanedWindow左窗格） ──
        right = tk.Frame(paned, bg=BG)
        paned.add(right)  # PDF预览区占主要空间
        paned.paneconfigure(right, minsize=200, width=600)
        self._reader_right_frame = right
        right.rowconfigure(0, weight=0)
        right.rowconfigure(1, weight=0)
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)

        # 预览标题栏（标题 + 功能按钮）。搜索框/目录/标记按钮现在统一挪到了
        # PDF1自己的预览面板【内部】（和PDF2/PDF3的"标题行外、预览面板内"
        # 结构保持一致），不再单独占外层一整行；放大镜这类PDF1专属的强力
        # 功能按钮、以及高亮色/多开/主题/Edge/AI助手等，因为PDF2/PDF3没有
        # 对应物，仍然放在这个外层导航栏里。
        nav = tk.Frame(right, bg=PANEL, height=30)
        nav.grid(row=0, column=0, sticky="ew")
        nav.grid_propagate(False)
        nav.columnconfigure(0, weight=0)   # 标题自适应宽
        nav.columnconfigure(1, weight=1)   # 按钮区撑满剩余，按钮多时水平滚动

        self._reader_title = tk.StringVar(value="PDF预览: —")
        _title_icon_lbl = tk.Label(nav, text="📄", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 12),
                 anchor="w", padx=10, cursor="hand2")
        _title_icon_lbl.grid(row=0, column=0, sticky="w")
        # 悬停图标显示完整文件名（长文件名不再占用导航栏空间）
        self._bind_tooltip(_title_icon_lbl, lambda: self._reader_title.get())

        # 导航按钮（放在导航栏右侧；按钮多时水平滚动显示全部）
        nav_scroll = tk.Canvas(nav, bg=PANEL, height=30,
                               highlightthickness=0, bd=0)
        nav_scroll.grid(row=0, column=1, sticky="ew")
        nav_btn_frame = tk.Frame(nav_scroll, bg=PANEL)
        nav_btn_frame_id = nav_scroll.create_window(
            (0, 0), window=nav_btn_frame, anchor="nw")
        nav_btn_frame.bind(
            "<Configure>",
            lambda e: nav_scroll.configure(scrollregion=nav_scroll.bbox("all")))
        nav_scroll.bind(
            "<Configure>",
            lambda e: nav_scroll.itemconfigure(nav_btn_frame_id, height=e.height))

        def _nav_wheel(e):
            # 只在鼠标位于按钮行内部时才水平滚动。按钮（RoundedButton）会吞掉
            # 滚轮事件，用 master 链判断鼠标下的 widget 是否属于 nav_btn_frame。
            w = e.widget
            while w is not None:
                if w is nav_btn_frame:
                    direction = 1 if e.delta > 0 else -1
                    steps = max(40, abs(e.delta))
                    nav_scroll.xview_scroll(direction * steps, "units")
                    return "break"
                w = w.master
            return None
        self.bind_all("<MouseWheel>", _nav_wheel, add="+")
        self.bind_all("<Shift-MouseWheel>", _nav_wheel, add="+")

        # ── 滚动速度：预设下拉 + 微调按钮 ──
        scroll_box = tk.Frame(nav_btn_frame, bg=PANEL)
        scroll_box.pack(side=tk.RIGHT, padx=(4, 2))

        # 预设速度下拉框
        self._reader_speed_presets = [
            ("慢速", 1), ("较慢", 2), ("标准", 3),
            ("较快", 5), ("快速", 8), ("极速", 15), ("飞跃", 20),
        ]
        self._reader_speed_combo_var = tk.StringVar(value="标准")
        self._reader_speed_combo = ttk.Combobox(
            scroll_box, textvariable=self._reader_speed_combo_var,
            values=[p[0] for p in self._reader_speed_presets],
            state="readonly", font=("Microsoft YaHei UI", 8), width=4)
        self._reader_speed_combo.pack(side=tk.LEFT, padx=(0, 2))
        self._reader_speed_combo.bind("<<ComboboxSelected>>", self._on_reader_speed_preset)

        RoundedButton(scroll_box, text="－", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                   font=("Microsoft YaHei UI", 8, "bold"), padx=4, cursor="hand2",
                   command=lambda: self._reader_change_scroll_lines(-1)).pack(side=tk.LEFT)
        self._reader_scroll_lbl = tk.Label(
            scroll_box, textvariable=tk.StringVar(value=str(self._reader_scroll_lines)),
            bg=PANEL, fg=TEXT, font=("Microsoft YaHei UI", 8), width=2)
        self._reader_scroll_lbl.pack(side=tk.LEFT, padx=1)
        # Bind the label text to a StringVar that can be updated
        self._reader_scroll_lbl_var = tk.StringVar(value=str(self._reader_scroll_lines))
        self._reader_scroll_lbl.config(textvariable=self._reader_scroll_lbl_var)
        RoundedButton(scroll_box, text="＋", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                   font=("Microsoft YaHei UI", 8, "bold"), padx=4, cursor="hand2",
                   command=lambda: self._reader_change_scroll_lines(1)).pack(side=tk.LEFT)

        self._reader_fit_btn = RoundedButton(nav_btn_frame, text="🔲 适应宽度", bg=ACCENT, fg="white",
                                          relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
                                          padx=6, cursor="hand2",
                                          command=self._reader_toggle_fit_width)
        self._reader_fit_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_highlight_btn = RoundedButton(
            nav_btn_frame, text="▣ 高亮色", bg=self._reader_highlight_hex.get(),
            fg="#111111", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=6,
            cursor="hand2", command=self._pick_reader_highlight_color)
        self._reader_highlight_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_annotate_btn = RoundedButton(
            nav_btn_frame, text="✏ 标注", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_annotate_mode)
        self._reader_annotate_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_text_btn = RoundedButton(
            nav_btn_frame, text="T 文本", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_text_mode)
        self._reader_text_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_text_hl_btn = RoundedButton(
            nav_btn_frame, text="🖍 高亮收藏", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_text_hl_mode)
        self._reader_text_hl_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_annotate_color_btn = RoundedButton(
            nav_btn_frame, text="  ", bg=self._reader_annotate_color,
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=4, cursor="hand2", command=self._pick_reader_annotate_color)
        self._reader_annotate_color_btn.pack(side=tk.RIGHT, padx=(0, 1))

        self._reader_crop_btn = RoundedButton(
            nav_btn_frame, text="✂ 智能裁边", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_crop_margins)
        self._reader_crop_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_body_crop_btn = RoundedButton(
            nav_btn_frame, text="⛶ 框选正文", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_body_crop_mode)
        self._reader_body_crop_btn.pack(side=tk.RIGHT, padx=1)

        # 颜色主题选择按钮
        self._theme_btn = RoundedButton(nav_btn_frame, text="🎨 主题", bg="#454b63", fg="white",
                                     relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
                                     padx=6, cursor="hand2",
                                     command=self._show_theme_menu)
        self._theme_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_dual_btn = RoundedButton(
            nav_btn_frame, text="▦ 多开", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._show_reader_dual_menu)
        self._reader_dual_btn.pack(side=tk.RIGHT, padx=1)

        self._reader_match_nav_btn = RoundedButton(
            nav_btn_frame, text="≡ 匹配", bg=ACCENT, fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_match_col)
        self._reader_match_nav_btn.pack(side=tk.RIGHT, padx=1)

        # 联动搜索切换按钮
        _ll_colors = {True: GREEN, False: "#454b63"}
        _ll_labels = {True: "🔗 联动搜索:开", False: "🔗 联动搜索:关"}
        self._reader_linked_btn = RoundedButton(
            nav_btn_frame, text=_ll_labels[self._reader_linked_search.get()],
            bg=_ll_colors[self._reader_linked_search.get()],
            fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=6, cursor="hand2",
            command=self._toggle_reader_linked_search)
        self._reader_linked_btn.pack(side=tk.RIGHT, padx=1)

        # 始终可见的文件栏开关；左侧栏完全隐藏后仍可从这里恢复。
        self._reader_left_nav_btn = RoundedButton(
            nav_btn_frame, text="◀ 文件栏", bg=ACCENT, fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_left_panel)
        self._reader_left_nav_btn.pack(side=tk.RIGHT, padx=1)

        # 阅读窗口最大化按钮
        self._reader_max_btn = RoundedButton(
            nav_btn_frame, text="⛶ 最大化阅读", bg="#454b63", fg="white",
            relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
            padx=6, cursor="hand2", command=self._toggle_reader_maximize)
        self._reader_max_btn.pack(side=tk.RIGHT, padx=1)

        RoundedButton(nav_btn_frame, text="\u25c0", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9, "bold"), padx=6,
                  cursor="hand2", command=lambda: self._reader_active_go_page(-1)).pack(side=tk.RIGHT, padx=1)
        self._reader_page_var = tk.StringVar(value="— / —")
        tk.Label(nav_btn_frame, textvariable=self._reader_page_var, bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8), width=14).pack(side=tk.RIGHT)
        RoundedButton(nav_btn_frame, text="\u25b6", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 9, "bold"), padx=6,
                  cursor="hand2", command=lambda: self._reader_active_go_page(+1)).pack(side=tk.RIGHT, padx=1)

        # ── 页码导航条 ──
        thumb_frame = tk.Frame(right, bg=PANEL, height=48,
                               highlightbackground=BORDER, highlightthickness=1)
        thumb_frame.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        thumb_frame.grid_propagate(False)
        thumb_frame.rowconfigure(0, weight=1)
        thumb_frame.columnconfigure(0, weight=1)

        self._reader_thumb_canvas = tk.Canvas(thumb_frame, bg=PANEL,
                                               highlightthickness=0, height=40)
        thumb_hsb = ttk.Scrollbar(thumb_frame, orient=tk.HORIZONTAL,
                                  command=self._reader_thumb_canvas.xview)
        self._reader_thumb_canvas.configure(xscrollcommand=thumb_hsb.set)

        self._reader_thumb_canvas.grid(row=0, column=0, sticky="nsew")
        thumb_hsb.grid(row=1, column=0, sticky="ew")

        # 鼠标滚轮横向滚动
        def _thumb_wheel(event):
            self._reader_thumb_canvas.xview_scroll(int(-1 * event.delta / 120), "units")
        self._reader_thumb_canvas.bind("<MouseWheel>", _thumb_wheel)

        # PDF画布 + 匹配页码列（右侧）
        canvas_wrapper = tk.Frame(right, bg=BG)
        canvas_wrapper.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
        canvas_wrapper.columnconfigure(0, weight=1)
        canvas_wrapper.rowconfigure(0, weight=1)

        reader_view_paned = tk.PanedWindow(
            canvas_wrapper, orient=tk.HORIZONTAL, bg=BORDER,
            sashwidth=3, sashrelief=tk.FLAT, borderwidth=0)
        reader_view_paned.grid(row=0, column=0, sticky="nsew")
        self._reader_view_paned = reader_view_paned

        # ── canvas_container：PDF1真正的"预览面板"（作为一个窗格整体加入
        #    reader_view_paned），内部自带搜索行 + 目录/标记侧栏 + 内容区，
        #    结构和PDF2/PDF3的自带面板保持一致 ──
        canvas_container = tk.Frame(
            reader_view_paned, bg=CANVAS,
            highlightbackground=BORDER, highlightthickness=1)
        reader_view_paned.add(canvas_container, minsize=260, stretch="always")
        self._reader_pdf1_pane = canvas_container  # PDF1 在 PanedWindow 里的直接子窗格（横滚隐藏/恢复用）
        canvas_container.rowconfigure(0, weight=0)
        canvas_container.rowconfigure(1, weight=1)
        canvas_container.columnconfigure(0, weight=1)

        # ── 搜索行：和PDF2/PDF3一样，放在预览面板【内部】顶端 ──
        search_row = tk.Frame(canvas_container, bg=PANEL, height=28)
        search_row.grid(row=0, column=0, sticky="ew")
        search_row.pack_propagate(False)

        search_frame = tk.Frame(search_row, bg=PANEL)
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0), pady=2)
        self._reader_toc_btn = RoundedButton(
            search_frame, text="☰", bg=ACCENT, fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 9, "bold"), padx=5, cursor="hand2",
            command=self._toggle_reader_toc)
        self._reader_toc_btn.pack(side=tk.LEFT, padx=(0, 2))
        self._reader_bookmark_btn = RoundedButton(
            search_frame, text="📌", bg="#454b63", fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 9, "bold"), padx=5, cursor="hand2",
            command=self._toggle_reader_bookmarks)
        self._reader_bookmark_btn.pack(side=tk.LEFT, padx=(0, 2))
        self._reader_preview_search = tk.StringVar(value="")
        preview_search_entry = tk.Entry(search_frame, textvariable=self._reader_preview_search,
                                        bg=ENTRY, fg=TEXT, insertbackground=TEXT,
                                        relief=tk.FLAT, font=("Microsoft YaHei UI", 9))
        preview_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        self._reader_search_entry = preview_search_entry
        self._bind_reader_text_input(preview_search_entry)
        preview_search_entry.bind("<Return>", lambda e: self._reader_preview_search_exec())
        RoundedButton(search_frame, text="\u25b6", bg=ACCENT, fg="white", relief=tk.FLAT,
                  font=("Microsoft YaHei UI", 8, "bold"), padx=6, cursor="hand2",
                  command=self._reader_preview_search_exec).pack(side=tk.LEFT, padx=(2, 0))
        tk.Label(search_frame, text="\u9884\u89c8\u641c\u7d22", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 7)).pack(side=tk.LEFT, padx=(4, 0))
        self._reader_anytxt_hint = tk.StringVar(value="")
        anytxt_hint_lbl = tk.Label(search_frame, textvariable=self._reader_anytxt_hint,
                                   bg=PANEL, fg=ACCENT, cursor="hand2",
                                   font=("Microsoft YaHei UI", 7))
        anytxt_hint_lbl.pack(side=tk.LEFT, padx=(4, 0))

        # ── 主体：目录/标记侧栏 + 实际内容区（画布/文本/代码/网页预览） ──
        body_frame = tk.Frame(canvas_container, bg=CANVAS)
        body_frame.grid(row=1, column=0, sticky="nsew")
        body_frame.rowconfigure(0, weight=1)
        body_frame.columnconfigure(0, weight=0)
        body_frame.columnconfigure(1, weight=1)

        # PDF自带书签目录（可折叠，紧凑样式，和PDF2/PDF3保持一致：
        # 没有单独的标题栏，靠搜索行里的☰按钮统一控制显示/隐藏）
        toc_frame = tk.Frame(
            body_frame, bg=PANEL, width=260,
            highlightbackground=BORDER, highlightthickness=1)
        toc_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        toc_frame.pack_propagate(False)
        self._reader_toc_frame = toc_frame
        self._reader_toc_visible = False
        toc_frame.grid_remove()  # 需求5.1：目录默认关闭，点☰才显示

        # ── 视图A：PDF自带目录（可折叠层级树）──
        toc_body = tk.Frame(toc_frame, bg=PANEL)
        toc_body.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self._reader_toc_body = toc_body
        self._reader_toc_tree = ttk.Treeview(
            toc_body, show="tree", selectmode="browse")
        # 列宽固定足够宽 + 水平滚动条：长标题不再被侧栏宽度截断
        self._reader_toc_tree.column("#0", width=420, stretch=False)
        toc_vsb = ttk.Scrollbar(toc_body, orient=tk.VERTICAL,
                                command=self._reader_toc_tree.yview)
        toc_hsb = ttk.Scrollbar(toc_body, orient=tk.HORIZONTAL,
                                command=self._reader_toc_tree.xview)
        self._reader_toc_tree.configure(yscrollcommand=toc_vsb.set,
                                        xscrollcommand=toc_hsb.set)
        toc_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        toc_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._reader_toc_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._reader_toc_tree.bind("<<TreeviewSelect>>", self._reader_toc_jump)
        self._reader_toc_page = {}
        self._reader_toc_tree.insert("", "end", text="  打开PDF后显示目录")

        # ── 视图B：我的标记/笔记（默认不显示，由_toggle_reader_bookmarks控制）──
        bookmark_body = tk.Frame(toc_frame, bg=PANEL)
        self._reader_bookmark_body = bookmark_body
        RoundedButton(bookmark_body, text="📌 标记当前页", command=self._reader_add_bookmark,
                  bg=GREEN, fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 8, "bold"), padx=4, pady=4,
                  wraplength=110).pack(fill=tk.X, pady=(0, 4))
        RoundedButton(bookmark_body, text="📤 导出摘录", command=self._reader_export_excerpt,
                  bg=ACCENT, fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 8, "bold"), padx=4, pady=4,
                  wraplength=110).pack(fill=tk.X, pady=(0, 4))
        bookmark_lb_frame = tk.Frame(bookmark_body, bg=PANEL)
        bookmark_lb_frame.pack(fill=tk.BOTH, expand=True)
        self._reader_bookmark_lb = tk.Listbox(
            bookmark_lb_frame, bg=CARD, fg=TEXT, selectbackground=ACCENT,
            font=("Microsoft YaHei UI", 9), relief=tk.FLAT,
            bd=0, highlightthickness=0, activestyle="none")
        bookmark_vsb = ttk.Scrollbar(bookmark_lb_frame, command=self._reader_bookmark_lb.yview)
        self._reader_bookmark_lb.config(yscrollcommand=bookmark_vsb.set)
        self._reader_bookmark_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bookmark_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._reader_bookmark_lb.bind("<<ListboxSelect>>", self._reader_bookmark_jump)
        self._reader_bookmark_lb.bind("<Button-3>", self._reader_show_bookmark_context_menu)

        self._reader_bookmark_ctx_menu = tk.Menu(
            self, tearoff=0, bg=PANEL, fg=TEXT,
            activebackground=ACCENT, activeforeground="white",
            font=("Microsoft YaHei UI", 9))
        self._reader_bookmark_ctx_menu.add_command(
            label="✎ 编辑笔记", command=self._reader_edit_bookmark_note)
        self._reader_bookmark_ctx_menu.add_command(
            label="🗑 删除标记", command=self._reader_delete_bookmark)

        self._reader_refresh_bookmark_list()

        # ── 内层内容区：画布/文本/代码/网页预览（原有内部结构完全不变，
        #    只是父容器从canvas_container换成了这个内层pdf1_content_area） ──
        pdf1_content_area = tk.Frame(body_frame, bg=CANVAS)
        pdf1_content_area.grid(row=0, column=1, sticky="nsew")
        pdf1_content_area.rowconfigure(0, weight=1)
        pdf1_content_area.rowconfigure(2, weight=0)
        pdf1_content_area.columnconfigure(0, weight=1)
        canvas_container = pdf1_content_area  # 后续代码原样使用canvas_container这个名字即可
        self._reader_canvas_container = canvas_container  # 存储引用用于resize绑定

        self._reader_canvas = tk.Canvas(
            canvas_container, bg=CANVAS, highlightthickness=0,
            yscrollincrement=1)  # 像素级滚动，由惯性动画控制
        h_sc = ttk.Scrollbar(canvas_container, orient=tk.HORIZONTAL, command=self._reader_canvas.xview)
        v_sc = SlimScrollbar(canvas_container, orient="vertical",
                             command=self._reader_main_yview)
        self._reader_canvas.config(xscrollcommand=h_sc.set, yscrollcommand=v_sc.set)
        self._reader_canvas.grid(row=0, column=0, sticky="nsew")
        v_sc.grid(row=0, column=1, sticky="ns")
        h_sc.grid(row=1, column=0, sticky="ew")

        # ── 文本预览区域（用于非PDF文件，默认隐藏）──
        self._reader_text_frame = tk.Frame(canvas_container, bg=CARD)
        self._reader_text_frame.grid(row=0, column=0, sticky="nsew")
        self._reader_text_frame.grid_remove()
        self._reader_text_frame.rowconfigure(0, weight=1)
        self._reader_text_frame.columnconfigure(0, weight=1)

        self._reader_text_preview = tk.Text(self._reader_text_frame, bg=CARD, fg=TEXT,
                                            font=("Consolas", 10), wrap=tk.WORD,
                                            relief=tk.FLAT, padx=12, pady=10,
                                            insertbackground=TEXT,
                                            selectbackground=ACCENT,
                                            selectforeground="white")
        text_vsb = ttk.Scrollbar(self._reader_text_frame, command=self._reader_text_preview.yview)
        self._reader_text_preview.config(yscrollcommand=text_vsb.set)
        self._reader_text_preview.grid(row=0, column=0, sticky="nsew")
        text_vsb.grid(row=0, column=1, sticky="ns")
        # 联动搜索：左键点击取词（add="+" 保留默认行为）

        # ── 代码预览区域（用于.c/.cpp/.py/.html等代码与结构化文本文件，
        #    VSCode Dark+ 风格：深色背景 + 行号 + 语法高亮，默认隐藏）──
        self._reader_code_frame = tk.Frame(canvas_container, bg=CODE_BG)
        self._reader_code_frame.grid(row=0, column=0, sticky="nsew")
        self._reader_code_frame.grid_remove()
        self._reader_code_frame.rowconfigure(0, weight=0)
        self._reader_code_frame.rowconfigure(1, weight=1)
        self._reader_code_frame.columnconfigure(1, weight=1)

        # 代码视图通用工具条：字体大小调节（所有代码文件都显示）；
        # HTML/HTM文件额外显示"网页预览/浏览器打开"按钮组（用pack/pack_forget控制）
        self._reader_code_toolbar = tk.Frame(self._reader_code_frame, bg="#252526")
        self._reader_code_toolbar.grid(row=0, column=0, columnspan=3, sticky="ew")

        font_box = tk.Frame(self._reader_code_toolbar, bg="#252526")
        font_box.pack(side=tk.LEFT, padx=8, pady=4)
        tk.Label(font_box, text="字号", bg="#252526", fg=CODE_FG,
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        RoundedButton(font_box, text="A-", command=lambda: self._reader_change_code_font_size(-1),
                  bg="#3a3f56", fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 9, "bold"), width=3).pack(side=tk.LEFT, padx=2)
        self._reader_code_font_lbl = tk.Label(
            font_box, textvariable=self._reader_code_font_size, bg="#252526",
            fg=CODE_FG, font=("Microsoft YaHei UI", 9), width=2)
        self._reader_code_font_lbl.pack(side=tk.LEFT, padx=2)
        RoundedButton(font_box, text="A+", command=lambda: self._reader_change_code_font_size(1),
                  bg="#3a3f56", fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 9, "bold"), width=3).pack(side=tk.LEFT, padx=2)

        # 仅HTML/HTM文件显示的操作条：源码 ↔ 网页预览 切换 + 用系统浏览器打开
        self._reader_html_actionbar = tk.Frame(self._reader_code_toolbar, bg="#252526")
        # 默认不pack，由 _reader_show_code_preview 按文件类型决定是否显示
        tk.Label(self._reader_html_actionbar, text="|  💻 当前为源码视图", bg="#252526",
                 fg=CODE_FG, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=(8, 4))
        RoundedButton(self._reader_html_actionbar, text="🌐 内嵌网页预览",
                  command=self._reader_toggle_html_preview,
                  bg="#3a3f56", fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 9), padx=8, pady=2).pack(side=tk.LEFT, padx=4, pady=4)
        RoundedButton(self._reader_html_actionbar, text="🌐 用系统浏览器打开",
                  command=self._reader_open_html_in_browser,
                  bg=ACCENT, fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 9), padx=8, pady=2).pack(side=tk.LEFT, padx=4, pady=4)

        self._reader_code_lines = tk.Text(self._reader_code_frame, width=5,
                                           bg=CODE_GUTTER_BG, fg=CODE_GUTTER_FG,
                                           font=("Consolas", 10), relief=tk.FLAT,
                                           padx=6, pady=10, takefocus=0,
                                           highlightthickness=0, wrap=tk.NONE,
                                           cursor="arrow")
        self._reader_code_lines.tag_configure("ln", justify=tk.RIGHT)
        self._reader_code_lines.config(state=tk.DISABLED)
        self._reader_code_lines.grid(row=1, column=0, sticky="ns")

        self._reader_code_text = tk.Text(self._reader_code_frame, bg=CODE_BG, fg=CODE_FG,
                                          font=("Consolas", 10), wrap=tk.NONE,
                                          relief=tk.FLAT, padx=8, pady=10,
                                          insertbackground=CODE_FG,
                                          selectbackground=CODE_SELECT_BG,
                                          selectforeground="#ffffff",
                                          undo=False)
        code_vsb = ttk.Scrollbar(self._reader_code_frame, command=self._reader_code_yview)
        code_hsb = ttk.Scrollbar(self._reader_code_frame, orient=tk.HORIZONTAL,
                                  command=self._reader_code_text.xview)

        def _on_code_text_scroll(*args):
            code_vsb.set(*args)
            self._reader_code_lines.yview_moveto(args[0])
        self._reader_code_text.config(yscrollcommand=_on_code_text_scroll,
                                       xscrollcommand=code_hsb.set)
        self._reader_code_text.grid(row=1, column=1, sticky="nsew")
        code_vsb.grid(row=1, column=2, sticky="ns")
        code_hsb.grid(row=2, column=1, sticky="ew")
        self._reader_code_text.config(state=tk.DISABLED)

        # 代码区鼠标滚轮：同时驱动行号栏同步滚动（yscrollcommand回调已处理同步）
        self._reader_code_text.bind(
            "<MouseWheel>",
            lambda e: self._reader_code_text.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        # 联动搜索：左键点击取词

        # ── 网页预览区域（仅当检测到 tkinterweb 库时可用，用于HTML文件的
        #    内嵌可视化预览；默认隐藏，由 _reader_toggle_html_preview 控制）──
        self._reader_html_frame = tk.Frame(canvas_container, bg="white")
        self._reader_html_frame.grid(row=0, column=0, sticky="nsew")
        self._reader_html_frame.grid_remove()
        self._reader_html_frame.rowconfigure(1, weight=1)
        self._reader_html_frame.columnconfigure(0, weight=1)
        html_bar = tk.Frame(self._reader_html_frame, bg="#252526")
        html_bar.grid(row=0, column=0, sticky="ew")
        tk.Label(html_bar, text="🌐 网页预览", bg="#252526", fg="white",
                 font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=8, pady=4)
        RoundedButton(html_bar, text="💻 查看源码", command=self._reader_toggle_html_preview,
                  bg="#3a3f56", fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 9), padx=8, pady=2).pack(side=tk.LEFT, padx=4, pady=4)
        self._reader_html_body = tk.Frame(self._reader_html_frame, bg="white")
        self._reader_html_body.grid(row=1, column=0, sticky="nsew")
        self._reader_html_widget = None  # 实际的tkinterweb.HtmlFrame，首次使用时按需创建
        self._reader_apply_code_font_size()  # 确保控件字体与IntVar初始值一致

        # 绑定窗口大小变化事件 → 自动重排页面宽度
        self._reader_resize_job = None
        canvas_container.bind("<Configure>", self._reader_on_container_resize)

        self._reader_canvas.bind("<MouseWheel>", self._reader_zoom)
        self._reader_canvas.bind("<ButtonPress-1>", lambda e: self._reader_on_canvas_press(1, e))
        self._reader_canvas.bind("<B1-Motion>", lambda e: self._reader_on_canvas_drag(1, e))
        self._reader_canvas.bind("<Double-Button-1>", lambda e: self._reader_on_canvas_double(1, e))
        self._reader_canvas.bind("<Button-3>", lambda e: self._reader_on_canvas_right(1, e))

        # ── 第二、第三阅读区：目录/搜索/点击高亮均为独立会话 ──
        self._build_reader_aux_panel(reader_view_paned, 2)
        # 保留旧属性名，兼容已有清理与外部调用。
        self._reader_dual_frame = self._reader_aux[2]["frame"]
        self._reader_canvas2 = self._reader_aux[2]["canvas"]

        # ── PDF 1匹配页码块（置于预览下方，释放横向空间） ──
        # 高度只留"页码本身 + 少量内边距"：原来的标题栏、分隔线、居中计数标签
        # 三层横带（约56px）全部并进页码这一行，省下的纵向空间还给上方的
        # PDF正文预览，让预览一直延伸到页码区顶边。
        match_col = tk.Frame(canvas_container, bg=PANEL, height=60,
                             highlightbackground=BORDER, highlightthickness=1)
        match_col.grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(3, 0))
        match_col.pack_propagate(False)
        match_col.rowconfigure(1, weight=1)
        self._reader_match_frame = match_col
        self._reader_match_visible = True
        match_col.grid()  # 默认显示：匹配页码块默认展开（需求1.3）

        self._reader_match_count_var = tk.StringVar(value="")

        match_lb_frame = tk.Frame(match_col, bg=PANEL)
        match_lb_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=(2, 3))

        # 计数标签固定在匹配页码区【最左侧】，与页码同一行（宽度写死，
        # 页码起始位置不会随"12 页 / 345 处"字数变化而左右抖动）
        match_count_box = tk.Frame(match_lb_frame, bg=PANEL, width=56)
        match_count_box.pack(side=tk.LEFT, fill=tk.Y)
        match_count_box.pack_propagate(False)
        tk.Label(match_count_box, textvariable=self._reader_match_count_var,
                 bg=PANEL, fg=DIM, anchor="nw", justify=tk.LEFT, wraplength=54,
                 font=("Microsoft YaHei UI", 6)).pack(fill=tk.BOTH, expand=True)

        # 折叠按钮跟到同一行的最右侧
        self._match_toggle_btn = RoundedButton(match_lb_frame, text="▶", bg=PANEL, fg=DIM,
                                           relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
                                           padx=2, cursor="hand2", bd=0,
                                           command=self._toggle_match_col)
        self._match_toggle_btn.pack(side=tk.RIGHT, padx=(2, 0))

        self._reader_match_lb = tk.Listbox(match_lb_frame, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                                            font=("Consolas", 10), relief=tk.FLAT,
                                            bd=0, highlightthickness=0, activestyle="none",
                                            width=10)
        self._reader_match_canvas = tk.Canvas(
            match_lb_frame, bg=CARD, highlightthickness=0,
            yscrollincrement=18, height=52)
        match_vsb = ttk.Scrollbar(
            match_lb_frame, command=self._reader_match_canvas.yview)
        self._reader_match_canvas.config(yscrollcommand=match_vsb.set)
        match_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._reader_match_canvas.pack(
            side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._reader_match_lb.bind("<<ListboxSelect>>", self._reader_match_jump)
        self._reader_match_canvas.bind(
            "<Configure>", lambda event: self._reader_refresh_main_match_flow())
        self._reader_match_canvas.bind(
            "<MouseWheel>",
            lambda event: self._reader_match_canvas.yview_scroll(
                int(-event.delta / 120), "units"))
        for widget in (
                self._reader_canvas, self._reader_toc_tree,
                self._reader_thumb_canvas, self._reader_match_canvas):
            widget.bind(
                "<ButtonPress-1>",
                lambda event: self._reader_set_active_slot(1), add="+")

        # 存储当前匹配页码列表
        self._reader_match_pages = []
        self._reader_match_counts = {}

        # 左侧搜索栏历史：{栏key: {关键词: 搜索次数}}，下拉按次数降序
        if not hasattr(self, "_search_histories"):
            self._search_histories = {}
        self._search_history_max = 50
        self._search_history_win = None
        self._search_history_lb = None
        self._search_history_items = []
        self._search_history_apply = None
        # 点击窗口任意处关闭搜索历史下拉（点击落在下拉内部时忽略）
        self.bind_all("<Button-1>", self._on_global_click, add="+")

        self.after_idle(lambda: self._reader_set_active_slot(1))

        # 如果已有路径，自动扫描
        if self._reader_dir and os.path.isdir(self._reader_dir):
            self._reader_scan()

    def _build_reader_aux_panel(self, parent_paned, slot):
        """构建功能对等的辅助PDF阅读会话（第2/3阅读区）。"""
        frame = tk.Frame(
            parent_paned, bg=CANVAS,
            highlightbackground=BORDER, highlightthickness=1)
        frame.rowconfigure(2, weight=1)
        frame.rowconfigure(3, weight=0)
        frame.columnconfigure(0, weight=1)

        title_var = tk.StringVar(value=f"PDF {slot}: —")
        page_var = tk.StringVar(value="— / —")
        search_var = tk.StringVar(value="")
        search_info = tk.StringVar(value="")

        header = tk.Frame(frame, bg=PANEL, height=28)
        header.grid(row=0, column=0, sticky="ew")
        header.pack_propagate(False)
        tk.Label(
            header, textvariable=title_var, bg=PANEL, fg=TEXT,
            anchor="w", padx=6,
            font=("Microsoft YaHei UI", 8, "bold")).pack(
                side=tk.LEFT, fill=tk.X, expand=True)
        RoundedButton(
            header, text="◀", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=4,
            command=lambda s=slot: self._reader_aux_go_page(s, -1)).pack(
                side=tk.LEFT, padx=1)
        tk.Label(
            header, textvariable=page_var, bg=PANEL, fg=DIM,
            width=7, font=("Microsoft YaHei UI", 8)).pack(side=tk.LEFT)
        RoundedButton(
            header, text="▶", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=4,
            command=lambda s=slot: self._reader_aux_go_page(s, 1)).pack(
                side=tk.LEFT, padx=1)
        RoundedButton(
            header, text="✕", bg="#454b63", fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=4,
            command=lambda s=slot: self._reader_aux_close(s)).pack(
                side=tk.RIGHT, padx=2)

        tools = tk.Frame(frame, bg=PANEL)
        tools.grid(row=1, column=0, sticky="ew", pady=(1, 0))
        toc_btn = RoundedButton(
            tools, text="☰", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=5,
            command=lambda s=slot: self._reader_aux_toggle_toc(s))
        toc_btn.pack(side=tk.LEFT, padx=(2, 2), pady=2)
        bookmark_btn = RoundedButton(
            tools, text="📌", bg="#454b63", fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=5,
            command=lambda s=slot: self._reader_aux_toggle_bookmarks(s))
        bookmark_btn.pack(side=tk.LEFT, padx=(0, 2), pady=2)
        search_entry = tk.Entry(
            tools, textvariable=search_var, bg=ENTRY, fg=TEXT,
            insertbackground=TEXT, relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8))
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        self._bind_reader_text_input(search_entry)
        search_entry.bind(
            "<Return>", lambda event, s=slot: self._reader_aux_search(s))
        RoundedButton(
            tools, text="搜索", bg=ACCENT, fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=5,
            command=lambda s=slot: self._reader_aux_search(s)).pack(
                side=tk.LEFT, padx=2)
        search_hint = tk.StringVar(value="")
        anytxt_hint_lbl = tk.Label(tools, textvariable=search_hint, bg=PANEL, fg=ACCENT,
                                   cursor="hand2", font=("Microsoft YaHei UI", 7))
        anytxt_hint_lbl.pack(side=tk.RIGHT, padx=(4, 0))
        tk.Label(
            tools, textvariable=search_info, bg=PANEL, fg=DIM,
            font=("Microsoft YaHei UI", 7), width=9).pack(side=tk.RIGHT)

        body = tk.Frame(frame, bg=CANVAS)
        body.grid(row=2, column=0, sticky="nsew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        toc_frame = tk.Frame(
            body, bg=PANEL, width=260,
            highlightbackground=BORDER, highlightthickness=1)
        toc_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        toc_frame.pack_propagate(False)
        toc_frame.grid_remove()  # 需求5.1：目录默认关闭，点☰才显示

        # ── 视图A：PDF自带目录 ──
        toc_body = tk.Frame(toc_frame, bg=PANEL)
        toc_body.pack(fill=tk.BOTH, expand=True)
        toc_lb = tk.Listbox(
            toc_body, bg=CARD, fg=TEXT, selectbackground=ACCENT,
            font=("Microsoft YaHei UI", 8), relief=tk.FLAT,
            bd=0, highlightthickness=0, activestyle="none")
        toc_sb = ttk.Scrollbar(toc_body, command=toc_lb.yview)
        toc_hsb = ttk.Scrollbar(toc_body, orient=tk.HORIZONTAL,
                                command=toc_lb.xview)
        toc_lb.config(yscrollcommand=toc_sb.set, xscrollcommand=toc_hsb.set)
        toc_hsb.pack(side=tk.BOTTOM, fill=tk.X)
        toc_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        toc_sb.pack(side=tk.RIGHT, fill=tk.Y)
        toc_lb.insert(tk.END, "打开PDF后显示目录")
        toc_lb.bind(
            "<<ListboxSelect>>",
            lambda event, s=slot: self._reader_aux_toc_jump(s))

        # ── 视图B：我的标记/笔记（默认不显示，由_reader_aux_toggle_bookmarks控制）──
        bookmark_body = tk.Frame(toc_frame, bg=PANEL)
        RoundedButton(bookmark_body, text="📌 标记当前页",
                  command=lambda s=slot: self._reader_add_bookmark(s),
                  bg=GREEN, fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 8, "bold"), padx=4, pady=4,
                  wraplength=100).pack(fill=tk.X, pady=(2, 4), padx=2)
        RoundedButton(bookmark_body, text="📤 导出摘录",
                  command=lambda s=slot: self._reader_export_excerpt(s),
                  bg=ACCENT, fg="white", relief=tk.FLAT, cursor="hand2",
                  font=("Microsoft YaHei UI", 8, "bold"), padx=4, pady=4,
                  wraplength=100).pack(fill=tk.X, pady=(2, 4), padx=2)
        bookmark_lb_frame = tk.Frame(bookmark_body, bg=PANEL)
        bookmark_lb_frame.pack(fill=tk.BOTH, expand=True)
        bookmark_lb = tk.Listbox(
            bookmark_lb_frame, bg=CARD, fg=TEXT, selectbackground=ACCENT,
            font=("Microsoft YaHei UI", 8), relief=tk.FLAT,
            bd=0, highlightthickness=0, activestyle="none")
        bookmark_vsb = ttk.Scrollbar(bookmark_lb_frame, command=bookmark_lb.yview)
        bookmark_lb.config(yscrollcommand=bookmark_vsb.set)
        bookmark_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bookmark_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        bookmark_lb.bind(
            "<<ListboxSelect>>",
            lambda event, s=slot: self._reader_bookmark_jump(event, s))
        bookmark_lb.bind(
            "<Button-3>",
            lambda event, s=slot: self._reader_show_bookmark_context_menu(event, s))

        bookmark_ctx_menu = tk.Menu(
            self, tearoff=0, bg=PANEL, fg=TEXT,
            activebackground=ACCENT, activeforeground="white",
            font=("Microsoft YaHei UI", 9))
        bookmark_ctx_menu.add_command(
            label="✎ 编辑笔记", command=lambda s=slot: self._reader_edit_bookmark_note(s))
        bookmark_ctx_menu.add_command(
            label="🗑 删除标记", command=lambda s=slot: self._reader_delete_bookmark(s))

        canvas_frame = tk.Frame(body, bg=CANVAS)
        canvas_frame.grid(row=0, column=1, sticky="nsew")
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)
        canvas = tk.Canvas(
            canvas_frame, bg=CANVAS, highlightthickness=0,
            yscrollincrement=1)
        hsb = ttk.Scrollbar(
            canvas_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        vsb = SlimScrollbar(
            canvas_frame, orient="vertical",
            command=lambda *args, s=slot: self._reader_aux_yview(s, *args))
        canvas.config(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # ── EPUB内容显示区（默认隐藏，只有打开EPUB文件时才切换显示）──
        # 注意：这里的父容器是body（跟canvas_frame同级），不是canvas_frame
        # 本身——如果嵌套在canvas_frame里面，隐藏canvas_frame时这个区域会
        # 跟着一起被隐藏，没法在"画布视图"和"EPUB文本视图"之间正确切换。
        epub_frame = tk.Frame(body, bg=CARD)
        epub_text = tk.Text(
            epub_frame, wrap=tk.WORD, relief=tk.FLAT, bd=0,
            padx=16, pady=12, state=tk.DISABLED)
        epub_vsb = ttk.Scrollbar(epub_frame, command=epub_text.yview)
        epub_text.config(yscrollcommand=epub_vsb.set)
        epub_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        epub_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        epub_frame.grid(row=0, column=1, sticky="nsew")
        epub_frame.grid_remove()  # 默认隐藏，只有打开EPUB文件时才显示

        match_frame = tk.Frame(
            frame, bg=PANEL, height=60,
            highlightbackground=BORDER, highlightthickness=1)
        match_frame.grid(row=3, column=0, sticky="ew", pady=(3, 0))
        match_frame.pack_propagate(False)
        match_frame.grid_remove()  # 默认隐藏，由顶部「≡ 匹配」按钮统一控制三个预览区
        match_count_var = tk.StringVar(value="")
        match_body = tk.Frame(match_frame, bg=PANEL)
        match_body.pack(fill=tk.BOTH, expand=True, padx=4, pady=(2, 3))

        # 计数标签固定在匹配页码区【最左侧】，与页码同一行
        match_count_box = tk.Frame(match_body, bg=PANEL, width=56)
        match_count_box.pack(side=tk.LEFT, fill=tk.Y)
        match_count_box.pack_propagate(False)
        tk.Label(match_count_box, textvariable=match_count_var,
                 bg=PANEL, fg=DIM, anchor="nw", justify=tk.LEFT, wraplength=54,
                 font=("Microsoft YaHei UI", 6)).pack(fill=tk.BOTH, expand=True)

        match_lb = tk.Listbox(
            match_body, bg=CARD, fg=TEXT, selectbackground=ACCENT,
            font=("Consolas", 9), relief=tk.FLAT, height=3,
            bd=0, highlightthickness=0, activestyle="none")
        match_canvas = tk.Canvas(
            match_body, bg=CARD, highlightthickness=0,
            yscrollincrement=18, height=48)
        match_sb = ttk.Scrollbar(match_body, command=match_canvas.yview)
        match_canvas.config(yscrollcommand=match_sb.set)
        match_sb.pack(side=tk.RIGHT, fill=tk.Y)
        match_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        match_lb.bind(
            "<<ListboxSelect>>",
            lambda event, s=slot: self._reader_aux_match_jump(s))

        state = {
            "slot": slot, "frame": frame, "canvas": canvas,
            "canvas_frame": canvas_frame, "toc_frame": toc_frame,
            "epub_frame": epub_frame, "epub_text": epub_text,
            "toc_btn": toc_btn, "toc_lb": toc_lb, "toc_entries": [],
            "toc_visible": False, "toc_mode": "toc", "title_var": title_var,
            "toc_body": toc_body, "bookmark_body": bookmark_body,
            "bookmark_btn": bookmark_btn, "bookmark_lb": bookmark_lb,
            "bookmark_ctx_menu": bookmark_ctx_menu,
            "page_var": page_var, "search_var": search_var,
            "search_info": search_info, "anytxt_hint": search_hint, "doc": None, "path": "",
            "page": 0, "scale": 1.0, "fit_width": True,
            "tk_img": None, "tk_imgs": [], "image_x": 0,
            "page_xoffsets": {}, "zoom": 1.0,
            "page_crop_origins": {},
            "page_yoffsets": [], "render_start": 0, "render_end": 0,
            "scroll_velocity": 0.0, "scroll_anim_job": None,
            "word_cache": None,
            "match_frame": match_frame, "match_lb": match_lb,
            "match_canvas": match_canvas,
            "match_count_var": match_count_var, "match_pages": [],
            "match_counts": {},
            "visible": False, "resize_job": None,
            "keywords": [], "exact": False,
            "occurrence_window": None, "occurrence_geometry": None,
            "occurrence_generation": 0, "occurrence_locked": True,
            "occurrence_size": None,
        }
        self._reader_aux[slot] = state
        canvas.bind(
            "<MouseWheel>",
            lambda event, s=slot: self._reader_aux_wheel(s, event))
        canvas.bind(
            "<ButtonPress-1>",
            lambda event, s=slot: self._reader_on_canvas_press(s, event))
        canvas.bind(
            "<B1-Motion>",
            lambda event, s=slot: self._reader_on_canvas_drag(s, event))
        canvas.bind(
            "<Double-Button-1>",
            lambda event, s=slot: self._reader_on_canvas_double(s, event))
        canvas.bind(
            "<Button-3>",
            lambda event, s=slot: self._reader_on_canvas_right(s, event))
        canvas_frame.bind(
            "<Configure>",
            lambda event, s=slot: self._reader_aux_on_resize(s))
        header.bind(
            "<Button-1>",
            lambda event, s=slot: self._reader_set_active_slot(s))
        match_canvas.bind(
            "<Configure>",
            lambda event, s=slot: self._reader_refresh_aux_match_flow(s))
        match_canvas.bind(
            "<MouseWheel>",
            lambda event, c=match_canvas: c.yview_scroll(
                int(-event.delta / 120), "units"))
        for widget in (search_entry, toc_lb, match_canvas):
            widget.bind(
                "<ButtonPress-1>",
                lambda event, s=slot: self._reader_set_active_slot(s),
                add="+")
        return frame


    def _reader_layout_match_flow(self, canvas, pages, counts, command,
                                   doc=None, keywords=None, file_path=None):
        """
        从左到右绘制小号页码，宽度不足时自动换行。
        每10个页码一组换一种颜色（方便在很长的匹配列表里快速分辨大致位置），
        鼠标悬停在某个页码上时会弹出该页整页的预览图（带搜索关键词高亮）。

        字号【自适应缩小】：优先用较大的字号，若排不下所有匹配页码则逐级降到
        更小字号，尽量让全部页码在窗口内一屏显示完、无需滚动。
        """
        canvas.delete("all")
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        if cw < 10:
            cw = 150
        if ch < 10:
            ch = 280
        width = max(60, cw - 8)
        height = max(60, ch - 6)

        if not pages:
            canvas.config(scrollregion=(0, 0, width, 20))
            return

        # 命中密度条：条长 = 本页命中数 / 全篇最大命中数，归一化后一眼看出
        # 哪一页命中最多（而不是只知道"这页有没有"）。counts 为空时退回纯页码。
        max_count = max(counts.values(), default=0) if counts else 0
        has_bar = max_count > 0

        # 自适应字号：从大到小逐个尝试，直到所有匹配页码能排进一屏高度。
        import tkinter.font as tkfont
        widest_text = max((str(p + 1) for p in pages), key=len)
        font_size = 6
        line_h = 12
        bar_max_w = 12
        for cand in (8, 7, 6, 5):
            f = tkfont.Font(family="Microsoft YaHei UI", size=cand, weight="bold")
            widest_w = f.measure(widest_text)
            bar_zone = (cand + 9) if has_bar else 0   # 密度条宽度随字号缩放
            chip_w = widest_w + 4 + bar_zone + 3       # 数字宽 + 内边距 + 密度条 + 间距
            per_row = max(1, int(width // chip_w))
            rows = -(-len(pages) // per_row)           # ceil 除法
            lh = cand + 5
            if 4 + rows * lh + 4 <= height or cand == 5:
                font_size = cand
                line_h = lh
                bar_max_w = cand + 4
                break

        chip_font = ("Microsoft YaHei UI", font_size, "bold")
        x, y = 4, 4
        for i, page_idx in enumerate(pages):
            group_color = MATCH_CHIP_COLOR_GROUPS[(i // 10) % len(MATCH_CHIP_COLOR_GROUPS)]
            text = str(page_idx + 1)
            text_id = canvas.create_text(
                x, y + line_h / 2, anchor="w", text=text,
                fill=TEXT, font=chip_font)
            bbox = canvas.bbox(text_id) or (x, y, x + 8, y + 8)
            bar_zone = (bar_max_w + 5) if has_bar else 0
            pad_l, pad_r = 2, 2
            chip_left = bbox[0] - pad_l
            chip_top = y + 1
            chip_right = bbox[2] + pad_r + bar_zone
            chip_bottom = y + line_h - 1
            item_width = chip_right - chip_left
            if x > 4 and x + item_width > width:
                x = 4
                y += line_h
                canvas.coords(text_id, x, y + line_h / 2)
                bbox = canvas.bbox(text_id) or (x, y, x + 8, y + 8)
                chip_left = bbox[0] - pad_l
                chip_top = y + 1
                chip_right = bbox[2] + pad_r + bar_zone
                chip_bottom = y + line_h - 1
            rect_id = canvas.create_rectangle(
                chip_left, chip_top, chip_right, chip_bottom,
                fill=ENTRY, outline=group_color, width=1)
            canvas.tag_lower(rect_id, text_id)
            # 密度条画在页码右侧、芯片框内：灰色轨道 + 按命中比例填充的绿色条
            track_id = fill_id = None
            if has_bar:
                count = counts.get(page_idx, 0)
                ratio = (count / max_count) if count else 0.0
                bar_w = int(round(bar_max_w * ratio))
                bar_x0 = bbox[2] + 4
                bar_x1 = chip_right - 2
                bar_y0 = chip_top + 2
                bar_y1 = chip_bottom - 2
                track_id = canvas.create_rectangle(
                    bar_x0, bar_y0, bar_x1, bar_y1, fill="#3a3f52", outline="")
                if bar_w > 0:
                    fill_id = canvas.create_rectangle(
                        bar_x0, bar_y0, bar_x0 + bar_w, bar_y1,
                        fill=GREEN, outline="")
                canvas.tag_lower(track_id, text_id)
                if fill_id is not None:
                    canvas.tag_lower(fill_id, text_id)

            def set_hover(hover, tid=text_id, rid=rect_id, gc=group_color):
                # 悬停只切颜色、不改几何，避免"放大→Leave→缩小→Enter"抖动循环。
                canvas.itemconfigure(tid, fill="#ffffff" if hover else TEXT)
                canvas.itemconfigure(
                    rid, fill=ACCENT if hover else ENTRY,
                    outline=ACCENT if hover else gc)

            def click_page(event=None, pg=page_idx):
                command(pg)

            def enter_chip(event, fn=set_hover, pg=page_idx, rid=rect_id, cv=canvas):
                fn(True)
                self._reader_schedule_match_preview(doc, pg, keywords, cv, rid, file_path)

            def leave_chip(event, fn=set_hover, rid=rect_id, cv=canvas):
                # 在真正触发"隐藏"之前，先用鼠标的【真实屏幕坐标】核实一下
                # 是不是真的离开了这个页码标签的范围。
                #
                # 背景：预览窗弹出的位置如果恰好盖住了正在悬停的这个页码，
                # Tkinter会把"鼠标事件被上面盖住的窗口抢走了"误判成"鼠标
                # 离开了"——但鼠标物理上根本没有移动。如果照单全收地处理
                # 这次<Leave>，就会触发隐藏；隐藏后页码露出来，Tkinter又
                # 误判"鼠标进入了"，于是再次显示……不断循环、画面一直在
                # 闪。这里直接核实鼠标当前的真实屏幕坐标是否还落在该页码
                # 标签的屏幕范围内：如果还在，说明这是预览窗盖住造成的
                # 误判，直接忽略，不去关闭悬停态、也不触发隐藏。
                try:
                    bbox = cv.bbox(rid)
                    if bbox:
                        sx0 = cv.winfo_rootx() + (bbox[0] - cv.canvasx(0))
                        sy0 = cv.winfo_rooty() + (bbox[1] - cv.canvasy(0))
                        sx1 = cv.winfo_rootx() + (bbox[2] - cv.canvasx(0))
                        sy1 = cv.winfo_rooty() + (bbox[3] - cv.canvasy(0))
                        mx, my = self.winfo_pointerx(), self.winfo_pointery()
                        if sx0 <= mx <= sx1 and sy0 <= my <= sy1:
                            return  # 鼠标其实还在原地，不是真的离开，忽略
                except Exception:
                    pass
                fn(False)
                self._reader_cancel_match_preview()

            # 页码、芯片框、密度条（轨道+填充）都要响应悬停/点击，
            # 否则鼠标停在密度条上时不会触发预览（条在框的上面一层）。
            chip_items = [text_id, rect_id]
            if track_id is not None:
                chip_items.append(track_id)
                if fill_id is not None:
                    chip_items.append(fill_id)
            for item_id in chip_items:
                canvas.tag_bind(item_id, "<Enter>", enter_chip)
                canvas.tag_bind(item_id, "<Leave>", leave_chip)
                canvas.tag_bind(item_id, "<Button-1>", click_page)
            x += item_width + 3
        canvas.config(scrollregion=(0, 0, width, y + line_h + 2))

    def _reader_schedule_match_preview(self, doc, page_idx, keywords, canvas, item_id,
                                        file_path=None):
        """
        悬停预览的"停留意图"判断（防抖）：鼠标快速划过一长串页码标签时，
        每经过一个都会触发<Enter>事件——如果每次都立刻去渲染整页预览，
        快速移动时会连续触发好几次渲染，明显卡顿。这里改成"稍微等一下"：
        只有鼠标真的停留在某个页码上一小段时间（120ms）没有移动到别的
        标签，才会真正触发渲染；单纯路过的标签不会留下任何渲染开销。
        """
        # 进入新的页码，取消任何还没来得及执行的"延迟隐藏"任务——
        # 不然刚显示出来的预览可能被前一个页码留下的隐藏任务误关掉。
        if self._reader_match_preview_hide_job:
            try:
                self.after_cancel(self._reader_match_preview_hide_job)
            except Exception:
                pass
            self._reader_match_preview_hide_job = None
        self._reader_match_preview_pending = (doc, page_idx, keywords, canvas, item_id, file_path)
        if self._reader_match_preview_job:
            try:
                self.after_cancel(self._reader_match_preview_job)
            except Exception:
                pass
        self._reader_match_preview_job = self.after(120, self._reader_fire_match_preview)

    def _reader_fire_match_preview(self):
        """防抖计时器到点后执行：用最后一次悬停的页码渲染预览。"""
        self._reader_match_preview_job = None
        pending = self._reader_match_preview_pending
        if pending is None:
            return
        doc, page_idx, keywords, canvas, item_id, file_path = pending
        self._reader_show_match_preview(doc, page_idx, keywords, canvas, item_id, file_path)

    def _reader_cancel_match_preview(self):
        """
        鼠标离开页码标签：取消还没触发的预览计时；【延迟一点】才真正隐藏
        已经显示的预览窗，而不是立刻隐藏——如果鼠标只是在相邻的两个页码
        标签之间有一点很正常的细微移动（不是bug，只是手部动作本身就有
        抖动），不会看到预览"消失一下又冒出来"的闪烁感；只有鼠标真的
        离开了一小段时间没有进入任何新页码，才会真正隐藏。
        """
        if self._reader_match_preview_job:
            try:
                self.after_cancel(self._reader_match_preview_job)
            except Exception:
                pass
            self._reader_match_preview_job = None
        self._reader_match_preview_pending = None
        if self._reader_match_preview_hide_job:
            try:
                self.after_cancel(self._reader_match_preview_hide_job)
            except Exception:
                pass
        self._reader_match_preview_hide_job = self.after(100, self._reader_hide_match_preview)

    def _reader_show_match_preview(self, doc, page_idx, keywords, canvas, item_id, file_path=None):
        """
        鼠标悬停在匹配页码上时，在【整个匹配页码区域的正上方】弹出一个
        大尺寸的整页预览图（带搜索关键词高亮）。

        预览【始终】用裁剪后的内容区域来渲染（而不是整页）：同样大小的
        预览窗口，裁掉空白边距之后，正文内容能用更高的缩放比例显示出来，
        看起来更大、更清楚——这正是悬停预览的目的（快速看清该页内容），
        所以这里不再跟随阅读区的"智能裁边"开关，而是始终裁边。

        关键约束：预览窗的【下边界】永远不会超过"匹配页码区域本身"的
        顶部——也就是说，预览窗绝对不会盖住页码列表的任何一部分。
        这不是"哪个页码就贴哪个页码上方"那种局部定位，而是统一贴着
        整个页码区域的上边界，原因是：如果只贴着"当前这一个页码"的上方，
        当页码区域比较矮、或者预览图比较大时，预览窗的下半部分可能会
        伸到页码区域里面去，把其他页码盖住——这样鼠标就没法移动到被
        盖住的页码上触发新的悬停（事件会被预览窗这个置顶窗口拦截），
        会表现成"预览一直不消失、卡住了""挡住了没法看别的页码"。
        统一贴着整个区域上方显示，不管悬停的是哪一个页码，预览窗都不会
        进入页码区域的范围，这个问题就不存在了。

        定位和限制尺寸都以【主窗口自身】的位置/尺寸为参照，而不是用
        winfo_screenwidth()/winfo_screenheight()这种"屏幕尺寸"——多屏
        环境下"屏幕尺寸"汇报的有时候是所有显示器加起来的虚拟桌面大小，
        会导致预览跑到别的屏幕上去或者算出来的范围本身就不对。
        """
        if not doc or page_idx is None or page_idx >= doc.page_count:
            return
        try:
            page = doc[page_idx]

            # 预览始终按裁剪后的内容区域来算（用同一份按文件+页码缓存的
            # 检测结果，不会有额外开销）；检测失败（空白页/无干净边距）就
            # 回退成整页。
            clip_rect = None
            if file_path:
                clip_rect = self._reader_get_content_bbox(doc, page_idx, file_path)
            if clip_rect:
                page_w_pt = max(clip_rect.width, 1.0)
                page_h_pt = max(clip_rect.height, 1.0)
            else:
                page_w_pt = max(page.rect.width, 1.0)
                page_h_pt = max(page.rect.height, 1.0)

            app_w = max(400, self.winfo_width())
            border = 4
            gap = 12

            # 整个匹配页码区域（不只是当前悬停的这一个页码）在屏幕上的
            # 顶部位置——预览窗的下边界绝不能超过这里。
            chip_area_top = canvas.winfo_rooty()
            avail_h_above = chip_area_top - self.winfo_rooty() - gap - border
            if avail_h_above < 80:
                # 页码区域上方空间实在太小（比如这个匹配列表面板本身就
                # 紧贴着主窗口顶部），挤不出一个还有意义的预览尺寸——
                # 宁可这次不显示，也不要硬挤一个肯定会盖住页码区域本身
                # 的预览出来（那样反而违背了"不能重叠"这个前提）。
                return

            margin = 30
            # 预览窗允许出现的横向区间：左边可以略微探出主窗口（那一侧通常
            # 还在屏幕里），右边严格不越过主窗口右边界——主窗口最大化时
            # 它就是屏幕右边界，哪怕只多出几十个像素也会被裁掉、看不见。
            buffer = 30
            bound_x0 = self.winfo_rootx() - buffer
            bound_x1 = self.winfo_rootx() + self.winfo_width()
            # 宽度上限同时受两个约束：想要的放大比例，以及上面这个区间本身
            # 的宽度。少了第二个约束，宽窗口下算出来的图会比区间还宽，后面
            # 的钳位就只能在"贴左"和"贴右"里二选一，无论选哪边都有一侧溢出。
            max_w = max(300, min(int(app_w * 1.15) - margin * 2,
                                 (bound_x1 - bound_x0) - border))
            # 预览高度的上限直接取"页码区域上方还剩多少空间"，不再额外
            # 设一个"至少多少像素"的下限——这种下限在空间本来就紧张时，
            # 反而会比真实可用空间还大，导致后面的位置计算被迫在"不超出
            # 主窗口范围"和"不能重叠页码区域"两个约束之间二选一，结果
            # 选择了前者、牺牲了后者。直接用真实可用空间渲染，从根上
            # 保证图片本身就不会比这块空间还大。
            max_h = avail_h_above
            zoom = min(max_h / page_h_pt, max_w / page_w_pt, 8.0)

            # "尽可能清晰"：这是一次性的悬停预览，不像连续滚动阅读那样要兼顾
            # 性能，所以这里不用通用的_render_page_crisp（那个为了滚动流畅度，
            # 超采样倍数封顶在3倍），而是直接按目标显示尺寸的2倍分辨率去渲染，
            # 再用LANCZOS高质量算法缩小回目标尺寸——不管目标尺寸本身是多大，
            # 文字边缘都有充分的像素去做抗锯齿，效果比直接按目标尺寸渲染更清晰。
            SUPERSAMPLE = 2.0
            hi_zoom = zoom * SUPERSAMPLE
            pix = page.get_pixmap(matrix=fitz.Matrix(hi_zoom, hi_zoom),
                                   clip=clip_rect, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # 匹配页悬停预览跟随阅读主题背景色（暗黑/护眼绿等），否则预览
            # 仍是原始白底，与主阅读区背景不一致。先换背景再叠高亮，顺序与主渲染一致。
            img = self._apply_theme_to_image(img)
            if keywords:
                # clip_rect不是None时，img是从clip_rect.x0/y0开始渲染的，
                # 不是从页面的(0,0)开始——这里必须把这个偏移量传给高亮
                # 函数，否则算出来的高亮框位置会完全错位（这正是"预览开
                # 启智能裁边后高亮经常不准"的根本原因）。
                offset_x = clip_rect.x0 if clip_rect else 0.0
                offset_y = clip_rect.y0 if clip_rect else 0.0
                img = self._reader_highlight_page(
                    img, page, keywords, hi_zoom,
                    offset_x=offset_x, offset_y=offset_y)
            target_w = max(1, int(round(img.width / SUPERSAMPLE)))
            target_h = max(1, int(round(img.height / SUPERSAMPLE)))
            img = img.resize((target_w, target_h), Image.LANCZOS)

            # 最终保险：再确认一次图片不会超出"页码区域上方的可用空间"
            # 和主窗口宽度限制，确保万无一失。
            avail_w = max(200, max_w - border)
            avail_h = max_h
            if img.width > avail_w or img.height > avail_h:
                final_scale = min(avail_w / img.width, avail_h / img.height)
                img = img.resize(
                    (max(1, int(img.width * final_scale)),
                     max(1, int(img.height * final_scale))), Image.LANCZOS)

            tk_img = ImageTk.PhotoImage(img)
            win_w, win_h = img.size
        except Exception:
            return

        if (self._reader_match_preview_window is None or
                not self._reader_match_preview_window.winfo_exists()):
            win = tk.Toplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            frame = tk.Frame(win, bg=ACCENT, bd=2, relief=tk.SOLID)
            frame.pack()
            lbl = tk.Label(frame, bg="black")
            lbl.pack()
            self._reader_match_preview_window = win
            self._reader_match_preview_label = lbl

        self._reader_match_preview_label.config(image=tk_img)
        self._reader_match_preview_label.image = tk_img  # 保留引用，防止被垃圾回收

        # 横向位置仍然贴着当前悬停的这个页码（这个不影响"会不会盖住页码
        # 区域"，只影响左右对齐，体验上更贴合鼠标当前所在的位置）
        bbox = canvas.bbox(item_id) if item_id is not None else None
        if bbox:
            anchor_x = canvas.winfo_rootx() + (bbox[0] - canvas.canvasx(0))
        else:
            anchor_x = self.winfo_pointerx()

        # bound_x0 / bound_x1 在上面的 try 里已经算好（宽度上限要用到它们）
        bound_y0 = self.winfo_rooty() - buffer

        total_w, total_h = win_w + border, win_h + border
        # 纵向位置固定贴着"整个页码区域"的正上方——下边界=chip_area_top-gap，
        # 不会再有"放不下就改成显示在下方"这种可能导致遮挡页码区域的回退。
        target_y = max(bound_y0, chip_area_top - gap - total_h)
        target_x = max(bound_x0, min(anchor_x, bound_x1 - total_w))

        self._reader_match_preview_window.geometry(f"+{int(target_x)}+{int(target_y)}")
        self._reader_match_preview_window.deiconify()

    def _reader_hide_match_preview(self, event=None):
        """隐藏匹配页码悬停预览窗。"""
        if self._reader_match_preview_window is not None:
            try:
                self._reader_match_preview_window.withdraw()
            except Exception:
                pass

    def _reader_refresh_main_match_flow(self):
        if not hasattr(self, "_reader_match_canvas"):
            return
        self._reader_layout_match_flow(
            self._reader_match_canvas,
            self._reader_match_pages,
            self._reader_match_counts,
            self._reader_main_match_chip_click,
            doc=self._reader_pdf,
            keywords=self._reader_highlight_keywords,
            file_path=self._reader_active_file_path)

    def _reader_main_match_chip_click(self, page_idx):
        if not self._reader_pdf:
            return
        self._reader_set_active_slot(1)
        self._reader_page = page_idx
        self._reader_update_active_page_info()
        self._reader_scroll_to_page(page_idx)
        self._reader_highlight_thumb()

    def _reader_refresh_aux_match_flow(self, slot):
        state = self._reader_aux.get(slot)
        if not state:
            return
        self._reader_layout_match_flow(
            state["match_canvas"], state["match_pages"],
            state["match_counts"],
            lambda page_idx, s=slot: self._reader_aux_match_chip_click(
                s, page_idx),
            doc=state["doc"], keywords=state["keywords"],
            file_path=state["path"])

    def _reader_aux_match_chip_click(self, slot, page_idx):
        state = self._reader_aux[slot]
        if not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        state["page"] = page_idx
        self._reader_aux_render(
            slot, fit_width=False, target_page=page_idx)

    def _reader_get_active_pdf_context(self):
        """返回当前激活阅读栏的 (slot, doc, page, path)。"""
        slot = getattr(self, "_reader_active_slot", 1)
        if slot == 1:
            return 1, self._reader_pdf, self._reader_page, self._reader_current_path
        state = self._reader_aux.get(slot)
        if state:
            return slot, state["doc"], state["page"], state["path"]
        return slot, None, 0, ""

    def _reader_doc_display_name(self, path, doc, fallback="未命名PDF"):
        """
        安全地取得文档的显示用文件名。
        path 和 doc.name 在"内存中合成的PDF"（比如PPT通过COM渲染出来的临时
        文档，没有对应磁盘文件）的场景下都可能是 None 或空字符串，直接
        os.path.basename(None) 会抛 TypeError，这里统一兜底处理。
        """
        raw = path or getattr(doc, "name", None) or ""
        try:
            base = os.path.basename(raw)
        except Exception:
            base = ""
        return base or fallback

    def _reader_update_active_page_info(self):
        """让顶部页码显示跟随当前激活的PDF阅读栏。"""
        if not hasattr(self, "_reader_page_var"):
            return
        slot = getattr(self, "_reader_active_slot", 1)
        if slot == 1:
            if self._reader_pdf:
                self._reader_page_var.set(
                    f"PDF 1 · {self._reader_page + 1} / {self._reader_pdf.page_count}")
            elif self._reader_file_type:
                self._reader_page_var.set("PDF 1 · 非PDF")
            else:
                self._reader_page_var.set("PDF 1 · — / —")
        else:
            state = self._reader_aux.get(slot)
            if state and state["doc"]:
                self._reader_page_var.set(
                    f"PDF {slot} · {state['page'] + 1} / {state['doc'].page_count}")
            else:
                self._reader_page_var.set(f"PDF {slot} · — / —")
        self._reader_update_ai_source()

    def _reader_active_go_page(self, delta):
        """顶部翻页按钮跟随当前激活阅读栏。"""
        slot = getattr(self, "_reader_active_slot", 1)
        if slot == 1:
            self._reader_go_page(delta)
        else:
            self._reader_aux_go_page(slot, delta)

    def _reader_main_yview(self, *args):
        """主阅读栏滚动条拖动时，同步刷新顶部页码。"""
        if not hasattr(self, "_reader_canvas"):
            return
        self._reader_set_active_slot(1)
        self._reader_canvas.yview(*args)
        self._reader_detect_current_page()
        self._reader_check_lazy_load()

    def _reader_aux_yview(self, slot, *args):
        """辅助阅读栏滚动条拖动时，同步刷新顶部页码。"""
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        state["canvas"].yview(*args)
        self._reader_aux_detect_page(slot, render_missing=True)

    def _reader_set_active_slot(self, slot):
        """设置文件列表下一次要替换的PDF阅读栏。"""
        if slot not in (1, 2):
            return
        self._reader_active_slot = slot
        try:
            self._reader_canvas_container.config(
                highlightbackground=ACCENT if slot == 1 else BORDER,
                highlightthickness=2 if slot == 1 else 1)
        except Exception:
            pass
        for aux_slot, state in self._reader_aux.items():
            state["frame"].config(
                highlightbackground=ACCENT if aux_slot == slot else BORDER,
                highlightthickness=2 if aux_slot == slot else 1)
        if hasattr(self, "_reader_dual_btn"):
            self._reader_dual_btn.config(text=f"▦ 当前PDF {slot}")
        self._reader_update_active_page_info()

    # ── 滚轮横滚模式 ──

    def _reader_open_all_slots(self):
        """返回所有已打开（有文档）的阅读区序号，按 1,2,3,4 排序。"""
        slots = []
        if self._reader_pdf or self._reader_active_file_path:
            slots.append(1)
        for s in (2, 3, 4):
            st = self._reader_aux.get(s)
            if st and (st.get("doc") or st.get("path")):
                slots.append(s)
        return slots





    def _reader_apply_hscroll(self):
        """横滚模式下按「固定窗口数」显示/隐藏阅读区；PDF1 是否钉死由 pin1 决定。

        用 paneconfigure(hide=...) 增量显隐（不 forget/re-add），窗格不销毁、
        顺序天然稳定，避免每次横滚都整帧重建导致的页面闪烁。
        """
        all_slots = self._reader_open_all_slots()
        aux_slots = [s for s in all_slots if s != 1]
        visible = max(1, min(self._reader_hscroll_visible, len(all_slots) or 1))

        if self._reader_hscroll_pin1:
            aux_visible = max(0, min(visible - 1, len(aux_slots)))
            max_win = max(0, len(aux_slots) - aux_visible)
            self._reader_hscroll_win = max(0, min(self._reader_hscroll_win, max_win))
            visible_slots = [1] + aux_slots[
                self._reader_hscroll_win:self._reader_hscroll_win + aux_visible]
        else:
            max_win = max(0, len(all_slots) - visible)
            self._reader_hscroll_win = max(0, min(self._reader_hscroll_win, max_win))
            visible_slots = all_slots[
                self._reader_hscroll_win:self._reader_hscroll_win + visible]

        paned = self._reader_view_paned
        show1 = 1 in visible_slots
        if self._reader_pdf1_pane is not None and show1 != self._reader_hscroll_pdf1_visible:
            try:
                paned.paneconfigure(self._reader_pdf1_pane, hide=not show1)
                self._reader_hscroll_pdf1_visible = show1
            except Exception:
                pass
        for s in (2, 3, 4):
            st = self._reader_aux.get(s)
            if not st:
                continue
            should = (s in visible_slots) and bool(st.get("doc") or st.get("path"))
            if should != st["visible"]:
                try:
                    paned.paneconfigure(st["frame"], hide=not should)
                except Exception:
                    continue
                st["visible"] = should
        self._reader_update_multi_button()


    def _reader_hscroll_shift(self, delta):
        """横滚窗口平移（delta>0 向右，delta<0 向左），越界由 _apply 内部钳位。"""
        self._reader_hscroll_win += delta
        self._reader_apply_hscroll()

    def _fmt_match_count(self, pages, occ):
        """匹配区计数标签：页码在上、命中处在下（两行竖排，字号已缩小）。"""
        return f"{pages} 页\n{occ} 处"


    def _reader_open_in_active_slot_for_slot(self, full_path, rel_path, slot):
        """在指定阅读区打开文件（slot 在调用时就已确定，避免延迟执行时 race condition）。"""
        if slot == 1:
            self._reader_open(full_path, rel_path)
            return
        ext = os.path.splitext(full_path)[1].lower()
        if ext not in (".pdf", ".epub", ".ppt", ".pptx") \
           and ext not in IMAGE_EXTS:
            messagebox.showinfo(
                "替换阅读栏",
                f"阅读区 {slot} 支持 PDF / EPUB / PPT / PPTX / 图片文件，"
                f"代码、Markdown、网页等格式暂时只能在主阅读区(PDF1)打开。")
            return
        self._reader_dual_open_path(full_path, slot)

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 文件夹/文件管理
    # ════════════════════════════════════════════════════════════════

    def _reader_select_dir(self):
        import tkinter.filedialog as fd
        path = fd.askdirectory(title="选择PDF文件夹")
        if not path:
            return
        self._reader_dir = path
        self._reader_dir_var.set(path)
        self._save_config()
        self._reader_scan()

    def _reader_scan(self):
        """递归扫描文件夹下的所有支持文件"""
        self._reader_pdf_files = []
        self._reader_pdf_files_all = []
        if not self._reader_dir:
            return
        SUPPORTED_EXTS = READER_SUPPORTED_EXTS
        try:
            for root, dirs, files in os.walk(self._reader_dir):
                for f in files:
                    if any(f.lower().endswith(ext) for ext in SUPPORTED_EXTS):
                        full = os.path.join(root, f)
                        rel = os.path.relpath(full, self._reader_dir)
                        self._reader_pdf_files.append((full, rel))
            self._reader_pdf_files_all = list(self._reader_pdf_files)
        except Exception:
            pass
        self._reader_sort_files()
        self._reader_info.set(f"{len(self._reader_pdf_files)} 个文件")

    def _refresh_reader_list(self):
        self._reader_lb.delete(0, tk.END)
        filt = self._reader_filter.get().strip().lower()
        icons = {
            ".pdf": "📄", ".ppt": "📊", ".pptx": "📊", ".md": "📝", ".epub": "📚",
            ".py": "🐍", ".c": "💻", ".h": "💻", ".cpp": "💻", ".cc": "💻",
            ".cxx": "💻", ".hpp": "💻", ".hxx": "💻", ".java": "💻", ".cs": "💻",
            ".html": "🌐", ".htm": "🌐", ".xml": "🌐",
            ".css": "🎨", ".js": "📜", ".ts": "📜", ".jsx": "📜", ".tsx": "📜",
            ".json": "🔧", ".yaml": "🔧", ".yml": "🔧", ".ini": "🔧", ".cfg": "🔧",
            ".sh": "🐚", ".bat": "🐚", ".ps1": "🐚",
            ".txt": "📃", ".log": "📃", ".sql": "💻",
            ".doc": "📘", ".docx": "📘", ".xls": "📊", ".xlsx": "📊",
        }
        for full, rel in self._reader_pdf_files:
            if filt and filt not in rel.lower():
                continue
            ext = os.path.splitext(rel)[1].lower()
            icon = icons.get(ext, "📄")
            cnt = self._open_counts.get(full, 0)
            suffix = f"  ×{cnt}" if cnt else ""
            self._reader_lb.insert(tk.END, f"  {icon} {rel}{suffix}")

    def _on_reader_filter(self, *args):
        self._refresh_reader_list()

    def _reader_set_sort(self, mode):
        """设置排序方式并重新排序文件列表。"""
        if self._reader_sort_mode == mode:
            # 同一排序方式 → 切换升降序
            self._reader_sort_reverse = not self._reader_sort_reverse
        else:
            self._reader_sort_mode = mode
            self._reader_sort_reverse = False
        self._reader_update_sort_btn_styles()
        self._reader_sort_files()

    def _reader_toggle_sort_direction(self):
        """切换升降序。"""
        self._reader_sort_reverse = not self._reader_sort_reverse
        self._reader_update_sort_btn_styles()
        self._reader_sort_files()

    def _reader_update_sort_btn_styles(self):
        """更新排序按钮高亮状态。"""
        for mode, btn in self._reader_sort_btns.items():
            if mode == self._reader_sort_mode:
                btn.config(bg=ACCENT, fg="white")
            else:
                btn.config(bg=ENTRY, fg=TEXT)
        arrow = "↓" if self._reader_sort_reverse else "↑"
        self._reader_sort_dir_btn.config(text=arrow)

    def _reader_sort_files(self):
        """根据当前排序方式对文件列表排序并刷新列表。"""
        mode = self._reader_sort_mode
        reverse = self._reader_sort_reverse
        if mode == "name":
            self._reader_pdf_files.sort(key=lambda x: x[1].lower(), reverse=reverse)
        elif mode == "date":
            def get_mtime(path_rel):
                try:
                    return os.path.getmtime(path_rel[0])
                except Exception:
                    return 0
            self._reader_pdf_files.sort(key=get_mtime, reverse=not reverse)
        elif mode == "size":
            def get_size(path_rel):
                try:
                    return os.path.getsize(path_rel[0])
                except Exception:
                    return 0
            self._reader_pdf_files.sort(key=get_size, reverse=not reverse)
        elif mode == "type":
            self._reader_pdf_files.sort(
                key=lambda x: os.path.splitext(x[1])[1].lower(), reverse=reverse)
        elif mode == "count":
            self._reader_pdf_files.sort(
                key=lambda x: self._open_counts.get(x[0], 0), reverse=not reverse)
        self._refresh_reader_list()

    def _on_reader_select(self, event=None):
        """单击文件名 → 立即在激活的阅读区打开。"""
        sel = self._reader_lb.curselection()
        if not sel:
            return
        idx = self._get_reader_actual_index(sel[0])
        if idx < 0:
            return
        full_path, rel_path = self._reader_pdf_files[idx]
        # 在点击瞬间捕获当前激活槽
        target_slot = getattr(self, "_reader_active_slot", 1)
        self._reader_open_in_active_slot_for_slot(full_path, rel_path, target_slot)

    def _on_reader_open(self, event=None):
        """双击 → 立即在激活的阅读区打开。"""
        sel = self._reader_lb.curselection()
        if not sel:
            return
        idx = self._get_reader_actual_index(sel[0])
        if idx < 0:
            return
        full_path, rel_path = self._reader_pdf_files[idx]
        target_slot = getattr(self, "_reader_active_slot", 1)
        self._reader_open_in_active_slot_for_slot(full_path, rel_path, target_slot)

    def _reader_show_context_menu(self, event):
        """右键点击文件列表 → 弹出上下文菜单"""
        idx = self._reader_lb.nearest(event.y)
        if idx < 0:
            return
        self._reader_lb.selection_clear(0, tk.END)
        self._reader_lb.selection_set(idx)
        self._reader_lb.activate(idx)
        self._reader_ctx_menu.post(event.x_root, event.y_root)

    def _reader_ctx_open(self):
        """右键菜单：用系统默认程序打开原文档"""
        sel = self._reader_lb.curselection()
        if not sel:
            return
        idx = self._get_reader_actual_index(sel[0])
        if idx < 0:
            return
        full_path, _ = self._reader_pdf_files[idx]
        if not os.path.isfile(full_path):
            messagebox.showerror("错误", f"文件不存在:\n{full_path}")
            return
        try:
            os.startfile(full_path)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开文件:\n{e}")

    def _reader_ctx_rename(self):
        """右键菜单：重命名文件"""
        sel = self._reader_lb.curselection()
        if not sel:
            return
        idx = self._get_reader_actual_index(sel[0])
        if idx < 0:
            return
        full_path, rel_path = self._reader_pdf_files[idx]
        dir_name = os.path.dirname(full_path)
        old_name = os.path.basename(full_path)
        from tkinter import simpledialog
        new_name = simpledialog.askstring("重命名", f"原文件名:\n{old_name}\n\n请输入新文件名:",
                                            initialvalue=old_name, parent=self)
        if not new_name or new_name == old_name:
            return
        forbidden = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        if any(c in new_name for c in forbidden):
            messagebox.showerror("错误", f"文件名不能包含以下字符:\n{' '.join(forbidden)}")
            return
        new_path = os.path.join(dir_name, new_name)
        if os.path.exists(new_path):
            messagebox.showerror("错误", f"文件已存在:\n{new_name}")
            return
        try:
            os.rename(full_path, new_path)
            self._reader_pdf_files[idx] = (new_path, new_name)
            self._refresh_reader_list()
            if self._reader_pdf and old_name in self._reader_title.get():
                self._reader_title.set(f"📄 {new_name}")
        except Exception as e:
            messagebox.showerror("重命名失败", f"无法重命名文件:\n{e}")

    def _reader_ctx_open_folder(self):
        """右键菜单：打开文件所在文件夹并选中该文件"""
        sel = self._reader_lb.curselection()
        if not sel:
            return
        idx = self._get_reader_actual_index(sel[0])
        if idx < 0:
            return
        full_path, _ = self._reader_pdf_files[idx]
        folder = os.path.dirname(full_path)
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"文件夹不存在:\n{folder}")
            return
        try:
            import subprocess
            subprocess.run(f'explorer /select,"{os.path.normpath(full_path)}"', shell=True)
        except Exception as e:
            try:
                os.startfile(folder)
            except Exception as e2:
                messagebox.showerror("打开失败", f"无法打开文件夹:\n{e2}")

    def _get_reader_actual_index(self, listbox_idx):
        """将 Listbox 索引映射到真实的 PDF 文件列表索引"""
        filt = self._reader_filter.get().strip().lower()
        actual_idx = -1
        for full, rel in self._reader_pdf_files:
            if filt and filt not in rel.lower():
                continue
            actual_idx += 1
            if actual_idx == listbox_idx:
                for i, (f, r) in enumerate(self._reader_pdf_files):
                    if f == full:
                        return i
        return -1

    def _record_open(self, full_path):
        """记录一次文件打开次数并持久化（用于"打开次数"排序）。"""
        if not full_path:
            return
        self._open_counts[full_path] = self._open_counts.get(full_path, 0) + 1
        self._save_config()

    def _reader_save_resume(self):
        """把主阅读区(PDF1)当前阅读位置（页码/缩放/自适应）写入续读记录。"""
        if not self._reader_pdf:
            return
        path = self._reader_active_file_path or self._reader_current_path
        if not path:
            return
        self._reader_resume[path] = {
            "page": int(self._reader_page),
            "scale": float(self._reader_scale),
            "fit_width": bool(self._reader_fit_width),
        }

    def _reader_save_aux_resume(self, slot):
        """把辅助阅读区(PDF2/3)当前阅读位置写入续读记录。"""
        state = self._reader_aux.get(slot)
        if not state or not state.get("doc") or not state.get("path"):
            return
        self._reader_resume[state["path"]] = {
            "page": int(state.get("page", 0)),
            "scale": float(state.get("scale", 1.0)),
            "fit_width": bool(state.get("fit_width", True)),
        }

    def _reader_open(self, full_path, rel_path):
        # ── 重入保护 + 文件存在性检查 ──
        if self._reader_opening:
            return
        if not full_path or not os.path.isfile(full_path):
            messagebox.showerror("打开失败", f"文件不存在或路径无效:\n{full_path}")
            return
        self._reader_save_resume()
        self._record_open(full_path)
        try:
            self._reader_opening = True
            self._reader_doc_generation += 1
            generation = self._reader_doc_generation
            if self._reader_lazy_job:
                try:
                    self.after_cancel(self._reader_lazy_job)
                except Exception:
                    pass
                self._reader_lazy_job = None
            if self._reader_scroll_anim_job:
                try:
                    self.after_cancel(self._reader_scroll_anim_job)
                except Exception:
                    pass
                self._reader_scroll_anim_job = None
                self._reader_scroll_velocity = 0.0
            if self._reader_pdf:
                self._reader_pdf.close()
                self._reader_pdf = None
            self._reader_load_toc()
            self._reader_page_cache.clear()
            self._reader_text_cache = None
            self._reader_word_index = None
            self._reader_yoffset_estimated = False
            self._reader_file_type = ""
            self._reader_match_lb.delete(0, tk.END)
            self._reader_match_pages = []
            self._reader_match_counts = {}
            self._reader_refresh_main_match_flow()
            self._reader_match_count_var.set("")
            self._reader_preview_search.set("")
            self._reader_canvas.grid()

            ext = os.path.splitext(full_path)[1].lower()
            if ext != ".pdf":
                messagebox.showerror("打开失败", f"仅支持 PDF 文件:\n{rel_path}")
                return
            self._reader_file_type = ext
            self._reader_current_path = full_path
            self._reader_active_file_path = full_path
            doc = fitz.open(full_path)
            self._reader_apply_pdf_doc(doc, rel_path, icon="📄", generation=generation)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开文件: {e}")
            self._reader_pdf = None
            self._reader_load_toc()
        finally:
            self._reader_opening = False





    def _reader_apply_pdf_doc(self, doc, rel_path, icon="📄", generation=None):
        """
        把一个fitz.Document设为当前阅读器的活动文档，并完成与普通PDF完全
        一致的初始化（缩放/高亮重置、标题、目录、缩略图、渲染、文本缓存）。
        普通.pdf文件和"PPT通过COM渲染成功后得到的内存PDF"都走这同一套逻辑，
        这样PPT一旦渲染成功，搜索/缩略图/翻页等功能无需任何特殊处理就能直接生效。
        """
        if generation is None:
            generation = self._reader_doc_generation
        self._reader_pdf = doc
        self._reader_page = 0
        self._reader_scale = 1.0
        self._reader_fit_width = True
        # 续读：恢复上次的页码/缩放/自适应状态
        resume = self._reader_resume.get(self._reader_active_file_path)
        resume_page = None
        if resume:
            try:
                resume_page = max(0, min(int(resume.get("page", 0)),
                                         self._reader_pdf.page_count - 1))
                self._reader_page = resume_page
                self._reader_scale = float(resume.get("scale", 1.0))
                self._reader_fit_width = bool(resume.get("fit_width", True))
            except (TypeError, ValueError):
                resume_page = None
        self._reader_highlight_keywords = []
        self._reader_highlight_exact = False
        self._reader_canvas.grid()
        if self._reader_text_frame:
            self._reader_text_frame.grid_remove()
        if self._reader_code_frame:
            self._reader_code_frame.grid_remove()
        if self._reader_html_frame:
            self._reader_html_frame.grid_remove()
        self._reader_title.set(f"{icon} {rel_path}")
        self._reader_page_var.set(f"{self._reader_page + 1} / {self._reader_pdf.page_count}")
        self._reader_update_active_page_info()
        self._reader_refresh_bookmark_list()
        self._reader_load_toc()
        self._reader_build_thumbnails()
        self._reader_render()
        if resume_page is not None and resume_page > 0:
            self._reader_scroll_to_page(resume_page)
        self.after_idle(lambda g=generation: self._reader_build_text_cache(g))


    def _read_text_file_smart(self, full_path):
        """
        按常见编码顺序尝试读取文本文件，提升对GBK等非UTF-8源码/文本文件的兼容性
        （国内很多旧的.c/.py源码文件是用GBK/GB2312保存的）。
        返回 (content, 实际使用的编码说明)。
        """
        for enc in ("utf-8", "utf-8-sig", "gbk", "gb2312", "big5"):
            try:
                with open(full_path, "r", encoding=enc) as f:
                    return f.read(), enc
            except (UnicodeDecodeError, UnicodeError, LookupError):
                continue
        # 最后兜底：用utf-8读取，无法解码的字节用占位符替换，保证不崩溃
        with open(full_path, "r", encoding="utf-8", errors="replace") as f:
            return f.read(), "utf-8(部分字符无法识别)"



    def _reader_extract_ppt_via_com(self, full_path):
        """
        用 PowerPoint（或WPS演示）COM自动化把PPT每页导出为图片并提取文字位置。
        这是整个PPT可视化预览里"慢"的部分（要启动PowerPoint/WPS、逐页导出），
        所以单独拆出来，返回的素材会被缓存，同一文件第二次打开时可以跳过
        这一步，直接用 _reader_build_ppt_doc 快速重建文档。

        注意：本功能依赖本机已安装 PowerPoint 或 WPS演示，且需要pywin32。
        在没有图形界面的环境下无法测试COM调用本身，如果你的WPS版本对应的
        自动化ProgID不是"Kwpp.Application"，请告诉我实际报错，我再调整。

        返回 (slide_w, slide_h, slides_data, 错误信息或None)
        slides_data: [(png_bytes, [(text, left, top, w, h), ...]), ...]  每项对应一页幻灯片
        """
        if win32 is None:
            return 0, 0, [], "未安装 pywin32"

        import tempfile
        import shutil
        try:
            import pythoncom
        except ImportError:
            return 0, 0, [], "未安装 pywin32(pythoncom)"

        pythoncom.CoInitialize()
        app = None
        presentation = None
        tmpdir = None
        try:
            last_err = None
            used_name = None
            for prog_id, name in (("PowerPoint.Application", "PowerPoint"),
                                   ("Kwpp.Application", "WPS演示")):
                try:
                    app = win32.Dispatch(prog_id)
                    used_name = name
                    break
                except Exception as e:
                    last_err = e
                    app = None
            if app is None:
                return 0, 0, [], f"未检测到PowerPoint或WPS演示：{last_err}"

            try:
                # ReadOnly=True, Untitled=False, WithWindow=False：只读、不新建窗口打开
                presentation = app.Presentations.Open(full_path, True, False, False)
            except Exception as e:
                return 0, 0, [], f"{used_name}打开文件失败: {e}"

            slide_w = float(presentation.PageSetup.SlideWidth)
            slide_h = float(presentation.PageSetup.SlideHeight)
            if slide_w <= 0 or slide_h <= 0:
                slide_w, slide_h = 960.0, 540.0  # 兜底默认尺寸（16:9，单位：点）

            target_scale = 150.0 / 72.0  # 约150 DPI
            px_w = max(1, int(slide_w * target_scale))
            px_h = max(1, int(slide_h * target_scale))

            tmpdir = tempfile.mkdtemp(prefix="ppt_render_")
            slides_data = []

            n_slides = presentation.Slides.Count
            for i in range(1, n_slides + 1):
                try:
                    slide = presentation.Slides(i)
                except Exception:
                    continue
                img_path = os.path.join(tmpdir, f"slide_{i:04d}.png")
                try:
                    slide.Export(img_path, "PNG", px_w, px_h)
                except Exception:
                    continue

                png_bytes = None
                if os.path.isfile(img_path):
                    try:
                        with open(img_path, "rb") as f:
                            png_bytes = f.read()
                    except Exception:
                        png_bytes = None

                texts = []
                try:
                    for shp in slide.Shapes:
                        try:
                            if not shp.HasTextFrame:
                                continue
                            tf = shp.TextFrame
                            if not tf.HasText:
                                continue
                            text = tf.TextRange.Text
                            if not text or not text.strip():
                                continue
                            left, top = float(shp.Left), float(shp.Top)
                            w, h = float(shp.Width), float(shp.Height)
                            if w <= 1 or h <= 1:
                                continue
                            texts.append((text, left, top, w, h))
                        except Exception:
                            continue
                except Exception:
                    pass

                slides_data.append((png_bytes, texts))

            if not slides_data:
                return 0, 0, [], "导出幻灯片图片失败（0页）"
            return slide_w, slide_h, slides_data, None
        except Exception as e:
            return 0, 0, [], str(e)
        finally:
            try:
                if presentation is not None:
                    presentation.Close()
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            try:
                if tmpdir and os.path.isdir(tmpdir):
                    shutil.rmtree(tmpdir, ignore_errors=True)
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _reader_build_ppt_doc(self, slide_w, slide_h, slides_data):
        """
        用已经提取好的素材（每页图片字节+文字位置）快速拼出一个全新的
        fitz.Document。纯本地操作，不涉及COM，缓存命中时调用这个即可，
        每次都会创建一个全新对象，不会和之前已经关闭的文档产生冲突。
        """
        new_doc = fitz.open()
        for png_bytes, texts in slides_data:
            page = new_doc.new_page(width=slide_w, height=slide_h)
            if png_bytes:
                try:
                    page.insert_image(page.rect, stream=png_bytes)
                except Exception:
                    pass
            for text, left, top, w, h in texts:
                try:
                    rect = fitz.Rect(left, top, left + w, top + h)
                    page.insert_textbox(rect, text, fontsize=8,
                                         render_mode=3, overlay=True)
                except Exception:
                    continue
        return new_doc


    def _epub_parse_chapters(self, full_path):
        """
        解析EPUB文件的全部章节内容，返回[(章节标题, 正文), ...]。
        结果按文件路径缓存，避免每次切换样式/重新打开时都重新解析一遍。
        和旧版本的区别：旧版本只取前10个文档项、每项截断5000字符，这次
        改成读取全部章节的完整正文。
        解析失败（没装ebooklib）返回None；正常解析但没有文本内容返回[]。
        """
        if full_path in self._epub_chapters_cache:
            return self._epub_chapters_cache[full_path]
        try:
            import ebooklib
            from ebooklib import epub
            from html.parser import HTMLParser

            class TextExtractor(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.texts = []

                def handle_data(self, data):
                    self.texts.append(data)

                def get_text(self):
                    return "".join(self.texts)

            book = epub.read_epub(full_path)
            chapters = []
            chapter_num = 0
            for item in book.get_items():
                if item.get_type() == ebooklib.ITEM_DOCUMENT:
                    try:
                        extractor = TextExtractor()
                        content = item.get_content()
                        if isinstance(content, bytes):
                            content = content.decode("utf-8", errors="ignore")
                        extractor.feed(content)
                        text = extractor.get_text().strip()
                        text = re.sub(r"\n{3,}", "\n\n", text)
                        if text:
                            chapter_num += 1
                            chapters.append((f"第 {chapter_num} 章", text))
                    except Exception:
                        pass
            self._epub_chapters_cache[full_path] = chapters
            return chapters
        except ImportError:
            return None
        except Exception:
            return []

    def _epub_text_color_for_bg(self, bg_hex):
        """根据背景色的明暗程度，自动选用黑色或白色文字，保证任何背景下都看得清。"""
        try:
            r, g, b = _hex_to_rgb(bg_hex)
        except Exception:
            return "#e9eaf5"
        brightness = 0.299 * r + 0.587 * g + 0.114 * b
        return "#1a1a1a" if brightness > 150 else "#e9eaf5"

    def _epub_apply_style(self, text_widget):
        """把当前的EPUB样式设置（背景色/字号/字体/行间距）应用到指定的Text控件上。"""
        fg = self._epub_text_color_for_bg(self._epub_bg_color)
        try:
            text_widget.config(
                bg=self._epub_bg_color, fg=fg, insertbackground=fg,
                font=(self._epub_font_family, self._epub_font_size),
                spacing1=self._epub_line_spacing,
                spacing2=self._epub_line_spacing,
                spacing3=self._epub_line_spacing)
            text_widget.tag_configure(
                "epub_chapter_title",
                font=(self._epub_font_family, self._epub_font_size + 4, "bold"),
                foreground=ACCENT, spacing3=self._epub_line_spacing + 6)
        except Exception:
            pass

    def _epub_render_into_text_widget(self, text_widget, chapters, rel_path="", full_path=None):
        """
        把解析出来的EPUB章节内容，按当前样式设置填充进指定的Text控件，
        并叠加显示已保存的标注（黄色背景高亮，按"章节+正文内字符偏移量"
        定位，而不是Text控件的行列索引——字号/行间距/字体changes之后
        行列索引会整个错位，但字符偏移量只要原文不变就稳定）。
        """
        was_disabled = text_widget.cget("state") == tk.DISABLED
        text_widget.config(state=tk.NORMAL)
        text_widget.delete("1.0", tk.END)
        text_widget.tag_remove("epub_mark", "1.0", tk.END)
        # 记录每个章节正文在Text控件里的起始位置（"start_index"）和长度，
        # 后面把"章节+字符偏移"换算成Text索引、以及处理"选中文字后添加
        # 标注"时都要用到。
        chapter_anchors = []  # [(start_index_str, text_length), ...]
        if chapters is None:
            text_widget.insert(
                tk.END,
                f"[EPUB文件预览]\n\n文件: {rel_path}\n\n"
                f"提示: 安装 ebooklib 库可获得完整的章节正文预览\n  pip install ebooklib")
        elif not chapters:
            text_widget.insert(tk.END, "[这本EPUB里没有解析出可显示的文本内容]")
        else:
            for i, (title, text) in enumerate(chapters):
                if i > 0:
                    text_widget.insert(tk.END, "\n\n")
                text_widget.insert(tk.END, title + "\n", ("epub_chapter_title",))
                text_widget.insert(tk.END, "\n")
                # 在插入正文【之前】先取一次END索引——这正是"即将插入的
                # 正文的起点"，不存在任何歧义。比起"插入完一个\n之后再
                # 用END-1c往回找"这种写法更直接可靠（后者在不同Tk版本/
                # 边界情况下的行为容易让人没有把握，干脆不用）。
                start_index = text_widget.index(tk.END)
                text_widget.insert(tk.END, text)
                chapter_anchors.append((start_index, len(text)))
                text_widget.insert(tk.END, "\n")
        self._epub_apply_style(text_widget)
        text_widget._epub_chapter_anchors = chapter_anchors  # 挂在控件上，方便后续直接取用
        text_widget._epub_full_path = full_path
        if full_path:
            self._epub_draw_marks(text_widget, full_path)
        if was_disabled:
            text_widget.config(state=tk.DISABLED)

    def _epub_index_for_offset(self, text_widget, chapter_idx, char_offset):
        """把"第chapter_idx章正文里的第char_offset个字符"换算成Text控件的实际索引(如'12.34')。"""
        anchors = getattr(text_widget, "_epub_chapter_anchors", None)
        if not anchors or chapter_idx >= len(anchors):
            return None
        start_index, length = anchors[chapter_idx]
        char_offset = max(0, min(char_offset, length))
        return text_widget.index(f"{start_index}+{char_offset}c")

    def _epub_draw_marks(self, text_widget, full_path):
        """把指定文件已保存的全部标注，按当前的章节锚点换算成Text索引后画出来（黄色背景高亮）。"""
        marks = self._epub_marks.get(full_path)
        if not marks:
            return
        text_widget.tag_configure("epub_mark", background="#ffdc00", foreground="#1a1a1a")
        for mark in marks:
            idx0 = self._epub_index_for_offset(text_widget, mark["chapter"], mark["start"])
            idx1 = self._epub_index_for_offset(text_widget, mark["chapter"], mark["end"])
            if idx0 and idx1:
                text_widget.tag_add("epub_mark", idx0, idx1)

    def _epub_show_selection_menu(self, text_widget, event, full_path):
        """
        在EPUB文本上右键时弹出菜单：如果当前有选中文字，提供"添加标注"；
        如果右键点击位置正好落在已有标注内，提供"删除这个标注"。
        两种情况可能同时出现（菜单会同时列出两个选项）。
        """
        if not full_path:
            return
        menu = tk.Menu(text_widget, tearoff=0, bg=CARD, fg=TEXT,
                        activebackground=ACCENT, activeforeground="white")
        added_any = False
        try:
            sel_start = text_widget.index(tk.SEL_FIRST)
            sel_end = text_widget.index(tk.SEL_LAST)
            has_selection = text_widget.compare(sel_start, "<", sel_end)
        except tk.TclError:
            has_selection = False
        if has_selection:
            menu.add_command(
                label="🖍 标注选中文字",
                command=lambda: self._epub_add_mark_from_selection(
                    text_widget, sel_start, sel_end, full_path))
            added_any = True

        click_index = text_widget.index(f"@{event.x},{event.y}")
        hit_mark = self._epub_find_mark_at_index(text_widget, full_path, click_index)
        if hit_mark:
            menu.add_command(
                label="🗑 删除此处标注",
                command=lambda: self._epub_delete_mark(text_widget, full_path, hit_mark))
            added_any = True

        if added_any:
            menu.tk_popup(event.x_root, event.y_root)

    def _epub_find_mark_at_index(self, text_widget, full_path, click_index):
        """找出鼠标点击位置落在哪一个已保存的标注范围内（找不到返回None）。"""
        marks = self._epub_marks.get(full_path)
        if not marks:
            return None
        for mark in marks:
            idx0 = self._epub_index_for_offset(text_widget, mark["chapter"], mark["start"])
            idx1 = self._epub_index_for_offset(text_widget, mark["chapter"], mark["end"])
            if idx0 and idx1 and text_widget.compare(idx0, "<=", click_index) \
                    and text_widget.compare(click_index, "<", idx1):
                return mark
        return None

    def _epub_offset_for_index(self, text_widget, text_index):
        """
        把Text控件的实际索引（比如'12.34'）反过来换算成"第几章+正文内第
        几个字符"——用户选中文字添加标注时，选区给的是Text索引，必须先
        换算成不受样式影响的"章节+字符偏移"才能持久化保存。
        """
        anchors = getattr(text_widget, "_epub_chapter_anchors", None)
        if not anchors:
            return None
        for chapter_idx, (start_index, length) in enumerate(anchors):
            end_index = text_widget.index(f"{start_index}+{length}c")
            if text_widget.compare(start_index, "<=", text_index) and \
                    text_widget.compare(text_index, "<=", end_index):
                # Tcl的Text.count返回的是两个索引之间隔着多少个字符
                offset = len(text_widget.get(start_index, text_index))
                return chapter_idx, offset
        return None

    def _epub_add_mark_from_selection(self, text_widget, sel_start, sel_end, full_path):
        """把当前选中的文字范围保存为一条新标注。"""
        start_info = self._epub_offset_for_index(text_widget, sel_start)
        end_info = self._epub_offset_for_index(text_widget, sel_end)
        if not start_info or not end_info or start_info[0] != end_info[0]:
            messagebox.showinfo("无法标注", "暂不支持跨章节的标注范围，请在同一章节内选择文字。")
            return
        chapter_idx, start_offset = start_info
        _, end_offset = end_info
        if end_offset <= start_offset:
            return
        marks = self._epub_marks.setdefault(full_path, [])
        marks.append({
            "chapter": chapter_idx, "start": start_offset, "end": end_offset,
            "color": self._epub_annotate_color,
        })
        self._save_config()
        text_widget.tag_remove(tk.SEL, "1.0", tk.END)
        self._epub_draw_marks(text_widget, full_path)

    def _epub_delete_mark(self, text_widget, full_path, mark):
        """删除一条标注，并刷新所有正在显示这个文件的面板。"""
        marks = self._epub_marks.get(full_path)
        if marks and mark in marks:
            marks.remove(mark)
            if not marks:
                del self._epub_marks[full_path]
            self._save_config()
        self._epub_redraw_marks_everywhere(full_path)

    def _epub_redraw_marks_everywhere(self, full_path):
        """同一个EPUB文件可能同时在PDF1/2/3里都打开着，标注变化后统一刷新。"""
        if self._reader_file_type == ".epub" and \
                getattr(self._reader_text_preview, "_epub_full_path", None) == full_path:
            self._reader_text_preview.tag_remove("epub_mark", "1.0", tk.END)
            self._epub_draw_marks(self._reader_text_preview, full_path)
        for state in self._reader_aux.values():
            epub_text = state.get("epub_text")
            if epub_text is not None and getattr(epub_text, "_epub_full_path", None) == full_path:
                epub_text.tag_remove("epub_mark", "1.0", tk.END)
                self._epub_draw_marks(epub_text, full_path)

    def _epub_bind_annotate_events(self, text_widget):
        """
        给指定的EPUB文本控件绑定右键标注交互（添加/删除）。
        注意：Tkinter的bind()对同一个事件类型("<Button-3>")是【覆盖】
        而不是叠加——所以PDF1每次重新打开EPUB都重复调用这个函数是安全
        的，不会导致右键菜单弹出多次；PDF2/3那边额外加了"只绑定一次"的
        标记，是因为它们的Text控件是常驻的（不会重新创建），单纯是为了
        避免重复绑定造成代码逻辑上的混淆，而不是因为不这样做会出BUG。
        """
        text_widget.bind(
            "<Button-3>",
            lambda event, tw=text_widget: self._epub_show_selection_menu(
                tw, event, getattr(tw, "_epub_full_path", None)))
        # 联动搜索：滑动选中关键词后松开鼠标触发搜索
        text_widget.bind(
            "<ButtonRelease-1>",
            lambda event, tw=text_widget: self._reader_epub_selection_search(tw, event),
            add="+")



    # ════════════════════════════════════════════════════════════════
    # 文件阅读器 - 代码文件（VSCode风格深色阅读 + 行号 + 语法高亮）
    # ════════════════════════════════════════════════════════════════



    def _reader_open_html_in_browser(self):
        """用系统默认浏览器打开当前HTML文件，获得最可靠的真实网页效果。"""
        if not self._reader_html_path or not os.path.isfile(self._reader_html_path):
            messagebox.showwarning("无法打开", "未找到当前HTML文件的路径。")
            return
        import webbrowser
        try:
            webbrowser.open("file:///" + self._reader_html_path.replace("\\", "/"))
            self._reader_status_safe_set("已用系统浏览器打开网页")
        except Exception as e:
            messagebox.showerror("打开失败", f"无法启动系统浏览器: {e}")

    def _reader_status_safe_set(self, text):
        """部分场景没有底部状态变量，安全地尝试设置，没有就忽略。"""
        if hasattr(self, "_reader_info"):
            try:
                self._reader_info.set(text)
            except Exception:
                pass

    def _reader_toggle_html_preview(self):
        """
        在"源码视图"和"内嵌网页预览"之间切换（仅HTML文件可用）。
        内嵌预览依赖可选库 tkinterweb；没安装的话会提示并引导改用
        "用系统浏览器打开"按钮，不会报错或卡死。
        """
        if not self._reader_html_path:
            return

        if self._reader_html_mode == "code":
            # 切换到网页预览
            self._reader_code_frame.grid_remove()
            self._reader_html_frame.grid()
            self._reader_html_mode = "preview"
            self._reader_render_html_preview()
        else:
            # 切回源码视图
            self._reader_html_frame.grid_remove()
            self._reader_code_frame.grid()
            self._reader_html_mode = "code"

    def _reader_render_html_preview(self):
        """实际渲染内嵌网页预览（首次调用时按需创建tkinterweb控件）。"""
        for w in self._reader_html_body.winfo_children():
            w.destroy()

        try:
            from tkinterweb import HtmlFrame
        except ImportError:
            tip = tk.Frame(self._reader_html_body, bg="white")
            tip.pack(fill=tk.BOTH, expand=True)
            tk.Label(tip, text="📦 内嵌网页预览需要安装可选库 tkinterweb",
                     bg="white", fg="#333", font=("Microsoft YaHei UI", 11, "bold"),
                     pady=20).pack()
            tk.Label(tip, text="pip install tkinterweb\n\n安装后重新打开此文件即可启用内嵌预览。\n"
                               "现在可以点击下面的按钮，用系统浏览器查看真实网页效果：",
                     bg="white", fg="#666", font=("Microsoft YaHei UI", 9),
                     justify=tk.CENTER).pack(pady=(0, 12))
            RoundedButton(tip, text="🌐 用系统浏览器打开", command=self._reader_open_html_in_browser,
                      bg=ACCENT, fg="white", relief=tk.FLAT, cursor="hand2",
                      font=("Microsoft YaHei UI", 10, "bold"), padx=14, pady=8).pack()
            return

        try:
            html_widget = HtmlFrame(self._reader_html_body, messages_enabled=False)
            html_widget.pack(fill=tk.BOTH, expand=True)
            html_widget.load_file(self._reader_html_path)
            self._reader_html_widget = html_widget
        except Exception as e:
            err_label = tk.Label(self._reader_html_body,
                                  text=f"内嵌预览加载失败: {e}\n\n请改用下面的按钮用系统浏览器打开。",
                                  bg="white", fg="#a33", font=("Microsoft YaHei UI", 9),
                                  justify=tk.CENTER, pady=20)
            err_label.pack()
            RoundedButton(self._reader_html_body, text="🌐 用系统浏览器打开",
                      command=self._reader_open_html_in_browser,
                      bg=ACCENT, fg="white", relief=tk.FLAT, cursor="hand2",
                      font=("Microsoft YaHei UI", 10, "bold"), padx=14, pady=8).pack()



    def _reader_code_yview(self, *args):
        """代码区垂直滚动条统一入口：同时驱动正文与行号栏滚动。"""
        self._reader_code_text.yview(*args)
        self._reader_code_lines.yview(*args)

    def _reader_change_code_font_size(self, delta: int):
        """调整代码阅读字体大小（A-/A+按钮），限制在6~32之间，并立即生效+保存。"""
        new_size = self._reader_code_font_size.get() + delta
        new_size = max(6, min(32, new_size))
        self._reader_code_font_size.set(new_size)
        self._reader_apply_code_font_size()
        self._save_config()

    def _reader_apply_code_font_size(self):
        """把当前字体大小应用到代码正文和行号栏控件上。"""
        size = self._reader_code_font_size.get()
        try:
            self._reader_code_text.config(font=("Consolas", size))
            self._reader_code_lines.config(font=("Consolas", size))
        except Exception:
            pass



    def _reader_get_active_text_widget(self):
        """
        返回当前"非PDF文本类预览"应使用的Text控件，及其对应的搜索域：
        - 代码文件 → self._reader_code_text
        - Markdown原文/PPT/EPUB文本 → self._reader_text_preview
        - PDF / 不支持类型 → None
        """
        if self._reader_pdf:
            return None
        if self._reader_file_type in CODE_VIEW_EXTS:
            return self._reader_code_text
        if self._reader_file_type in TEXT_VIEW_EXTS or self._reader_file_type in OFFICE_EXTS:
            return self._reader_text_preview
        return None

    def _reader_build_text_cache(self, generation=None):
        """构建文本缓存：每页words列表 + 词→页面倒排索引，异步执行不阻塞UI"""
        if not self._reader_pdf:
            return
        if generation is None:
            generation = self._reader_doc_generation
        if generation != self._reader_doc_generation:
            return
        doc = self._reader_pdf
        # 清空旧缓存
        self._reader_text_cache = [None] * doc.page_count
        self._reader_word_index = {}
        word_index = self._reader_word_index
        cache = self._reader_text_cache

        # 分批处理，每批处理若干页后用after处理下一批，保持UI响应
        BATCH_SIZE = 5

        def _process_batch(start):
            if generation != self._reader_doc_generation or self._reader_pdf is not doc:
                return
            end = min(start + BATCH_SIZE, doc.page_count)
            for i in range(start, end):
                page = doc[i]
                words = page.get_text("words")
                # 缓存原始数据
                cache[i] = words
                # 构建倒排索引
                seen_in_page = set()
                for w in words:
                    cleaned = self._clean_word(w[4])
                    if len(cleaned) < 1:
                        continue
                    key = cleaned.lower()
                    if key not in word_index:
                        word_index[key] = []
                    # 同一页去重（只记录第一次出现）
                    if key not in seen_in_page:
                        word_index[key].append(i)
                        seen_in_page.add(key)
                # 释放page引用，让gc回收
                del page
            if end < doc.page_count:
                self.after(10, lambda: _process_batch(end))

        # 从第0页开始分批处理
        self.after(20, lambda: _process_batch(0))

    def _reader_search_cached(self, word):
        """使用缓存快速搜索：返回 [(page_idx, fitz.Rect), ...]"""
        if not self._reader_word_index or not self._reader_text_cache:
            return None  # 缓存不可用，需回退到慢搜索

        key = word.lower()
        if key not in self._reader_word_index:
            return []

        results = []
        for pg_idx in self._reader_word_index[key]:
            words = self._reader_text_cache[pg_idx]
            for w in words:
                w_clean = self._clean_word(w[4])
                if w_clean and w_clean.lower() == key:
                    results.append((pg_idx, fitz.Rect(w[0], w[1], w[2], w[3])))
        return results

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 预览内容搜索与匹配页码列
    # ════════════════════════════════════════════════════════════════

    def _reader_preview_search_exec(self):
        """执行预览内容搜索：在右侧匹配页码列显示包含搜索文本的位置，并高亮内容。"""
        query = self._reader_preview_search.get().strip()
        if not query:
            return
        # 搜索栏输入 → 跨程序广播到勾选的目标实例（落在它们的预览搜索栏执行搜索）
        self._reader_broadcast_search(query)

        # 如果当前是非PDF的文本类文件（代码/Markdown/PPT/EPUB），在对应文本控件中搜索高亮
        active_widget = self._reader_get_active_text_widget()
        if active_widget is not None:
            active_widget.tag_remove("search", "1.0", tk.END)
            active_widget.tag_config("search", background="#ff5050", foreground="white")
            idx = "1.0"
            count = 0
            while True:
                idx = active_widget.search(query, idx, nocase=1, stopindex=tk.END)
                if not idx:
                    break
                end_idx = f"{idx}+{len(query)}c"
                active_widget.tag_add("search", idx, end_idx)
                count += 1
                idx = end_idx
            if count:
                active_widget.see(active_widget.search(query, "1.0", nocase=1, stopindex=tk.END))
            self._reader_match_count_var.set(self._fmt_match_count(0, count))
            self._reader_match_lb.delete(0, tk.END)
            self._reader_match_pages = []
            self._reader_match_counts = {}
            self._reader_refresh_main_match_flow()
            for i in range(min(count, 50)):
                self._reader_match_lb.insert(tk.END, f"  匹配 {i+1}")
            return

        if not self._reader_pdf:
            return

        # 存储高亮关键词
        self._reader_highlight_keywords = [query]
        self._reader_highlight_exact = False

        self._reader_match_lb.delete(0, tk.END)
        self._reader_match_pages = []
        self._reader_match_counts = {}
        query_lower = query.lower()

        # 尝试使用缓存搜索（精确匹配）
        if self._reader_word_index and query_lower in self._reader_word_index:
            # 精确匹配：找到该词出现的所有页面，并按缓存逐页数出真实命中数
            matched_pages = self._reader_word_index[query_lower]
            matched_pages = sorted(set(matched_pages))
            for pg in matched_pages:
                self._reader_match_pages.append(pg)
                self._reader_match_lb.insert(tk.END, f"  {pg + 1}")
                cnt = sum(
                    1 for w in (self._reader_text_cache[pg] or [])
                    if self._clean_word(w[4]).lower() == query_lower)
                self._reader_match_counts[pg] = cnt or 1
            self._reader_match_count_var.set(self._fmt_match_count(len(matched_pages), sum(self._reader_match_counts.values())))
        else:
            # 模糊匹配：逐页扫描文本内容（包含子串匹配），并统计每页命中数
            matched_pages = []
            for i in range(self._reader_pdf.page_count):
                try:
                    page = self._reader_pdf[i]
                    text = page.get_text()
                    text_lower = text.lower()
                    if query_lower in text_lower:
                        matched_pages.append(i)
                        self._reader_match_counts[i] = text_lower.count(query_lower)
                    del page
                except Exception:
                    pass
            for pg in matched_pages:
                self._reader_match_pages.append(pg)
                self._reader_match_lb.insert(tk.END, f"  {pg + 1}")
            self._reader_match_count_var.set(self._fmt_match_count(len(matched_pages), sum(self._reader_match_counts.values())))

        # ── 合并「页面添加文本」命中（搜索栏也要能搜到添加的文本）──
        tm_matches = self._reader_text_marks_matches(
            self._reader_active_file_path, query)
        if tm_matches:
            for pg, cnt in tm_matches.items():
                if pg < 0 or (self._reader_pdf and pg >= self._reader_pdf.page_count):
                    continue
                if pg not in self._reader_match_pages:
                    self._reader_match_pages.append(pg)
                    self._reader_match_lb.insert(tk.END, f"  {pg + 1}")
                self._reader_match_counts[pg] = self._reader_match_counts.get(pg, 0) + cnt
            self._reader_match_pages = sorted(set(self._reader_match_pages))
            self._reader_match_count_var.set(
                self._fmt_match_count(
                    len(self._reader_match_pages),
                    sum(self._reader_match_counts.values())))

        self._reader_refresh_main_match_flow()

        # 联动搜索开启时，把关键词（子串模糊匹配）同步高亮到其它面板
        if self._reader_linked_search.get():
            self._reader_propagate_search(query, 1)

        # 重新渲染带高亮，并跳转到第一个匹配页
        if self._reader_match_pages:
            self._reader_match_lb.selection_set(0)
            first_page = self._reader_match_pages[0]
            self._reader_page = first_page
            self._reader_render(keywords=[query])
            self._reader_update_active_page_info()
            self._reader_highlight_thumb()
            self._reader_scroll_to_page(first_page)
            self._reader_center_match_on_page(query, first_page)
        else:
            # 无匹配，也重新渲染（清除高亮）
            self._reader_highlight_keywords = []
            self._reader_render(keywords=[])
            self._reader_match_count_var.set(self._fmt_match_count(0, 0))

        # 跨文件命中提示：后台用 AnyTXT 查其它文件（不阻塞本文件高亮/跳转）
        self._reader_anytxt_cross_hint(
            1, query, self._reader_active_file_path, self._reader_anytxt_hint)

    def _reader_broadcast_search(self, query):
        """搜索栏搜索的跨程序广播：把关键词发给勾选的目标实例（落在它们的预览搜索栏）。

        IPC 收到文本后触发搜索时会置 _reader_search_broadcast_suppressed，避免
        A→B→A 无限互发（见 _ipc_poll）。
        """
        if getattr(self, "_reader_search_broadcast_suppressed", False):
            return
        if not self._cross_send_enabled.get():
            return
        self._ipc_send_text(query)

    def _reader_propagate_search(self, query, source_slot=1):
        """搜索栏的跨栏传播：把关键词（子串模糊匹配）高亮到除 source_slot 外的面板。

        搜索本身已在 source_slot 做了本地搜索 + 跳转，这里只负责把同一关键词
        同步到其它已打开的面板，让它们的匹配列和高亮一起更新——效果等价于
        "点文本联动高亮其它面板"，只是这里用子串而非整词匹配，和搜索栏的模糊
        语义保持一致。source_slot 既可以是 PDF1（预览搜索栏），也可以是 PDF2/3
        （各自的检索预览栏）。
        """
        for slot in (1, 2, 3, 4):
            if slot == source_slot:
                continue
            if slot != 1:
                state = self._reader_aux.get(slot)
                if not state or not state.get("visible"):
                    continue
            path, file_type, matches = self._reader_search_in_slot(
                slot, query, substring=True)
            if path and matches is not None:
                self._reader_update_other_slot_highlight(
                    slot, query,
                    {"type": file_type, "path": path, "matches": matches},
                    exact=False)

    def _reader_match_jump(self, event=None):
        """点击匹配页码列表 → 跳转到对应位置并高亮居中。"""
        sel = self._reader_match_lb.curselection()
        if not sel:
            return
        idx = sel[0]

        # 非PDF文本类文件（代码/Markdown/PPT/EPUB）：按第idx个匹配项跳转
        active_widget = self._reader_get_active_text_widget()
        if active_widget is not None:
            query = self._reader_preview_search.get().strip()
            if query:
                active_widget.tag_remove("search", "1.0", tk.END)
                active_widget.tag_config("search", background="#ff5050", foreground="white")
                all_indices = []
                search_idx = "1.0"
                while True:
                    search_idx = active_widget.search(query, search_idx, nocase=1, stopindex=tk.END)
                    if not search_idx:
                        break
                    all_indices.append(search_idx)
                    search_idx = f"{search_idx}+{len(query)}c"
                if 0 <= idx < len(all_indices):
                    target = all_indices[idx]
                    active_widget.see(target)
                    end_idx = f"{target}+{len(query)}c"
                    active_widget.tag_add("search", target, end_idx)
            return

        if idx < 0 or idx >= len(self._reader_match_pages):
            return
        page_idx = self._reader_match_pages[idx]

        if not self._reader_pdf or page_idx >= self._reader_pdf.page_count:
            return

        # 跳转到页面
        self._reader_page = page_idx
        query = self._reader_preview_search.get().strip()

        # 关键：先粗略滚动到目标页附近（用已知或估算的偏移量），这样接下来
        # _reader_render() 判断"可见页面范围"时才会以目标页为中心计算，
        # 渲染出来的Y偏移才是该页的精确值——而不是用跳转前的旧滚动位置
        # 算出的窗口、对目标页只能给一个估算值，导致居中不准。
        self._reader_rough_scroll_to_page(page_idx)

        # 更新高亮关键词并重新渲染（此时会基于目标页位置计算准确窗口）
        if query:
            self._reader_highlight_keywords = [query]
            self._reader_render(keywords=[query])
        else:
            self._reader_highlight_keywords = []
            self._reader_render()
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()
        # 用刚渲染出的精确偏移量，再滚动到页面顶部
        self._reader_scroll_to_page(self._reader_page)
        # 居中显示匹配区域
        self._reader_center_match_on_page(query, page_idx)

    def _reader_rough_scroll_to_page(self, page_idx):
        """
        用已知/估算的偏移量【直接】把画布滚动位置移过去，不触发任何渲染或
        异步任务——只是为了让接下来的 _reader_render() 在计算可见窗口时，
        是以目标页所在位置为基准，而不是用跳转前的旧滚动位置。
        """
        if not self._reader_pdf:
            return
        if self._reader_page_yoffsets and page_idx < len(self._reader_page_yoffsets):
            rough_y = self._reader_page_yoffsets[page_idx]
        else:
            zoom = self.PDF_DPI / 72.0 * self._reader_scale
            est_page_h = int(self._reader_pdf[0].rect.height * zoom) + 8
            rough_y = page_idx * est_page_h
        scrollregion = self._reader_canvas.cget("scrollregion")
        total_h = int(scrollregion.split()[-1]) if scrollregion else 0
        if total_h > 0:
            self._reader_canvas.yview_moveto(max(0.0, min(1.0, rough_y / total_h)))

    def _reader_center_match_on_page(self, query, page_idx):
        """
        在指定页面中查找匹配词，把匹配区域同时在【垂直和水平】两个方向都
        滚动到画布中心（适配多页连续渲染模式）。
        之所以要同时处理水平方向：当页面被放大到比画布更宽时，只做垂直
        居中并不够——匹配文字有可能落在画布左右可视范围之外，需要水平
        滚动条配合才能真正把它带到屏幕中间。
        """
        if not self._reader_pdf or not query:
            return
        try:
            page = self._reader_pdf[page_idx]
            words = page.get_text("words")
            query_lower = query.lower()
            zoom = self._reader_scale * (self.PDF_DPI / 72.0)
            crop_x, crop_y = (
                self._reader_page_crop_origins[page_idx]
                if page_idx < len(self._reader_page_crop_origins) else (0.0, 0.0))
            match_rects = []
            for w in words:
                word_text = self._clean_word(w[4])
                if query_lower in word_text.lower():
                    x0 = (w[0] - crop_x) * zoom
                    y0 = (w[1] - crop_y) * zoom
                    x1 = (w[2] - crop_x) * zoom
                    y1 = (w[3] - crop_y) * zoom
                    match_rects.append([x0, y0, x1, y1])
            del page
            if not match_rects:
                return

            scrollregion = self._reader_canvas.cget("scrollregion")
            parts = scrollregion.split() if scrollregion else []
            total_w = int(parts[-2]) if len(parts) >= 4 else 1
            total_h = int(parts[-1]) if len(parts) >= 4 else 1
            canvas_w = self._reader_canvas.winfo_width()
            canvas_h = self._reader_canvas.winfo_height()

            # ── 垂直居中 ──
            page_y_offset = self._reader_page_yoffsets[page_idx] if page_idx < len(self._reader_page_yoffsets) else 0
            all_y0 = min(r[1] for r in match_rects) + page_y_offset
            all_y1 = max(r[3] for r in match_rects) + page_y_offset
            match_center_y = (all_y0 + all_y1) / 2
            if total_h > 0:
                target_top = max(0, match_center_y - canvas_h / 2)
                self._reader_canvas.yview_moveto(min(1.0, max(0.0, target_top / total_h)))

            # ── 水平居中（放大后页面比画布宽时才需要）──
            if total_w > canvas_w:
                # 读取这一页图片在画布上的真实位置，而不是重新假设按整页
                # 宽度居中——开启"智能裁边"后实际显示的是裁剪过的窄图。
                bbox = self._reader_canvas.bbox(f"page_{page_idx}")
                x_offset = bbox[0] if bbox else 0
                all_x0 = min(r[0] for r in match_rects) + x_offset
                all_x1 = max(r[2] for r in match_rects) + x_offset
                match_center_x = (all_x0 + all_x1) / 2
                target_left = max(0, match_center_x - canvas_w / 2)
                self._reader_canvas.xview_moveto(min(1.0, max(0.0, target_left / total_w)))
            else:
                # 页面没有比画布宽（比如适应宽度模式），水平方向回到最左对齐
                self._reader_canvas.xview_moveto(0.0)
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 渲染与导航
    # ════════════════════════════════════════════════════════════════

    def _reader_toggle_fit_width(self):
        """
        切换自适应宽度模式——PDF1/PDF2/PDF3共用同一个开关，跟"智能裁边"
        是同样的思路：点一次同时对三个阅读区生效，不用分别在每个面板上
        各点一次。切到"适应宽度"时三个面板都会重新按各自画布宽度计算
        缩放比例并重新渲染；切到"自由缩放"时只是停止"以后自动按宽度
        计算缩放"，三个面板各自当前的视角都保留不变（不会被重置）。
        """
        self._reader_fit_width = not self._reader_fit_width
        self._save_config()
        active = self._reader_fit_width
        if active:
            self._reader_fit_btn.config(text="🔲 适应宽度", bg=ACCENT)
        else:
            self._reader_fit_btn.config(text="🔓 自由缩放", bg="#454b63")

        # PDF1
        if self._reader_pdf:
            if active:
                self._reader_render(force_fit_width=True, keywords=self._reader_highlight_keywords)
                self._reader_page = min(self._reader_page, self._reader_pdf.page_count - 1)
                self._reader_scroll_to_page(self._reader_page)
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            # 自由缩放模式下保留当前视角，不需要重新渲染

        # PDF2 / PDF3
        for slot, state in self._reader_aux.items():
            if not state["doc"]:
                continue
            if active:
                self._reader_aux_render(slot, fit_width=True, target_page=state["page"])
            else:
                state["fit_width"] = False
                # 同样保留当前视角，不重新渲染

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 放大镜（不放大整页的情况下查看局部高清细节）
    # ════════════════════════════════════════════════════════════════



    def _reader_change_scroll_lines(self, delta):
        """调整滚轮每次滚动的行数（1-20 行）。"""
        self._reader_scroll_lines = max(1, min(20, self._reader_scroll_lines + delta))
        self._reader_scroll_lbl_var.set(str(self._reader_scroll_lines))
        self._reader_sync_speed_combo()
        self._save_config()

    def _on_reader_speed_preset(self, event=None):
        """从下拉框选择预设滚动速度。"""
        selected = self._reader_speed_combo_var.get()
        for label, lines in self._reader_speed_presets:
            if label == selected:
                self._reader_scroll_lines = lines
                self._reader_scroll_lbl_var.set(str(lines))
                self._save_config()
                break

    def _reader_sync_speed_combo(self):
        """将当前滚动行数同步到预设下拉框（精确匹配时显示预设名，否则留空）。"""
        current = self._reader_scroll_lines
        matched = None
        for label, lines in self._reader_speed_presets:
            if lines == current:
                matched = label
                break
        if matched:
            self._reader_speed_combo_var.set(matched)
        else:
            self._reader_speed_combo_var.set(f"{current}行")

    def _reader_change_magnifier_ratio(self, delta):
        """调整放大镜倍率（－/＋按钮），限制在1.5~8倍之间。"""
        new_ratio = round(self._reader_magnifier_ratio + delta, 1)
        new_ratio = max(1.5, min(8.0, new_ratio))
        self._reader_magnifier_ratio = new_ratio
        self._reader_magnifier_ratio_text.set(f"{new_ratio:.1f}x")

    def _reader_magnifier_wheel_zoom(self, event):
        """放大镜开启时，Ctrl+滚轮调节放大镜倍率并立即重渲染。返回 True 表示已处理。"""
        if not self._reader_magnifier_enabled:
            return False
        self._reader_change_magnifier_ratio(0.5 if event.delta > 0 else -0.5)
        # 立即重渲染当前可见的放大镜：锁定→按锁定位置；弹窗缩略图→按悬停位置；跟随→按最新鼠标位置
        if self._reader_magnifier_pinned:
            self._reader_magnifier_render_pinned()
        elif self._reader_magnifier_thumb_pending is not None:
            if self._reader_magnifier_thumb_job is not None:
                self.after_cancel(self._reader_magnifier_thumb_job)
                self._reader_magnifier_thumb_job = None
            self._reader_magnifier_thumb_fire()
        elif self._reader_magnifier_is_visible and self._reader_magnifier_pending is not None:
            if self._reader_magnifier_job is not None:
                self.after_cancel(self._reader_magnifier_job)
                self._reader_magnifier_job = None
            self._reader_magnifier_fire()
        return True


    def _reader_magnifier_fire(self):
        """节流计时器到点后执行：用最新一次记录的鼠标位置渲染放大镜。"""
        self._reader_magnifier_job = None
        pending = self._reader_magnifier_pending
        if pending is None or not self._reader_magnifier_enabled:
            return
        slot, ex, ey, rx, ry = pending
        self._reader_magnifier_update(slot, ex, ey, rx, ry)

    def _reader_magnifier_hide(self, event=None):
        """隐藏放大镜悬浮窗（鼠标移出画布，或关闭放大镜模式时）。"""
        if self._reader_magnifier_window is not None:
            try:
                self._reader_magnifier_window.withdraw()
            except Exception:
                pass
        self._reader_magnifier_is_visible = False
        self._reader_magnifier_last_pos = None
        self._reader_magnifier_cache = None
        self._reader_magnifier_smooth = None






    def _reader_magnifier_get_context(self, slot):
        """
        统一获取某个阅读区(1/2/3)做放大镜坐标换算所需的上下文：
        (doc, canvas, zoom, page_yoffsets, get_page_info函数)。
        get_page_info(page_idx) 返回 (x_offset, 显示宽度, 显示高度, 裁剪偏移)。
        没有打开文档/无效slot时返回None。
        """
        if slot == 1:
            doc = self._reader_pdf
            canvas = self._reader_canvas
            yoffsets = self._reader_page_yoffsets
            if not doc or not yoffsets:
                return None
            zoom = self.PDF_DPI / 72.0 * self._reader_scale

            def get_page_info(page_idx, _canvas=canvas, _doc=doc, _zoom=zoom):
                # 直接读取这一页图片在画布上的真实边界框——开启"智能裁边"后
                # 实际显示的是裁剪过的窄图，不能再假设"按整页宽度居中"。
                bbox = _canvas.bbox(f"page_{page_idx}")
                if bbox:
                    x_offset = bbox[0]
                    disp_w = bbox[2] - bbox[0]
                    disp_h = bbox[3] - bbox[1]
                else:
                    page_w_px = _doc[page_idx].rect.width * _zoom
                    x_offset = max(0, (_canvas.winfo_width() - page_w_px) / 2)
                    disp_w = page_w_px
                    disp_h = _doc[page_idx].rect.height * _zoom
                crop_origin = (
                    self._reader_page_crop_origins[page_idx]
                    if page_idx < len(self._reader_page_crop_origins) else (0.0, 0.0))
                return x_offset, disp_w, disp_h, crop_origin

            return doc, canvas, zoom, yoffsets, get_page_info
        else:
            state = self._reader_aux.get(slot)
            if not state or not state["doc"] or not state["page_yoffsets"]:
                return None
            doc = state["doc"]
            canvas = state["canvas"]
            zoom = state["zoom"]
            yoffsets = state["page_yoffsets"]

            def get_page_info(page_idx, _state=state, _canvas=canvas, _doc=doc, _zoom=zoom):
                # 同样优先读取真实渲染边界框——开启"智能裁边"后实际显示
                # 的是裁剪过的窄图，不能假设是整页尺寸。
                bbox = _canvas.bbox(f"aux_{slot}_page_{page_idx}")
                if bbox:
                    x_offset = bbox[0]
                    disp_w = bbox[2] - bbox[0]
                    disp_h = bbox[3] - bbox[1]
                else:
                    x_offset = _state["page_xoffsets"].get(page_idx, 0)
                    disp_w = _doc[page_idx].rect.width * _zoom
                    disp_h = _doc[page_idx].rect.height * _zoom
                crop_origin = _state.get("page_crop_origins", {}).get(page_idx, (0.0, 0.0))
                return x_offset, disp_w, disp_h, crop_origin

            return doc, canvas, zoom, yoffsets, get_page_info

    def _reader_magnifier_pdf_pos(self, slot, canvas_x, canvas_y):
        """把画布坐标换算成 (页码, PDF x, PDF y)；鼠标不在页面图像内时返回 None。"""
        ctx = self._reader_magnifier_get_context(slot)
        if ctx is None:
            return None
        doc, canvas, zoom, yoffsets, get_page_info = ctx
        try:
            import bisect
            cx = canvas.canvasx(canvas_x)
            cy = canvas.canvasy(canvas_y)

            page_idx = bisect.bisect_right(yoffsets, cy) - 1
            page_idx = max(0, min(page_idx, doc.page_count - 1))

            x_offset, disp_w, disp_h, (crop_x, crop_y) = get_page_info(page_idx)
            local_x_px = cx - x_offset
            local_y_px = cy - yoffsets[page_idx]
            # 鼠标不在页面图像范围内（比如停在页间空隙或左右留白），不显示放大镜
            if not (0 <= local_x_px <= disp_w and 0 <= local_y_px <= disp_h):
                return None

            # local_x_px/local_y_px是相对【当前显示的图片】（可能是裁剪后的）
            # 的坐标，要先除以缩放、再加上裁剪偏移，才是相对【整页】PDF的坐标。
            pdf_x = local_x_px / zoom + crop_x
            pdf_y = local_y_px / zoom + crop_y
            return page_idx, pdf_x, pdf_y
        except Exception:
            return None

    def _reader_scroll_to_anchor(self, slot, pdf_pos, view_xy):
        """
        Ctrl+滚轮缩放后，把 pdf_pos 这个 PDF 坐标回卷到 view_xy 这个画布
        视口坐标，实现「缩放锚定光标」——缩放前后光标指的内容保持在原地
        不动。算法复用 _reader_center_match_on_page 的 moveto 分数回卷思路。
        """
        ctx = self._reader_magnifier_get_context(slot)
        if ctx is None:
            return
        doc, canvas, zoom, yoffsets, get_page_info = ctx
        page_idx, pdf_x, pdf_y = pdf_pos
        view_x, view_y = view_xy
        try:
            if page_idx < 0 or page_idx >= doc.page_count:
                return
            x_offset, _disp_w, _disp_h, (crop_x, crop_y) = get_page_info(page_idx)
            # 缩放后同一 PDF 点在画布上的新坐标
            new_cx = x_offset + (pdf_x - crop_x) * zoom
            new_cy = yoffsets[page_idx] + (pdf_y - crop_y) * zoom

            scrollregion = canvas.cget("scrollregion")
            parts = scrollregion.split() if scrollregion else []
            total_w = int(parts[-2]) if len(parts) >= 4 else 1
            total_h = int(parts[-1]) if len(parts) >= 4 else 1
            canvas_w = canvas.winfo_width()

            # 垂直：让 PDF 点落在原视口的 Y 位置（yview_moveto 的分数是相对 total_h）
            if total_h > 0:
                target_top = max(0.0, new_cy - view_y)
                canvas.yview_moveto(min(1.0, max(0.0, target_top / total_h)))
            # 水平：仅当内容比画布宽（可横向滚动）时锚定，否则回最左对齐
            if total_w > canvas_w:
                target_left = max(0.0, new_cx - view_x)
                canvas.xview_moveto(min(1.0, max(0.0, target_left / total_w)))
            else:
                canvas.xview_moveto(0.0)
        except Exception:
            pass

    def _reader_magnifier_update(self, slot, canvas_x, canvas_y, root_x, root_y):
        """
        核心逻辑：把鼠标当前位置换算成PDF页面上的坐标，直接从PDF矢量数据
        按"当前缩放 × 放大倍率"重新渲染一块小区域（而不是把已经渲染好的
        低清图片简单放大），保证放大镜里看到的内容始终是高清的。
        支持PDF1/PDF2/PDF3三个阅读区，由传入的slot决定取哪一份数据。
        锁定（pinned）状态下不跟随鼠标，位置交由方向键微调逻辑直接渲染。
        """
        if not self._reader_magnifier_enabled or self._reader_magnifier_pinned:
            return
        pos = self._reader_magnifier_pdf_pos(slot, canvas_x, canvas_y)
        if pos is None:
            self._reader_magnifier_hide()
            return
        page_idx, pdf_x, pdf_y = pos
        # 位置平滑：衰减手部高频抖动，避免高倍率下放大镜内容"飘"。
        # 同页内做指数滑动平均；翻页立即跳转（不沿用上一页的平滑位置）。
        sm = self._reader_magnifier_smooth
        if sm is not None and sm[0] == slot and sm[1] == page_idx:
            alpha = 0.5
            pdf_x = sm[2] + alpha * (pdf_x - sm[2])
            pdf_y = sm[3] + alpha * (pdf_y - sm[3])
        self._reader_magnifier_smooth = (slot, page_idx, pdf_x, pdf_y)
        self._reader_magnifier_render_at(slot, page_idx, pdf_x, pdf_y, root_x, root_y)

    def _reader_magnifier_render_at(self, slot, page_idx, pdf_x, pdf_y, root_x, root_y):
        """在指定PDF位置渲染放大镜图像并更新悬浮窗（鼠标跟随与方向键微调共用）。"""
        ctx = self._reader_magnifier_get_context(slot)
        if ctx is None:
            return
        doc, canvas, zoom, yoffsets, get_page_info = ctx
        # PDF1用 _reader_highlight_keywords，PDF2/3用各自 aux state 的 keywords
        if slot == 1:
            active_kw = self._reader_highlight_keywords
        else:
            aux = self._reader_aux.get(slot)
            active_kw = aux["keywords"] if aux else []
        self._reader_magnifier_render_doc(
            doc, page_idx, pdf_x, pdf_y, root_x, root_y, zoom, active_kw)

    def _reader_magnifier_render_doc(self, doc, page_idx, pdf_x, pdf_y, root_x, root_y,
                                     base_zoom, active_kw=()):
        """
        通用放大镜渲染：对任意 doc/页/PDF坐标，按 base_zoom×倍率 渲染局部高清图。
        base_zoom = 内容当前的"显示缩放"（像素/PDF点）：
          主阅读区传画布显示缩放；点击弹窗缩略图传「缩略图显示宽 / 裁剪宽」。
        含 overscan 缓存：小范围移动只做廉价 crop，越界才重新 get_pixmap，
        保证高倍率下稳定不「飘」。
        """
        try:
            if page_idx < 0 or page_idx >= doc.page_count:
                return
            page = doc[page_idx]

            ratio = self._reader_magnifier_ratio
            hi_zoom = base_zoom * ratio
            half = (self._reader_magnifier_size / 2) / hi_zoom
            clip = fitz.Rect(pdf_x - half, pdf_y - half, pdf_x + half, pdf_y + half)
            clip = clip & page.rect  # 裁剪到页面范围内，避免越界

            if clip.is_empty or clip.width <= 0 or clip.height <= 0:
                self._reader_magnifier_hide()
                return

            # overscan 区域：以当前视口为中心向外扩一倍（线性 2× = 面积 4×），
            # 鼠标在区域内小范围移动只需从缓存 crop，不必每帧重新 get_pixmap。
            oc_half = half * 2.0
            oc = fitz.Rect(pdf_x - oc_half, pdf_y - oc_half,
                           pdf_x + oc_half, pdf_y + oc_half) & page.rect

            cache = self._reader_magnifier_cache
            key = (id(doc), page_idx, round(hi_zoom, 4), tuple(active_kw or []),
                   self._reader_theme.get())
            cache_hit = False
            coc = None
            if cache is not None and cache["key"] == key:
                coc = cache["oc"]
                if (coc.x0 <= clip.x0 and clip.x1 <= coc.x1
                        and coc.y0 <= clip.y0 and clip.y1 <= coc.y1):
                    cache_hit = True

            if cache_hit:
                big = cache["img"]
                box = (int(round((clip.x0 - coc.x0) * hi_zoom)),
                       int(round((clip.y0 - coc.y0) * hi_zoom)),
                       int(round((clip.x1 - coc.x0) * hi_zoom)),
                       int(round((clip.y1 - coc.y0) * hi_zoom)))
                img = big.crop(box)
            else:
                pix = page.get_pixmap(matrix=fitz.Matrix(hi_zoom, hi_zoom), clip=oc, alpha=False)
                big = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                big = self._apply_theme_to_image(big)
                if active_kw:
                    big = self._reader_highlight_page(
                        big, page, active_kw, hi_zoom, offset_x=oc.x0, offset_y=oc.y0)
                self._reader_magnifier_cache = {"key": key, "oc": oc, "img": big}
                box = (int(round((clip.x0 - oc.x0) * hi_zoom)),
                       int(round((clip.y0 - oc.y0) * hi_zoom)),
                       int(round((clip.x1 - oc.x0) * hi_zoom)),
                       int(round((clip.y1 - oc.y0) * hi_zoom)))
                img = big.crop(box)

            img = self._make_circular_magnifier_image(img)
            # 中心对准点：标记放大镜正中心对应的PDF位置（即鼠标指向处/微调锚点）
            mx = (pdf_x - clip.x0) * hi_zoom
            my = (pdf_y - clip.y0) * hi_zoom
            img = self._draw_magnifier_crosshair(img, mx, my)
            tk_img = ImageTk.PhotoImage(img)

            self._reader_show_magnifier_window(tk_img, root_x, root_y)
        except Exception:
            pass

    def _reader_magnifier_thumb_fire(self):
        """点击弹窗缩略图悬停放大镜的节流执行。"""
        self._reader_magnifier_thumb_job = None
        p = self._reader_magnifier_thumb_pending
        if p is None or not self._reader_magnifier_enabled:
            return
        doc, page_idx, pdf_x, pdf_y, root_x, root_y, base_zoom, active_kw = p
        self._reader_magnifier_render_doc(
            doc, page_idx, pdf_x, pdf_y, root_x, root_y, base_zoom, active_kw)

    def _reader_magnifier_render_pinned(self):
        """用当前锁定的位置重新渲染放大镜（方向键微调后调用）。"""
        if not self._reader_magnifier_pinned or not self._reader_magnifier_pin:
            return
        slot, page_idx, pdf_x, pdf_y, root_x, root_y = self._reader_magnifier_pin
        self._reader_magnifier_render_at(slot, page_idx, pdf_x, pdf_y, root_x, root_y)



    def _make_circular_magnifier_image(self, img):
        """
        把放大镜内容裁剪成圆形（无外框）；圆形外的区域填充成一个特殊的
        "透明色"，配合悬浮窗的 -transparentcolor 属性，Windows会把这个
        颜色当成真正的透明，从而呈现出圆形放大镜的效果（而不是方形）。

        注意：鼠标靠近页面边缘时，裁剪区域可能不是正方形（被页面边界
        切掉一部分），这里分别用宽、高各自的尺寸画椭圆，避免在这种情况
        下把圆形拉伸变形。

        性能优化：圆形遮罩(mask)只跟图像的宽高有关——绝大多数时候这个
        尺寸是固定不变的（只有贴着页面边缘移动时才会变小），每一帧重新
        画一次纯属浪费，这里按(宽,高)缓存遮罩，避免重复创建。
        """
        w, h = img.size
        mask = self._magnifier_mask_cache.get((w, h))
        if mask is None:
            mask = Image.new("L", (w, h), 0)
            ImageDraw.Draw(mask).ellipse([0, 0, w - 1, h - 1], fill=255)
            if len(self._magnifier_mask_cache) > 8:
                self._magnifier_mask_cache.clear()
            self._magnifier_mask_cache[(w, h)] = mask

        bg_rgb = self._magnifier_transparent_rgb()
        result = Image.new("RGB", (w, h), bg_rgb)
        result.paste(img, (0, 0), mask)
        # 无外框：不画强调色圆环边框，放大镜边缘就是圆形内容边界。
        return result

    def _magnifier_transparent_rgb(self):
        """放大镜悬浮窗使用的"透明色"RGB值，和 MAGNIFIER_TRANSPARENT_COLOR 保持一致。"""
        c = MAGNIFIER_TRANSPARENT_COLOR
        return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)

    def _draw_magnifier_crosshair(self, img, cx, cy):
        """在放大镜图像上画一个高对比十字对准点，帮助对准具体文本。

        红+黑双色（先粗黑描边、再细红线），深浅背景都看得清。
        cx/cy 是「鼠标正指着的 PDF 点」在放大镜图像里的像素坐标；未贴边时
        正好是图像中心，贴页边时会被裁剪偏移，所以显式传入而不是写死中心。
        """
        draw = ImageDraw.Draw(img)
        cx, cy = int(cx), int(cy)
        arm, gap, dot_r = 7, 2, 2
        dark = (20, 20, 20)
        red = (255, 70, 70)
        for color, wd in ((dark, 3), (red, 1)):
            draw.line([cx - arm, cy, cx - gap, cy], fill=color, width=wd)
            draw.line([cx + gap, cy, cx + arm, cy], fill=color, width=wd)
            draw.line([cx, cy - arm, cx, cy - gap], fill=color, width=wd)
            draw.line([cx, cy + gap, cx, cy + arm], fill=color, width=wd)
        draw.ellipse([cx - dot_r, cy - dot_r, cx + dot_r, cy + dot_r],
                     fill=red, outline=dark, width=1)
        return img

    def _monitor_rect_containing(self, x, y):
        """返回包含点(x,y)的显示器矩形 (left, top, right, bottom)；失败返回 None。

        用 MonitorFromPoint/GetMonitorInfo 拿到光标所在显示器的真实边界，
        替代 winfo_screenwidth/height（多屏下后者只返回主屏尺寸，会把放大镜
        clamp 回主屏，导致放大镜与 GUI 不在同一屏）。
        """
        try:
            user32 = ctypes.windll.user32
            user32.MonitorFromPoint.argtypes = [_MONITOR_POINT, ctypes.c_ulong]
            user32.MonitorFromPoint.restype = ctypes.c_void_p
            user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.POINTER(_MONITORINFO)]
            user32.GetMonitorInfoW.restype = ctypes.c_int
            pt = _MONITOR_POINT(int(x), int(y))
            hmon = user32.MonitorFromPoint(pt, 2)  # MONITOR_DEFAULTTONEAREST
            if not hmon:
                return None
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            if not user32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
                return None
            r = mi.rcMonitor
            return (r.left, r.top, r.right, r.bottom)
        except Exception:
            return None

    def _all_monitor_rects(self):
        """枚举所有显示器的矩形 (left, top, right, bottom)；失败/单屏时返回 None。"""
        try:
            _MONITOR_ENUM_PROC = ctypes.WINFUNCTYPE(
                ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p,
                ctypes.POINTER(_MONITOR_RECT), ctypes.c_long)
            rects = []

            def _cb(hmon, hdc, lprc, data):
                rects.append((lprc.contents.left, lprc.contents.top,
                              lprc.contents.right, lprc.contents.bottom))
                return True

            proc = _MONITOR_ENUM_PROC(_cb)
            user32 = ctypes.windll.user32
            user32.EnumDisplayMonitors.argtypes = [
                ctypes.c_void_p, ctypes.c_void_p, _MONITOR_ENUM_PROC,
                ctypes.c_long]
            user32.EnumDisplayMonitors.restype = ctypes.c_bool
            if not user32.EnumDisplayMonitors(None, None, proc, 0):
                return None
            return rects if rects else None
        except Exception:
            return None

    def _other_monitor_rect(self, own_rect):
        """选一块与 own_rect 不同的显示器矩形；仅单屏/取不到时返回 None。

        优先主屏右侧相邻屏（视觉上最顺），其次左侧，再任意其它。
        """
        rects = self._all_monitor_rects()
        if not rects or len(rects) <= 1:
            return None
        if not own_rect:
            return rects[0]
        ol, ot, orr, ob = own_rect

        def _is_own(r):
            return (abs(r[0] - ol) < 2 and abs(r[1] - ot) < 2 and
                    abs(r[2] - orr) < 2 and abs(r[3] - ob) < 2)

        others = [r for r in rects if not _is_own(r)]
        if not others:
            return None
        right = [r for r in others if r[0] >= orr - 2]
        if right:
            return min(right, key=lambda r: r[0] - orr)
        left = [r for r in others if r[2] <= ol + 2]
        if left:
            return max(left, key=lambda r: r[2])
        return others[0]

    def _reader_show_magnifier_window(self, tk_img, root_x, root_y):
        """
        创建（如果还没有）并更新跟随鼠标的圆形放大镜悬浮窗。

        性能要点：移动悬浮窗的位置（geometry）和反复调用deiconify，
        在Windows上每次都要走一遍真实的窗口管理器调用，开销比想象中大，
        每秒做几十次就很容易感觉"卡"。这里做了两个节流：
          1) deiconify只在窗口从"隐藏"变成"显示"那一刻调一次，不会
             每帧都重复调用。
          2) 鼠标只移动了几个像素时，跳过窗口重新定位（人眼根本看不出
             这几像素的差异，但能省下大量不必要的窗口移动调用）。
        """
        if self._reader_magnifier_window is None or not self._reader_magnifier_window.winfo_exists():
            win = tk.Toplevel(self)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            try:
                win.attributes("-transparentcolor", MAGNIFIER_TRANSPARENT_COLOR)
            except Exception:
                # 非Windows环境不支持-transparentcolor，退化为普通方形悬浮窗
                try:
                    win.attributes("-alpha", 0.97)
                except Exception:
                    pass
            lbl = tk.Label(win, bg=MAGNIFIER_TRANSPARENT_COLOR, bd=0, highlightthickness=0)
            lbl.pack()
            self._reader_magnifier_window = win
            self._reader_magnifier_label = lbl
            self._reader_magnifier_is_visible = False
            self._reader_magnifier_last_pos = None

        # 放在鼠标右下方一点，避免遮住鼠标指着的位置
        offset = 24
        target_x, target_y = root_x + offset, root_y + offset
        # 边缘避让：放大镜靠近屏幕右/下边缘会显示不全，翻转到鼠标左上，
        # 再兜底 clamp 到屏幕内，保证放大镜始终完整可见。
        size = self._reader_magnifier_size
        mwin = self._reader_magnifier_window
        mon = self._monitor_rect_containing(root_x, root_y)
        if mon:
            # 多屏：按「光标所在显示器」的真实边界翻转+clamp，避免被 clamp 回主屏
            ml, mt, mr, mb = mon
            if target_x + size > mr:
                target_x = root_x - offset - size
            if target_y + size > mb:
                target_y = root_y - offset - size
            target_x = max(ml, min(target_x, mr - size))
            target_y = max(mt, min(target_y, mb - size))
        else:
            # 回退：无法取到显示器边界时，沿用 winfo_screenwidth/height
            sw = mwin.winfo_screenwidth()
            sh = mwin.winfo_screenheight()
            if target_x + size > sw:
                target_x = root_x - offset - size
            if target_y + size > sh:
                target_y = root_y - offset - size
            target_x = max(0, min(target_x, sw - size))
            target_y = max(0, min(target_y, sh - size))
        last = self._reader_magnifier_last_pos
        moved_enough = (last is None or
                         abs(target_x - last[0]) >= 3 or abs(target_y - last[1]) >= 3)
        if moved_enough:
            self._reader_magnifier_window.geometry(f"+{target_x}+{target_y}")
            self._reader_magnifier_last_pos = (target_x, target_y)

        self._reader_magnifier_label.config(image=tk_img)
        self._reader_magnifier_label.image = tk_img  # 保留引用，防止被垃圾回收

        if not self._reader_magnifier_is_visible:
            self._reader_magnifier_window.deiconify()
            self._reader_magnifier_is_visible = True

    def _pick_reader_annotate_color(self):
        """选择标注高亮框的颜色。"""
        color = colorchooser.askcolor(
            color=self._reader_annotate_color, title="选择标注颜色")
        if not color or not color[1]:
            return
        self._reader_annotate_color = color[1].lower()
        self._reader_annotate_color_btn.config(bg=self._reader_annotate_color)



    def _toggle_reader_crop_margins(self):
        """
        开启/关闭"智能裁边"——PDF1/PDF2/PDF3共用同一个开关，点一次同时
        对三个阅读区生效，并自动切到"适应宽度"模式（裁边之后页面变窄，
        配合适应宽度才能让内容真正撑满画布、达到"小窗口显示更大内容"
        的效果）。

        之所以做成共用一个开关而不是三个面板各自独立：核心的"检测内容
        边界"结果本来就是按(文件路径,页码)缓存、PDF1/2/3之间天然共享的，
        三个独立开关并不会减少这部分计算量，唯一区别只是"能不能只裁
        某一个面板"这种精细控制——换来的是一键三个面板一起生效、还顺手
        切到适应宽度的更顺手体验。
        """
        self._reader_crop_margins_mode = not self._reader_crop_margins_mode
        self._save_config()
        active = self._reader_crop_margins_mode
        if active:
            self._reader_crop_btn.config(text="✂ 智能裁边(开)", bg=ACCENT)
        else:
            self._reader_crop_btn.config(text="✂ 智能裁边", bg="#454b63")

        # PDF1
        if self._reader_pdf:
            if active:
                self._reader_fit_width = True
                self._reader_fit_btn.config(text="🔲 适应宽度", bg=ACCENT)
            self._reader_render(
                force_fit_width=active, keywords=self._reader_highlight_keywords)
            self._reader_scroll_to_page(self._reader_page)

        # PDF2 / PDF3
        for slot, state in self._reader_aux.items():
            if not state["doc"]:
                continue
            self._reader_aux_render(
                slot, fit_width=True if active else None,
                target_page=state["page"])


    def _reader_sync_scroll_others(self, source_slot):
        """
        把 source_slot 当前的页码同步到其它已打开文档的阅读区（按页码，越界钳位）。
        用 _reader_sync_active 做重入保护，防止 A→B→A 递归回传。
        """
        if not self._reader_sync_scroll or self._reader_sync_active:
            return
        src_page = self._reader_page_for_slot(source_slot)
        self._reader_sync_active = True
        try:
            for s in (1, 2, 3, 4):
                if s == source_slot:
                    continue
                doc = self._reader_doc_for_slot(s)
                if not doc:
                    continue
                target = max(0, min(src_page, doc.page_count - 1))
                if s == 1:
                    if self._reader_page != target:
                        self._reader_page = target
                        self._reader_scroll_to_page(target)
                        self._reader_update_active_page_info()
                        self._reader_highlight_thumb()
                else:
                    state = self._reader_aux[s]
                    if state["page"] != target:
                        state["page"] = target
                        self._reader_aux_render(s, fit_width=False, target_page=target)
        finally:
            self._reader_sync_active = False

    def _reader_sync_scale_others(self, source_slot):
        """把 source_slot 当前的缩放比例同步到其它已打开文档的阅读区。"""
        if not self._reader_sync_scroll or self._reader_sync_active:
            return
        if source_slot == 1:
            scale = self._reader_scale
        else:
            state = self._reader_aux.get(source_slot)
            if not state:
                return
            scale = state["scale"]
        self._reader_sync_active = True
        try:
            for s in (1, 2, 3, 4):
                if s == source_slot:
                    continue
                doc = self._reader_doc_for_slot(s)
                if not doc:
                    continue
                if s == 1:
                    self._reader_scale = scale
                    self._reader_fit_width = False
                    self._reader_render(
                        force_fit_width=False,
                        keywords=self._reader_highlight_keywords, quick=True)
                else:
                    state = self._reader_aux[s]
                    state["scale"] = scale
                    state["fit_width"] = False
                    self._reader_aux_render(
                        s, fit_width=False, target_page=state["page"], quick=True)
        finally:
            self._reader_sync_active = False

    def _toggle_reader_annotate_mode(self):
        """
        开启/关闭"标注模式"。开启后，在PDF1画布上按住左键拖动会画出一个
        半透明的高亮框（松开鼠标后保存），用来标记页面上的重点区域——
        和搜索关键词自动高亮不同，这是用户自己手动画的标注。
        开启标注模式期间会暂时接管画布的左键事件（平时左键拖动是"拖动
        画布"、单击是"点字搜索"），关闭后恢复原来的行为。
        """
        self._reader_annotate_mode = not self._reader_annotate_mode
        if self._reader_annotate_mode:
            if self._reader_text_mode:
                self._reader_text_mode = False
                if getattr(self, "_reader_text_btn", None):
                    self._reader_text_btn.config(text="T 文本", bg="#454b63")
            self._reader_annotate_btn.config(text="✏ 标注(开)", bg=ACCENT)
            self._reader_canvas.config(cursor="tcross")
            self._reader_canvas.bind("<ButtonPress-1>", self._reader_annotate_press)
            self._reader_canvas.bind("<B1-Motion>", self._reader_annotate_drag)
            self._reader_canvas.bind("<ButtonRelease-1>", self._reader_annotate_release)
        else:
            self._reader_annotate_btn.config(text="✏ 标注", bg="#454b63")
            self._reader_canvas.config(cursor="arrow")
            self._reader_canvas.bind("<ButtonPress-1>", lambda e: self._reader_on_canvas_press(1, e))
            self._reader_canvas.bind("<B1-Motion>", lambda e: self._reader_on_canvas_drag(1, e))
            self._reader_canvas.unbind("<ButtonRelease-1>")

    def _reader_annotate_press(self, event):
        """标注模式：按下左键，记录起点，开始画一个临时的矩形。"""
        x = self._reader_canvas.canvasx(event.x)
        y = self._reader_canvas.canvasy(event.y)
        self._reader_annotate_start_xy = (x, y)
        self._reader_annotate_rect_id = self._reader_canvas.create_rectangle(
            x, y, x, y, outline=self._reader_annotate_color, width=2,
            fill=self._reader_annotate_color, stipple="gray25")

    def _reader_annotate_drag(self, event):
        """标注模式：拖动鼠标，实时更新矩形大小。"""
        if not self._reader_annotate_rect_id or not self._reader_annotate_start_xy:
            return
        x = self._reader_canvas.canvasx(event.x)
        y = self._reader_canvas.canvasy(event.y)
        x0, y0 = self._reader_annotate_start_xy
        self._reader_canvas.coords(self._reader_annotate_rect_id, x0, y0, x, y)

    def _reader_annotate_release(self, event):
        """标注模式：松开左键，把矩形换算成PDF页面坐标后保存下来。"""
        if not self._reader_annotate_rect_id or not self._reader_annotate_start_xy:
            return
        x0, y0 = self._reader_annotate_start_xy
        x1 = self._reader_canvas.canvasx(event.x)
        y1 = self._reader_canvas.canvasy(event.y)
        self._reader_annotate_start_xy = None
        # 太小的矩形基本上就是误触的单击，不保存
        if abs(x1 - x0) < 5 and abs(y1 - y0) < 5:
            self._reader_canvas.delete(self._reader_annotate_rect_id)
            self._reader_annotate_rect_id = None
            return
        # 先删掉临时矩形，避免点「取消」时残留（保存时由 on_confirm 重新渲染正式框）
        self._reader_canvas.delete(self._reader_annotate_rect_id)
        self._reader_annotate_rect_id = None
        # 画框后弹标签/备注框，确认后保存（标签可留空，取消则放弃本次画框）
        def _on_confirm(label, note):
            self._reader_save_annotation_rect(x0, y0, x1, y1, label, note)
            # 用正式渲染流程重新画一遍，这样标注会作为统一管理的画布元素出现
            # （支持后续右键删除/左键拖拽），而不是停留在刚才临时绘制的那个矩形上
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords)

        self._reader_show_bookmark_dialog("标注本页位置", on_confirm=_on_confirm)

    def _reader_save_annotation_rect(self, cx0, cy0, cx1, cy1, label="", note=""):
        """把画布像素坐标的矩形换算成PDF页面点坐标，按文件+页码保存。"""
        import bisect
        full_path = self._reader_active_file_path
        if not full_path or not self._reader_pdf:
            return
        yoffsets = self._reader_page_yoffsets
        if not yoffsets:
            return
        top_y = min(cy0, cy1)
        page_idx = bisect.bisect_right(yoffsets, top_y) - 1
        page_idx = max(0, min(page_idx, self._reader_pdf.page_count - 1))
        zoom = self.PDF_DPI / 72.0 * self._reader_scale

        # 直接读取这一页图片在画布上的实际位置，而不是重新假设"按整页宽度
        # 居中"去计算——开启"智能裁边"后实际显示的是裁剪过的窄图，重新假设
        # 整页宽度算出来的位置会跟真实渲染的位置不一致。
        bbox = self._reader_canvas.bbox(f"page_{page_idx}")
        x_offset = bbox[0] if bbox else 0
        page_y_offset = yoffsets[page_idx]
        crop_x, crop_y = (
            self._reader_page_crop_origins[page_idx]
            if page_idx < len(self._reader_page_crop_origins) else (0.0, 0.0))

        # 减去画布偏移、除以缩放，得到的是"裁剪后图片"内部的相对坐标；
        # 再加上crop_origin才是相对于整页PDF左上角的坐标（标注是按整页
        # 坐标系存储的，这样不管之后开不开裁边，标注位置都一致、可还原）。
        local_x0 = (min(cx0, cx1) - x_offset) / zoom + crop_x
        local_x1 = (max(cx0, cx1) - x_offset) / zoom + crop_x
        local_y0 = (min(cy0, cy1) - page_y_offset) / zoom + crop_y
        local_y1 = (max(cy0, cy1) - page_y_offset) / zoom + crop_y

        by_page = self._reader_page_marks.setdefault(full_path, {})
        marks = by_page.setdefault(str(page_idx), [])
        marks.append({
            "rect": [local_x0, local_y0, local_x1, local_y1],
            "color": self._reader_annotate_color,
            "label": label,
            "note": note,
        })
        self._save_config()

    def _reader_draw_page_annotations(self, canvas, page_idx, x_offset, y_cursor, zoom,
                                       full_path=None, crop_origin=(0.0, 0.0)):
        """
        把指定页面已保存的标注框，按当前缩放/位置换算成画布坐标后画出来。
        在每次渲染某一页之后调用——标注是单独的画布矩形元素叠加在页面图片
        上面，不会"烫"进图片本身，所以缩放、主题切换都不需要重新生成标注。

        crop_origin：如果"智能裁边"把这一页的图片裁掉了一部分空白（图片
        左上角不再对应PDF页面的(0,0)，而是对应crop_origin这个点），标注
        框保存的还是相对于【整页】的PDF坐标，所以这里要先减去crop_origin
        再乘以缩放，才能落在裁剪后图片上的正确位置。
        """
        path = full_path if full_path is not None else self._reader_active_file_path
        if not path:
            return
        crop_x, crop_y = crop_origin
        # 页面添加的文本（覆盖层）与高亮原文，随每页渲染一起叠加
        self._reader_draw_page_text_marks(
            canvas, page_idx, x_offset, y_cursor, zoom, path, crop_origin)
        self._reader_draw_page_text_hl(
            canvas, page_idx, x_offset, y_cursor, zoom, path, crop_origin)
        by_page = self._reader_page_marks.get(path, {})
        marks = by_page.get(str(page_idx))
        if not marks:
            return
        for mark in marks:
            rect = mark.get("rect")
            if not rect or len(rect) != 4:
                continue
            x0, y0, x1, y1 = rect
            rect_id = canvas.create_rectangle(
                x_offset + (x0 - crop_x) * zoom, y_cursor + (y0 - crop_y) * zoom,
                x_offset + (x1 - crop_x) * zoom, y_cursor + (y1 - crop_y) * zoom,
                outline=mark.get("color", "#ffdc00"), width=2,
                fill=mark.get("color", "#ffdc00"), stipple="gray25",
                tags=("reader_annotation",))
            canvas.tag_bind(
                rect_id, "<Button-3>",
                lambda event, p=path, pi=page_idx, m=mark:
                    self._reader_delete_annotation(p, pi, m))
            # 左键拖拽移动（press/motion 返回 "break" 阻止冒泡到画布的文字点击搜索）
            canvas.tag_bind(
                rect_id, "<ButtonPress-1>",
                lambda event, p=path, pi=page_idx, m=mark, rid=rect_id, z=zoom:
                    self._reader_annot_rect_press(event, p, pi, m, rid, z))
            canvas.tag_bind(
                rect_id, "<B1-Motion>",
                lambda event, rid=rect_id: self._reader_annot_rect_motion(event, rid))
            canvas.tag_bind(
                rect_id, "<ButtonRelease-1>", self._reader_annot_rect_release)

    def _reader_delete_annotation(self, full_path, page_idx, mark):
        """右键点击某个标注框 → 删除它。"""
        by_page = self._reader_page_marks.get(full_path, {})
        marks = by_page.get(str(page_idx))
        if marks and mark in marks:
            marks.remove(mark)
            if not marks:
                del by_page[str(page_idx)]
            self._save_config()
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords)

    # ── 页面添加文本（覆盖层，不写入PDF）──

    def _toggle_reader_text_mode(self):
        """开启/关闭「添加文本」模式：点击页面弹输入框，在该位置放一个文本标注。

        与「✏ 标注」模式互斥，开启文本模式会先关闭标注模式。"""
        self._reader_text_mode = not self._reader_text_mode
        if self._reader_text_mode:
            if self._reader_annotate_mode:
                self._toggle_reader_annotate_mode()
            self._reader_text_btn.config(text="T 文本(开)", bg=ACCENT)
            self._reader_canvas.config(cursor="tcross")
            self._reader_canvas.bind("<ButtonPress-1>", self._reader_add_text_click)
        else:
            self._reader_text_btn.config(text="T 文本", bg="#454b63")
            self._reader_canvas.config(cursor="arrow")
            self._reader_canvas.bind("<ButtonPress-1>", lambda e: self._reader_on_canvas_press(1, e))

    def _reader_add_text_click(self, event):
        """添加文本模式：点击页面，换算成PDF坐标后弹输入框，保存文本标注。"""
        import bisect
        full_path = self._reader_active_file_path
        if not full_path or not self._reader_pdf:
            return
        canvas = self._reader_canvas
        x = canvas.canvasx(event.x)
        y = canvas.canvasy(event.y)
        yoffsets = self._reader_page_yoffsets
        if not yoffsets:
            return
        page_idx = bisect.bisect_right(yoffsets, y) - 1
        page_idx = max(0, min(page_idx, self._reader_pdf.page_count - 1))
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        bbox = canvas.bbox(f"page_{page_idx}")
        x_offset = bbox[0] if bbox else 0
        page_y_offset = yoffsets[page_idx]
        crop_x, crop_y = (
            self._reader_page_crop_origins[page_idx]
            if page_idx < len(self._reader_page_crop_origins) else (0.0, 0.0))
        pdf_x = (x - x_offset) / zoom + crop_x
        pdf_y = (y - page_y_offset) / zoom + crop_y

        text = simpledialog.askstring("添加文本", "在此处添加的文字：", parent=self)
        if text is None or not text.strip():
            return
        by_page = self._reader_text_marks.setdefault(full_path, {})
        marks = by_page.setdefault(str(page_idx), [])
        marks.append({"x": pdf_x, "y": pdf_y, "text": text.strip(),
                      "color": self._reader_text_color})
        self._save_config()
        self._reader_render(
            force_fit_width=False, keywords=self._reader_highlight_keywords)

    def _reader_draw_page_text_marks(self, canvas, page_idx, x_offset, y_cursor, zoom,
                                     full_path, crop_origin=(0.0, 0.0)):
        """把该页已添加的文本标注画到画布上（色块底 + 黑字），右键可编辑/删除。"""
        by_page = self._reader_text_marks.get(full_path, {})
        marks = by_page.get(str(page_idx))
        if not marks:
            return
        crop_x, crop_y = crop_origin
        for tm in marks:
            x = tm.get("x"); y = tm.get("y")
            text = tm.get("text", "")
            color = tm.get("color", "#ffdc00")
            if x is None or y is None or not text:
                continue
            cx = x_offset + (x - crop_x) * zoom
            cy = y_cursor + (y - crop_y) * zoom
            tid = canvas.create_text(
                cx, cy, text=text, anchor="nw", fill="#111111",
                font=("Microsoft YaHei UI", 10, "bold"),
                tags=("reader_text_mark",))
            bbox = canvas.bbox(tid)
            if bbox:
                bg_id = canvas.create_rectangle(
                    bbox[0] - 2, bbox[1] - 1, bbox[2] + 2, bbox[3] + 1,
                    fill=color, outline="", tags=("reader_text_mark",))
                canvas.tag_lower(bg_id, tid)
            canvas.tag_bind(
                tid, "<Button-3>",
                lambda e, p=full_path, pi=page_idx, m=tm:
                    self._reader_text_mark_menu(e, p, pi, m))

    def _reader_text_mark_menu(self, event, full_path, page_idx, mark):
        """文本标注的右键菜单：编辑 / 删除。"""
        menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground="white",
                       font=("Microsoft YaHei UI", 9))
        menu.add_command(label="✏ 编辑文字",
                         command=lambda: self._reader_edit_text_mark(full_path, page_idx, mark))
        menu.add_command(label="🗑 删除",
                         command=lambda: self._reader_delete_text_mark(full_path, page_idx, mark))
        menu.tk_popup(event.x_root, event.y_root)

    def _reader_edit_text_mark(self, full_path, page_idx, mark):
        new = simpledialog.askstring(
            "编辑文字", "文字内容：", initialvalue=mark.get("text", ""), parent=self)
        if new is None or not new.strip():
            return
        mark["text"] = new.strip()
        self._save_config()
        self._reader_render(
            force_fit_width=False, keywords=self._reader_highlight_keywords)

    def _reader_delete_text_mark(self, full_path, page_idx, mark):
        by_page = self._reader_text_marks.get(full_path, {})
        marks = by_page.get(str(page_idx))
        if marks and mark in marks:
            marks.remove(mark)
            if not marks:
                del by_page[str(page_idx)]
            self._save_config()
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords)

    def _reader_text_marks_matches(self, full_path, query):
        """返回「页面添加文本」中匹配 query 的 {页索引: 命中次数}（供搜索栏搜添加文本）。"""
        if not full_path or not query:
            return {}
        by_page = self._reader_text_marks.get(full_path, {})
        if not by_page:
            return {}
        q = query.lower()
        matches = {}
        for page_str, marks in by_page.items():
            cnt = 0
            for tm in marks:
                txt = str(tm.get("text", "")).lower()
                if q in txt:
                    cnt += txt.count(q)
            if cnt:
                try:
                    matches[int(page_str)] = cnt
                except (TypeError, ValueError):
                    pass
        return matches

    # ── 高亮原文（点选文字 → 持久高亮 + 存原文文本）──

    def _toggle_reader_text_hl_mode(self):
        """开启/关闭「高亮收藏」：点击文字 → 持久高亮并把原文文本存入标记索引。"""
        self._reader_text_hl_mode = not self._reader_text_hl_mode
        if self._reader_text_hl_mode:
            self._reader_text_hl_btn.config(text="🖍 高亮收藏(开)", bg=ACCENT)
        else:
            self._reader_text_hl_btn.config(text="🖍 高亮收藏", bg="#454b63")

    def _reader_save_text_hl(self, page_idx, rect, word):
        """把点选的文字保存为持久高亮（含原文文本），供全局标记搜索反向定位。"""
        full_path = self._reader_active_file_path
        if not full_path:
            return
        by_page = self._reader_text_hl.setdefault(full_path, {})
        marks = by_page.setdefault(str(page_idx), [])
        rect_f = [float(v) for v in rect]
        for m in marks:
            if m.get("text", "").lower() == word.lower() and m.get("rect") == rect_f:
                return  # 去重
        marks.append({"rect": rect_f, "text": word, "color": self._reader_text_hl_color})
        self._save_config()
        self._reader_render(
            force_fit_width=False, keywords=self._reader_highlight_keywords)

    def _reader_draw_page_text_hl(self, canvas, page_idx, x_offset, y_cursor, zoom,
                                  full_path, crop_origin=(0.0, 0.0)):
        """把该页的「高亮原文」矩形叠加到画布上，右键可删除。"""
        by_page = self._reader_text_hl.get(full_path, {})
        marks = by_page.get(str(page_idx))
        if not marks:
            return
        crop_x, crop_y = crop_origin
        for hl in marks:
            rect = hl.get("rect")
            if not rect or len(rect) != 4:
                continue
            x0, y0, x1, y1 = rect
            rid = canvas.create_rectangle(
                x_offset + (x0 - crop_x) * zoom, y_cursor + (y0 - crop_y) * zoom,
                x_offset + (x1 - crop_x) * zoom, y_cursor + (y1 - crop_y) * zoom,
                outline=hl.get("color", "#ff5050"), width=1,
                fill=hl.get("color", "#ff5050"), stipple="gray25",
                tags=("reader_text_hl",))
            canvas.tag_bind(
                rid, "<Button-3>",
                lambda e, p=full_path, pi=page_idx, m=hl:
                    self._reader_delete_text_hl(p, pi, m))

    def _reader_delete_text_hl(self, full_path, page_idx, mark):
        by_page = self._reader_text_hl.get(full_path, {})
        marks = by_page.get(str(page_idx))
        if marks and mark in marks:
            marks.remove(mark)
            if not marks:
                del by_page[str(page_idx)]
            self._save_config()
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords)


    def _pick_reader_highlight_color(self):
        """选择PDF阅读器点击文字与预览搜索的高亮颜色。"""
        color = colorchooser.askcolor(
            color=self._reader_highlight_hex.get(), title="选择PDF文本高亮颜色")
        if not color or not color[0] or not color[1]:
            return
        self._reader_highlight_color = tuple(int(v) for v in color[0])
        self._reader_highlight_hex.set(color[1].lower())
        self._reader_highlight_btn.config(bg=color[1], fg="#111111")
        if self._reader_pdf and self._reader_highlight_keywords:
            self._reader_render(
                force_fit_width=False,
                keywords=self._reader_highlight_keywords)
        for slot, state in self._reader_aux.items():
            if state["doc"] and state["keywords"]:
                self._reader_aux_render(
                    slot, fit_width=False, target_page=state["page"])
        self._save_config()



    def _reader_dual_open_selected(self, slot=2):
        """将文件列表中选中的文件在指定阅读区打开（支持所有支持的文件格式）。"""
        selection = self._reader_lb.curselection()
        if not selection:
            messagebox.showinfo("双开阅读", "请先在左侧文件列表选择一份文件")
            return
        idx = self._get_reader_actual_index(selection[0])
        if idx < 0:
            return
        path, _ = self._reader_pdf_files[idx]
        self._reader_dual_open_path(path, slot)


    def _reader_dual_open_path(self, path, slot=2):
        """在PDF 2或PDF 3会话中打开文档——支持PDF/EPUB/PPT/PPTX四种格式。"""
        if slot not in self._reader_aux:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext == ".epub":
            self._reader_dual_open_epub(path, slot)
            return
        if ext in (".ppt", ".pptx"):
            self._reader_dual_open_ppt(path, slot)
            return
        try:
            new_doc = fitz.open(path)
            if new_doc.page_count < 1:
                new_doc.close()
                raise ValueError("PDF没有可显示页面")
        except Exception as exc:
            messagebox.showerror(f"PDF {slot} 打开失败", str(exc))
            return
        self._reader_dual_apply_doc(new_doc, path, slot)

    def _reader_dual_apply_doc(self, new_doc, path, slot, icon="📄"):
        """
        把已经准备好的fitz.Document（直接打开的PDF，或者EPUB/PPT转换出来
        的"伪PDF"）应用到指定阅读区——这部分逻辑是PDF/PPT共用的，因为
        PPT在转换完成后就是一个普通的fitz.Document，跟原生PDF用同一套
        画布渲染流程显示。
        """
        state = self._reader_aux[slot]
        self._reader_save_aux_resume(slot)
        self._record_open(path)
        self._reader_clear_merged_slot(slot)
        for key in ("resize_job", "scroll_anim_job"):
            if state[key]:
                try:
                    self.after_cancel(state[key])
                except Exception:
                    pass
                state[key] = None
        state["scroll_velocity"] = 0.0
        if state["doc"]:
            try:
                state["doc"].close()
            except Exception:
                pass
        # 切到画布视图（如果之前显示的是EPUB内容）
        if state.get("epub_frame") is not None:
            state["epub_frame"].grid_remove()
        state["canvas_frame"].grid()
        state.update({
            "doc": new_doc, "path": path, "page": 0, "scale": 1.0,
            "fit_width": True, "keywords": [], "exact": False,
            "word_cache": None,
        })
        # 续读：恢复上次的页码/缩放/自适应状态
        resume = self._reader_resume.get(path)
        if resume:
            try:
                state["page"] = max(0, min(int(resume.get("page", 0)),
                                           new_doc.page_count - 1))
                state["scale"] = float(resume.get("scale", 1.0))
                state["fit_width"] = bool(resume.get("fit_width", True))
            except (TypeError, ValueError):
                pass
        state["search_var"].set("")
        state["search_info"].set("")
        state["match_pages"] = []
        state["match_counts"] = {}
        state["match_lb"].delete(0, tk.END)
        state["match_count_var"].set("")
        self._reader_refresh_aux_match_flow(slot)
        if self._reader_match_visible:
            state["match_frame"].grid()
        if not state["visible"]:
            self._reader_view_paned.add(
                state["frame"], minsize=250, stretch="always")
            state["visible"] = True
        self._reader_active_slot = slot
        self._reader_aux_load_toc(slot)
        self._reader_refresh_bookmark_list(slot)
        self._reader_set_active_slot(slot)
        self._reader_update_multi_button()
        if slot == 2:
            self._reader_pdf2 = new_doc
            self._reader_path2 = path
        rel_path = os.path.basename(path)
        state["title_var"].set(f"{icon} PDF {slot}: {rel_path}")
        resume_fit = state["fit_width"]
        resume_page = state["page"]
        self.after_idle(
            lambda s=slot, f=resume_fit, p=resume_page: self._reader_aux_render(
                s, fit_width=f, target_page=p))

    def _reader_dual_open_epub(self, path, slot):
        """在PDF2/PDF3里打开EPUB文件——跟PDF1一样，完整解析章节正文，
        套用当前的背景色/字号/字体/行间距设置，显示在专门的文本区域里
        （不是画布渲染，所以不需要fitz.Document）。"""
        state = self._reader_aux[slot]
        chapters = self._epub_parse_chapters(path)
        self._reader_save_aux_resume(slot)
        self._record_open(path)
        self._reader_clear_merged_slot(slot)
        if state["doc"]:
            try:
                state["doc"].close()
            except Exception:
                pass
        state.update({
            "doc": None, "path": path, "page": 0, "scale": 1.0,
            "keywords": [], "exact": False, "word_cache": None,
        })
        state["canvas_frame"].grid_remove()
        state["epub_frame"].grid(row=0, column=1, sticky="nsew")
        self._epub_render_into_text_widget(
            state["epub_text"], chapters, os.path.basename(path), full_path=path)
        if not getattr(state["epub_text"], "_epub_annotate_bound", False):
            self._epub_bind_annotate_events(state["epub_text"])
            state["epub_text"]._epub_annotate_bound = True
        state["match_pages"] = []
        state["match_counts"] = {}
        state["match_lb"].delete(0, tk.END)
        state["match_count_var"].set("")
        self._reader_refresh_aux_match_flow(slot)
        if self._reader_match_visible:
            state["match_frame"].grid()
        if not state["visible"]:
            self._reader_view_paned.add(
                state["frame"], minsize=250, stretch="always")
            state["visible"] = True
        self._reader_active_slot = slot
        self._reader_set_active_slot(slot)
        self._reader_update_multi_button()
        state["title_var"].set(f"📚 PDF {slot}: {os.path.basename(path)}")
        state["page_var"].set("EPUB文件")

    def _reader_dual_open_ppt(self, full_path, slot):
        """
        在PDF2/PDF3里打开PPT/PPTX文件——通过COM把每一页渲染成图片后
        拼成一个内存中的"伪PDF"（fitz.Document），之后就跟普通PDF走
        同一套画布渲染流程。跟PDF1共用同一份转换函数和缓存
        （_reader_extract_ppt_via_com / _reader_ppt_data_cache），
        同一个PPT文件不会因为在PDF1还是PDF2/3打开就重复转换一次。
        """
        state = self._reader_aux[slot]
        state["title_var"].set(f"⏳ PDF {slot}: 正在渲染PPT…")

        mtime = None
        try:
            mtime = os.path.getmtime(full_path)
        except Exception:
            pass
        cached = self._reader_ppt_data_cache.get(full_path)
        cache_hit = cached is not None and mtime is not None and cached[0] == mtime

        generation = self._reader_doc_generation

        def worker():
            if cache_hit:
                _mtime, slide_w, slide_h, slides_data = cached
                err = None
            else:
                slide_w, slide_h, slides_data, err = self._reader_extract_ppt_via_com(full_path)
                if slides_data and mtime is not None:
                    self._reader_ppt_data_cache[full_path] = (mtime, slide_w, slide_h, slides_data)
                    self._reader_ppt_data_cache.move_to_end(full_path)
                    while len(self._reader_ppt_data_cache) > self._reader_ppt_cache_max:
                        self._reader_ppt_data_cache.popitem(last=False)

            doc = None
            if slides_data:
                try:
                    doc = self._reader_build_ppt_doc(slide_w, slide_h, slides_data)
                except Exception as e:
                    doc = None
                    err = err or str(e)

            def apply():
                if self._reader_doc_generation != generation:
                    if doc is not None:
                        try:
                            doc.close()
                        except Exception:
                            pass
                    return
                if doc is not None and doc.page_count > 0:
                    self._reader_dual_apply_doc(doc, full_path, slot, icon="📊")
                else:
                    messagebox.showerror(
                        f"PDF {slot} 打开失败",
                        f"PPT渲染失败：{err or '未知错误'}")
                    state["title_var"].set(f"❌ PDF {slot}: 渲染失败")

            self.after(0, apply)

        threading.Thread(target=worker, daemon=True).start()


    def _reader_aux_close(self, slot):
        state = self._reader_aux.get(slot)
        if not state:
            return
        self._reader_clear_merged_slot(slot)
        for key in ("resize_job", "scroll_anim_job"):
            if state[key]:
                try:
                    self.after_cancel(state[key])
                except Exception:
                    pass
                state[key] = None
        state["scroll_velocity"] = 0.0
        state["occurrence_generation"] += 1
        if state["occurrence_window"]:
            try:
                state["occurrence_window"].destroy()
            except Exception:
                pass
            state["occurrence_window"] = None
        if state["doc"]:
            try:
                state["doc"].close()
            except Exception:
                pass
        state["doc"] = None
        state["path"] = ""
        state["tk_img"] = None
        state["keywords"] = []
        state["canvas"].delete("all")
        if state["visible"]:
            try:
                self._reader_view_paned.remove(state["frame"])
            except Exception:
                pass
        state["visible"] = False
        state["title_var"].set(f"PDF {slot}: —")
        state["page_var"].set("— / —")
        state["search_info"].set("")
        state["match_pages"] = []
        state["match_counts"] = {}
        state["match_lb"].delete(0, tk.END)
        state["match_count_var"].set("")
        self._reader_refresh_aux_match_flow(slot)
        if slot == 2:
            self._reader_pdf2 = None
            self._reader_path2 = ""
        if self._reader_active_slot == slot:
            self._reader_set_active_slot(1)
        self._reader_update_multi_button()

    def _reader_update_multi_button(self):
        opened = sum(
            1 for state in self._reader_aux.values() if state["visible"])
        self._reader_dual_visible = opened > 0
        if hasattr(self, "_reader_dual_btn"):
            if opened:
                self._reader_dual_btn.config(
                    bg=ACCENT, text=f"▦ 当前PDF {self._reader_active_slot}")
            else:
                self._reader_dual_btn.config(
                    bg="#454b63", text="▦ 当前PDF 1")

    def _reader_aux_load_toc(self, slot):
        state = self._reader_aux[slot]
        toc_lb = state["toc_lb"]
        toc_lb.delete(0, tk.END)
        state["toc_entries"] = []
        doc = state["doc"]
        if not doc:
            toc_lb.insert(tk.END, "未打开PDF")
            return
        try:
            toc = doc.get_toc(simple=True) or []
        except Exception:
            toc = []
        for entry in toc:
            if len(entry) < 3:
                continue
            level, title, page_number = entry[:3]
            try:
                page_idx = int(page_number) - 1
                level = max(1, int(level))
            except (TypeError, ValueError):
                continue
            if not 0 <= page_idx < doc.page_count:
                continue
            state["toc_entries"].append(page_idx)
            indent = "  " * min(level - 1, 5)
            clean_title = " ".join(str(title).split()) or f"第{page_idx + 1}页"
            toc_lb.insert(tk.END, f"{indent}{clean_title} · {page_idx + 1}")
        if not state["toc_entries"]:
            toc_lb.insert(tk.END, "此PDF没有目录")

    def _reader_aux_toggle_toc(self, slot):
        """显示「目录」视图：如果面板已经在显示目录就收起，否则切到目录视图。"""
        state = self._reader_aux[slot]
        if state["toc_visible"] and state.get("toc_mode", "toc") == "toc":
            state["toc_frame"].grid_remove()
            state["toc_btn"].config(bg="#454b63")
            state["toc_visible"] = False
            self.after_idle(lambda s=slot: self._reader_aux_render(
                s, fit_width=True, target_page=self._reader_aux[s]["page"]))
            return
        state["toc_mode"] = "toc"
        state["bookmark_body"].pack_forget()
        state["toc_body"].pack(fill=tk.BOTH, expand=True)
        state["toc_frame"].grid()
        state["toc_btn"].config(bg=ENTRY)
        state["bookmark_btn"].config(bg="#454b63")
        state["toc_visible"] = True
        self.after_idle(lambda s=slot: self._reader_aux_render(
            s, fit_width=True, target_page=self._reader_aux[s]["page"]))

    def _reader_aux_toggle_bookmarks(self, slot):
        """显示「标记」视图：如果面板已经在显示标记就收起，否则切到标记视图。"""
        state = self._reader_aux[slot]
        if state["toc_visible"] and state.get("toc_mode", "toc") == "bookmark":
            state["toc_frame"].grid_remove()
            state["bookmark_btn"].config(bg="#454b63")
            state["toc_visible"] = False
            self.after_idle(lambda s=slot: self._reader_aux_render(
                s, fit_width=True, target_page=self._reader_aux[s]["page"]))
            return
        state["toc_mode"] = "bookmark"
        state["toc_body"].pack_forget()
        state["bookmark_body"].pack(fill=tk.BOTH, expand=True)
        state["toc_frame"].grid()
        state["bookmark_btn"].config(bg=ENTRY)
        state["toc_btn"].config(bg="#454b63")
        state["toc_visible"] = True
        self._reader_refresh_bookmark_list(slot)
        self.after_idle(lambda s=slot: self._reader_aux_render(
            s, fit_width=True, target_page=self._reader_aux[s]["page"]))

    def _reader_aux_toc_jump(self, slot):
        state = self._reader_aux[slot]
        selection = state["toc_lb"].curselection()
        if not selection or selection[0] >= len(state["toc_entries"]):
            return
        self._reader_set_active_slot(slot)
        state["page"] = state["toc_entries"][selection[0]]
        self._reader_aux_render(
            slot, fit_width=False, target_page=state["page"])

    def _reader_aux_render(self, slot, fit_width=None, target_page=None, quick=False):
        """连续页虚拟渲染：只保留当前视口附近页面。
        quick=True 时只渲染刚好填满视口的页数（Ctrl+滚轮缩放用），尽快出图。"""
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        if fit_width is not None:
            state["fit_width"] = bool(fit_width)
        doc = state["doc"]
        canvas = state["canvas"]
        canvas.update_idletasks()
        canvas_w = max(180, canvas.winfo_width())
        canvas_h = max(300, canvas.winfo_height())
        reference_page = doc[state["page"]]
        if state["fit_width"]:
            # 开启了"智能裁边"时，改用裁剪后的内容宽度来算，而不是整页
            # 宽度——理由跟PDF1的_reader_auto_fit_scale一样：否则裁掉
            # 空白边距之后，内容反而比画布窄了，没有真正撑满画布。
            ref_w = reference_page.rect.width
            if self._reader_crop_margins_mode:
                content_rect = self._reader_get_content_bbox(
                    doc, state["page"], state["path"])
                if content_rect and content_rect.width > 10:
                    ref_w = content_rect.width
            state["scale"] = (
                max(100, canvas_w - 8) / ref_w *
                (72.0 / self.PDF_DPI))
            state["scale"] = max(0.08, min(state["scale"], 6.0))
        zoom = self.PDF_DPI / 72.0 * state["scale"]
        total_pages = doc.page_count
        page_gap = 8
        est_page_h = int(doc[0].rect.height * zoom) + page_gap
        total_h_est = max(1, total_pages * est_page_h)

        if target_page is not None:
            target_page = max(0, min(int(target_page), total_pages - 1))
            state["page"] = target_page
            start_page = max(0, target_page - 2)
        else:
            yview = canvas.yview()
            view_top = (yview[0] if yview else 0.0) * total_h_est
            start_page = max(0, int(view_top / max(1, est_page_h)) - 2)
        if quick:
            # 缩放时只渲染「向上取整的视口页数 + 3 尾缓冲」，跟手不卡顿；
            # 用 floor 会在「一页比视口高」时漏渲染顶部一页，露出黑屏
            view_pages = max(1, (canvas_h + max(1, est_page_h) - 1) // max(1, est_page_h))
            visible_count = view_pages + 3
        else:
            visible_count = max(8, int(canvas_h / max(1, est_page_h)) + 6)
        end_page = min(total_pages, start_page + visible_count)

        state["zoom"] = zoom
        theme_page_bg = self._reader_themes[self._reader_theme.get()]["bg"]
        canvas.config(bg=theme_page_bg)
        canvas.delete("all")
        state["tk_imgs"] = []
        state["page_xoffsets"] = {}
        state["page_crop_origins"] = {}
        offsets = [i * est_page_h for i in range(total_pages)]
        y_cursor = start_page * est_page_h
        max_w = canvas_w
        for page_idx in range(start_page, end_page):
            page = doc[page_idx]
            try:
                img = self._render_page_crisp(page, zoom)
                img_w, img_h = img.size
                themed = self._apply_theme_to_image(img)
                if themed is not img:
                    img.close()
                    img = themed
                crop_origin = (0.0, 0.0)
                if self._reader_crop_margins_mode:
                    content_rect = self._reader_get_content_bbox(
                        doc, page_idx, state["path"])
                    if content_rect:
                        cbox = (int(content_rect.x0 * zoom), int(content_rect.y0 * zoom),
                                int(content_rect.x1 * zoom), int(content_rect.y1 * zoom))
                        cbox = (max(0, cbox[0]), max(0, cbox[1]),
                                min(img.width, cbox[2]), min(img.height, cbox[3]))
                        if cbox[2] > cbox[0] and cbox[3] > cbox[1]:
                            img = img.crop(cbox)
                            img_w, img_h = img.size
                            crop_origin = (content_rect.x0, content_rect.y0)
                state["page_crop_origins"][page_idx] = crop_origin
                tk_img = ImageTk.PhotoImage(img)
                img.close()
            except Exception as exc:
                state["title_var"].set(
                    f"PDF {slot} 渲染失败: {exc}")
                return
            offsets[page_idx] = y_cursor
            x_offset = max(0, (canvas_w - img_w) // 2)
            state["page_xoffsets"][page_idx] = x_offset
            canvas.create_image(
                x_offset, y_cursor, anchor="nw", image=tk_img,
                tags=f"aux_{slot}_page_{page_idx}")
            state["tk_imgs"].append(tk_img)
            if state["keywords"]:
                words = page.get_text("words")
                self._reader_draw_highlight_overlay(
                    canvas, words, x_offset, y_cursor, zoom,
                    state["keywords"], exact=state["exact"],
                    crop_origin=crop_origin, tag=f"hl_overlay_{slot}")
            y_cursor += img_h + page_gap
            max_w = max(max_w, img_w)
        for page_idx in range(end_page, total_pages):
            offsets[page_idx] = y_cursor + (page_idx - end_page) * est_page_h
        total_h = max(
            total_h_est,
            y_cursor + (total_pages - end_page) * est_page_h)
        state["page_yoffsets"] = offsets
        state["render_start"] = start_page
        state["render_end"] = end_page
        canvas.config(scrollregion=(0, 0, max_w, total_h))
        if target_page is not None:
            canvas.yview_moveto(offsets[target_page] / max(1, total_h))
        state["page_var"].set(
            f"{state['page'] + 1} / {state['doc'].page_count}")
        if getattr(self, "_reader_active_slot", 1) == slot:
            self._reader_update_active_page_info()
        state["title_var"].set(
            f"PDF {slot}: {os.path.basename(state['path'])}")

        # 翻页预取：后台线程预渲染视口前后各 2 页到 PIL 缓存（quick 缩放路径不预取）
        if not quick:
            self._reader_prefetch_around(
                doc, start_page, end_page, zoom, total_pages)

    def _reader_aux_go_page(self, slot, delta):
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        new_page = state["page"] + delta
        if 0 <= new_page < state["doc"].page_count:
            state["page"] = new_page
            self._reader_aux_render(
                slot, fit_width=False, target_page=new_page)
            self._reader_sync_scroll_others(slot)

    def _reader_aux_wheel(self, slot, event):
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        if self._reader_hscroll_on and (event.state & 0x1):
            self._reader_hscroll_shift(1 if event.delta < 0 else -1)
            return "break"
        if event.state & 0x4:
            if self._reader_magnifier_wheel_zoom(event):
                return
            state["fit_width"] = False
            # 记录光标下的 PDF 锚点，缩放重渲染后回卷保持光标内容不动
            anchor_pos = self._reader_magnifier_pdf_pos(slot, event.x, event.y)
            anchor_view = (event.x, event.y) if anchor_pos is not None else None
            state["scale"] *= 1.08 if event.delta > 0 else 0.92
            state["scale"] = max(0.08, min(state["scale"], 6.0))
            self._reader_aux_render(
                slot, fit_width=False, target_page=state["page"], quick=True)
            if anchor_pos is not None and anchor_view is not None:
                self._reader_scroll_to_anchor(slot, anchor_pos, anchor_view)
            self._reader_sync_scale_others(slot)
            return
        # 逐行滚动：每次滚轮滚动固定行数，不叠加惯性动画
        notches = int(event.delta / 120.0) if abs(event.delta) >= 120 else (1 if event.delta > 0 else -1)
        lines = -notches * self._reader_scroll_lines
        state["scroll_velocity"] = 0.0
        if state["scroll_anim_job"]:
            self.after_cancel(state["scroll_anim_job"])
            state["scroll_anim_job"] = None
        state["canvas"].yview_scroll(lines, "units")
        self._reader_aux_detect_page(slot, render_missing=True)
        self._reader_sync_scroll_others(slot)


    def _reader_aux_detect_page(self, slot, render_missing=True):
        state = self._reader_aux[slot]
        if not state["doc"] or not state["page_yoffsets"]:
            return
        canvas = state["canvas"]
        region = canvas.cget("scrollregion").split()
        total_h = float(region[3]) if len(region) == 4 else 1.0
        yview = canvas.yview()
        view_y = (yview[0] if yview else 0.0) * total_h
        view_y += canvas.winfo_height() * 0.3
        import bisect
        page_idx = bisect.bisect_right(
            state["page_yoffsets"], view_y) - 1
        page_idx = max(0, min(page_idx, state["doc"].page_count - 1))
        state["page"] = page_idx
        state["page_var"].set(
            f"{page_idx + 1} / {state['doc'].page_count}")
        if getattr(self, "_reader_active_slot", 1) == slot:
            self._reader_update_active_page_info()
        if render_missing and not (
                state["render_start"] <= page_idx < state["render_end"]):
            self._reader_aux_render(
                slot, fit_width=False, target_page=page_idx)

    def _reader_aux_on_resize(self, slot):
        state = self._reader_aux.get(slot)
        if not state or not state["doc"] or not state["fit_width"]:
            return
        if state["resize_job"]:
            self.after_cancel(state["resize_job"])
        state["resize_job"] = self.after(
            120, lambda s=slot: self._reader_aux_resize_render(s))

    def _reader_aux_resize_render(self, slot):
        state = self._reader_aux.get(slot)
        if not state:
            return
        state["resize_job"] = None
        if state["doc"] and state["fit_width"]:
            self._reader_aux_render(
                slot, fit_width=True, target_page=state["page"])

    def _reader_aux_get_words(self, slot):
        """一次解析并缓存辅助会话全部页面文字。"""
        state = self._reader_aux[slot]
        if state["word_cache"] is not None:
            return state["word_cache"]
        cache = []
        for page in state["doc"]:
            page_words = []
            for word in page.get_text("words"):
                clean = self._clean_word(word[4])
                if clean:
                    page_words.append((
                        word[0], word[1], word[2], word[3], clean))
            cache.append(page_words)
        state["word_cache"] = cache
        return cache

    def _reader_aux_search(self, slot, query=None, exact=False,
                           show_preview=False):
        """在指定PDF会话内独立搜索、高亮并按需显示命中页。"""
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        if query is None:
            query = state["search_var"].get().strip()
            exact = False
        else:
            query = query.strip()
            state["search_var"].set(query)
        if not query:
            return
        # 检索预览栏输入（非点击文本路径）→ 跨程序广播到勾选的目标实例
        if not show_preview:
            self._reader_broadcast_search(query)
        query_lower = query.lower()
        occurrences = []
        for page_idx, words in enumerate(self._reader_aux_get_words(slot)):
            for x0, y0, x1, y1, word in words:
                word_lower = word.lower()
                matched = (word_lower == query_lower if exact
                           else query_lower in word_lower)
                if matched:
                    occurrences.append(
                        (page_idx, fitz.Rect(x0, y0, x1, y1)))
        state["keywords"] = [query]
        state["exact"] = exact
        pages = sorted({page_idx for page_idx, _ in occurrences})
        counts = {}
        for page_idx, _ in occurrences:
            counts[page_idx] = counts.get(page_idx, 0) + 1
        # ── 合并「页面添加文本」命中（搜索栏也要能搜到添加的文本）──
        tm_matches = self._reader_text_marks_matches(state["path"], query)
        tm_total = 0
        if tm_matches:
            doc_page_count = state["doc"].page_count if state["doc"] else 0
            tm_valid = {}
            for pg, cnt in tm_matches.items():
                if pg < 0 or pg >= doc_page_count:
                    continue
                tm_valid[pg] = cnt
                counts[pg] = counts.get(pg, 0) + cnt
                tm_total += cnt
            if tm_valid:
                pages = sorted(set(pages) | set(tm_valid.keys()))
        total_occurrences = len(occurrences) + tm_total
        state["match_pages"] = pages
        state["match_counts"] = counts
        state["match_lb"].delete(0, tk.END)
        for page_idx in pages:
            state["match_lb"].insert(
                tk.END, f"第{page_idx + 1}页  {counts[page_idx]}处")
        state["match_count_var"].set(
            self._fmt_match_count(len(pages), total_occurrences))
        self._reader_refresh_aux_match_flow(slot)
        # 联动搜索开启 + 检索预览栏输入 → 把关键词（子串模糊匹配）同步到其它面板
        if not show_preview and self._reader_linked_search.get():
            self._reader_propagate_search(query, slot)
        if state["page"] in pages:
            match_index = pages.index(state["page"])
            state["match_lb"].selection_set(match_index)
            state["match_lb"].see(match_index)
        state["search_info"].set(
            f"{len(pages)}页/{len(occurrences)}处")
        # 跨文件命中提示（仅搜索栏输入触发，点击文本不触发）：后台用 AnyTXT 查其它文件
        if not show_preview:
            self._reader_anytxt_cross_hint(
                slot, query, state["path"], state["anytxt_hint"])
        if not occurrences:
            self._reader_aux_render(
                slot, fit_width=False, target_page=state["page"])
            return
        if not show_preview:
            state["page"] = pages[0]
            state["match_lb"].selection_clear(0, tk.END)
            state["match_lb"].selection_set(0)
        if show_preview:
            # 点击文本：页面不变，仅重画高亮覆盖层，避免 delete("all") 整页重渲染导致黑闪
            self._reader_aux_redraw_highlight_overlay(slot)
        else:
            self._reader_aux_render(
                slot, fit_width=False, target_page=state["page"])
        if show_preview:
            self._reader_aux_show_occurrences(
                slot, query, occurrences)

    def _reader_aux_match_jump(self, slot):
        state = self._reader_aux[slot]
        selection = state["match_lb"].curselection()
        if not selection or selection[0] >= len(state["match_pages"]):
            return
        self._reader_set_active_slot(slot)
        page_idx = state["match_pages"][selection[0]]
        state["page"] = page_idx
        self._reader_aux_render(
            slot, fit_width=False, target_page=page_idx)
        query = state["search_var"].get().strip()
        self._reader_aux_center_match_on_page(slot, query, page_idx)

    def _reader_aux_jump_to_page(self, slot, page_idx):
        """PDF2/PDF3专用：跳转到指定页（供简略弹窗点击使用）。"""
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        state["page"] = page_idx
        self._reader_aux_render(slot, fit_width=False, target_page=page_idx)

    def _reader_aux_center_match_on_page(self, slot, query, page_idx):
        """
        PDF2/PDF3专用：把匹配区域同时在垂直和水平方向都滚动到画布中心
        （和PDF1的 _reader_center_match_on_page 是同一套思路，因为PDF2/3
        的渲染状态保存在各自的 state 字典里，所以单独写一份，取数方式不同）。
        """
        state = self._reader_aux.get(slot)
        if not state or not state["doc"] or not query:
            return
        try:
            canvas = state["canvas"]
            doc = state["doc"]
            zoom = state["zoom"]
            page = doc[page_idx]
            words = page.get_text("words")
            query_lower = query.lower()
            crop_x, crop_y = state.get("page_crop_origins", {}).get(page_idx, (0.0, 0.0))
            match_rects = []
            for w in words:
                word_text = self._clean_word(w[4])
                if query_lower in word_text.lower():
                    x0 = (w[0] - crop_x) * zoom
                    y0 = (w[1] - crop_y) * zoom
                    x1 = (w[2] - crop_x) * zoom
                    y1 = (w[3] - crop_y) * zoom
                    match_rects.append([x0, y0, x1, y1])
            del page
            if not match_rects:
                return

            scrollregion = canvas.cget("scrollregion")
            parts = scrollregion.split() if scrollregion else []
            total_w = int(parts[-2]) if len(parts) >= 4 else 1
            total_h = int(parts[-1]) if len(parts) >= 4 else 1
            canvas_w = canvas.winfo_width()
            canvas_h = canvas.winfo_height()

            # ── 垂直居中 ──
            yoffsets = state["page_yoffsets"]
            page_y_offset = yoffsets[page_idx] if page_idx < len(yoffsets) else 0
            all_y0 = min(r[1] for r in match_rects) + page_y_offset
            all_y1 = max(r[3] for r in match_rects) + page_y_offset
            match_center_y = (all_y0 + all_y1) / 2
            if total_h > 0:
                target_top = max(0, match_center_y - canvas_h / 2)
                canvas.yview_moveto(min(1.0, max(0.0, target_top / total_h)))

            # ── 水平居中（放大后页面比画布宽时才需要）──
            if total_w > canvas_w:
                x_offset = state["page_xoffsets"].get(page_idx, 0)
                all_x0 = min(r[0] for r in match_rects) + x_offset
                all_x1 = max(r[2] for r in match_rects) + x_offset
                match_center_x = (all_x0 + all_x1) / 2
                target_left = max(0, match_center_x - canvas_w / 2)
                canvas.xview_moveto(min(1.0, max(0.0, target_left / total_w)))
            else:
                canvas.xview_moveto(0.0)
        except Exception:
            pass

    def _reader_aux_canvas_click(self, slot, event):
        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        canvas = state["canvas"]
        canvas_y = canvas.canvasy(event.y)
        import bisect
        page_idx = bisect.bisect_right(state["page_yoffsets"], canvas_y) - 1
        if not 0 <= page_idx < state["doc"].page_count:
            canvas.scan_mark(event.x, event.y)
            return
        page_x = state["page_xoffsets"].get(page_idx, 0)
        crop_x, crop_y = state.get("page_crop_origins", {}).get(page_idx, (0.0, 0.0))
        pdf_x = (canvas.canvasx(event.x) - page_x) / state["zoom"] + crop_x
        pdf_y = (canvas_y - state["page_yoffsets"][page_idx]) / state["zoom"] + crop_y
        state["page"] = page_idx
        state["page_var"].set(f"{page_idx + 1} / {state['doc'].page_count}")
        self._reader_update_active_page_info()
        clicked_word = None
        words = self._reader_aux_get_words(slot)[page_idx]
        for x0, y0, x1, y1, word in words:
            if x0 <= pdf_x <= x1 and y0 <= pdf_y <= y1:
                clicked_word = word
                break
        if clicked_word:
            self._reader_last_clicked_word = clicked_word
        else:
            canvas.scan_mark(event.x, event.y)
    def _reader_capture_merged_geometry(self, win):
        if not win or self._reader_merged_occurrence_window is not win:
            return
        try:
            win.update_idletasks()
            # 关键（混合 DPI 多屏，主屏 150%/副屏 100%）：winfo_x/winfo_y
            # 与 geometry("+X+Y") 存在固定偏差（实测副屏 X+511/Y+248），
            # 用 winfo 值逐次重建会累积漂移（每点一次文本窗口下移一截）；
            # 而 win.geometry() 字符串能稳定往返。故统一从 geometry 字符串
            # 解析位置与尺寸，作为唯一真相源。
            g = win.geometry()
            self._reader_merged_occurrence_geometry = g
            m = re.match(r"^(\d+)x(\d+)([+-]\d+)([+-]\d+)$", g)
            if m:
                self._reader_merged_occurrence_size = (
                    int(m.group(1)), int(m.group(2)))
                self._reader_merged_occurrence_position = (
                    int(m.group(3)), int(m.group(4)))
            # 位置/尺寸变化后立即持久化，重启后仍固定到该屏幕位置
            self._save_config()
        except Exception:
            pass

    def _reader_schedule_merged_geometry_capture(self, win):
        if self._reader_merged_geometry_job:
            try:
                self.after_cancel(self._reader_merged_geometry_job)
            except Exception:
                pass
        self._reader_merged_geometry_job = self.after(
            80, lambda: self._reader_finish_merged_geometry_capture(win))

    def _reader_finish_merged_geometry_capture(self, win):
        self._reader_merged_geometry_job = None
        self._reader_capture_merged_geometry(win)

    def _reader_clear_merged_slot(self, slot):
        self._reader_merged_occurrence_results.pop(slot, None)
        self._reader_merged_occurrence_generation += 1
        win = self._reader_merged_occurrence_window
        if win:
            try:
                self._reader_capture_merged_geometry(win)
                win.destroy()
            except Exception:
                pass
            self._reader_merged_occurrence_window = None
            # 清空列级增量状态，下次重新创建窗口
            self._reader_merged_columns = {}
            self._reader_merged_summary_label = None
            self._reader_merged_lock_btn = None
            self._reader_merged_columns_frame = None
            self._reader_merged_active_slots = []

    def _reader_show_merged_occurrences(
            self, slot, doc, path, word, occurrences):
        """在一个窗口中按PDF 1/2/3顺序展示各自最后一次点击结果。"""
        if not doc or not occurrences:
            return
        self._reader_merged_occurrence_results[slot] = {
            "doc": doc, "path": path, "word": word,
            "occurrences": list(occurrences),
        }
        self._reader_merged_occurrence_generation += 1
        generation = self._reader_merged_occurrence_generation

        # 复用或首次创建窗口
        win = self._reader_merged_occurrence_window
        if win is not None:
            try:
                if not win.winfo_exists():
                    win = None
            except Exception:
                win = None
        if win is None:
            win = self._reader_create_merged_window()
            if win is None:
                return

        # 更新 header（复用 summary label + lock 按钮，不重建）
        self._reader_update_merged_header(win)

        # 列集合变化才重建列骨架；否则复用，只刷新本次点击的那一栏
        active_slots = [
            result_slot for result_slot in (1, 2, 3, 4)
            if self._reader_merged_occurrence_results.get(result_slot)]
        rebuilt = list(active_slots) != self._reader_merged_active_slots
        if rebuilt:
            self._reader_rebuild_merged_columns(win, active_slots)
            self._reader_merged_active_slots = list(active_slots)

        win.update_idletasks()
        if rebuilt:
            # 新建/增删列时骨架全空，需把每一栏都渲染一遍
            for result_slot in active_slots:
                self._reader_refresh_merged_column(win, result_slot, generation)
        else:
            # 复用骨架：只刷新本次点击的那一栏，其它栏（PDF1/PDF3）不动
            self._reader_refresh_merged_column(win, slot, generation)

    def _reader_create_merged_window(self):
        """首次创建点击文本弹窗窗口（定位到其它屏并最大化）。"""
        win = tk.Toplevel(self)
        self._reader_merged_occurrence_window = win
        win.title("PDF 1 / PDF 2 / PDF 3 / PDF 4 · 点击文本命中页面")
        win.configure(bg=BG)
        # 弹窗固定最大化，铺满「主程序所在屏之外」的另一块屏工作区，避免
        # 盖住 GUI。主程序所在屏用主窗口中心点判定；单屏时回退主屏最大化。
        cx = self.winfo_rootx() + max(1, self.winfo_width()) // 2
        cy = self.winfo_rooty() + max(1, self.winfo_height()) // 2
        own_rect = self._monitor_rect_containing(cx, cy)
        target = self._other_monitor_rect(own_rect)
        if target is None:
            target = own_rect
        if target:
            left, top, right, bottom = target
            win.geometry(
                f"{max(760, right - left)}x{max(500, bottom - top)}"
                f"+{left}+{top}")
        else:
            win.geometry("1180x780")
        win.update_idletasks()
        try:
            win.state("zoomed")
        except Exception:
            pass
        win.minsize(760, 500)
        win.resizable(
            not self._reader_merged_occurrence_locked,
            not self._reader_merged_occurrence_locked)
        win.bind(
            "<Configure>",
            lambda event: self._reader_schedule_merged_geometry_capture(win)
            if event.widget is win else None)

        def close_window():
            self._reader_merged_occurrence_generation += 1
            try:
                self._reader_capture_merged_geometry(win)
            except Exception:
                pass
            if self._reader_merged_occurrence_window is win:
                self._reader_merged_occurrence_window = None
            # 窗口销毁后清空列级增量状态，下次重新创建
            self._reader_merged_columns = {}
            self._reader_merged_summary_label = None
            self._reader_merged_lock_btn = None
            self._reader_merged_columns_frame = None
            self._reader_merged_active_slots = []
            try:
                win.destroy()
            except Exception:
                pass

        win.protocol("WM_DELETE_WINDOW", close_window)
        return win

    def _reader_update_merged_header(self, win):
        """更新弹窗顶部 summary 文本；首次创建 header（含锁定按钮）。"""
        summaries = []
        for result_slot in (1, 2, 3, 4):
            result = self._reader_merged_occurrence_results.get(result_slot)
            if result:
                summaries.append(
                    f"PDF {result_slot}: {result['word']} "
                    f"({len(result['occurrences'])}处)")
        text = "  |  ".join(summaries)
        label = self._reader_merged_summary_label
        if label is not None:
            try:
                if label.winfo_exists():
                    label.config(text=text)
                    return
            except Exception:
                pass
        header = tk.Frame(win, bg=PANEL)
        header.pack(fill=tk.X)
        self._reader_merged_summary_label = tk.Label(
            header, text=text, bg=PANEL, fg=TEXT, anchor="w", padx=12,
            pady=9, font=("Microsoft YaHei UI", 9, "bold"))
        self._reader_merged_summary_label.pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        full_btn = RoundedButton(
            header, text="⤢ 整页", bg="#3a3f56", fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=8,
            command=self._reader_popup_reset_full_page)
        full_btn.pack(side=tk.RIGHT, padx=(0, 4), pady=7)
        lock_btn = RoundedButton(
            header, fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=8)

        def toggle_lock():
            self._reader_merged_occurrence_locked = (
                not self._reader_merged_occurrence_locked)
            if self._reader_merged_occurrence_locked:
                self._reader_merged_occurrence_size = (
                    win.winfo_width(), win.winfo_height())
                self._reader_capture_merged_geometry(win)
                win.resizable(False, False)
                lock_btn.config(text="🔒 已锁定", bg=ACCENT)
            else:
                win.resizable(True, True)
                lock_btn.config(text="🔓 可调整", bg=ENTRY)

        lock_btn.config(
            text=("🔒 已锁定" if self._reader_merged_occurrence_locked
                  else "🔓 可调整"),
            bg=(ACCENT if self._reader_merged_occurrence_locked else ENTRY),
            command=toggle_lock)
        lock_btn.pack(side=tk.RIGHT, padx=10, pady=7)
        self._reader_merged_lock_btn = lock_btn

    def _reader_rebuild_merged_columns(self, win, active_slots):
        """按 active_slots 重建所有列骨架（列集合变化时调用）。"""
        old_frame = self._reader_merged_columns_frame
        if old_frame is not None:
            try:
                if old_frame.winfo_exists():
                    old_frame.destroy()
            except Exception:
                pass
        self._reader_merged_columns = {}

        column_count = max(1, len(active_slots))
        win.minsize(max(520, 310 * column_count), 500)
        columns = tk.Frame(win, bg=CANVAS)
        columns.pack(fill=tk.BOTH, expand=True)
        self._reader_merged_columns_frame = columns
        for column_index in range(column_count):
            columns.columnconfigure(
                column_index, weight=1, uniform="pdf_result_column")
        columns.rowconfigure(0, weight=1)

        for column_index, result_slot in enumerate(active_slots):
            result = self._reader_merged_occurrence_results.get(result_slot)
            source_name = (os.path.basename(result["path"])
                           or f"PDF {result_slot}")
            column = tk.Frame(
                columns, bg=PANEL, highlightbackground=ACCENT,
                highlightthickness=2)
            column.grid(
                row=0, column=column_index, sticky="nsew",
                padx=(5 if column_index == 0 else 3,
                      5 if column_index == column_count - 1 else 3),
                pady=6)
            column.rowconfigure(1, weight=1)
            column.columnconfigure(0, weight=1)
            header_label = tk.Label(
                column,
                text=(f"PDF {result_slot} · {source_name} · "
                      f"点击：{result['word']} · "
                      f"{len(result['occurrences'])}处"),
                bg=PANEL, fg=TEXT, anchor="w", padx=10, pady=7,
                font=("Microsoft YaHei UI", 9, "bold"))
            header_label.grid(row=0, column=0, columnspan=2, sticky="ew")

            column_scroll = tk.Canvas(
                column, bg=CANVAS, highlightthickness=0,
                yscrollincrement=20)
            column_bar = ttk.Scrollbar(
                column, command=column_scroll.yview)
            column_scroll.config(yscrollcommand=column_bar.set)
            column_inner = tk.Frame(column_scroll, bg=CANVAS)
            column_inner_id = column_scroll.create_window(
                (0, 0), window=column_inner, anchor="nw")
            column_scroll.grid(row=1, column=0, sticky="nsew")
            column_bar.grid(row=1, column=1, sticky="ns")
            column_scroll.bind(
                "<Configure>",
                lambda event, c=column_scroll, iid=column_inner_id:
                c.itemconfigure(iid, width=max(180, event.width - 4)))

            def wheel_column(event, c=column_scroll):
                if event.state & 0x0004:  # Ctrl+滚轮
                    # 放大镜开着优先调放大镜倍率；否则调弹窗裁剪范围（缩到整页）
                    if self._reader_magnifier_enabled:
                        self._reader_magnifier_wheel_zoom(event)
                    else:
                        self._reader_popup_zoom(event.delta)
                else:
                    c.yview_scroll(int(-event.delta / 120) * 3, "units")

            column_scroll.bind("<MouseWheel>", wheel_column)
            column_inner.bind("<MouseWheel>", wheel_column)

            self._reader_merged_columns[result_slot] = {
                "column": column,
                "header_label": header_label,
                "scroll": column_scroll,
                "bar": column_bar,
                "inner": column_inner,
                "inner_id": column_inner_id,
                "wheel": wheel_column,
                "images": [],
            }

    def _reader_refresh_merged_column(self, win, slot, generation):
        """刷新某一栏缩略图（清空该栏后分帧重渲染），其它栏不动。"""
        col = self._reader_merged_columns.get(slot)
        if col is None:
            return
        result = self._reader_merged_occurrence_results.get(slot)
        if not result:
            return
        source_name = os.path.basename(result["path"]) or f"PDF {slot}"
        try:
            col["header_label"].config(
                text=(f"PDF {slot} · {source_name} · "
                      f"点击：{result['word']} · "
                      f"{len(result['occurrences'])}处"))
        except Exception:
            pass
        for child in col["inner"].winfo_children():
            child.destroy()
        col["images"].clear()

        grouped = OrderedDict()
        for page_idx, rect in result["occurrences"]:
            grouped.setdefault(page_idx, []).append(rect)
        tasks = list(grouped.items())
        column_scroll = col["scroll"]
        column_inner = col["inner"]
        wheel_column = col["wheel"]

        def render_column(index=0):
            if (generation != self._reader_merged_occurrence_generation or
                    self._reader_merged_occurrence_window is not win or
                    not win.winfo_exists()):
                return
            if index >= len(tasks):
                column_inner.update_idletasks()
                column_scroll.config(
                    scrollregion=column_scroll.bbox("all"))
                return
            page_idx, rects = tasks[index]
            try:
                page = result["doc"][page_idx]
                # 以命中文字为中心的局部区域：pad 控制上下文大小，Ctrl+滚轮缩放。
                # 放大 = 缩小 pad（更聚焦文字），缩小 = 放大 pad（更多上下文）。
                pad = self._reader_popup_pad
                x0 = min(r.x0 for r in rects)
                y0 = min(r.y0 for r in rects)
                x1 = max(r.x1 for r in rects)
                y1 = max(r.y1 for r in rects)
                # 正文区裁剪：手框正文区 > 空白边距检测 > 整页，切掉表头/外框。
                # 缩放"缩小到底"时 crop 会收敛到整个正文区，即"缩到整页"。
                body = self._reader_preview_region(result["doc"], page_idx, result["path"])
                crop = fitz.Rect(x0 - pad, y0 - pad, x1 + pad, y1 + pad) & body
                if crop.width < 2 or crop.height < 2:
                    crop = fitz.Rect(x0, y0, x1, y1) & body
                zoom = 150 / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=crop, alpha=False)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img = self._apply_theme_to_image(img)
                draw = ImageDraw.Draw(img, "RGBA")
                hr, hg, hb = self._reader_highlight_color
                for r in rects:
                    draw.rectangle(
                        [int((r.x0 - crop.x0) * zoom),
                         int((r.y0 - crop.y0) * zoom),
                         int((r.x1 - crop.x0) * zoom),
                         int((r.y1 - crop.y0) * zoom)],
                        fill=(hr, hg, hb, 120))
                del draw
                target_w = max(
                    200, min(900, column_scroll.winfo_width() - 18))
                if img.width > target_w:
                    ratio = target_w / img.width
                    resized = img.resize(
                        (target_w, max(1, int(img.height * ratio))),
                        Image.LANCZOS)
                    img.close()
                    img = resized
                tk_img = ImageTk.PhotoImage(img)
                img.close()
            except Exception:
                self.after(10, lambda: render_column(index + 1))
                return
            col["images"].append(tk_img)
            row = tk.Frame(
                column_inner, bg=CARD, highlightbackground=BORDER,
                highlightthickness=1, cursor="hand2")
            row.pack(fill=tk.X, padx=3, pady=3)
            tk.Label(
                row,
                text=f"PDF {slot} · 第{page_idx + 1}页 · {len(rects)}处",
                bg=CARD, fg=TEXT, anchor="w", padx=7, pady=3,
                font=("Microsoft YaHei UI", 8, "bold")).pack(fill=tk.X)
            image_label = tk.Label(
                row, image=tk_img, bg=CARD, cursor="hand2")
            image_label.pack(padx=4, pady=(0, 4))
            # 放大镜：悬停弹窗缩略图时也显示放大图（激活放大镜后生效）
            _disp_w = tk_img.width()
            _disp_h = tk_img.height()
            _base_zoom = (_disp_w / crop.width) if crop.width > 0 else 2.0
            _thumb_kw = [result["word"]]

            def _thumb_motion(e, d=result["doc"], pg=page_idx, cr=crop,
                              dw=_disp_w, dh=_disp_h, bz=_base_zoom, kw=_thumb_kw):
                if not self._reader_magnifier_enabled or dw <= 0 or dh <= 0:
                    return
                nx = max(0.0, min(1.0, e.x / dw))
                ny = max(0.0, min(1.0, e.y / dh))
                self._reader_magnifier_thumb_pending = (
                    d, pg, cr.x0 + nx * cr.width, cr.y0 + ny * cr.height,
                    e.x_root, e.y_root, bz, kw)
                if self._reader_magnifier_thumb_job is None:
                    self._reader_magnifier_thumb_job = self.after(
                        30, self._reader_magnifier_thumb_fire)

            def _thumb_leave(e):
                if self._reader_magnifier_thumb_job is not None:
                    self.after_cancel(self._reader_magnifier_thumb_job)
                    self._reader_magnifier_thumb_job = None
                self._reader_magnifier_thumb_pending = None
                self._reader_magnifier_hide()

            image_label.bind("<Motion>", _thumb_motion)
            image_label.bind("<Leave>", _thumb_leave)

            def jump(event=None, s=slot, pg=page_idx):
                self._reader_set_active_slot(s)
                if s == 1 and self._reader_pdf:
                    self._reader_page = pg
                    self._reader_scroll_to_page(pg)
                elif s in self._reader_aux and self._reader_aux[s]["doc"]:
                    self._reader_aux[s]["page"] = pg
                    self._reader_aux_render(
                        s, fit_width=False, target_page=pg)

            row.bind("<Button-1>", jump)
            image_label.bind("<Button-1>", jump)
            row.bind("<MouseWheel>", wheel_column)
            image_label.bind("<MouseWheel>", wheel_column)
            column_inner.update_idletasks()
            column_scroll.config(
                scrollregion=column_scroll.bbox("all"))
            self.after(12, lambda: render_column(index + 1))

        self.after_idle(render_column)

    def _reader_popup_zoom(self, delta):
        """点击弹窗内 Ctrl+滚轮：以命中文字为中心缩放局部区域（调 pad，防抖）。
        pad 上限放宽到 10000，配合渲染里的 `& 正文区`，缩小到底时 crop 收敛到
        整个正文区，实现"缩小到整页"。"""
        pad = self._reader_popup_pad * (0.8 if delta > 0 else 1.25)
        pad = max(8.0, min(10000.0, pad))
        if abs(pad - self._reader_popup_pad) < 0.5:
            return
        self._reader_popup_pad = pad
        if self._reader_popup_zoom_job:
            self.after_cancel(self._reader_popup_zoom_job)
        self._reader_popup_zoom_job = self.after(80, self._reader_popup_rezoom)

    def _reader_popup_reset_full_page(self):
        """把弹窗缩放复位到整页（pad 设大，crop 收敛到整个正文区）。"""
        self._reader_popup_pad = 10000.0
        self._reader_popup_rezoom()

    def _reader_popup_rezoom(self):
        """Ctrl+滚轮缩放防抖后，用新的 pad 重渲染所有命中列。"""
        self._reader_popup_zoom_job = None
        win = self._reader_merged_occurrence_window
        if not win:
            return
        self._reader_merged_occurrence_generation += 1
        generation = self._reader_merged_occurrence_generation
        for slot in (1, 2, 3, 4):
            if self._reader_merged_occurrence_results.get(slot):
                self._reader_refresh_merged_column(win, slot, generation)

    def _reader_aux_show_occurrences(self, slot, word, occurrences):
        """显示带PDF来源标识的辅助会话命中页面预览（支持弹窗模式控制）。"""
        mode = self._reader_click_popup_mode.get()
        if mode == "hidden":
            return
        state = self._reader_aux[slot]
        if mode == "minimal":
            self._reader_show_minimal_popup(word, occurrences, slot=slot)
            return
        # mode == "full"
        self._reader_show_merged_occurrences(
            slot, state["doc"], state["path"], word, occurrences)
        return
        state = self._reader_aux[slot]
        state["occurrence_generation"] += 1
        generation = state["occurrence_generation"]
        old = state["occurrence_window"]
        if old:
            try:
                if old.winfo_exists():
                    state["occurrence_geometry"] = old.geometry()
                    if state["occurrence_locked"]:
                        state["occurrence_size"] = (
                            old.winfo_width(), old.winfo_height())
                    old.destroy()
            except Exception:
                pass

        source_name = os.path.basename(state["path"])
        win = tk.Toplevel(self)
        state["occurrence_window"] = win
        win.title(
            f"PDF {slot} · {source_name} · {word} · {len(occurrences)}处")
        win.configure(bg=BG)
        default_w, default_h = 960, 700
        if state["occurrence_size"] is None:
            state["occurrence_size"] = (default_w, default_h)
        width, height = state["occurrence_size"]
        position = f"+{max(0, self.winfo_rootx() + 80 * slot)}+{max(0, self.winfo_rooty() + 55)}"
        saved = state["occurrence_geometry"]
        if saved:
            match = re.match(r"^\d+x\d+([+-]\d+)([+-]\d+)$", saved)
            if match:
                position = f"{match.group(1)}{match.group(2)}"
        win.geometry(f"{width}x{height}{position}")
        win.minsize(620, 420)
        win.resizable(
            not state["occurrence_locked"],
            not state["occurrence_locked"])

        header = tk.Frame(win, bg=PANEL)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text=(f"来源：PDF {slot} · {source_name}\n"
                  f"点击文本：{word} · {len(occurrences)}处"),
            bg=PANEL, fg=TEXT, justify=tk.LEFT, anchor="w",
            padx=12, pady=7,
            font=("Microsoft YaHei UI", 9, "bold")).pack(
                side=tk.LEFT, fill=tk.X, expand=True)
        lock_btn = RoundedButton(
            header, fg="white", relief=tk.FLAT,
            font=("Microsoft YaHei UI", 8, "bold"), padx=8)

        def toggle_lock():
            state["occurrence_locked"] = not state["occurrence_locked"]
            if state["occurrence_locked"]:
                state["occurrence_size"] = (
                    win.winfo_width(), win.winfo_height())
                win.resizable(False, False)
                lock_btn.config(text="🔒 已锁定", bg=ACCENT)
            else:
                win.resizable(True, True)
                lock_btn.config(text="🔓 可调整", bg=ENTRY)

        lock_btn.config(
            text="🔒 已锁定" if state["occurrence_locked"] else "🔓 可调整",
            bg=ACCENT if state["occurrence_locked"] else ENTRY,
            command=toggle_lock)
        lock_btn.pack(side=tk.RIGHT, padx=10, pady=7)

        scroll = tk.Canvas(win, bg=CANVAS, highlightthickness=0)
        scrollbar = ttk.Scrollbar(win, command=scroll.yview)
        scroll.config(yscrollcommand=scrollbar.set)
        inner = tk.Frame(scroll, bg=CANVAS)
        inner_id = scroll.create_window((0, 0), window=inner, anchor="nw")
        scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        scroll.bind(
            "<Configure>",
            lambda event: scroll.itemconfigure(
                inner_id, width=max(300, event.width - 4)))
        scroll.bind(
            "<MouseWheel>",
            lambda event: scroll.yview_scroll(
                int(-event.delta / 120) * 3, "units"))

        page_rects = OrderedDict()
        for page_idx, rect in occurrences:
            page_rects.setdefault(page_idx, []).append(rect)
        pages = list(page_rects)
        win._images = []

        def render_page(index=0):
            if (generation != state["occurrence_generation"] or
                    state["occurrence_window"] is not win or
                    not win.winfo_exists()):
                return
            if index >= len(pages):
                inner.update_idletasks()
                scroll.config(scrollregion=scroll.bbox("all"))
                return
            page_idx = pages[index]
            page = state["doc"][page_idx]
            zoom = 120 / 72.0
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            draw = ImageDraw.Draw(img, "RGBA")
            hr, hg, hb = self._reader_highlight_color
            for rect in page_rects[page_idx]:
                draw.rectangle(
                    [int(rect.x0 * zoom), int(rect.y0 * zoom),
                     int(rect.x1 * zoom), int(rect.y1 * zoom)],
                    fill=(hr, hg, hb, 120))
            del draw
            target_w = max(420, min(900, win.winfo_width() - 70))
            if img.width > target_w:
                ratio = target_w / img.width
                resized = img.resize(
                    (target_w, max(1, int(img.height * ratio))), Image.LANCZOS)
                img.close()
                img = resized
            tk_img = ImageTk.PhotoImage(img)
            img.close()
            win._images.append(tk_img)
            row = tk.Frame(
                inner, bg=PANEL, highlightbackground=BORDER,
                highlightthickness=1, cursor="hand2")
            row.pack(fill=tk.X, padx=5, pady=4)
            tk.Label(
                row, text=f"PDF {slot} · 第{page_idx + 1}页 · {len(page_rects[page_idx])}处",
                bg=PANEL, fg=TEXT, anchor="w", padx=8, pady=4,
                font=("Microsoft YaHei UI", 8, "bold")).pack(fill=tk.X)
            image_label = tk.Label(row, image=tk_img, bg=PANEL, cursor="hand2")
            image_label.pack(padx=5, pady=(0, 5))

            def jump(event=None, pg=page_idx):
                state["page"] = pg
                self._reader_aux_render(
                    slot, fit_width=False, target_page=pg)

            row.bind("<Button-1>", jump)
            image_label.bind("<Button-1>", jump)
            inner.update_idletasks()
            scroll.config(scrollregion=scroll.bbox("all"))
            self.after(15, lambda: render_page(index + 1))

        def close_window():
            state["occurrence_generation"] += 1
            try:
                state["occurrence_geometry"] = win.geometry()
                if state["occurrence_locked"]:
                    state["occurrence_size"] = (
                        win.winfo_width(), win.winfo_height())
            except Exception:
                pass
            if state["occurrence_window"] is win:
                state["occurrence_window"] = None
            try:
                win.destroy()
            except Exception:
                pass

        win.protocol("WM_DELETE_WINDOW", close_window)
        self.after_idle(render_page)

    # 旧的PDF 2方法保留为通用会话包装。





    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 左侧面板折叠/展开
    # ════════════════════════════════════════════════════════════════

    def _bind_reader_text_input(self, entry):
        """确保嵌入Edge后，Tk输入框点击时能可靠取回键盘焦点。"""
        def reclaim_focus(event):
            widget = event.widget
            widget.after_idle(widget.focus_force)

        entry.bind("<ButtonPress-1>", reclaim_focus, add="+")

    def _toggle_reader_left_panel(self):
        """折叠/展开左侧搜索面板"""
        if self._reader_left_visible:
            self._reader_left_frame.grid_remove()
            if getattr(self, "_reader_left_expand_btn", None):
                self._reader_left_expand_btn.destroy()
                self._reader_left_expand_btn = None
            if getattr(self, "_reader_left_nav_btn", None):
                self._reader_left_nav_btn.config(text="▶ 文件栏", bg="#454b63")
            self._reader_left_visible = False
        else:
            if getattr(self, '_reader_left_expand_btn', None):
                self._reader_left_expand_btn.destroy()
                self._reader_left_expand_btn = None
            self._reader_left_frame.grid()
            self._reader_left_toggle_btn.config(text="◀")
            if getattr(self, "_reader_left_nav_btn", None):
                self._reader_left_nav_btn.config(text="◀ 文件栏", bg=ACCENT)
            self._reader_left_visible = True

    def _toggle_reader_toc(self):
        """显示「目录」视图：如果面板已经在显示目录就收起，否则切到目录视图。"""
        if self._reader_toc_visible and self._reader_toc_mode == "toc":
            self._reader_toc_frame.grid_remove()
            self._reader_toc_btn.config(bg="#454b63")
            self._reader_toc_visible = False
            return
        self._reader_toc_mode = "toc"
        self._reader_bookmark_body.pack_forget()
        self._reader_toc_body.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self._reader_toc_frame.grid()
        self._reader_toc_btn.config(bg=ACCENT)
        self._reader_bookmark_btn.config(bg="#454b63")
        self._reader_toc_visible = True

    def _toggle_reader_bookmarks(self):
        """显示「标记」视图：如果面板已经在显示标记就收起，否则切到标记视图。"""
        if self._reader_toc_visible and self._reader_toc_mode == "bookmark":
            self._reader_toc_frame.grid_remove()
            self._reader_bookmark_btn.config(bg="#454b63")
            self._reader_toc_visible = False
            return
        self._reader_toc_mode = "bookmark"
        self._reader_toc_body.pack_forget()
        self._reader_bookmark_body.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self._reader_toc_frame.grid()
        self._reader_bookmark_btn.config(bg=ACCENT)
        self._reader_toc_btn.config(bg="#454b63")
        self._reader_toc_visible = True
        self._reader_refresh_bookmark_list()



    # ════════════════════════════════════════════════════════════════
    # PDF标记/笔记：给当前页打个标记，可选填一句简短笔记，按文件区分保存
    # ════════════════════════════════════════════════════════════════

    def _reader_show_bookmark_dialog(self, title, initial_label="", initial_note="", on_confirm=None):
        """添加/编辑标记的小对话框：标签输入框 + 笔记多行文本框。"""
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=BG)
        win.geometry("360x260")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        form = tk.Frame(win, bg=BG)
        form.pack(fill=tk.BOTH, expand=True, padx=18, pady=14)

        tk.Label(form, text="标签（可选，支持多级如 拓扑/LLC）:", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        label_var = tk.StringVar(value=initial_label)
        label_entry = tk.Entry(form, textvariable=label_var, bg=ENTRY, fg=TEXT,
                                insertbackground=TEXT, relief=tk.FLAT,
                                font=("Microsoft YaHei UI", 9))
        label_entry.pack(fill=tk.X, pady=(2, 10), ipady=4)

        tk.Label(form, text="笔记（可选）:", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        note_text = tk.Text(form, bg=ENTRY, fg=TEXT, insertbackground=TEXT,
                             relief=tk.FLAT, font=("Microsoft YaHei UI", 9),
                             height=6, wrap=tk.WORD)
        note_text.pack(fill=tk.BOTH, expand=True, pady=(2, 10))
        if initial_note:
            note_text.insert("1.0", initial_note)
        self._bind_reader_text_input(note_text)

        label_entry.focus_set()

        def confirm():
            label = label_var.get().strip()
            note = note_text.get("1.0", "end-1c").strip()
            if on_confirm:
                on_confirm(label, note)
            win.destroy()

        btn_row = tk.Frame(win, bg=BG)
        btn_row.pack(pady=(0, 12))
        RoundedButton(btn_row, text="保存", command=confirm, bg=ACCENT, fg="white",
                  relief=tk.FLAT, padx=18, pady=5, cursor="hand2",
                  font=("Microsoft YaHei UI", 9, "bold")).pack(side=tk.LEFT, padx=6)
        RoundedButton(btn_row, text="取消", command=win.destroy, bg="#454b63", fg="white",
                  relief=tk.FLAT, padx=18, pady=5, cursor="hand2",
                  font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=6)

    def _reader_path_for_slot(self, slot):
        """返回指定阅读区(1/2/3)当前打开的文件路径。"""
        if slot == 1:
            return self._reader_active_file_path
        state = self._reader_aux.get(slot)
        return state["path"] if state else ""

    def _reader_doc_for_slot(self, slot):
        """返回指定阅读区(1/2/3)当前的文档对象。"""
        if slot == 1:
            return self._reader_pdf
        state = self._reader_aux.get(slot)
        return state["doc"] if state else None

    def _reader_page_for_slot(self, slot):
        """返回指定阅读区(1/2/3)当前所在页码。"""
        if slot == 1:
            return self._reader_page
        state = self._reader_aux.get(slot)
        return state["page"] if state else 0

    def _reader_bookmark_lb_for_slot(self, slot):
        """返回指定阅读区对应的标记列表框控件。"""
        if slot == 1:
            return getattr(self, "_reader_bookmark_lb", None)
        state = self._reader_aux.get(slot)
        return state.get("bookmark_lb") if state else None

    def _reader_bookmark_ctx_menu_for_slot(self, slot):
        """返回指定阅读区对应的标记右键菜单。"""
        if slot == 1:
            return getattr(self, "_reader_bookmark_ctx_menu", None)
        state = self._reader_aux.get(slot)
        return state.get("bookmark_ctx_menu") if state else None

    def _reader_epub_text_widget_for_slot(self, slot):
        """如果指定阅读区当前正在显示EPUB内容，返回对应的Text控件；否则返回None。"""
        if slot == 1:
            if (self._reader_file_type == ".epub" and self._reader_text_frame and
                    self._reader_text_frame.winfo_ismapped()):
                return self._reader_text_preview
            return None
        state = self._reader_aux.get(slot)
        if state and state.get("epub_frame") and state["epub_frame"].winfo_ismapped():
            return state.get("epub_text")
        return None

    def _reader_add_bookmark(self, slot=1):
        """
        把当前页（PDF/PPT）或当前阅读位置（EPUB）加入标记列表，支持
        PDF1/PDF2/PDF3任意阅读区。EPUB没有"页码"这个概念，所以标记的是
        当前可视区域顶部所在的文本位置（行号），并顺手截取一小段附近的
        文字作为预览片段，方便在标记列表里一眼认出是哪一段。
        """
        full_path = self._reader_path_for_slot(slot)
        if not full_path:
            messagebox.showinfo("未打开文件", "请先打开一个文件再添加标记。")
            return

        epub_widget = self._reader_epub_text_widget_for_slot(slot)
        if epub_widget is not None:
            pos = epub_widget.index("@0,0")
            line_no = pos.split(".")[0]
            try:
                snippet = epub_widget.get(pos, f"{line_no}.end").strip()[:30]
            except Exception:
                snippet = ""

            def on_confirm(label, note):
                items = self._reader_annotations.setdefault(full_path, [])
                items.append({
                    "epub_pos": pos, "snippet": snippet,
                    "label": label, "note": note,
                })
                self._save_config()
                self._reader_refresh_bookmark_list(slot)

            self._reader_show_bookmark_dialog("标记当前阅读位置", on_confirm=on_confirm)
            return

        doc = self._reader_doc_for_slot(slot)
        if not doc:
            messagebox.showinfo("未打开PDF", "请先打开一个PDF（或可视化渲染成功的PPT）再添加标记。")
            return
        page = self._reader_page_for_slot(slot)

        def on_confirm(label, note):
            items = self._reader_annotations.setdefault(full_path, [])
            items.append({"page": page, "label": label, "note": note})
            items.sort(key=lambda it: it.get("page", 0))
            self._save_config()
            self._reader_refresh_bookmark_list(slot)

        self._reader_show_bookmark_dialog(f"标记第 {page + 1} 页", on_confirm=on_confirm)

    def _reader_refresh_bookmark_list(self, slot=1):
        """刷新标记列表显示，展示的是【该阅读区当前打开文件】对应的标记。"""
        lb = self._reader_bookmark_lb_for_slot(slot)
        if lb is None:
            return
        lb.delete(0, tk.END)
        full_path = self._reader_path_for_slot(slot)
        items = self._reader_annotations.get(full_path, []) if full_path else []
        if not items:
            lb.insert(tk.END, "  暂无标记")
            return
        for it in items:
            label = it.get("label") or "(无标签)"
            note_mark = " 📝" if it.get("note") else ""
            if "epub_pos" in it:
                snippet = it.get("snippet", "")
                lb.insert(tk.END, f"  📖 {label}{note_mark} {snippet}")
            else:
                lb.insert(tk.END, f"  第{it.get('page', 0) + 1}页 {label}{note_mark}")

    def _reader_current_bookmark_items(self, slot=1):
        """返回指定阅读区当前打开文件对应的标记列表（用于索引选中项）。"""
        full_path = self._reader_path_for_slot(slot)
        return self._reader_annotations.get(full_path, []) if full_path else []

    # ── 全局标记搜索（标签/笔记、页面文本、高亮原文 → 反向定位文件）──















    def _reader_annot_rect_press(self, event, path, page_idx, mark, rect_id, zoom):
        """左键按住标注框：记录拖拽起点，准备移动。"""
        cx = self._reader_canvas.canvasx(event.x)
        cy = self._reader_canvas.canvasy(event.y)
        self._reader_mark_drag = {
            "path": path, "page_idx": page_idx, "mark": mark,
            "rect_id": rect_id, "zoom": zoom,
            "start_cx": cx, "start_cy": cy,
            "last_cx": cx, "last_cy": cy,
            "orig_rect": list(mark.get("rect", [0, 0, 0, 0])),
        }
        return "break"

    def _reader_annot_rect_motion(self, event, rect_id):
        """拖拽移动标注框（画布像素增量）。"""
        drag = getattr(self, "_reader_mark_drag", None)
        if not drag or drag["rect_id"] != rect_id:
            return
        cx = self._reader_canvas.canvasx(event.x)
        cy = self._reader_canvas.canvasy(event.y)
        self._reader_canvas.move(rect_id, cx - drag["last_cx"], cy - drag["last_cy"])
        drag["last_cx"] = cx
        drag["last_cy"] = cy
        return "break"

    def _reader_annot_rect_release(self, event):
        """松开左键：把拖拽位移换算成 PDF 点坐标，写回并保存。"""
        drag = getattr(self, "_reader_mark_drag", None)
        if not drag:
            return
        self._reader_mark_drag = None
        cx = self._reader_canvas.canvasx(event.x)
        cy = self._reader_canvas.canvasy(event.y)
        zoom = drag["zoom"]
        if zoom:
            dx_pdf = (cx - drag["start_cx"]) / zoom
            dy_pdf = (cy - drag["start_cy"]) / zoom
            x0, y0, x1, y1 = drag["orig_rect"]
            drag["mark"]["rect"] = [x0 + dx_pdf, y0 + dy_pdf, x1 + dx_pdf, y1 + dy_pdf]
            self._save_config()
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords)
        return "break"

    def _reader_bookmark_jump(self, event=None, slot=1):
        """点击标记列表中的某一项 → 跳转到对应页面（PDF/PPT）或文本位置（EPUB）。"""
        lb = self._reader_bookmark_lb_for_slot(slot)
        if lb is None:
            return
        sel = lb.curselection()
        if not sel:
            return
        items = self._reader_current_bookmark_items(slot)
        if sel[0] >= len(items):
            return
        item = items[sel[0]]

        if "epub_pos" in item:
            epub_widget = self._reader_epub_text_widget_for_slot(slot)
            if epub_widget is not None:
                try:
                    epub_widget.see(item["epub_pos"])
                except Exception:
                    pass
            return

        doc = self._reader_doc_for_slot(slot)
        if not doc:
            return
        page = max(0, min(item.get("page", 0), doc.page_count - 1))
        if slot == 1:
            self._reader_page = page
            self._reader_scroll_to_page(page)
            self._reader_update_active_page_info()
            self._reader_highlight_thumb()
        else:
            state = self._reader_aux[slot]
            state["page"] = page
            self._reader_aux_render(slot, fit_width=False, target_page=page)

    def _reader_show_bookmark_context_menu(self, event, slot=1):
        """标记列表的右键菜单：编辑笔记 / 删除标记。"""
        lb = self._reader_bookmark_lb_for_slot(slot)
        menu = self._reader_bookmark_ctx_menu_for_slot(slot)
        if lb is None or menu is None:
            return
        idx = lb.nearest(event.y)
        items = self._reader_current_bookmark_items(slot)
        if idx < 0 or idx >= len(items):
            return
        lb.selection_clear(0, tk.END)
        lb.selection_set(idx)
        lb.activate(idx)
        menu.post(event.x_root, event.y_root)

    def _reader_edit_bookmark_note(self, slot=1):
        """编辑选中标记的标签/笔记。"""
        lb = self._reader_bookmark_lb_for_slot(slot)
        if lb is None:
            return
        sel = lb.curselection()
        items = self._reader_current_bookmark_items(slot)
        if not sel or sel[0] >= len(items):
            return
        idx = sel[0]
        item = items[idx]

        def on_confirm(label, note):
            item["label"] = label
            item["note"] = note
            self._save_config()
            self._reader_refresh_bookmark_list(slot)

        self._reader_show_bookmark_dialog(
            f"编辑第 {item['page'] + 1} 页的标记",
            initial_label=item.get("label", ""), initial_note=item.get("note", ""),
            on_confirm=on_confirm)

    def _reader_delete_bookmark(self, slot=1):
        """删除选中的标记。"""
        lb = self._reader_bookmark_lb_for_slot(slot)
        if lb is None:
            return
        sel = lb.curselection()
        full_path = self._reader_path_for_slot(slot)
        items = self._reader_current_bookmark_items(slot)
        if not sel or sel[0] >= len(items) or not full_path:
            return
        idx = sel[0]
        del self._reader_annotations[full_path][idx]
        if not self._reader_annotations[full_path]:
            del self._reader_annotations[full_path]
        self._save_config()
        self._reader_refresh_bookmark_list(slot)

    def _reader_export_excerpt(self, slot=1):
        """
        把当前阅读区打开文件的「标记」导出成一个 Markdown 摘录 + 页面缩略图。
        每条标记渲染一页缩略图（智能裁边后只保留内容区），图文并茂，方便归档
        进 Obsidian。导出目录每次弹框手动选择；图片统一写到目标目录的 assets/
        子目录。标准 Markdown 语法（![](assets/xxx.png)），Obsidian 可直接识别。
        """
        full_path = self._reader_path_for_slot(slot)
        doc = self._reader_doc_for_slot(slot)
        if not full_path:
            messagebox.showinfo("未打开文件", "请先打开一个文件再导出摘录。")
            return
        items = self._reader_annotations.get(full_path, []) if full_path else []
        if not items:
            messagebox.showinfo("没有标记", "当前文件还没有添加任何标记，先「📌 标记当前页」再导出。")
            return
        if not doc:
            messagebox.showinfo("暂不支持", "当前文件不是 PDF/PPT，暂不支持摘录导出。")
            return

        target_dir = filedialog.askdirectory(title="选择摘录导出目录")
        if not target_dir:
            return

        self.config(cursor="watch")
        self.update_idletasks()
        try:
            base_name = os.path.splitext(os.path.basename(full_path))[0]
            safe_base = re.sub(r'[\\/:*?"<>|]+', "_", base_name)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            md_name = f"摘录-{safe_base}-{stamp}.md"
            assets_dir = os.path.join(target_dir, "assets")
            os.makedirs(assets_dir, exist_ok=True)

            # 缩略图用约 200 DPI，配合智能裁边只保留内容区，兼顾清晰度和文件大小
            zoom = 200.0 / 72.0
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            lines = [
                f"# 阅读摘录：{base_name}",
                "",
                f"- 文件：`{full_path}`",
                f"- 总页数：{doc.page_count}",
                f"- 标记数：{len(items)}",
                f"- 导出时间：{now_str}",
                "",
                "---",
                "",
            ]
            for idx, item in enumerate(items):
                page_idx = item.get("page", 0)
                label = item.get("label") or "(无标签)"
                note = (item.get("note") or "").strip()
                lines.append(f"## {idx + 1}. 第 {page_idx + 1} 页 — {label}")
                lines.append("")
                img_rel = None
                try:
                    page = doc[page_idx]
                    img = self._render_page_crisp(page, zoom)
                    crop = self._reader_get_content_bbox(doc, page_idx, full_path)
                    if crop is not None:
                        cbox = (
                            max(0, int(crop.x0 * zoom)), max(0, int(crop.y0 * zoom)),
                            min(img.width, int(crop.x1 * zoom)),
                            min(img.height, int(crop.y1 * zoom)))
                        if cbox[2] > cbox[0] and cbox[3] > cbox[1]:
                            img = img.crop(cbox)
                    img_name = f"{safe_base}_p{page_idx + 1}_{stamp}.png"
                    img.save(os.path.join(assets_dir, img_name))
                    img.close()
                    img_rel = f"assets/{img_name}"
                except Exception:
                    img_rel = None
                if img_rel:
                    lines.append(f"![第{page_idx + 1}页]({img_rel})")
                    lines.append("")
                if note:
                    lines.append(f"> {note}")
                    lines.append("")
                lines.append("")

            md_path = os.path.join(target_dir, md_name)
            with open(md_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as exc:
            self.config(cursor="")
            messagebox.showerror("导出失败", f"导出摘录时出错：{exc}")
            return
        self.config(cursor="")

        messagebox.showinfo("导出完成", f"已导出 {len(items)} 条标记：\n{md_path}")

    def _reader_load_toc(self):
        """读取当前PDF内置书签，填充左侧目录为可折叠层级树。"""
        self._reader_toc_page = {}
        tree = self._reader_toc_tree
        for child in tree.get_children():
            tree.delete(child)
        if not self._reader_pdf:
            tree.insert("", "end", text="  当前文件没有PDF目录")
            return
        try:
            toc = self._reader_pdf.get_toc(simple=True) or []
        except Exception:
            toc = []
        if not toc:
            tree.insert("", "end", text="  此PDF没有内置书签目录")
            return
        stack = []  # [(level, item_id)]：维护各级父节点，用于构建层级树
        for entry in toc:
            if len(entry) < 3:
                continue
            level, title, page_number = entry[:3]
            try:
                page_index = int(page_number) - 1
            except (TypeError, ValueError):
                continue
            if not 0 <= page_index < self._reader_pdf.page_count:
                continue
            try:
                level = max(1, int(level))
            except (TypeError, ValueError):
                level = 1
            title = " ".join(str(title).split()) or f"第 {page_index + 1} 页"
            while stack and stack[-1][0] >= level:
                stack.pop()
            parent = stack[-1][1] if stack else ""
            item = tree.insert(
                parent, "end", text=f"{title}  ·  {page_index + 1}")
            self._reader_toc_page[item] = page_index
            stack.append((level, item))
        # 默认展开前两层，避免长目录一开始铺满整屏
        for child in tree.get_children():
            tree.item(child, open=True)
            for grand in tree.get_children(child):
                tree.item(grand, open=True)

    def _reader_toc_jump(self, event=None):
        """点击目录树节点后跳转到对应PDF页面。"""
        sel = self._reader_toc_tree.selection()
        if not sel or not self._reader_pdf:
            return
        page_index = self._reader_toc_page.get(sel[0])
        if page_index is None:
            return
        self._reader_page = page_index
        self._reader_update_active_page_info()
        self._reader_scroll_to_page(page_index)
        self._reader_highlight_thumb()

    def _toggle_match_col(self):
        """折叠/展开三个PDF预览区(PDF1/2/3)下方的匹配页码块（同一按钮统一控制）。"""
        if self._reader_match_visible:
            self._reader_match_frame.grid_remove()
            for slot in (2, 3, 4):
                st = self._reader_aux.get(slot)
                if st:
                    st["match_frame"].grid_remove()
            self._reader_match_nav_btn.config(bg="#454b63")
            self._reader_match_visible = False
        else:
            self._reader_match_frame.grid()
            for slot in (2, 3, 4):
                st = self._reader_aux.get(slot)
                if st:
                    st["match_frame"].grid()
            self._reader_match_nav_btn.config(bg=ACCENT)
            self._reader_match_visible = True

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - AI阅读助手（本地文本分析）
    # ════════════════════════════════════════════════════════════════


    def _reader_update_ai_source(self):
        """刷新AI面板显示的当前分析来源。"""
        if not hasattr(self, "_reader_ai_source_var"):
            return
        slot, doc, page_idx, path = self._reader_get_active_pdf_context()
        if doc:
            name = self._reader_doc_display_name(path, doc, "未命名PDF")
            self._reader_ai_source_var.set(
                f"当前来源：PDF {slot} · {name} · 第 {page_idx + 1} / {doc.page_count} 页")
        else:
            self._reader_ai_source_var.set(
                f"当前来源：PDF {slot} · 未打开PDF")

























    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - Edge浏览器集成
    # ════════════════════════════════════════════════════════════════


    # ════════════════════════════════════════════════════════════════
    # 全局快捷键 - 任何时候按下都能隐藏/恢复主窗口
    # ════════════════════════════════════════════════════════════════


    # ════════════════════════════════════════════════════════════════
    # 应用内快捷键体系（区别于"全局热键"——这些只在本程序窗口内生效）
    # ════════════════════════════════════════════════════════════════

    def _shortcut_is_text_focused(self):
        """
        检查当前键盘焦点是否落在文本输入控件（Entry/Text）上。
        大多数快捷键在这种情况下应该自动避让，避免用户正常打字、或者在
        笔记框里用方向键移动光标时，被意外当成"翻页"之类的指令。
        """
        try:
            w = self.focus_get()
        except Exception:
            return False
        return isinstance(w, (tk.Entry, tk.Text))

    def _shortcut_reader_tab_active(self):
        """检查当前选中的页签是不是"📂 文件阅读"。"""
        try:
            return str(self._notebook.select()) == str(self._tab_reader)
        except Exception:
            return False

    def _setup_keyboard_shortcuts(self):
        """应用内快捷键：翻页 / 缩放 / 聚焦搜索 / 添加标记。"""
        self._shortcut_bound_seqs = set()
        self._shortcut_apply_bindings()

    def _shortcut_registry(self):
        """快捷键注册表：返回 [(id, 说明, 默认绑定序列, 处理函数), ...]。"""
        reg = [
            ("focus_search", "聚焦搜索框", ["<Control-f>", "<Control-F>"], self._shortcut_focus_search),
            ("add_bookmark", "添加标记", ["<Control-b>", "<Control-B>"], self._shortcut_add_bookmark),
            ("toggle_bookmarks", "显示/隐藏标记面板", ["<Control-Shift-B>", "<Control-Shift-b>"], self._shortcut_toggle_bookmarks),
            ("page_prev", "上一页", ["<Prior>", "<Left>"], lambda e: self._shortcut_reader_page(-1)),
            ("page_next", "下一页", ["<Next>", "<Right>"], lambda e: self._shortcut_reader_page(1)),
            ("first_page", "跳到第一页", ["<Home>"], self._shortcut_reader_first_page),
            ("last_page", "跳到最后一页", ["<End>"], self._shortcut_reader_last_page),
            ("zoom_in", "放大", ["<Control-equal>", "<Control-plus>"], lambda e: self._shortcut_reader_zoom(1.15)),
            ("zoom_out", "缩小", ["<Control-minus>"], lambda e: self._shortcut_reader_zoom(1 / 1.15)),
            ("zoom_reset", "恢复适应宽度", ["<Control-Key-0>"], self._shortcut_reader_zoom_reset),
            ("toggle_crop", "开启/关闭智能裁边", ["<Control-e>", "<Control-E>"], self._shortcut_toggle_crop),
            ("toggle_fit_width", "开启/关闭适应宽度", ["<Control-w>", "<Control-W>"], self._shortcut_toggle_fit_width),
        ]
        for s in (1, 2):
            reg.append((f"slot_{s}", f"切换阅读区·PDF{s}",
                        [f"<Alt-Key-{s}>"],
                        lambda e, s=s: self._shortcut_set_active_slot(s)))
        return reg

    def _shortcut_apply_bindings(self):
        """按 _shortcut_bindings 重新注册所有快捷键（先解绑旧序列，避免改键后残留）。"""
        if not hasattr(self, "_shortcut_bound_seqs"):
            self._shortcut_bound_seqs = set()
        for seq in self._shortcut_bound_seqs:
            try:
                self.unbind_all(seq)
            except Exception:
                pass
        self._shortcut_bound_seqs.clear()
        for sid, _label, defaults, handler in self._shortcut_registry():
            for b in self._shortcut_bindings.get(sid, list(defaults)):
                if not b:
                    continue
                try:
                    self.bind_all(b, handler)
                    self._shortcut_bound_seqs.add(b)
                except tk.TclError:
                    continue






    def _shortcut_focus_search(self, event=None):
        """Ctrl+F：切到文件阅读页签并聚焦预览搜索框（随时可用，不避让文本框）。"""
        if not hasattr(self, "_tab_reader"):
            return
        self._notebook.select(self._tab_reader)
        if hasattr(self, "_reader_search_entry"):
            self._reader_search_entry.focus_set()
            self._reader_search_entry.select_range(0, tk.END)
        return "break"

    def _shortcut_add_bookmark(self, event=None):
        """Ctrl+B：给当前页添加标记。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        self._reader_add_bookmark()
        return "break"

    def _shortcut_toggle_bookmarks(self, event=None):
        """Ctrl+Shift+B：显示/隐藏标记面板。"""
        if not self._shortcut_reader_tab_active():
            return
        self._toggle_reader_bookmarks()
        return "break"


    def _shortcut_reader_page(self, delta):
        """PageUp/PageDown/左右方向键：翻页。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        self._reader_go_page(delta)
        return "break"

    def _shortcut_reader_first_page(self, event=None):
        """Home：跳到第一页。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        if not self._reader_pdf:
            return
        self._reader_page = 0
        self._reader_scroll_to_page(0)
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()
        return "break"

    def _shortcut_reader_last_page(self, event=None):
        """End：跳到最后一页。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        if not self._reader_pdf:
            return
        last = self._reader_pdf.page_count - 1
        self._reader_page = last
        self._reader_scroll_to_page(last)
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()
        return "break"

    def _shortcut_reader_zoom(self, factor):
        """Ctrl+ / Ctrl-：放大/缩小当前PDF。"""
        if not self._shortcut_reader_tab_active() or not self._reader_pdf:
            return
        self._reader_fit_width = False
        new_scale = max(0.2, min(self._reader_scale * factor, 6.0))
        self._reader_scale = new_scale
        self._reader_render(force_fit_width=False, keywords=self._reader_highlight_keywords)
        return "break"

    def _shortcut_reader_zoom_reset(self, event=None):
        """Ctrl+0：恢复"适应宽度"。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        if not self._reader_pdf:
            return
        self._reader_fit_width = True
        self._reader_fit_btn.config(text="🔲 适应宽度", bg=ACCENT)
        self._reader_render(force_fit_width=True, keywords=self._reader_highlight_keywords)
        self._reader_scroll_to_page(self._reader_page)
        return "break"



    def _shortcut_toggle_crop(self, event=None):
        """Ctrl+E：开启/关闭智能裁边。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        self._toggle_reader_crop_margins()
        return "break"

    def _shortcut_toggle_fit_width(self, event=None):
        """Ctrl+W：开启/关闭适应宽度。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        self._toggle_reader_fit_width()
        return "break"

    def _shortcut_set_active_slot(self, slot):
        """Alt+1~4：切换当前 PDF 阅读区（下次打开文件 / 放大镜等操作的目标栏）。"""
        if self._shortcut_is_text_focused() or not self._shortcut_reader_tab_active():
            return
        self._reader_set_active_slot(slot)
        return "break"

























    # ── 键盘转发：tkinter <Key> 绑定，安全转发键盘输入到 Edge ──
    # 注意：绝不能用 Win32 WH_GETMESSAGE 钩子，因为钩子回调在 Windows
    # 消息处理上下文中运行，此时调用 tkinter 方法会造成 Tcl/Tk 事件循环
    # 重入，导致栈溢出崩溃。用 bind_all 绑定 <Key> 事件则完全在 tkinter
    # 自己的事件循环中运行，不存在重入风险。















    def _record_search_history(self, key, keyword):
        """记录一次搜索，关键词次数 +1（用于左侧搜索栏历史）。"""
        keyword = (keyword or "").strip()
        if not keyword:
            return
        hist = self._search_histories.setdefault(key, {})
        hist[keyword] = hist.get(keyword, 0) + 1
        # 即时持久化：搜索历史需跨重启保留；搜索是低频动作，全量保存开销可忽略
        try:
            self._save_config()
        except Exception:
            pass

    def _bind_search_history(self, entry, key, apply_cb):
        """给搜索框绑定历史下拉：点击(ButtonRelease)或按↓弹出。

        用 ButtonRelease 而非 FocusIn：选中记录后焦点回弹到 Entry 且保持聚焦，
        FocusIn 不会再触发（导致「再点不弹」）；ButtonRelease 每次点击都触发，
        且晚于 bind_all 的 <Button-1> 全局关闭，不会刚弹出就被关掉。
        """
        entry.bind(
            "<ButtonRelease-1>",
            lambda e: self._show_search_history(entry, key, apply_cb))
        entry.bind(
            "<Down>",
            lambda e: self._show_search_history(entry, key, apply_cb))

    def _show_search_history(self, entry, key, apply_cb):
        """在搜索框下方弹出历史下拉，按搜索次数降序。"""
        hist = self._search_histories.get(key, {})
        if not hist:
            return
        self._hide_search_history()
        items = sorted(hist.items(), key=lambda kv: (-kv[1], kv[0]))
        items = items[:self._search_history_max]
        win = tk.Toplevel(self)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        w = max(entry.winfo_width(), 200)
        h = min(240, 26 * len(items) + 6)
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        win.geometry(f"{w}x{h}+{x}+{y}")
        lb = tk.Listbox(win, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                        selectforeground="white", relief=tk.SOLID, bd=1,
                        font=("Microsoft YaHei UI", 9), highlightthickness=0,
                        activestyle="none")
        lb.pack(fill=tk.BOTH, expand=True)
        self._search_history_items = []
        for kw, cnt in items:
            lb.insert(tk.END, f"{kw}   ({cnt}次)")
            self._search_history_items.append(kw)
        self._search_history_win = win
        self._search_history_lb = lb
        self._search_history_apply = apply_cb

        def _pick(event):
            idx = lb.nearest(event.y)
            if 0 <= idx < len(self._search_history_items):
                kw = self._search_history_items[idx]
                self._hide_search_history()
                apply_cb(kw)

        lb.bind("<Button-1>", _pick)
        lb.bind("<Escape>", lambda e: self._hide_search_history())

    def _hide_search_history(self):
        """关闭搜索历史下拉。"""
        win = getattr(self, "_search_history_win", None)
        if win is not None and win.winfo_exists():
            try:
                win.destroy()
            except Exception:
                pass
        self._search_history_win = None
        self._search_history_lb = None



    # ════════════════════════════════════════════════════════════════
    # 常用文件 - 关闭程序后仍会保留，默认收起不占地方，点头部展开/收起
    # ════════════════════════════════════════════════════════════════

    def _bind_tooltip(self, widget, text_getter, delay=450):
        """给 widget 绑定悬停提示（无边框 Toplevel + Label）。

        text_getter: 无参函数，返回要显示的字符串；返回空串则不显示。
        用途：常用文件下拉框里长文件名被截断，悬停显示完整文件名+路径。
        """
        tip_win = [None]
        tip_after = [None]

        def _hide():
            if tip_after[0] is not None:
                try:
                    widget.after_cancel(tip_after[0])
                except Exception:
                    pass
                tip_after[0] = None
            if tip_win[0] is not None:
                try:
                    tip_win[0].destroy()
                except Exception:
                    pass
                tip_win[0] = None

        def _show():
            try:
                text = text_getter()
                if not text or not widget.winfo_exists():
                    _hide()
                    return
                _hide()
                win = tk.Toplevel(widget)
                tip_win[0] = win
                win.wm_overrideredirect(True)
                win.attributes("-topmost", True)
                tk.Label(
                    win, text=text, justify=tk.LEFT, anchor="w",
                    bg="#2a2d3a", fg="#ffffff", relief=tk.SOLID,
                    borderwidth=1, padx=8, pady=6,
                    font=("Microsoft YaHei UI", 9),
                    wraplength=620).pack()
                win.update_idletasks()
                x = widget.winfo_rootx() + 12
                y = widget.winfo_rooty() + widget.winfo_height() + 8
                w = win.winfo_reqwidth()
                h = win.winfo_reqheight()
                sw = win.winfo_screenwidth()
                sh = win.winfo_screenheight()
                if x + w > sw - 8:
                    x = sw - w - 8
                if y + h > sh - 8:
                    y = widget.winfo_rooty() - h - 8
                win.wm_geometry(f"+{max(0, x)}+{max(0, y)}")
            except Exception:
                pass

        def _on_enter(event):
            _hide()
            tip_after[0] = widget.after(delay, _show)

        def _on_leave(event):
            _hide()

        def _on_motion(event):
            # 已显示时刷新内容（跟随列表不同行），未显示时不打断 Enter 的延迟
            if tip_win[0] is not None:
                _show()

        widget.bind("<Enter>", _on_enter, add="+")
        widget.bind("<Leave>", _on_leave, add="+")
        widget.bind("<Motion>", _on_motion, add="+")
        return _hide















    def _show_theme_menu(self):
        """显示颜色主题选择菜单"""
        menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground="white",
                       font=("Microsoft YaHei UI", 9))
        for key, theme in self._reader_themes.items():
            menu.add_radiobutton(label=f"{theme['name']}", variable=self._reader_theme,
                                value=key, command=self._on_theme_change,
                                background=PANEL, selectcolor=CARD)

        # 在主题按钮下方显示菜单
        x = self._theme_btn.winfo_rootx()
        y = self._theme_btn.winfo_rooty() + self._theme_btn.winfo_height()
        menu.post(x, y)

    def _on_theme_change(self):
        """主题变更时重新渲染PDF"""
        theme = self._reader_theme.get()
        theme_name = self._reader_themes[theme]["name"]
        self._theme_btn.config(text=f"🎨 {theme_name}")
        self._save_config()

        # 更新画布背景色
        canvas_bg = self._reader_themes[theme].get("canvas_bg", CANVAS)
        self._reader_canvas.config(bg=canvas_bg)

        # 清除缓存并重新渲染
        self._reader_page_cache.clear()
        self._reader_popup_page_cache.clear()
        if self._reader_pdf:
            self._reader_render(keywords=self._reader_highlight_keywords)
        for slot, state in self._reader_aux.items():
            if state["doc"]:
                self._reader_aux_render(
                    slot, fit_width=False, target_page=state["page"])

    def _reader_get_content_bbox(self, doc, page_idx, file_path):
        """
        检测指定页面"实际内容"所在的边界框（PDF点坐标），用于智能裁边——
        把页面渲染成低分辨率灰度图，找到非空白像素的边界，再转换回PDF点
        坐标系，加一点边距后裁剪到页面实际范围内。
        结果按(文件路径,页码)缓存，避免每次渲染都重新分析一遍像素。
        检测失败、空白页、或检测结果过小（可能是误判）时返回None——
        调用方应该把None当成"不裁剪，显示整页"来处理。
        """
        cache_key = (file_path, page_idx)
        if cache_key in self._reader_content_bbox_cache:
            return self._reader_content_bbox_cache[cache_key]
        result = None
        try:
            page = doc[page_idx]
            analyze_zoom = 1.0  # 72 DPI足够判断"哪里有内容"，不需要高清，保持检测速度快
            pix = page.get_pixmap(matrix=fitz.Matrix(analyze_zoom, analyze_zoom), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            gray = img.convert("L")
            bw = gray.point(lambda p: 255 if p < 250 else 0)
            bbox = bw.getbbox()
            if bbox is not None:
                x0, y0, x1, y1 = bbox
                pad = 8.0  # 留一点边距，避免裁得太紧贴着文字
                page_rect = page.rect
                candidate = fitz.Rect(
                    max(page_rect.x0, x0 / analyze_zoom - pad),
                    max(page_rect.y0, y0 / analyze_zoom - pad),
                    min(page_rect.x1, x1 / analyze_zoom + pad),
                    min(page_rect.y1, y1 / analyze_zoom + pad))
                if candidate.width >= 10 and candidate.height >= 10:
                    result = candidate
        except Exception:
            result = None
        self._reader_content_bbox_cache[cache_key] = result
        return result

    def _reader_page_offset_table(self, zoom):
        """每页顶部 Y 的估算偏移表（按每页真实 rect.height 累加，纯整数不渲染）。
        返回 (offsets, 总估算高度)。混合页型文档比统一页高估计更准。"""
        doc = self._reader_pdf
        PAGE_GAP = 8
        if getattr(self, "_reader_page_heights_doc_id", None) != id(doc):
            self._reader_page_heights = [doc[i].rect.height for i in range(doc.page_count)]
            self._reader_page_heights_doc_id = id(doc)
        offs = []
        acc = 0
        for h in self._reader_page_heights:
            offs.append(acc)
            acc += int(h * zoom) + PAGE_GAP
        return offs, acc

    def _render_page_crisp(self, page, zoom):
        """
        按目标zoom渲染页面为PIL图像。当目标有效分辨率偏低时（比如"自适应
        宽度"模式下同时打开多个阅读区，画布变窄导致缩放比例很小），会按
        更高分辨率采样后用LANCZOS高质量算法缩小到目标尺寸——文字边缘的
        抗锯齿效果明显更好，看起来更清晰，而不是直接在低分辨率下渲染、
        发虚发糊。zoom本身已经足够高时（比如手动放大查看细节）不会做
        任何多余的超采样工作，不影响那种场景下的性能。

        优先命中「预渲染 PIL 缓存」：预取后台线程已把相邻页光栅化好，
        这里只需一次廉价内存拷贝（copy()），把耗时的 PyMuPDF 光栅化从
        主线程卸载到后台线程，翻页/滚动时 UI 不卡顿。
        """
        key = (id(page.parent), page.number, round(zoom, 4))
        with self._reader_pil_lock:
            cached = self._reader_pil_cache.get(key)
            if cached is not None:
                return cached.copy()
        img = self._render_page_raw(page, zoom)
        result = img.copy()                      # 调用方拿到独立副本，可自由 close()
        self._reader_pil_cache_put(key, img)     # 缓存持有原图
        return result

    def _render_page_raw(self, page, zoom):
        """无缓存的原始渲染（_render_page_crisp 的底层实现，预取线程直接调用）。"""
        MIN_RENDER_DPI = 200.0
        current_dpi = zoom * 72.0
        if current_dpi >= MIN_RENDER_DPI:
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ss = min(3.0, MIN_RENDER_DPI / max(current_dpi, 1.0))
        hi_zoom = zoom * ss
        pix = page.get_pixmap(matrix=fitz.Matrix(hi_zoom, hi_zoom), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        target_w = max(1, int(round(pix.width / ss)))
        target_h = max(1, int(round(pix.height / ss)))
        return img.resize((target_w, target_h), Image.LANCZOS)

    def _reader_pil_cache_put(self, key, img):
        """把预渲染结果写入缓存，超出上限按 LRU 淘汰（淘汰时释放底层位图）。"""
        with self._reader_pil_lock:
            self._reader_pil_cache[key] = img
            self._reader_pil_cache.move_to_end(key)
            while len(self._reader_pil_cache) > self._reader_pil_cap:
                _, old = self._reader_pil_cache.popitem(last=False)
                try:
                    old.close()
                except Exception:
                    pass

    def _reader_prefetch_start(self):
        """惰性启动预取后台线程（daemon，随主程序退出自动回收）。"""
        if self._reader_prefetch_thread is not None:
            return
        self._reader_prefetch_queue = queue.Queue()
        self._reader_prefetch_thread = threading.Thread(
            target=self._reader_prefetch_worker,
            daemon=True, name="pdf-prefetch")
        self._reader_prefetch_thread.start()

    def _reader_prefetch_worker(self):
        """后台预取线程：渲染相邻页到 PIL 缓存，主线程翻页时直接命中。"""
        while True:
            doc, page_idx, zoom = self._reader_prefetch_queue.get()
            try:
                key = (id(doc), page_idx, round(zoom, 4))
                with self._reader_pil_lock:
                    if key in self._reader_pil_cache:
                        continue
                page = doc[page_idx]
                img = self._render_page_raw(page, zoom)
                self._reader_pil_cache_put(key, img)
            except Exception:
                pass
            finally:
                self._reader_prefetch_queue.task_done()

    def _reader_prefetch_around(self, doc, start, end, zoom, total_pages):
        """按滚动方向预取：前方多、后方少（尚未渲染的）加入预取队列。"""
        if doc is None:
            return
        self._reader_prefetch_start()
        last = getattr(self, "_reader_last_prefetch_start", None)
        if last is None or start == last:
            front, back = 2, 2
        elif start > last:
            front, back = 4, 1   # 向下滚（往后翻）
        else:
            front, back = 1, 4   # 向上滚（往前翻）
        self._reader_last_prefetch_start = start

        def enq(p):
            key = (id(doc), p, round(zoom, 4))
            with self._reader_pil_lock:
                if key in self._reader_pil_cache:
                    return
            self._reader_prefetch_queue.put((doc, p, zoom))

        for p in range(end, min(total_pages, end + front)):
            enq(p)
        for p in range(max(0, start - back), start):
            enq(p)

    def _apply_theme_to_image(self, img):
        """
        根据当前主题对PIL Image应用颜色变换。
        支持的主题：normal, dark, green, sepia, black_white, blue
        """
        theme = self._reader_theme.get()

        if theme == "normal":
            return img

        elif theme in ("dark", "black_white"):
            # 暗黑/黑白主题：反转颜色（白→黑，黑→白）
            from PIL import ImageOps
            return ImageOps.invert(img)

        elif theme == "green":
            # 护眼绿：将白色背景改为淡绿色
            return self._replace_near_white(img, (199, 237, 204))

        elif theme == "blue":
            # 护眼蓝：将白色背景改为淡蓝色
            return self._replace_near_white(img, (232, 240, 248))

        elif theme == "sepia":
            # 由Pillow底层执行矩阵变换，比创建多份浮点数组更省内存。
            matrix = (
                0.393, 0.769, 0.189, 0,
                0.349, 0.686, 0.168, 0,
                0.272, 0.534, 0.131, 0,
            )
            return img.convert("RGB", matrix)

        else:
            return img

    @staticmethod
    def _replace_near_white(img, color):
        """只替换RGB三个通道都接近白色的像素。"""
        from PIL import ImageChops
        channels = img.split()
        masks = [channel.point(lambda value: 255 if value > 240 else 0)
                 for channel in channels]
        first_and_second = ImageChops.darker(masks[0], masks[1])
        mask = ImageChops.darker(first_and_second, masks[2])
        result = img.copy()
        result.paste(color, mask=mask)
        first_and_second.close()
        mask.close()
        for item in channels + tuple(masks):
            item.close()
        return result

    def _reader_on_container_resize(self, event):
        """容器大小变化时，延迟重排（防抖动）"""
        if not self._reader_fit_width or not self._reader_pdf:
            return
        # 防抖动：取消上一次的定时任务
        if self._reader_resize_job:
            self.after_cancel(self._reader_resize_job)
        self._reader_resize_job = self.after(500, self._reader_on_resize)

    def _reader_on_resize(self):
        """窗口大小变化后重新自适应渲染，保持高亮。"""
        if not self._reader_fit_width or not self._reader_pdf:
            return
        # 记住当前可见的页面，渲染后恢复
        old_page = self._reader_page
        self._reader_render(force_fit_width=True, keywords=self._reader_highlight_keywords)
        # 恢复页码显示和滚动位置
        self._reader_page = min(old_page, self._reader_pdf.page_count - 1)
        self._reader_scroll_to_page(self._reader_page)
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()

    def _reader_auto_fit_scale(self):
        """
        自动计算scale使PDF页面宽度填满画布可视区域（保持比例）。

        开启了"智能裁边"时，改用【裁剪后的内容宽度】来算，而不是整页
        宽度——否则缩放比例还是按整页算出来的，裁掉空白边距之后，
        实际显示的内容反而比画布窄了，留出一截空白没有撑满，没有达到
        "裁完边之后内容完全填满窗口、放大阅读"的效果。
        """
        if not self._reader_pdf or self._reader_pdf.page_count == 0:
            return
        # 确保画布布局已完成
        self._reader_canvas.update_idletasks()
        canvas_w = self._reader_canvas.winfo_width()
        if canvas_w <= 20:  # 画布还没布局好
            return
        # 取第一页宽度（多数PDF所有页同宽）
        page_w = self._reader_pdf[0].rect.width
        if self._reader_crop_margins_mode:
            content_rect = self._reader_get_content_bbox(
                self._reader_pdf, 0, self._reader_active_file_path)
            if content_rect and content_rect.width > 10:
                page_w = content_rect.width
        # canvas_w = page_w * (PDF_DPI/72) * scale
        self._reader_scale = (canvas_w / page_w) * (72.0 / self.PDF_DPI)
        # 放宽上限到6.0，让大屏幕能充分利用
        self._reader_scale = max(0.2, min(self._reader_scale, 6.0))

    # 首次渲染的最大页面数（最少10页，避免小PDF也懒加载）
    _READER_INITIAL_PAGES = 10

    def _reader_render(self, force_fit_width=None, keywords=None, quick=False):
        """
        优化渲染：只渲染可见区域附近的页面，使用快速Y偏移计算。
        force_fit_width: True=强制自适应, False=保持当前scale, None=使用self._reader_fit_width
        keywords: 高亮关键词列表
        quick: True时不强制"至少渲染10页"，只渲染刚好够填满当前视口的页数，
               让这一次同步渲染尽快完成（避免卡顿），渲染完成后会自动安排
               后台懒加载把缓冲页面补齐。用于"快速滚动后跳出已渲染范围"
               这种需要尽快出图的场景；正常渲染仍使用默认False，保持原有
               的"多渲染几页减少重复渲染次数"的效率优化。
        """
        if not self._reader_pdf:
            theme_page_bg = self._reader_themes[self._reader_theme.get()]["bg"]
            self._reader_canvas.config(bg=theme_page_bg)
            self._reader_canvas.delete("all")
            self._reader_canvas.create_rectangle(
                0, 0, 2000, 2000, fill=theme_page_bg, outline="", tags="bg_fill")
            return
        generation = self._reader_doc_generation

        # 取消可能存在的渲染防抖任务
        if self._reader_render_job:
            self.after_cancel(self._reader_render_job)
            self._reader_render_job = None
        if self._reader_lazy_job:
            try:
                self.after_cancel(self._reader_lazy_job)
            except Exception:
                pass
            self._reader_lazy_job = None

        # 存储高亮关键词供懒加载使用
        if keywords is not None:
            self._reader_highlight_keywords = keywords
        active_keywords = self._reader_highlight_keywords

        # 决定是否自适应宽度
        do_fit = self._reader_fit_width if force_fit_width is None else force_fit_width
        if do_fit:
            self._reader_auto_fit_scale()

        self._reader_tk_imgs = []
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        PAGE_GAP = 8

        # 获取画布尺寸
        self._reader_canvas.update_idletasks()
        canvas_w = self._reader_canvas.winfo_width()
        if canvas_w <= 20:
            canvas_w = 800
        canvas_h = self._reader_canvas.winfo_height()
        if canvas_h <= 20:
            canvas_h = 600

        # 估算页面高度（用于快速计算可见范围）
        total_pages = self._reader_pdf.page_count
        est_page_h = int(self._reader_pdf[0].rect.height * zoom) + PAGE_GAP
        # 每页真实高度累加的偏移表（混合页型文档滚动位置更准）
        _offsets, _total_h_est = self._reader_page_offset_table(zoom)

        # 获取当前滚动位置
        yview = self._reader_canvas.yview()
        scroll_y = yview[0] if yview else 0.0
        total_h_est = _total_h_est

        # 确定可见页面范围（前缓冲2页）。quick 模式（缩放/快速滚动补渲染）用
        # 向上取整的视口页数 + 3 尾缓冲，保证视口完整覆盖——用 floor 会在
        # 「一页比视口高」或页高不一累计误差大时漏渲染顶部一页，露出黑屏。
        view_top = scroll_y * total_h_est
        import bisect
        start_page = max(0, bisect.bisect_right(_offsets, view_top) - 1 - 2)
        if quick:
            view_pages = max(1, (canvas_h + est_page_h - 1) // est_page_h)
            end_page = min(total_pages, start_page + view_pages + 3)
        else:
            end_page = min(total_pages, start_page + int(canvas_h / est_page_h) + 5)
        # 确保最少渲染10页——但"快速重渲染"模式下不强制这个下限：
        # 这个最小渲染量是为了减少平时阅读时的重复渲染次数，但如果是
        # 快速滚动到很远的位置后触发的"补渲染"，一次性同步渲染10页正是
        # 卡顿的来源，这种场景应该尽量少渲染、快速出图，剩下的页面交给
        # 已有的后台懒加载机制慢慢补上，而不是卡住主线程。
        if not quick and end_page - start_page < 10:
            end_page = min(total_pages, start_page + 10)

        # 清除旧内容：先把画布底色同步为主题页面背景色，
        # 这样 delete("all") 后到 create_rectangle 之间的间隙也不会闪。
        theme_page_bg = self._reader_themes[self._reader_theme.get()]["bg"]
        self._reader_canvas.config(bg=theme_page_bg)
        self._reader_canvas.delete("all")
        self._reader_canvas.create_rectangle(
            0, 0, max(canvas_w, 2000), total_h_est + 2000,
            fill=theme_page_bg, outline="", tags="bg_fill")

        # 快速预计算所有页面的Y偏移（用每页真实高度的偏移表）
        self._reader_page_yoffsets = list(_offsets)
        self._reader_page_crop_origins = [(0.0, 0.0)] * total_pages

        # 渲染可见区域
        y_cursor = _offsets[start_page]
        max_w = 0
        rendered_count = 0
        CACHE_MAX = 25  # LRU缓存最多25页

        for i in range(start_page, end_page):
            # 缓存键包含主题和裁边模式，状态不同时不能复用旧缓存
            cache_key = (i, round(zoom, 4), self._reader_theme.get(),
                         self._reader_crop_margins_mode)
            cached = self._reader_page_cache.get(cache_key)

            if cached:
                # ✅ 缓存命中（高亮为覆盖层，页面图始终可复用缓存）
                _, tk_img, crop_origin = cached
                self._reader_page_cache.move_to_end(cache_key)
                # 从缓存的tk_img获取尺寸
                pix_w = tk_img.width()
                pix_h = tk_img.height()
                self._reader_tk_imgs.append(tk_img)
                self._reader_page_yoffsets[i] = y_cursor
                self._reader_page_crop_origins[i] = crop_origin
                x_offset = max(0, (canvas_w - pix_w) // 2)
                self._reader_canvas.create_image(x_offset, y_cursor, anchor="nw",
                                                  image=tk_img, tags=f"page_{i}")
                self._reader_draw_page_annotations(
                    self._reader_canvas, i, x_offset, y_cursor, zoom, crop_origin=crop_origin)
                if active_keywords:
                    words = self._reader_words_for_page(i, self._reader_pdf[i])
                    self._reader_draw_highlight_overlay(
                        self._reader_canvas, words, x_offset, y_cursor, zoom,
                        active_keywords, exact=self._reader_highlight_exact,
                        crop_origin=crop_origin)
                y_cursor += pix_h + PAGE_GAP
                if pix_w > max_w:
                    max_w = pix_w
            else:
                # ❌ 缓存未命中：重新渲染页面图（高亮由覆盖层叠加）
                page = self._reader_pdf[i]
                img = self._render_page_crisp(page, zoom)
                img_w, img_h = img.size
                # 应用主题变换
                img = self._apply_theme_to_image(img)
                crop_origin = (0.0, 0.0)
                if self._reader_crop_margins_mode:
                    content_rect = self._reader_get_content_bbox(
                        self._reader_pdf, i, self._reader_active_file_path)
                    if content_rect:
                        cbox = (int(content_rect.x0 * zoom), int(content_rect.y0 * zoom),
                                int(content_rect.x1 * zoom), int(content_rect.y1 * zoom))
                        cbox = (max(0, cbox[0]), max(0, cbox[1]),
                                min(img.width, cbox[2]), min(img.height, cbox[3]))
                        if cbox[2] > cbox[0] and cbox[3] > cbox[1]:
                            img = img.crop(cbox)
                            img_w, img_h = img.size
                            crop_origin = (content_rect.x0, content_rect.y0)
                self._reader_page_crop_origins[i] = crop_origin
                tk_img = ImageTk.PhotoImage(img)
                img.close()
                # 页面图始终写入缓存（高亮是覆盖层，不影响复用）
                self._reader_page_cache[cache_key] = (zoom, tk_img, crop_origin)
                while len(self._reader_page_cache) > CACHE_MAX:
                    self._reader_page_cache.popitem(last=False)
                self._reader_tk_imgs.append(tk_img)
                self._reader_page_yoffsets[i] = y_cursor
                x_offset = max(0, (canvas_w - img_w) // 2)
                self._reader_canvas.create_image(x_offset, y_cursor, anchor="nw",
                                                  image=tk_img, tags=f"page_{i}")
                self._reader_draw_page_annotations(
                    self._reader_canvas, i, x_offset, y_cursor, zoom, crop_origin=crop_origin)
                if active_keywords:
                    words = self._reader_words_for_page(i, self._reader_pdf[i])
                    self._reader_draw_highlight_overlay(
                        self._reader_canvas, words, x_offset, y_cursor, zoom,
                        active_keywords, exact=self._reader_highlight_exact,
                        crop_origin=crop_origin)
                y_cursor += img_h + PAGE_GAP
                if img_w > max_w:
                    max_w = img_w
            rendered_count += 1

        # 更新后续未渲染页面的Y偏移（用每页真实高度估算）
        for i in range(end_page, total_pages):
            self._reader_page_yoffsets[i] = _offsets[i]

        # scrollregion 使用估算总高度
        remaining = (total_h_est - _offsets[end_page]) if end_page < total_pages else 0
        total_h = max(y_cursor + remaining, total_h_est)
        self._reader_canvas.config(scrollregion=(0, 0, max(canvas_w, max_w), total_h))
        self._reader_render_upto = end_page
        self._reader_render_start = start_page

        # 只预取一批相邻页面；后续随滚动按需加载，避免长PDF全部驻留内存。
        if end_page < total_pages:
            self._reader_lazy_job = self.after(
                100, lambda g=generation: self._reader_lazy_load_pages(g))

        # 翻页预取：后台线程预渲染视口前后各 2 页到 PIL 缓存。quick（缩放）
        # 路径不预取——zoom 每格都在变、命中率低，且会与同步缩放渲染抢 CPU。
        if not quick:
            self._reader_prefetch_around(
                self._reader_pdf, start_page, end_page, zoom, total_pages)

    def _reader_lazy_load_pages(self, generation=None):
        """后台懒加载：分批渲染后续页面，不阻塞UI，同时保持高亮。"""
        self._reader_lazy_job = None
        if generation is None:
            generation = self._reader_doc_generation
        if generation != self._reader_doc_generation:
            return
        if not self._reader_pdf:
            return
        if self._reader_render_upto >= self._reader_pdf.page_count:
            return  # 已全部渲染完毕

        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        PAGE_GAP = 8
        canvas_w = self._reader_canvas.winfo_width()
        if canvas_w <= 20:
            canvas_w = 800

        start = self._reader_render_upto
        BATCH = 5  # 减小批次，降低UI卡顿感
        end = min(start + BATCH, self._reader_pdf.page_count)

        y_cursor = self._reader_page_yoffsets[start] if start < len(self._reader_page_yoffsets) else 0
        max_w = int(self._reader_canvas.cget("scrollregion").split()[2]) if self._reader_canvas.cget("scrollregion") else canvas_w

        active_keywords = self._reader_highlight_keywords

        CACHE_MAX = 25
        for i in range(start, end):
            # 缓存键包含主题和裁边模式，状态不同时不能复用旧缓存
            cache_key = (i, round(zoom, 4), self._reader_theme.get(),
                         self._reader_crop_margins_mode)
            cached = self._reader_page_cache.get(cache_key)
            if cached:
                # ✅ 缓存命中（高亮为覆盖层，页面图始终可复用缓存）
                _, tk_img, crop_origin = cached
                self._reader_page_cache.move_to_end(cache_key)
                # 从缓存的tk_img获取尺寸
                pix_w = tk_img.width()
                pix_h = tk_img.height()
                self._reader_tk_imgs.append(tk_img)
                self._reader_page_yoffsets[i] = y_cursor
                self._reader_page_crop_origins[i] = crop_origin
                x_offset = max(0, (canvas_w - pix_w) // 2)
                self._reader_canvas.create_image(x_offset, y_cursor, anchor="nw",
                                                  image=tk_img, tags=f"page_{i}")
                self._reader_draw_page_annotations(
                    self._reader_canvas, i, x_offset, y_cursor, zoom, crop_origin=crop_origin)
                if active_keywords:
                    words = self._reader_words_for_page(i, self._reader_pdf[i])
                    self._reader_draw_highlight_overlay(
                        self._reader_canvas, words, x_offset, y_cursor, zoom,
                        active_keywords, exact=self._reader_highlight_exact,
                        crop_origin=crop_origin)
                y_cursor += pix_h + PAGE_GAP
                if pix_w > max_w:
                    max_w = pix_w
            else:
                # ❌ 缓存未命中：重新渲染页面图（高亮由覆盖层叠加）
                page = self._reader_pdf[i]
                img = self._render_page_crisp(page, zoom)
                img_w, img_h = img.size
                # 应用主题变换
                img = self._apply_theme_to_image(img)
                crop_origin = (0.0, 0.0)
                if self._reader_crop_margins_mode:
                    content_rect = self._reader_get_content_bbox(
                        self._reader_pdf, i, self._reader_active_file_path)
                    if content_rect:
                        cbox = (int(content_rect.x0 * zoom), int(content_rect.y0 * zoom),
                                int(content_rect.x1 * zoom), int(content_rect.y1 * zoom))
                        cbox = (max(0, cbox[0]), max(0, cbox[1]),
                                min(img.width, cbox[2]), min(img.height, cbox[3]))
                        if cbox[2] > cbox[0] and cbox[3] > cbox[1]:
                            img = img.crop(cbox)
                            img_w, img_h = img.size
                            crop_origin = (content_rect.x0, content_rect.y0)
                self._reader_page_crop_origins[i] = crop_origin
                tk_img = ImageTk.PhotoImage(img)
                img.close()
                self._reader_page_cache[cache_key] = (zoom, tk_img, crop_origin)
                while len(self._reader_page_cache) > CACHE_MAX:
                    self._reader_page_cache.popitem(last=False)
                self._reader_tk_imgs.append(tk_img)
                self._reader_page_yoffsets[i] = y_cursor
                x_offset = max(0, (canvas_w - img_w) // 2)
                self._reader_canvas.create_image(x_offset, y_cursor, anchor="nw",
                                                  image=tk_img, tags=f"page_{i}")
                self._reader_draw_page_annotations(
                    self._reader_canvas, i, x_offset, y_cursor, zoom, crop_origin=crop_origin)
                if active_keywords:
                    words = self._reader_words_for_page(i, self._reader_pdf[i])
                    self._reader_draw_highlight_overlay(
                        self._reader_canvas, words, x_offset, y_cursor, zoom,
                        active_keywords, exact=self._reader_highlight_exact,
                        crop_origin=crop_origin)
                y_cursor += img_h + PAGE_GAP
                if img_w > max_w:
                    max_w = img_w

        self._reader_render_upto = end
        # 更新scrollregion：已渲染的精确高度 + 剩余页面的估算高度
        remaining = self._reader_pdf.page_count - end
        est_page_h = int(self._reader_pdf[0].rect.height * zoom) + PAGE_GAP
        total_h = y_cursor + remaining * est_page_h
        self._reader_canvas.config(scrollregion=(0, 0, max(canvas_w, max_w), total_h))

        # 只在用户已接近窗口尾部时继续预取，不再自动渲染整份文档。
        yview = self._reader_canvas.yview()
        view_bottom = yview[1] * total_h if yview else 0
        canvas_h = max(1, self._reader_canvas.winfo_height())
        if end < self._reader_pdf.page_count and view_bottom + canvas_h * 2 > y_cursor:
            self._reader_lazy_job = self.after(
                80, lambda g=generation: self._reader_lazy_load_pages(g))

    def _reader_highlight_page(self, img, page, keywords, zoom, exact=None,
                                offset_x=0.0, offset_y=0.0):
        """
        在页面图片上绘制匹配关键词的高亮框（黄色半透明）。

        offset_x/offset_y：默认是0，代表"img是从页面左上角(0,0)开始渲染
        的"——目前绝大多数调用方都是这种情况（整页渲染后再决定要不要
        裁剪，高亮永远是在裁剪【之前】画在整页坐标系上，裁剪只是事后把
        图片切掉一部分，高亮位置自然还是对的）。

        只有一种场景需要传非零的offset：渲染的时候【一开始就】用clip
        参数只截取了页面的一部分区域（比如悬停预览开启智能裁边时，为了
        省去渲染整页的开销，直接从裁剪区域开始渲染）——这时候img的
        左上角对应的不是页面的(0,0)，而是裁剪区域的左上角，所有单词的
        坐标都要先减去这个偏移量才能落在img上正确的位置，否则高亮框会
        完全错位（之前预览功能开启裁边后高亮经常不准，就是因为忘了减
        这个偏移量）。
        """
        if not keywords or not page:
            return img
        kw_set = {k.strip().lower() for k in keywords if k.strip()}
        if not kw_set:
            return img
        if exact is None:
            exact = self._reader_highlight_exact
        try:
            draw = ImageDraw.Draw(img, "RGBA")
            hr, hg, hb = self._reader_highlight_color
            fill = (hr, hg, hb, 120)
            for w in page.get_text("words"):
                word = self._clean_word(w[4])
                if not word:
                    continue
                word_lower = word.lower()
                # 点击文字时精确匹配；手动预览搜索仍支持子串包含。
                matched = word_lower in kw_set
                if not matched and not exact:
                    for kw in kw_set:
                        if kw in word_lower:
                            matched = True
                            break
                if matched:
                    x0 = int((w[0] - offset_x) * zoom)
                    y0 = int((w[1] - offset_y) * zoom)
                    x1 = int((w[2] - offset_x) * zoom)
                    y1 = int((w[3] - offset_y) * zoom)
                    draw.rectangle([x0, y0, x1, y1], fill=fill)
            del draw
        except Exception:
            pass
        return img

    def _reader_highlight_color_hex(self):
        """把高亮 RGB 元组转成 canvas 可用的 #rrggbb 字符串。"""
        try:
            r, g, b = self._reader_highlight_color
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return "#ffc832"

    def _reader_words_for_page(self, i, page):
        """优先取文本缓存，缓存未就绪时回退到 get_text('words')。仅用于主阅读区。"""
        cache = self._reader_text_cache
        if cache and 0 <= i < len(cache) and cache[i] is not None:
            return cache[i]
        return page.get_text("words")

    def _reader_draw_highlight_overlay(self, canvas, words, x_offset, y_cursor, zoom,
                                       keywords, exact=None, crop_origin=(0.0, 0.0),
                                       tag="hl_overlay"):
        """把匹配关键词的高亮框画成 canvas 覆盖层矩形（stipple 半透明），
        叠加在页面图片之上，而不是把高亮烘焙进页面图片重新渲染——这样点击
        文本时只需 delete+重画这些矩形，页面图片原样不动，就不会黑闪/卡顿。
        返回画出的矩形数量。"""
        if not keywords or not words:
            return 0
        kw_set = {k.strip().lower() for k in keywords if k.strip()}
        if not kw_set:
            return 0
        if exact is None:
            exact = self._reader_highlight_exact
        color = self._reader_highlight_color_hex()
        crop_x, crop_y = crop_origin
        n = 0
        for w in words:
            word = self._clean_word(w[4])
            if not word:
                continue
            word_lower = word.lower()
            matched = word_lower in kw_set
            if not matched and not exact:
                for kw in kw_set:
                    if kw in word_lower:
                        matched = True
                        break
            if not matched:
                continue
            canvas.create_rectangle(
                x_offset + (w[0] - crop_x) * zoom,
                y_cursor + (w[1] - crop_y) * zoom,
                x_offset + (w[2] - crop_x) * zoom,
                y_cursor + (w[3] - crop_y) * zoom,
                fill=color, outline="", stipple="gray50",
                tags=(tag,))
            n += 1
        return n

    def _reader_redraw_highlight_overlay(self):
        """点击文本后仅重画主阅读区高亮覆盖层（不重渲染页面，消除黑闪）。"""
        canvas = self._reader_canvas
        canvas.delete("hl_overlay")
        if not self._reader_pdf:
            return
        kw = self._reader_highlight_keywords
        if not kw:
            return
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        start = self._reader_render_start
        end = min(self._reader_render_upto, self._reader_pdf.page_count)
        for i in range(start, end):
            bbox = canvas.bbox(f"page_{i}")
            if not bbox:
                continue
            page = self._reader_pdf[i]
            crop_origin = (self._reader_page_crop_origins[i]
                           if i < len(self._reader_page_crop_origins) else (0.0, 0.0))
            words = self._reader_words_for_page(i, page)
            self._reader_draw_highlight_overlay(
                canvas, words, bbox[0], bbox[1], zoom, kw,
                exact=self._reader_highlight_exact, crop_origin=crop_origin)

    def _reader_aux_redraw_highlight_overlay(self, slot):
        """点击文本后仅重画指定辅助阅读区的高亮覆盖层。"""
        state = self._reader_aux.get(slot)
        if not state or not state.get("doc"):
            return
        canvas = state["canvas"]
        tag = f"hl_overlay_{slot}"
        canvas.delete(tag)
        kw = state.get("keywords") or []
        if not kw:
            return
        zoom = state.get("zoom", 1.0)
        doc = state["doc"]
        for i in range(state.get("render_start", 0),
                       min(state.get("render_end", 0), doc.page_count)):
            bbox = canvas.bbox(f"aux_{slot}_page_{i}")
            if not bbox:
                continue
            page = doc[i]
            crop_origin = state.get("page_crop_origins", {}).get(i, (0.0, 0.0))
            words = page.get_text("words")
            self._reader_draw_highlight_overlay(
                canvas, words, bbox[0], bbox[1], zoom, kw,
                exact=state.get("exact", False), crop_origin=crop_origin, tag=tag)

    def _reader_check_lazy_load(self):
        """滚动时检查：如果接近未渲染区域，触发懒加载"""
        if not self._reader_pdf:
            return

        # 获取当前可视区域底部位置
        yview = self._reader_canvas.yview()
        total_h = int(self._reader_canvas.cget("scrollregion").split()[-1]) if self._reader_canvas.cget("scrollregion") else 1
        if total_h <= 0:
            return
        view_bottom_y = yview[1] * total_h

        # 最后一个已渲染页面的底部Y坐标
        last_y = self._reader_page_yoffsets[self._reader_render_upto - 1] if self._reader_render_upto > 0 and self._reader_render_upto <= len(self._reader_page_yoffsets) else 0
        # 估算最后一个已渲染页面的高度
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        est_page_h = int(self._reader_pdf[0].rect.height * zoom) + 8
        view_top_y = yview[0] * total_h
        import bisect
        _offsets, _ = self._reader_page_offset_table(zoom)
        visible_page = max(0, min(
            self._reader_pdf.page_count - 1,
            bisect.bisect_right(_offsets, view_top_y) - 1))
        if (visible_page < self._reader_render_start or
                visible_page >= self._reader_render_upto):
            self._reader_render(
                force_fit_width=False, keywords=self._reader_highlight_keywords,
                quick=True)
            return
        if self._reader_render_upto >= self._reader_pdf.page_count:
            return
        last_page_bottom = last_y + est_page_h

        # 如果可视区域底部距离未渲染区域小于3屏高度，触发懒加载
        canvas_h = self._reader_canvas.winfo_height()
        if view_bottom_y + canvas_h * 3 > last_page_bottom:
            self.after_idle(self._reader_lazy_load_pages)

    def _reader_go_page(self, delta):
        if not self._reader_pdf:
            return
        new = self._reader_page + delta
        if 0 <= new < self._reader_pdf.page_count:
            self._reader_page = new
            self._reader_scroll_to_page(new)
            self._reader_update_active_page_info()
            self._reader_highlight_thumb()
            self._reader_sync_scroll_others(1)

    def _reader_scroll_to_page(self, pg_idx):
        """滚动到指定页面的顶部（使用缓存Y偏移加速）"""
        if not self._reader_pdf:
            return
        # 优先使用缓存Y偏移
        if self._reader_page_yoffsets and pg_idx < len(self._reader_page_yoffsets):
            y = self._reader_page_yoffsets[pg_idx]
        else:
            # 快速估算：使用第一页高度 * 页码
            zoom = self.PDF_DPI / 72.0 * self._reader_scale
            PAGE_GAP = 8
            est_page_h = int(self._reader_pdf[0].rect.height * zoom) + PAGE_GAP
            y = pg_idx * est_page_h
        canvas_h = self._reader_canvas.winfo_height()
        total_h = int(self._reader_canvas.cget("scrollregion").split()[-1]) if self._reader_canvas.cget("scrollregion") else 1
        if total_h > 0:
            ratio = y / total_h
            self._reader_canvas.yview_moveto(ratio)
            if (pg_idx < self._reader_render_start or
                    pg_idx >= self._reader_render_upto):
                self.after_idle(lambda: self._reader_render(
                    force_fit_width=False,
                    keywords=self._reader_highlight_keywords))

    def _reader_detect_current_page(self):
        """根据当前滚动位置检测当前页码（使用缓存Y偏移加速）"""
        if not self._reader_pdf:
            return
        yview = self._reader_canvas.yview()
        total_h = int(self._reader_canvas.cget("scrollregion").split()[-1]) if self._reader_canvas.cget("scrollregion") else 1
        if total_h <= 0:
            return
        view_top = yview[0] * total_h
        canvas_h = self._reader_canvas.winfo_height()
        view_center = view_top + canvas_h * 0.3

        # 使用缓存Y偏移加速
        yoffsets = self._reader_page_yoffsets
        if yoffsets:
            import bisect
            # 找到第一个y偏移 > view_center的项
            idx = bisect.bisect_right(yoffsets, view_center) - 1
            i = max(0, idx)
        else:
            # 回退：使用估算，避免遍历所有页面
            zoom = self.PDF_DPI / 72.0 * self._reader_scale
            PAGE_GAP = 8
            est_page_h = int(self._reader_pdf[0].rect.height * zoom) + PAGE_GAP
            i = max(0, int(view_center / est_page_h))
            i = min(i, self._reader_pdf.page_count - 1)

        if self._reader_page != i:
            self._reader_page = i
            self._reader_update_active_page_info()
            self._reader_highlight_thumb()


    def _reader_zoom(self, event):
        self._reader_set_active_slot(1)
        if self._reader_hscroll_on and (event.state & 0x1):
            self._reader_hscroll_shift(1 if event.delta < 0 else -1)
            return "break"
        # 如果当前打开的是纯文本类文件（Markdown原文/PPT/EPUB），用Text控件滚动
        # （代码文件走独立的_reader_code_text，已单独绑定滚轮，不在此处处理）
        if not self._reader_pdf and self._reader_file_type in (".md", ".ppt", ".pptx", ".epub"):
            if self._reader_text_preview:
                self._reader_text_preview.yview_scroll(int(-1 * event.delta / 120) * 3, "units")
            return
        if not self._reader_pdf:
            return
        # Ctrl+滚轮 = 缩放，普通滚轮 = 翻页滚动
        if event.state & 0x4:  # Ctrl 键
            if self._reader_magnifier_wheel_zoom(event):
                return
            if self._reader_scroll_anim_job:
                self.after_cancel(self._reader_scroll_anim_job)
                self._reader_scroll_anim_job = None
                self._reader_scroll_velocity = 0.0
            self._reader_fit_width = False  # 手动缩放时关闭自适应
            # 首次滚轮记录光标下的 PDF 锚点（缩放后回卷，保持光标内容不动）
            if self._reader_zoom_anchor is None:
                anchor_pos = self._reader_magnifier_pdf_pos(1, event.x, event.y)
                if anchor_pos is not None:
                    self._reader_zoom_anchor = anchor_pos
                    self._reader_zoom_anchor_view = (event.x, event.y)
            # 累积缩放变化，防抖后统一重渲染
            new_scale = self._reader_scale * (1.1 if event.delta > 0 else 0.9)
            new_scale = max(0.2, min(new_scale, 6.0))
            self._reader_scale = new_scale
            # 取消旧的渲染任务
            if self._reader_zoom_job:
                self.after_cancel(self._reader_zoom_job)
            # 延迟50ms后重渲染，合并极速滚轮的同时保持跟手
            self._reader_zoom_job = self.after(50, self._reader_zoom_render)
        else:
            # 逐行滚动：每次滚轮滚动固定行数，不叠加惯性动画
            notches = int(event.delta / 120.0) if abs(event.delta) >= 120 else (1 if event.delta > 0 else -1)
            lines = -notches * self._reader_scroll_lines
            self._reader_scroll_velocity = 0.0  # 清除惯性速度
            if self._reader_scroll_anim_job:
                self.after_cancel(self._reader_scroll_anim_job)
                self._reader_scroll_anim_job = None
            self._reader_canvas.yview_scroll(lines, "units")
            self._reader_detect_current_page()
            self._reader_check_lazy_load()
            self._reader_sync_scroll_others(1)

    def _reader_zoom_render(self):
        """防抖后的缩放重渲染"""
        self._reader_zoom_job = None
        anchor = self._reader_zoom_anchor
        anchor_view = self._reader_zoom_anchor_view
        self._reader_zoom_anchor = None
        self._reader_zoom_anchor_view = None
        if not self._reader_pdf:
            return
        # 使用增量渲染策略：先清空并重新计算可见区域，后台懒加载剩余。
        # quick=True 只渲染刚好填满视口的页数，缩放跟手、不卡顿。
        self._reader_render(force_fit_width=False, keywords=self._reader_highlight_keywords, quick=True)
        # 缩放后把光标锚点回卷到原视口位置，实现「缩放锚定光标」
        if anchor is not None and anchor_view is not None:
            self._reader_scroll_to_anchor(1, anchor, anchor_view)
        self._reader_sync_scale_others(1)


    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 点击文字 → 全文档搜索 → 裁剪缩略图
    # ════════════════════════════════════════════════════════════════

    # ════════════════════════════════════════════════════════════════
    # PDF阅读器 - 框选正文区 / 圆盘菜单 / 单击·双击·右键判定
    # ════════════════════════════════════════════════════════════════

    def _reader_all_canvases(self):
        """返回所有阅读区画布 [(slot, canvas), ...]。"""
        out = [(1, self._reader_canvas)]
        for slot, state in self._reader_aux.items():
            out.append((slot, state["canvas"]))
        return out

    def _reader_restore_normal_click(self):
        """恢复所有阅读区画布为普通左键行为（点选搜索/拖拽平移）。"""
        for slot, canvas in self._reader_all_canvases():
            canvas.config(cursor="arrow")
            canvas.bind("<ButtonPress-1>", lambda e, s=slot: self._reader_on_canvas_press(s, e))
            canvas.bind("<B1-Motion>", lambda e, s=slot: self._reader_on_canvas_drag(s, e))
            canvas.unbind("<ButtonRelease-1>")

    def _reader_canvas_point_to_pdf(self, slot, canvas_x, canvas_y):
        """把画布坐标换算成 (page_idx, pdf_x, pdf_y)。找不到页时返回 None。"""
        import bisect
        if slot == 1:
            doc = self._reader_pdf
            if not doc or not self._reader_page_yoffsets:
                return None
            yoffsets = self._reader_page_yoffsets
            idx = bisect.bisect_right(yoffsets, canvas_y) - 1
            if not 0 <= idx < doc.page_count:
                return None
            bbox = self._reader_canvas.bbox(f"page_{idx}")
            x_offset = bbox[0] if bbox else 0
            crop_x, crop_y = (self._reader_page_crop_origins[idx]
                              if idx < len(self._reader_page_crop_origins) else (0.0, 0.0))
            zoom = self.PDF_DPI / 72.0 * self._reader_scale
            return idx, (canvas_x - x_offset) / zoom + crop_x, (canvas_y - yoffsets[idx]) / zoom + crop_y
        state = self._reader_aux.get(slot)
        if not state or not state["doc"] or not state["page_yoffsets"]:
            return None
        yoffsets = state["page_yoffsets"]
        idx = bisect.bisect_right(yoffsets, canvas_y) - 1
        if not 0 <= idx < state["doc"].page_count:
            return None
        page_x = state["page_xoffsets"].get(idx, 0)
        crop_x, crop_y = state.get("page_crop_origins", {}).get(idx, (0.0, 0.0))
        return idx, (canvas_x - page_x) / state["zoom"] + crop_x, (canvas_y - yoffsets[idx]) / state["zoom"] + crop_y

    def _reader_preview_region(self, doc, page_idx, file_path):
        """弹窗/匹配页预览用的「正文区」：手框正文区优先，否则整页。
        手框一次后弹窗只显示框内正文、切掉表头/外框。"""
        page = doc[page_idx]
        region = self._reader_body_crop_regions.get(file_path)
        if region is not None:
            pr = page.rect
            rect = fitz.Rect(region[0] * pr.width, region[1] * pr.height,
                             region[2] * pr.width, region[3] * pr.height)
            return rect & page.rect
        return page.rect

    def _toggle_reader_body_crop_mode(self):
        """开启/关闭「框选正文区」：开启后在画布拖一个框，松手把框内区域记为
        当前文件的正文区（归一化坐标），后续弹窗/匹配页预览只显示这个区域、
        切掉表头和外框。一次框选全文档复用，再次框选覆盖；与标注/文本模式互斥。"""
        self._reader_body_crop_mode = not self._reader_body_crop_mode
        self._reader_body_crop_start = None
        self._reader_body_crop_slot = None
        self._reader_body_crop_clear_rect()
        if self._reader_body_crop_mode:
            if getattr(self, "_reader_annotate_mode", False):
                self._toggle_reader_annotate_mode()
            if getattr(self, "_reader_text_mode", False):
                self._toggle_reader_text_mode()
            self._reader_body_crop_btn.config(text="⛶ 框选正文(开)", bg=ACCENT)
            for slot, canvas in self._reader_all_canvases():
                canvas.config(cursor="crosshair")
                canvas.bind("<ButtonPress-1>", lambda e, s=slot: self._reader_body_crop_press(s, e))
                canvas.bind("<B1-Motion>", lambda e, s=slot: self._reader_body_crop_drag(s, e))
                canvas.bind("<ButtonRelease-1>", lambda e, s=slot: self._reader_body_crop_release(s, e))
        else:
            self._reader_body_crop_btn.config(text="⛶ 框选正文", bg="#454b63")
            self._reader_restore_normal_click()

    def _reader_body_crop_press(self, slot, event):
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        cx = canvas.canvasx(event.x)
        cy = canvas.canvasy(event.y)
        self._reader_body_crop_slot = slot
        self._reader_body_crop_start = (cx, cy)
        if self._reader_body_crop_rect_id is not None:
            try:
                canvas.delete(self._reader_body_crop_rect_id)
            except Exception:
                pass
        self._reader_body_crop_rect_id = canvas.create_rectangle(
            cx, cy, cx, cy, outline=ACCENT, width=2, dash=(4, 2))

    def _reader_body_crop_drag(self, slot, event):
        if self._reader_body_crop_start is None or self._reader_body_crop_rect_id is None:
            return
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        cx = canvas.canvasx(event.x)
        cy = canvas.canvasy(event.y)
        x0, y0 = self._reader_body_crop_start
        canvas.coords(self._reader_body_crop_rect_id, x0, y0, cx, cy)

    def _reader_body_crop_release(self, slot, event):
        start = self._reader_body_crop_start
        self._reader_body_crop_start = None
        if start is None:
            return
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        end = (canvas.canvasx(event.x), canvas.canvasy(event.y))
        self._reader_body_crop_clear_rect()
        ps = self._reader_canvas_point_to_pdf(slot, start[0], start[1])
        pe = self._reader_canvas_point_to_pdf(slot, end[0], end[1])
        doc = self._reader_doc_for_slot(slot)
        if not ps or not pe or not doc:
            return
        s_page, sx, sy = ps
        e_page, ex, ey = pe
        if s_page != e_page:
            # 跨页框选不支持，直接退出框选模式（不覆盖已有正文区）
            self._reader_body_crop_mode = False
            self._reader_body_crop_btn.config(text="⛶ 框选正文", bg="#454b63")
            self._reader_restore_normal_click()
            return
        pr = doc[s_page].rect
        nx0 = max(0.0, min(1.0, (min(sx, ex) - pr.x0) / pr.width))
        nx1 = max(0.0, min(1.0, (max(sx, ex) - pr.x0) / pr.width))
        ny0 = max(0.0, min(1.0, (min(sy, ey) - pr.y0) / pr.height))
        ny1 = max(0.0, min(1.0, (max(sy, ey) - pr.y0) / pr.height))
        if nx1 - nx0 < 0.02 or ny1 - ny0 < 0.02:
            # 框太小视为误触，不覆盖已有正文区
            self._reader_body_crop_mode = False
            self._reader_body_crop_btn.config(text="⛶ 框选正文", bg="#454b63")
            self._reader_restore_normal_click()
            return
        file_path = self._reader_path_for_slot(slot)
        if not file_path:
            return
        self._reader_body_crop_regions[file_path] = (nx0, ny0, nx1, ny1)
        self._save_config()
        self._reader_body_crop_mode = False
        self._reader_body_crop_btn.config(text="⛶ 框选正文", bg="#454b63")
        self._reader_restore_normal_click()

    def _reader_body_crop_clear_rect(self):
        rid = self._reader_body_crop_rect_id
        self._reader_body_crop_rect_id = None
        if rid is None:
            return
        slot = self._reader_body_crop_slot
        canvas = None
        if slot == 1:
            canvas = getattr(self, "_reader_canvas", None)
        elif slot:
            st = self._reader_aux.get(slot)
            canvas = st["canvas"] if st else None
        if canvas is not None:
            try:
                canvas.delete(rid)
            except Exception:
                pass

    # ── 单击/双击/右键判定 ──

    def _reader_cancel_pending_click(self):
        if self._reader_click_job is not None:
            try:
                self.after_cancel(self._reader_click_job)
            except Exception:
                pass
            self._reader_click_job = None
        self._reader_pending_click = None

    def _reader_on_canvas_press(self, slot, event):
        """画布左键按下：框选模式走框选；否则立即开始拖拽平移，并把「点选
        文字搜索」延后 260ms 执行，给双击留出判定窗口。"""
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        if self._reader_body_crop_mode:
            self._reader_body_crop_press(slot, event)
            return
        self._reader_cancel_pending_click()
        self._reader_press_moved = False
        self._reader_press_xy = (event.x, event.y)
        canvas.scan_mark(event.x, event.y)
        self._reader_pending_click = (slot, event)
        self._reader_click_job = self.after(260, self._reader_fire_pending_click)

    def _reader_on_canvas_drag(self, slot, event):
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        if self._reader_body_crop_mode:
            self._reader_body_crop_drag(slot, event)
            return
        if self._reader_press_xy is not None:
            px, py = self._reader_press_xy
            if abs(event.x - px) + abs(event.y - py) >= 4:
                self._reader_press_moved = True
        canvas.scan_dragto(event.x, event.y, gain=1)

    def _reader_fire_pending_click(self):
        self._reader_click_job = None
        if self._reader_pending_click is None:
            return
        slot, event = self._reader_pending_click
        self._reader_pending_click = None
        if self._reader_press_moved:
            return  # 是拖拽，不是点击
        if slot == 1:
            self._on_reader_canvas_click(event)
        else:
            self._reader_aux_canvas_click(slot, event)


    def _reader_on_canvas_double(self, slot, event):
        self._reader_cancel_pending_click()
    def _reader_on_canvas_right(self, slot, event):
        self._reader_cancel_pending_click()
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        try:
            current = canvas.find_withtag("current")
            if current:
                tags = canvas.gettags(current[0])
                if any(t in ("reader_annotation", "reader_text_mark", "reader_text_hl") for t in tags):
                    return
        except Exception:
            pass


    # ── UI 位置标注（给程序各区域起名，浮层显示，方便描述优化位置）──







    def _on_reader_canvas_click(self, event):
        """点击预览文字 → 高亮收藏模式时持久高亮该词。"""
        self._reader_set_active_slot(1)
        if not self._reader_pdf:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        canvas_x = self._reader_canvas.canvasx(event.x)
        canvas_y = self._reader_canvas.canvasy(event.y)
        yoffsets = self._reader_page_yoffsets
        if not yoffsets:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        import bisect
        idx = bisect.bisect_right(yoffsets, canvas_y) - 1
        target_page = idx if 0 <= idx < self._reader_pdf.page_count else None
        if target_page is None:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        bbox = self._reader_canvas.bbox(f"page_{target_page}")
        x_offset = bbox[0] if bbox else 0
        crop_x, crop_y = (
            self._reader_page_crop_origins[target_page]
            if target_page < len(self._reader_page_crop_origins) else (0.0, 0.0))
        pdf_x = (canvas_x - x_offset) / zoom + crop_x
        pdf_y = (canvas_y - yoffsets[target_page]) / zoom + crop_y
        self._reader_page = target_page
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()
        page = self._reader_pdf[target_page]
        clicked_word = None
        clicked_rect = None
        cache = self._reader_text_cache
        if cache and cache[target_page] is not None:
            for w in cache[target_page]:
                if w[0] <= pdf_x <= w[2] and w[1] <= pdf_y <= w[3]:
                    clicked_word = self._clean_word(w[4])
                    clicked_rect = [w[0], w[1], w[2], w[3]]
                    break
        else:
            for w in page.get_text("words"):
                if w[0] <= pdf_x <= w[2] and w[1] <= pdf_y <= w[3]:
                    clicked_word = self._clean_word(w[4])
                    clicked_rect = [w[0], w[1], w[2], w[3]]
                    break
        if not clicked_word or len(clicked_word) < 1:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        self._reader_last_clicked_word = clicked_word
        if self._reader_text_hl_mode and clicked_rect:
            self._reader_save_text_hl(target_page, clicked_rect, clicked_word)

        # 联动搜索：如果开启，同步搜索所有已打开文件
        if self._reader_linked_search.get():
            self._reader_linked_search_all(clicked_word, source_slot=1)
        # 匹配页码列：更新命中页面（不弹窗）
        self._reader_search_and_show(clicked_word)


    def _reader_show_minimal_popup(self, word, occurrences, slot=1):
        """简略弹窗：仅文字列表显示命中页码与次数，不渲染缩略图。
        slot=1→PDF1跳转，slot=2/3→对应辅助阅读区跳转。"""
        from collections import Counter
        page_counts = Counter(pg for pg, _ in occurrences)
        total = len(occurrences)
        page_list = sorted(page_counts.items())

        win = tk.Toplevel(self)
        slot_label = f"PDF{slot} · " if slot > 1 else ""
        win.title(f"{slot_label}「{word}」— {total}处 / {len(page_counts)}页")
        win.configure(bg=CARD)
        win.attributes("-topmost", True)
        win.minsize(220, 160)

        # 定位在主窗口右侧
        mx = self.winfo_rootx() + self.winfo_width() + 12
        my = self.winfo_rooty() + 80
        screen_w = win.winfo_screenwidth()
        if mx + 320 > screen_w:
            mx = self.winfo_rootx() - 340
        win.geometry(f"320x{min(520, max(180, len(page_list) * 28 + 60))}+{max(0, mx)}+{max(0, my)}")

        # 标题栏
        hdr = tk.Frame(win, bg=ACCENT, height=36)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"{slot_label}「{word}」命中 {total} 处 / {len(page_counts)} 页",
                 bg=ACCENT, fg="white", font=("Microsoft YaHei UI", 10, "bold"),
                 padx=12).pack(side=tk.LEFT, pady=6)

        # 可滚动列表
        canvas = tk.Canvas(win, bg=CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(win, orient=tk.VERTICAL, command=canvas.yview)
        canvas.config(yscrollcommand=scrollbar.set)
        inner = tk.Frame(canvas, bg=CARD, padx=8, pady=6)
        canvas.create_window((0, 0), window=inner, anchor="nw", tags="inner")
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=4)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=4)

        for pg, cnt in page_list:
            row = tk.Frame(inner, bg=CARD, cursor="hand2")
            row.pack(fill=tk.X, pady=1)
            tk.Label(row, text=f"第 {pg + 1} 页",
                     bg=CARD, fg=TEXT, font=("Microsoft YaHei UI", 10),
                     anchor="w", padx=4).pack(side=tk.LEFT)
            tk.Label(row, text=f"{cnt}处",
                     bg=CARD, fg=ACCENT, font=("Microsoft YaHei UI", 9, "bold"),
                     anchor="e").pack(side=tk.RIGHT, padx=8)

            if slot == 1:
                def make_handler(pg_idx):
                    return lambda e: self._reader_jump_to(pg_idx)
            else:
                def make_handler(pg_idx):
                    return lambda e: self._reader_aux_jump_to_page(slot, pg_idx)
            for w in row.winfo_children():
                w.bind("<Button-1>", make_handler(pg))

        def _on_wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_wheel)
        inner.bind("<MouseWheel>", _on_wheel)

        inner.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

        # 鼠标离开后自动关闭
        def _leave_close(e):
            try:
                win.destroy()
            except Exception:
                pass
        win.bind("<Leave>", _leave_close)
        win.bind("<Escape>", lambda e: win.destroy())

    def _reader_jump_to(self, pg_idx):
        """跳转到指定页（连续模式下滚动到页面顶部）"""
        if not self._reader_pdf:
            return
        if 0 <= pg_idx < self._reader_pdf.page_count:
            self._reader_page = pg_idx
            self._reader_scroll_to_page(pg_idx)
            self._reader_update_active_page_info()
            self._reader_highlight_thumb()

    def _reader_build_thumbnails(self):
        """生成纯数字页码导航条（分批创建，避免大PDF阻塞UI）"""
        if not self._reader_pdf:
            return
        canvas = self._reader_thumb_canvas
        canvas.delete("all")
        self._reader_thumb_images = []
        self._reader_thumb_positions = []

        total = self._reader_pdf.page_count
        BOX_W = 40
        BOX_H = 32
        PAD = 6
        GROUP_COLORS = [
            "#c0392b", "#1e8449", "#2471a3", "#d68910",
            "#7d3c98", "#148f77", "#ba4a00", "#1e8449"
        ]

        # 预计算所有位置（纯数学运算，非常快）
        x = PAD
        for i in range(total):
            self._reader_thumb_positions.append((x, BOX_W))
            x += BOX_W + PAD
        canvas.config(scrollregion=(0, 0, max(x, 1), BOX_H + 8))

        # 分批创建，每批最多30个，避免大PDF卡顿
        def _create_batch(start, end):
            if not self._reader_pdf:
                return
            end = min(end, total)
            for i in range(start, end):
                x_pos, w = self._reader_thumb_positions[i]
                color = GROUP_COLORS[(i // 10) % len(GROUP_COLORS)]
                canvas.create_rectangle(x_pos, 4, x_pos + w, 4 + BOX_H,
                                        fill=color, outline="", tags=f"thumb_{i}")
                canvas.create_text(x_pos + w // 2, 4 + BOX_H // 2,
                                   text=str(i + 1), fill="white",
                                   font=("Microsoft YaHei UI", 10, "bold"),
                                   anchor="center", tags=f"thumb_{i}")
                canvas.tag_bind(f"thumb_{i}", "<Button-1>",
                                lambda e, idx=i: self._reader_jump_to(idx))
            # 继续下一批
            if end < total:
                canvas.after(10, lambda: _create_batch(end, end + 30))
            else:
                self._reader_highlight_thumb()

        _create_batch(0, min(30, total))

    def _reader_highlight_thumb(self):
        """高亮当前页码并滚动到可见区域"""
        if not self._reader_pdf or not hasattr(self, '_reader_thumb_positions'):
            return
        canvas = self._reader_thumb_canvas
        canvas.delete("thumb_sel")
        i = self._reader_page
        if 0 <= i < len(self._reader_thumb_positions):
            x, w = self._reader_thumb_positions[i]
            BOX_H = 32
            # 白色高亮边框
            canvas.create_rectangle(x - 2, 2, x + w + 2, 6 + BOX_H + 2,
                                    outline="#ffffff", width=2,
                                    tags="thumb_sel")
            # 确保页码可见（自动滚动）
            bbox = canvas.bbox("all")
            if bbox:
                total_w = bbox[2]
                canvas_w = canvas.winfo_width()
                if canvas_w > 0:
                    left_ratio = max(0, (x - 10) / total_w)
                    right_ratio = min(1, (x + w + 10) / total_w)
                    cur_left = canvas.xview()[0]
                    cur_right = canvas.xview()[1]
                    if right_ratio > cur_right or left_ratio < cur_left:
                        canvas.xview_moveto(left_ratio)

    # ════════════════════════════════════════════════════════════════
    # Everything 全局搜索
    # ════════════════════════════════════════════════════════════════




    # ════════════════════════════════════════════════════════════════
    # 网络 PDF 搜索 — TI 官方直链 + 360 搜索 + 下载 + 预览 + 另存
    # ════════════════════════════════════════════════════════════════
    # 三级数据源，按可靠性排序：
    #   ① TI 官方 datasheet 直链 ti.com/lit/ds/symlink/<型号族>.pdf —— 逐级
    #      截短型号去 HEAD 探测，命中即最权威（TPS16417 的手册在 tps1641.pdf）。
    #   ② 360 搜索 www.so.com —— 国内可直连、结构干净、data-mdurl 直接给真实
    #      URL；需要先预热首页拿 cookie，否则回验证码页。
    #   ③ Bing 兜底 —— 仅在 ①② 都没结果时用。Bing 对无 JS 的 HTTP 请求会
    #      返回"诱饵页"（标题回显搜索词、结果体是随机无关内容）——实测标准
    #      库 urllib 和 curl_cffi 指纹伪装都无效，所以结果必须过一遍相关性
    #      校验才敢展示，详见 _reader_net_search_worker 的说明。
    # 结果按"型号是否出现在标题/URL 里"过滤掉凑数条目；一条都不剩时回退显示
    # 全部并在状态栏标注。下载到临时目录后用 fitz 预览（复用 PDF1 阅读区），
    # 再另存为 / 一键存入默认目录。搜索抓取用 curl_cffi 伪装浏览器指纹（对
    # 360 有效更稳，对 Bing 诱饵页无效）。TI 探测与下载仍用标准库 urllib。

    _NET_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
               "AppleWebKit/537.36 (KHTML, like Gecko) "
               "Chrome/124.0.0.0 Safari/537.36")
























    # ════════════════════════════════════════════════════════════════
    # 多实例联动 — 跨程序发送到其它实例的 Everything 搜索栏
    # ════════════════════════════════════════════════════════════════
    # 机制：每个实例启动时（1）用命名互斥体判断自己是不是"第一个打开的主程序"，
    # （2）在共享临时目录写一份 <pid>.json 注册文件供其它实例发现。点击文本时，
    # 若已启用"跨程序发送"，把文字复制到剪贴板，并原子写入每个勾选目标实例的
    # <pid>.inbox.json；目标实例通过 after() 轮询自己的收件箱，收到新文本就
    # 填入 Everything 搜索框并自动执行搜索。全程走本地文件、无网络、无额外依赖。

    def _ipc_dir(self):
        """共享临时目录：存放各实例的注册文件和收件箱。"""
        import tempfile
        return os.path.join(tempfile.gettempdir(), "bjzy_reader_instances")



    @staticmethod
    def _ipc_pid_alive(pid):
        """判断指定 pid 的进程是否仍存活（Windows：进程句柄未收到退出信号）。"""
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.OpenProcess.argtypes = [
                wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            kernel32.OpenProcess.restype = wintypes.HANDLE
            kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
            kernel32.WaitForSingleObject.restype = wintypes.DWORD
            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            SYNCHRONIZE = 0x00100000
            WAIT_TIMEOUT = 0x102
            h = kernel32.OpenProcess(SYNCHRONIZE, False, pid)
            if not h:
                return False
            ret = kernel32.WaitForSingleObject(h, 0)
            kernel32.CloseHandle(h)
            return ret == WAIT_TIMEOUT  # 仍运行中（未发出退出信号）
        except Exception:
            return False




    def _ipc_send_text(self, text):
        """把文本原子写入所有勾选目标实例的收件箱。"""
        import json
        import time
        import tempfile
        text = (text or "").strip()
        targets = list(self._cross_send_targets)
        if not text or not targets:
            return
        nonce = time.time_ns()
        for pid in targets:
            if not self._ipc_pid_alive(pid):
                continue
            inbox = os.path.join(self._ipc_dir(), f"{pid}.inbox.json")
            try:
                fd, tmp = tempfile.mkstemp(
                    dir=self._ipc_dir(), prefix=f"{pid}.", suffix=".tmp")
                try:
                    with os.fdopen(fd, "w", encoding="utf-8") as f:
                        json.dump({"nonce": nonce, "text": text}, f)
                    os.replace(tmp, inbox)  # 原子替换，避免目标读到半截文件
                except Exception:
                    try:
                        os.remove(tmp)
                    except Exception:
                        pass
            except Exception:
                pass



    def _reader_copy_and_send(self, text):
        """点击文本后：总是复制到剪贴板；若已启用跨程序发送，再发送到勾选的目标实例。"""
        text = (text or "").strip()
        if not text:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update_idletasks()
        except Exception:
            pass
        if self._cross_send_enabled.get():
            self._ipc_send_text(text)



    # ════════════════════════════════════════════════════════════════
    # PDF 阅读器 - 内容搜索
    # ════════════════════════════════════════════════════════════════


    def _reader_anytxt_cross_hint(self, slot, query, exclude_path, hint_var):
        """后台用 AnyTXT 查「当前文件夹里还有哪些文件也含该关键词」，回主线程更新提示。

        只做跨文件命中计数+文件列表，不动本文件内的页码高亮（那仍需 fitz 本地匹配）。
        AnyTXT 不可用（返回 None）或查询失败时静默置空提示，不影响本文件搜索。
        """
        if not query:
            return
        source = (list(self._reader_pdf_files_all) if self._reader_pdf_files_all
                  else list(self._reader_pdf_files))
        ex_norm = os.path.normcase(exclude_path or "")
        paths = [p for p, _ in source
                 if p and os.path.normcase(p) != ex_norm]
        if not paths:
            return
        self._reader_anytxt_seq += 1
        seq = self._reader_anytxt_seq
        hint_var.set("🔎 查找其他文件...")

        def worker():
            hits = None
            try:
                hits = AnyTxtSearch.content_filter(paths, query)
            except Exception:
                hits = None

            def done():
                if getattr(self, "_reader_anytxt_seq", 0) != seq:
                    return  # 过期结果，丢弃
                if hits is None:
                    hint_var.set("")
                elif hits:
                    self._reader_anytxt_hits[slot] = list(hits)
                    hint_var.set(f"🔎 另有 {len(hits)} 文件命中")
                else:
                    self._reader_anytxt_hits[slot] = []
                    hint_var.set("🔎 无其他文件命中")
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

        """点击「另有 N 文件命中」提示 → 弹出命中文件列表，双击/按钮打开。"""
        hits = self._reader_anytxt_hits.get(slot, [])
        if not hits:
            return
        win = tk.Toplevel(self)
        win.title(f"跨文件命中 ({len(hits)})")
        win.attributes("-topmost", True)
        win.configure(bg=PANEL)
        try:
            win.geometry(f"+{max(0, self.winfo_pointerx() - 40)}+{max(0, self.winfo_pointery() - 20)}")
        except Exception:
            pass
        tk.Label(win, text=f"以下 {len(hits)} 个文件也含该关键词：",
                 bg=PANEL, fg=TEXT, anchor="w", padx=8, pady=4,
                 font=("Microsoft YaHei UI", 9, "bold")).pack(fill=tk.X)
        lb = tk.Listbox(win, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                        selectforeground="white", relief=tk.FLAT, bd=0,
                        highlightthickness=0, font=("Microsoft YaHei UI", 9),
                        width=70, height=min(16, len(hits)))
        lb.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        for p in hits:
            lb.insert(tk.END, f"  {os.path.basename(p)}    {p}")

        def open_selected():
            sel = lb.curselection()
            if not sel:
                return
            path = hits[sel[0]]
            try:
                win.destroy()
            except Exception:
                pass
            self._reader_open(path, os.path.basename(path))

        lb.bind("<Double-Button-1>", lambda e: open_selected())
        RoundedButton(win, text="打开所选", bg=ACCENT, fg="white",
                      relief=tk.FLAT, cursor="hand2",
                      font=("Microsoft YaHei UI", 9), padx=8,
                      command=open_selected).pack(pady=(0, 8))



    # ════════════════════════════════════════════════════════════════
    # 联动搜索
    # ════════════════════════════════════════════════════════════════


    # ════════════════════════════════════════════════════════════════
    # 工作区快照（需求3）：保存/切换 多文档组合 + 阅读位置 + 主题
    # ════════════════════════════════════════════════════════════════
    # ════════════════════════════════════════════════════════════════
    # 工作区文件夹（可嵌套）+ 排序
    # ════════════════════════════════════════════════════════════════






















    def _reader_epub_selection_search(self, text_widget, event):
        """EPUB 滑动选中关键词后触发搜索（松开鼠标时检测选区）。"""
        # 推断 slot
        slot = None
        for s in (2, 3, 4):
            state = self._reader_aux.get(s)
            if state and state.get("epub_text") is text_widget:
                slot = s
                break
        if slot is None:
            slot = 1

        # 检测当前是否有选区
        try:
            sel_start = text_widget.index(tk.SEL_FIRST)
            sel_end = text_widget.index(tk.SEL_LAST)
            if not text_widget.compare(sel_start, "<", sel_end):
                return  # 没有选区，不处理
        except tk.TclError:
            return

        # 获取选中文字并清理
        selected = text_widget.get(sel_start, sel_end).strip()
        # 限制长度，避免选中大段文字
        if len(selected) > 80:
            # 取前 80 个字符中的第一个"词"
            selected = selected[:80]
        word = self._clean_word(selected)
        if not word or len(word) < 1:
            return

        # 跨程序发送：选中文字 → 复制到剪贴板 + 发送到勾选的目标实例
        self._reader_copy_and_send(word)

        if self._reader_linked_search.get():
            self._reader_linked_search_all(word, source_slot=slot)
        else:
            # 非联动模式：在当前 EPUB 中高亮匹配
            self._reader_search_text_file(word, text_widget)

    def _reader_search_text_file(self, word, text_widget):
        """非联动模式：在文本文件中搜索单词，在文本控件中高亮匹配。"""
        # 清除旧高亮
        text_widget.tag_remove("search_hl", "1.0", tk.END)
        text_widget.tag_configure("search_hl", background="#ff5050", foreground="#ffffff")

        # 搜索并高亮所有匹配项
        idx = "1.0"
        count = 0
        while True:
            idx = text_widget.search(word, idx, nocase=True, stopindex=tk.END)
            if not idx:
                break
            end_idx = f"{idx}+{len(word)}c"
            text_widget.tag_add("search_hl", idx, end_idx)
            count += 1
            idx = end_idx

        # 跳转到第一个匹配项
        if count > 0:
            first_idx = text_widget.search(word, "1.0", nocase=True, stopindex=tk.END)
            if first_idx:
                text_widget.see(first_idx)
            self._reader_preview_search.set(word)
            self._reader_highlight_keywords = [word]

    def _reader_search_cached_for_slot(self, slot, word):
        """对指定 slot（可能是 slot 1 或 aux slot）执行 PDF 单词搜索。"""
        if slot == 1:
            if not self._reader_pdf:
                return []
            occurrences = self._reader_search_cached(word)
            if occurrences is None:
                # 回退到逐页扫描
                word_lower = word.lower()
                occurrences = []
                for i in range(self._reader_pdf.page_count):
                    for w in self._reader_pdf[i].get_text("words"):
                        if self._clean_word(w[4]).lower() == word_lower:
                            occurrences.append((i, fitz.Rect(w[0], w[1], w[2], w[3])))
            return occurrences
        else:
            state = self._reader_aux.get(slot)
            if not state or not state.get("doc"):
                return []
            doc = state["doc"]
            word_lower = word.lower()
            occurrences = []
            for i in range(doc.page_count):
                for w in doc[i].get_text("words"):
                    if self._clean_word(w[4]).lower() == word_lower:
                        occurrences.append((i, fitz.Rect(w[0], w[1], w[2], w[3])))
            return occurrences

    def _search_text_for_word(self, text, word):
        """在文本内容中搜索单词，返回包含匹配行的行号列表。"""
        word_lower = word.lower()
        lines = text.split("\n")
        matched_lines = []
        for i, line in enumerate(lines):
            if word_lower in line.lower():
                matched_lines.append(i + 1)  # 1-based 行号
        return matched_lines

    def _reader_get_epub_text(self, path):
        """从缓存或文件中获取 EPUB 的纯文本内容。"""
        if path in self._epub_chapters_cache:
            chapters = self._epub_chapters_cache[path]
            return "\n".join(text for _, text in chapters)
        # 回退：重新解析
        try:
            chapters = self._epub_parse_chapters(path)
            return "\n".join(text for _, text in chapters)
        except Exception:
            return ""

    def _reader_search_substring_in_doc(self, doc, word):
        """在 PDF 文档里做子串匹配，返回 [(page_idx, None), ...]（每处命中一条）。

        和 _reader_search_cached_for_slot 的"精确整词"不同，这里按子串在整页
        文本里计数，专门给预览搜索栏的模糊匹配用。返回的 Rect 位置对联动高亮
        路径没用（高亮靠 keywords 重渲染，不靠 Rect），所以统一填 None；
        每处命中一条记录，是为了让下游按 page_idx 分组时能数出真实的每页命中数。
        """
        word_lower = word.lower()
        out = []
        for i in range(doc.page_count):
            try:
                c = doc[i].get_text().lower().count(word_lower)
            except Exception:
                c = 0
            for _ in range(c):
                out.append((i, None))
        return out

    def _reader_search_in_slot(self, slot, word, substring=False):
        """在指定 slot 中搜索单词。返回 (path, file_type, matches)。

        substring=True 时 PDF 走子串模糊匹配（预览搜索栏），False 走精确整词
        （点击文本 / 联动搜索）。
        """
        if slot == 1:
            doc = self._reader_pdf
            path = self._reader_active_file_path
            file_type = self._reader_file_type
        else:
            state = self._reader_aux.get(slot)
            if not state or not state.get("visible"):
                return None, None, None
            doc = state.get("doc")
            path = state.get("path", "")
            file_type = os.path.splitext(path)[1].lower() if path else ""

        if not path:
            return None, None, None

        ext = file_type if file_type.startswith(".") else os.path.splitext(path)[1].lower()

        if ext == ".pdf" and doc:
            if substring:
                matches = self._reader_search_substring_in_doc(doc, word)
            else:
                matches = self._reader_search_cached_for_slot(slot, word)
            return path, "pdf", matches

        elif ext == ".epub":
            text = self._reader_get_epub_text(path)
            if text:
                matches = self._search_text_for_word(text, word)
                return path, "epub", matches
            return path, "epub", []

        elif ext in (".html", ".htm"):
            try:
                content, _ = self._read_text_file_smart(path)
                matches = self._search_text_for_word(content, word)
                return path, "html", matches
            except Exception:
                return path, "html", []

        elif ext in CODE_VIEW_EXTS | {".md", ".txt"}:
            try:
                content, _ = self._read_text_file_smart(path)
                matches = self._search_text_for_word(content, word)
                return path, "text", matches
            except Exception:
                return path, "text", []

        elif ext in (".ppt", ".pptx") and doc:
            matches = self._reader_search_cached_for_slot(slot, word)
            return path, "ppt", matches

        return None, None, None

    def _reader_get_text_widget_for_slot(self, slot):
        """获取指定 slot 当前显示的文本控件（用于 EPUB/代码/文本高亮）。"""
        if slot == 1:
            if self._reader_file_type == ".epub" and self._reader_text_frame.winfo_ismapped():
                return self._reader_text_preview
            if self._reader_file_type in CODE_VIEW_EXTS and self._reader_code_frame.winfo_ismapped():
                return self._reader_code_text
            if self._reader_text_frame.winfo_ismapped():
                return self._reader_text_preview
            return None
        else:
            state = self._reader_aux.get(slot)
            if state and state.get("epub_text") and state.get("epub_frame"):
                if state["epub_frame"].winfo_ismapped():
                    return state["epub_text"]
            return None

    def _reader_highlight_in_text_widget(self, text_widget, word):
        """在 Text 控件中高亮所有匹配 word 的位置。返回匹配次数。"""
        text_widget.tag_remove("linked_hl", "1.0", tk.END)
        text_widget.tag_configure("linked_hl", background="#ffdc00", foreground="#1a1a1a")
        idx = "1.0"
        count = 0
        while True:
            idx = text_widget.search(word, idx, nocase=True, stopindex=tk.END)
            if not idx:
                break
            end_idx = f"{idx}+{len(word)}c"
            text_widget.tag_add("linked_hl", idx, end_idx)
            count += 1
            idx = end_idx
        # 滚动到第一个匹配
        if count > 0:
            first = text_widget.search(word, "1.0", nocase=True, stopindex=tk.END)
            if first:
                text_widget.see(first)
        return count

    def _reader_update_source_slot(self, slot, word, result):
        """更新触发搜索的源 slot 的 UI（高亮 + 匹配面板）。"""
        if not result:
            return

        if slot == 1:
            if result["type"] == "pdf" and self._reader_pdf:
                self._reader_highlight_keywords = [word]
                self._reader_highlight_exact = True
                self._reader_preview_search.set(word)
                occurrences = result["matches"]
                matched_pages = sorted({p for p, _ in occurrences})
                self._reader_match_pages = matched_pages
                self._reader_match_counts = {}
                for p, _ in occurrences:
                    self._reader_match_counts[p] = self._reader_match_counts.get(p, 0) + 1
                self._reader_match_lb.delete(0, tk.END)
                for p in matched_pages:
                    self._reader_match_lb.insert(tk.END, f"  {p + 1}")
                self._reader_match_count_var.set(
                    self._fmt_match_count(len(matched_pages), len(occurrences)))
                self._reader_refresh_main_match_flow()
                self._reader_redraw_highlight_overlay()
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(1)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    self._reader_preview_search.set(word)
                    self._reader_highlight_keywords = [word]
                    self._reader_match_count_var.set(self._fmt_match_count(0, count))
        else:
            state = self._reader_aux.get(slot)
            if not state:
                return
            state["keywords"] = [word]
            state["exact"] = True
            state["search_var"].set(word)
            if result["type"] == "pdf":
                occurrences = result["matches"]
                pages = sorted({p for p, _ in occurrences})
                counts = {}
                for p, _ in occurrences:
                    counts[p] = counts.get(p, 0) + 1
                state["match_pages"] = pages
                state["match_counts"] = counts
                state["match_lb"].delete(0, tk.END)
                for p in pages:
                    state["match_lb"].insert(tk.END, f"第{p + 1}页  {counts[p]}处")
                state["match_count_var"].set(self._fmt_match_count(len(pages), len(occurrences)))
                self._reader_refresh_aux_match_flow(slot)
                self._reader_aux_redraw_highlight_overlay(slot)
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(slot)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    state["match_count_var"].set(self._fmt_match_count(0, count))

    def _reader_update_other_slot_highlight(self, slot, word, result, exact=True):
        """更新其他 slot 的高亮并触发重渲染，使关键词高亮在画布上可见。

        exact=False 用于预览搜索栏的模糊子串传播——让 PDF2/3 的高亮和 PDF1 的
        子串高亮保持一致（点击文本的联动搜索仍走 exact=True 整词匹配）。
        """
        if not result:
            return

        if slot == 1:
            if result["type"] == "pdf" and self._reader_pdf:
                self._reader_highlight_keywords = [word]
                self._reader_highlight_exact = exact
                self._reader_preview_search.set(word)
                occurrences = result["matches"]
                matched_pages = sorted({p for p, _ in occurrences})
                self._reader_match_pages = matched_pages
                self._reader_match_counts = {}
                for p, _ in occurrences:
                    self._reader_match_counts[p] = self._reader_match_counts.get(p, 0) + 1
                self._reader_match_lb.delete(0, tk.END)
                for p in matched_pages:
                    self._reader_match_lb.insert(tk.END, f"  {p + 1}")
                self._reader_match_count_var.set(
                    self._fmt_match_count(len(matched_pages), len(occurrences)))
                self._reader_refresh_main_match_flow()
                # 触发重渲染使关键词高亮可见
                self._reader_redraw_highlight_overlay()
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(1)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    self._reader_highlight_keywords = [word]
                    self._reader_preview_search.set(word)
                    self._reader_match_count_var.set(self._fmt_match_count(0, count))
        else:
            state = self._reader_aux.get(slot)
            if not state:
                return
            state["keywords"] = [word]
            state["exact"] = exact
            state["search_var"].set(word)
            if result["type"] == "pdf":
                occurrences = result["matches"]
                pages = sorted({p for p, _ in occurrences})
                counts = {}
                for p, _ in occurrences:
                    counts[p] = counts.get(p, 0) + 1
                state["match_pages"] = pages
                state["match_counts"] = counts
                state["match_lb"].delete(0, tk.END)
                for p in pages:
                    state["match_lb"].insert(tk.END, f"第{p + 1}页  {counts[p]}处")
                state["match_count_var"].set(self._fmt_match_count(len(pages), len(occurrences)))
                self._reader_refresh_aux_match_flow(slot)
                # 触发重渲染使关键词高亮在画布上可见
                self._reader_aux_redraw_highlight_overlay(slot)
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(slot)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    state["match_count_var"].set(self._fmt_match_count(0, count))

    def _reader_linked_search_all(self, word, source_slot=1, quiet=False):
        """联动搜索：在全部已打开的阅读区中搜索关键词，展示统一结果。

        quiet=True 时（如 IPC 后台触发）找不到匹配只静默返回，不弹窗打断用户。
        """
        results = {}  # {slot: {"type": ..., "path": ..., "matches": ...}}

        # 遍历 3 个 slot，分别搜索
        for slot in (1, 2, 3, 4):
            path, file_type, matches = self._reader_search_in_slot(slot, word)
            if path and matches is not None:
                results[slot] = {
                    "type": file_type,
                    "path": path,
                    "matches": matches,
                }

        if not results:
            if not quiet:
                messagebox.showinfo("联动搜索", f"所有已打开文件中未找到: \"{word}\"")
            return

        # 更新源 slot 的高亮和匹配面板
        self._reader_update_source_slot(source_slot, word, results.get(source_slot))

        # 更新其他 slot 的高亮
        for slot in results:
            if slot != source_slot:
                self._reader_update_other_slot_highlight(slot, word, results[slot])

        # 不弹窗，通过各 slot 的匹配面板即可查看跨文件结果
        # 匹配信息已通过 _reader_update_source_slot / _reader_update_other_slot_highlight 更新



    # ════════════════════════════════════════════════════════════════
    # 关闭
    # ════════════════════════════════════════════════════════════════

    def on_close(self):
        self._reader_doc_generation += 1
        for job_name in ("_reader_lazy_job", "_reader_scroll_job", "_reader_render_job",
                         "_reader_zoom_job", "_reader_scroll_anim_job"):
            job = getattr(self, job_name, None)
            if job:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                setattr(self, job_name, None)
        self._reader_save_resume()
        for slot in self._reader_aux:
            self._reader_save_aux_resume(slot)
        self._save_config()
        if self._reader_pdf:
            self._reader_pdf.close()
        for state in self._reader_aux.values():
            for key in ("resize_job", "scroll_anim_job"):
                if state.get(key):
                    try:
                        self.after_cancel(state[key])
                    except Exception:
                        pass
            if state.get("doc"):
                try:
                    state["doc"].close()
                except Exception:
                    pass
        for image in self._diff_images:
            try:
                image.close()
            except Exception:
                pass
        self._diff_images.clear()
        self._reader_page_cache.clear()
        self.destroy()

    # ════════════════════════════════════════════════════════════════
    # 🔧 脚本工坊 — PS脚本管理 + 联动执行
    # ════════════════════════════════════════════════════════════════

    # ════════════════════════════════════════════════════════════════
    # 文本搜索页签（txt/xlsx/c/cpp/py 关键词全文搜索 + 内容预览）
    # 集成自 Python-excel-txt-search/auto_save_txt_5_2_3_1_2_9-final use.py
    # ════════════════════════════════════════════════════════════════

    _TXT_SEARCH_MAX_SELECTION = 100
    _TXT_SEARCH_TYPES = [
        (".txt", ".txt 文件"), (".xlsx", ".xlsx 文件"), (".c", ".c/.h 文件"),
        (".cpp", ".cpp/.hpp 文件"), (".py", ".py 文件"),
    ]


























    # ════════════════════════════════════════════════════════════════
    # Excel 内容搜索页签（xlsx/xls 结构化关键词搜索 + 层级展示）
    # 复刻 PowerShell「powershell-excel-搜索.md」的匹配语义：
    #   首词 = 文件名过滤（* 或单关键词 = 不过滤）
    #   关键词间 AND/OR 逻辑（默认 OR），-前缀 = 精确匹配（忽略首尾空格）
    #   .xlsx → openpyxl；.xls(97-2003) → xlrd
    # ════════════════════════════════════════════════════════════════

    # 关键词高亮配色池（循环取色，深浅主题均可读）
    _EXCEL_KW_COLORS = [
        "#ff6b6b", "#4ecdc4", "#ffd93d", "#6c5ce7", "#ff8c42",
        "#2dd4a8", "#ff77e9", "#5eb8e0", "#ff9ff3", "#7bed9f",
    ]

    def _build_excel_search_tab(self, parent):
        """构建「Excel 搜索」页签：结构化关键词搜索 + 缩进层级 + 关键词字级高亮。"""
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)

        # ── 标题栏 ──
        hdr = tk.Frame(parent, bg=PANEL, height=36)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📊 Excel 内容搜索 — xlsx / xls 结构化关键词搜索",
                 bg=PANEL, fg=TEXT, font=("Microsoft YaHei UI", 12, "bold"),
                 padx=12).pack(side=tk.LEFT, pady=6)
        RoundedButton(hdr, text="📁 选择路径", bg=ENTRY, fg=TEXT,
                      relief=tk.FLAT, font=("Microsoft YaHei UI", 8),
                      padx=8, cursor="hand2",
                      command=self._excel_search_select_path).pack(side=tk.RIGHT, padx=8, pady=6)
        RoundedButton(hdr, text="▾ 历史", bg=ENTRY, fg=TEXT,
                      relief=tk.FLAT, font=("Microsoft YaHei UI", 8),
                      padx=8, cursor="hand2",
                      command=self._excel_search_show_path_history).pack(side=tk.RIGHT, padx=2, pady=6)
        tk.Checkbutton(hdr, text="递归子目录", variable=self._excel_search_recursive,
                       bg=PANEL, fg=TEXT, selectcolor=ENTRY,
                       activebackground=PANEL, activeforeground=TEXT,
                       font=("Microsoft YaHei UI", 8)).pack(side=tk.RIGHT, padx=6)

        # ── 搜索栏 ──
        bar = tk.Frame(parent, bg=PANEL)
        bar.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 2))
        bar.columnconfigure(0, weight=1)
        self._excel_search_path_var = tk.StringVar(value=self._excel_search_search_path)
        tk.Label(bar, textvariable=self._excel_search_path_var, bg=PANEL, fg=TEXT,
                 font=("Microsoft YaHei UI", 8)).grid(row=0, column=0, sticky="w")
        # 文件名搜索栏
        tk.Label(bar, text="📄 文件名（支持 AND/OR，-前缀精确）", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 7)).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._excel_filename_entry = tk.Entry(bar, bg=ENTRY, fg=TEXT,
                                              insertbackground=TEXT, relief=tk.FLAT,
                                              font=("Microsoft YaHei UI", 10))
        self._excel_filename_entry.grid(row=2, column=0, sticky="ew", pady=(2, 0), ipady=3)
        self._excel_filename_entry.bind("<Return>", lambda e: self._excel_search_run())
        self._bind_search_history(
            self._excel_filename_entry, "excel_search_fn",
            lambda v: (self._excel_filename_entry.delete(0, tk.END),
                       self._excel_filename_entry.insert(0, v)))
        # 内容搜索栏
        tk.Label(bar, text="🔍 内容（支持 AND/OR，-前缀精确）", bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 7)).grid(row=3, column=0, sticky="w", pady=(4, 0))
        self._excel_search_entry = tk.Entry(bar, bg=ENTRY, fg=TEXT,
                                            insertbackground=TEXT, relief=tk.FLAT,
                                            font=("Microsoft YaHei UI", 10))
        self._excel_search_entry.grid(row=4, column=0, sticky="ew", pady=(2, 0), ipady=3)
        self._excel_search_entry.bind("<Return>", lambda e: self._excel_search_run())
        self._bind_search_history(
            self._excel_search_entry, "excel_search_kw",
            lambda v: (self._excel_search_entry.delete(0, tk.END),
                       self._excel_search_entry.insert(0, v)))
        btn_row = tk.Frame(bar, bg=PANEL)
        btn_row.grid(row=5, column=0, sticky="ew", pady=(4, 0))
        RoundedButton(btn_row, text="📁 列出全部", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=12, cursor="hand2",
                      command=self._excel_search_list_all).pack(side=tk.LEFT)
        RoundedButton(btn_row, text="🔍 搜索", bg=ACCENT, fg="white", relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=12, cursor="hand2",
                      command=self._excel_search_run).pack(side=tk.LEFT, padx=(6, 0))
        RoundedButton(btn_row, text="✕ 停止", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=12, cursor="hand2",
                      command=self._excel_search_stop).pack(side=tk.LEFT, padx=(6, 0))
        RoundedButton(btn_row, text="⬇ 导出全部", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=12, cursor="hand2",
                      command=lambda: self._excel_search_export(False)).pack(side=tk.LEFT, padx=(6, 0))
        RoundedButton(btn_row, text="⬇ 导出选中", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=12, cursor="hand2",
                      command=lambda: self._excel_search_export(True)).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(btn_row, text="首词=文件名(*=全部)  空格/AND/OR  -前缀=精确",
                 bg=PANEL, fg=DIM, font=("Microsoft YaHei UI", 7)).pack(side=tk.LEFT, padx=10)

        # ── 主区：单一结果 Text（层级缩进 + 关键词字级高亮，可选中复制）──
        main = tk.Frame(parent, bg=CARD)
        main.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 0))
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        self._excel_result_text = tk.Text(main, bg=CARD, fg=TEXT, insertbackground=TEXT,
                                          relief=tk.FLAT, wrap="none",
                                          font=("Microsoft YaHei UI", 9),
                                          padx=10, pady=8, cursor="arrow")
        rvsb = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self._excel_result_text.yview)
        rhsb = ttk.Scrollbar(main, orient=tk.HORIZONTAL, command=self._excel_result_text.xview)
        self._excel_result_text.configure(yscrollcommand=rvsb.set, xscrollcommand=rhsb.set)
        self._excel_result_text.grid(row=0, column=0, sticky="nsew")
        rvsb.grid(row=0, column=1, sticky="ns")
        rhsb.grid(row=1, column=0, sticky="ew")
        # 固定层级配色：文件名 / sheet / 表头 三色（关键词色在 render 动态配置）
        self._excel_result_text.tag_configure("file", foreground="#4ea8e0")    # 蓝 — 文件名
        self._excel_result_text.tag_configure("sheet", foreground="#4cc38a")   # 绿 — sheet
        self._excel_result_text.tag_configure("header", foreground="#e0a83e")  # 琥珀 — 表头
        self._excel_result_text.config(state=tk.DISABLED)

        # ── 状态栏 ──
        self._excel_search_status_var = tk.StringVar(value="输入关键词开始搜索（多词空格分隔）")
        tk.Label(parent, textvariable=self._excel_search_status_var, bg=PANEL, fg=DIM,
                 font=("Microsoft YaHei UI", 8), anchor="w",
                 padx=8).grid(row=3, column=0, sticky="ew", pady=(2, 4))

    def _excel_search_select_path(self):
        path = filedialog.askdirectory(title="选择要搜索的 Excel 目录")
        if not path:
            return
        self._excel_search_search_path = path
        self._excel_search_path_var.set(path)
        self._excel_search_add_recent(path)
        self._save_config()

    def _excel_search_add_recent(self, path):
        if path not in self._excel_search_recent_paths:
            self._excel_search_recent_paths.insert(0, path)
            self._excel_search_recent_paths = self._excel_search_recent_paths[:10]

    def _excel_search_show_path_history(self):
        """弹出最近路径下拉列表，选择填入。"""
        paths = self._excel_search_recent_paths
        if not paths:
            return
        win = tk.Toplevel(self)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        w, h = 500, min(240, 26 * len(paths) + 6)
        x = self.winfo_pointerx() - 100
        y = self.winfo_pointery() + 10
        win.geometry(f"{w}x{h}+{x}+{y}")
        lb = tk.Listbox(win, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                        selectforeground="white", relief=tk.SOLID, bd=1,
                        font=("Microsoft YaHei UI", 9), highlightthickness=0,
                        activestyle="none")
        lb.pack(fill=tk.BOTH, expand=True)
        for p in paths:
            lb.insert(tk.END, p)

        def _pick(event):
            idx = lb.nearest(event.y)
            if 0 <= idx < len(paths):
                p = paths[idx]
                win.destroy()
                self._excel_search_search_path = p
                self._excel_search_path_var.set(p)
                self._excel_search_add_recent(p)
                self._save_config()

        lb.bind("<Button-1>", _pick)
        lb.bind("<Escape>", lambda e: win.destroy())


    def _excel_parse_keywords(self, raw):
        """解析关键词列表 + AND/OR + 精确标记（不含首词文件名过滤，供文件名/内容两个搜索栏共用）。返回 (keywords, logic, exact_set)。"""
        tokens = [t for t in raw.split() if t]
        if not tokens:
            return [], "OR", set()
        has_and = "AND" in tokens
        has_or = "OR" in tokens
        if has_and and has_or:
            return None, "OR", set()  # 冲突，由调用方提示
        if has_and:
            logic = "AND"
            kw_list = [t for t in tokens if t != "AND"]
        elif has_or:
            logic = "OR"
            kw_list = [t for t in tokens if t != "OR"]
        else:
            logic = "OR"
            kw_list = tokens
        exact_set = set()
        cleaned = []
        seen = set()
        for kw in kw_list:
            exact = kw.startswith("-") and len(kw) > 1
            clean = kw[1:] if exact else kw
            low = clean.lower()
            if low in seen:
                continue
            seen.add(low)
            cleaned.append(clean)
            if exact:
                exact_set.add(clean)
        return cleaned, logic, exact_set

    def _excel_filename_matches(self, fname, keywords, logic, exact_set):
        """文件名是否匹配关键词（AND/OR 逻辑，- 前缀精确匹配文件名）。"""
        low = fname.lower()
        stem = os.path.splitext(low)[0]

        def match(kw):
            k = kw.lower()
            if kw in exact_set:
                return low == k or stem == k
            return k in low

        if not keywords:
            return True
        if logic == "AND":
            return all(match(kw) for kw in keywords)
        return any(match(kw) for kw in keywords)

    @staticmethod
    def _excel_search_row_to_text(row):
        return " | ".join(str(c) if c is not None and c != "" else "" for c in row)

    def _excel_search_row_matches(self, row, keywords, logic, exact_set):
        """返回该行命中的关键词列表（AND=全命中才返回全列表；OR=命中的所有）。"""
        cells = [str(c).strip() for c in row
                 if c is not None and str(c).strip() != ""]

        def cell_matches(cell, kw):
            c = cell.strip()
            if kw in exact_set:
                return c.lower() == kw.lower()
            return kw.lower() in c.lower()

        hits = [kw for kw in keywords if any(cell_matches(c, kw) for c in cells)]
        if logic == "AND":
            return hits if len(hits) == len(keywords) else []
        return hits

    def _excel_search_one_file(self, fp, keywords, logic, exact_set):
        """扫描单个 Excel 文件，返回 {file, sheets:[{sheet,header,rows}]} 或 None。"""
        ext = os.path.splitext(fp)[1].lower()
        sheets = []
        if ext == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(fp, read_only=True, data_only=True)
            try:
                for ws in wb:
                    header = None
                    rows = []
                    for ri, row in enumerate(ws.iter_rows(values_only=True)):
                        if ri == 0:
                            header = self._excel_search_row_to_text(row)
                            continue
                        hits = self._excel_search_row_matches(row, keywords, logic, exact_set)
                        if hits:
                            rows.append((self._excel_search_row_to_text(row), hits))
                    if rows:
                        sheets.append({"sheet": ws.title, "header": header or "", "rows": rows})
            finally:
                wb.close()
        elif ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(fp)
            for sh in wb.sheets():
                header = self._excel_search_row_to_text(sh.row_values(0)) if sh.nrows > 0 else ""
                rows = []
                for ri in range(1, sh.nrows):
                    row = sh.row_values(ri)
                    hits = self._excel_search_row_matches(row, keywords, logic, exact_set)
                    if hits:
                        rows.append((self._excel_search_row_to_text(row), hits))
                if rows:
                    sheets.append({"sheet": sh.name, "header": header, "rows": rows})
        else:
            return None
        if not sheets:
            return None
        return {"file": fp, "sheets": sheets}

    def _excel_search_run(self):
        fraw = self._excel_filename_entry.get().strip()
        craw = self._excel_search_entry.get().strip()
        if (not fraw and not craw) or self._excel_search_searching:
            return
        file_keywords, file_logic, file_exact = self._excel_parse_keywords(fraw)
        keywords, logic, exact_set = self._excel_parse_keywords(craw)
        if file_keywords is None or keywords is None:
            self._excel_search_status_var.set("❌ 不能同时使用 AND 和 OR")
            return
        if not file_keywords and not keywords:
            self._excel_search_status_var.set("❌ 请输入文件名或内容关键词")
            return
        # 记录搜索词历史（下拉列表选择用）
        if fraw:
            self._record_search_history("excel_search_fn", fraw)
        if craw:
            self._record_search_history("excel_search_kw", craw)
        self._excel_search_keywords = keywords
        self._excel_search_logic = logic
        self._excel_search_searching = True
        self._excel_search_id += 1
        search_id = self._excel_search_id
        self._excel_search_status_var.set(
            f"正在扫描...（文件名: {' '.join(file_keywords) or '全部'} | 内容: {' '.join(keywords) or '无'}）")
        self.update_idletasks()

        import threading

        def worker():
            try:
                root = self._excel_search_search_path
                files = []
                if self._excel_search_recursive.get():
                    for dirpath, _dirs, filenames in os.walk(root):
                        if self._excel_search_id != search_id:
                            return
                        for fn in filenames:
                            if fn.lower().endswith((".xlsx", ".xls")):
                                files.append(os.path.join(dirpath, fn))
                else:
                    try:
                        names = os.listdir(root)
                    except OSError:
                        names = []
                    for fn in names:
                        if fn.lower().endswith((".xlsx", ".xls")):
                            p = os.path.join(root, fn)
                            if os.path.isfile(p):
                                files.append(p)
                if file_keywords:
                    files = [f for f in files if self._excel_filename_matches(
                        os.path.basename(f), file_keywords, file_logic, file_exact)]

                def progress(idx, t, sid):
                    if self._excel_search_id == sid and self._excel_search_searching:
                        self._excel_search_status_var.set(f"扫描中... {idx}/{t}")

                results = []
                total = len(files)
                for i, fp in enumerate(files):
                    if self._excel_search_id != search_id:
                        return
                    if keywords:
                        try:
                            fr = self._excel_search_one_file(fp, keywords, logic, exact_set)
                        except Exception:
                            fr = None
                        if fr:
                            results.append(fr)
                    else:
                        # 只按文件名搜索，列出文件名（不搜内容）
                        results.append({"file": fp, "sheets": []})
                    if i % 5 == 0 or i == total - 1:
                        self.after(0, lambda idx=i + 1, t=total, sid=search_id: progress(idx, t, sid))
                self.after(0, lambda: self._excel_search_render(results, search_id))
            except Exception as e:
                def rescue(sid=search_id):
                    if self._excel_search_id == sid:
                        self._excel_search_searching = False
                        self._excel_search_status_var.set(f"搜索异常: {e}")
                self.after(0, rescue)

        threading.Thread(target=worker, daemon=True).start()

    def _excel_search_render(self, results, search_id):
        if self._excel_search_id != search_id or not self._excel_search_searching:
            return
        self._excel_search_searching = False
        self._excel_search_results = results
        keywords = self._excel_search_keywords
        logic = self._excel_search_logic
        colors = self._EXCEL_KW_COLORS
        kw_color = {kw: colors[i % len(colors)] for i, kw in enumerate(keywords)}
        self._excel_search_kw_idx = {kw: i for i, kw in enumerate(keywords)}

        t = self._excel_result_text
        t.config(state=tk.NORMAL)
        t.delete("1.0", tk.END)
        # 关键词字级高亮 tag（颜色 + 加粗）
        for i, kw in enumerate(keywords):
            t.tag_configure(f"kw{i}", foreground=kw_color[kw],
                            font=("Microsoft YaHei UI", 9, "bold"))

        total_rows = 0
        for fr in results:
            fname = os.path.basename(fr["file"])
            t.insert(tk.END, f"📁 {fname}\n", "file")
            for sh in fr["sheets"]:
                t.insert(tk.END, f"  📄 {sh['sheet']}\n", "sheet")
                t.insert(tk.END, f"    🏷 表头: {sh['header']}\n", "header")
                # ── 按关键词分组：先第一个关键词的匹配行，再第二个… ──
                if logic == "AND":
                    groups = [(" AND ".join(keywords), keywords,
                               [rt for rt, _ in sh["rows"]])]
                else:
                    buckets = {kw: [] for kw in keywords}
                    for rt, hits in sh["rows"]:
                        first = next((kw for kw in keywords if kw in hits), None)
                        if first is not None:
                            buckets[first].append(rt)
                    groups = [(kw, [kw], buckets[kw]) for kw in keywords if buckets[kw]]
                for gtitle, gkws, grows in groups:
                    gidx = self._excel_search_kw_idx.get(gkws[0], 0)
                    gtag = f"kw{gidx}"
                    t.insert(tk.END, f"      🔍 {gtitle} ({len(grows)})\n", gtag)
                    for rt in grows:
                        t.insert(tk.END, "        ")
                        self._excel_search_insert_line(t, rt, gkws)
                        t.insert(tk.END, "\n")
                        total_rows += 1
        t.config(state=tk.DISABLED)
        n_files = len(results)
        kw_str = " ".join(keywords)
        self._excel_search_status_var.set(
            f"✅ 匹配 {n_files} 个文件 / {total_rows} 行（关键词: {kw_str}）")

    def _excel_search_insert_line(self, t, line, hits):
        """插入一行，命中关键词做字级高亮（不同颜色；重叠时保留先出现的关键词）。"""
        if not hits:
            t.insert(tk.END, line)
            return
        spans = []
        lower = line.lower()
        for kw in hits:
            i = self._excel_search_kw_idx.get(kw)
            if i is None:
                continue
            kwl = kw.lower()
            start = 0
            while True:
                pos = lower.find(kwl, start)
                if pos < 0:
                    break
                spans.append((pos, pos + len(kw), f"kw{i}"))
                start = pos + len(kw)
        if not spans:
            t.insert(tk.END, line)
            return
        spans.sort(key=lambda s: (s[0], s[1]))
        merged = []
        for sp in spans:
            if merged and sp[0] < merged[-1][1]:
                continue  # 重叠：保留先出现的关键词
            merged.append(sp)
        pos = 0
        for s, e, tag in merged:
            if s > pos:
                t.insert(tk.END, line[pos:s])
            t.insert(tk.END, line[s:e], tag)
            pos = e
        t.insert(tk.END, line[pos:])

    def _excel_search_pick_files(self):
        """弹窗勾选要导出的文件，返回选中的 results 子集（取消返回 None）。"""
        results = self._excel_search_results
        dlg = tk.Toplevel(self)
        dlg.title("选择要导出的文件")
        dlg.configure(bg=PANEL)
        dlg.transient(self)
        dlg.grab_set()
        dlg.geometry("480x400")
        dlg.rowconfigure(0, weight=1)
        dlg.columnconfigure(0, weight=1)

        canvas = tk.Canvas(dlg, bg=PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(dlg, orient=tk.VERTICAL, command=canvas.yview)
        inner = tk.Frame(canvas, bg=PANEL)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        vsb.grid(row=0, column=1, sticky="ns", pady=8)

        vars_ = []
        for fr in results:
            v = tk.BooleanVar(value=True)
            vars_.append(v)
            fname = os.path.basename(fr["file"])
            nrows = sum(len(sh["rows"]) for sh in fr["sheets"])
            tk.Checkbutton(inner, text=f"{fname}  ({nrows} 行)",
                           variable=v, bg=PANEL, fg=TEXT, selectcolor=ENTRY,
                           activebackground=PANEL, activeforeground=TEXT,
                           anchor="w", font=("Microsoft YaHei UI", 9)).pack(fill=tk.X, pady=1)

        result = {"targets": None}

        def confirm():
            result["targets"] = [fr for fr, v in zip(results, vars_) if v.get()]
            dlg.destroy()

        def cancel():
            dlg.destroy()

        btn_row = tk.Frame(dlg, bg=PANEL)
        btn_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=8)
        RoundedButton(btn_row, text="全选", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=10, cursor="hand2",
                      command=lambda: [v.set(True) for v in vars_]).pack(side=tk.LEFT, padx=8)
        RoundedButton(btn_row, text="全不选", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=10, cursor="hand2",
                      command=lambda: [v.set(False) for v in vars_]).pack(side=tk.LEFT)
        RoundedButton(btn_row, text="取消", bg=ENTRY, fg=TEXT, relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=10, cursor="hand2",
                      command=cancel).pack(side=tk.RIGHT, padx=8)
        RoundedButton(btn_row, text="确定导出", bg=ACCENT, fg="white", relief=tk.FLAT,
                      font=("Microsoft YaHei UI", 9), padx=10, cursor="hand2",
                      command=confirm).pack(side=tk.RIGHT)

        self.wait_window(dlg)
        return result["targets"]

    @staticmethod
    def _excel_search_safe_sheet_name(name):
        name = os.path.splitext(name)[0]
        for ch in "[]:*?/\\":
            name = name.replace(ch, "_")
        return name[:31] or "Sheet"

    def _excel_search_export(self, selected_only):
        """导出搜索结果：全部 / 弹窗勾选若干文件，合并到一个 xlsx（每个源文件一个 sheet）。"""
        results = self._excel_search_results
        if not results:
            self._excel_search_status_var.set("没有可导出的结果")
            return
        if selected_only:
            targets = self._excel_search_pick_files()
            if targets is None:  # 用户取消
                return
            if not targets:
                self._excel_search_status_var.set("未选择任何文件")
                return
        else:
            targets = results
        out = filedialog.asksaveasfilename(
            title="导出搜索结果", defaultextension=".xlsx",
            initialfile="Excel搜索结果.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")])
        if not out:
            return
        try:
            from openpyxl import Workbook
            wb = Workbook()
            wb.remove(wb.active)
            used = set()
            for fr in targets:
                fname = os.path.basename(fr["file"])
                base = self._excel_search_safe_sheet_name(fname)[:28]
                title = base
                k = 2
                while title.lower() in used:
                    title = f"{base}_{k}"
                    k += 1
                used.add(title.lower())
                ws = wb.create_sheet(title=title)
                ws.append(["源文件", "源工作表", "匹配关键词", "表头（第一行）", "匹配行"])
                for sh in fr["sheets"]:
                    for rt, hits in sh["rows"]:
                        ws.append([fname, sh["sheet"], " ".join(hits), sh["header"], rt])
                for col, width in zip("ABCDE", [24, 16, 20, 44, 90]):
                    ws.column_dimensions[col].width = width
            wb.save(out)
            self._excel_search_status_var.set(
                f"✅ 已导出 {len(targets)} 个文件的结果 → {out}")
        except Exception as e:
            messagebox.showerror("导出失败", f"无法导出: {e}")

    def _excel_search_stop(self):
        self._excel_search_id += 1
        self._excel_search_searching = False
        self._excel_search_status_var.set("已停止搜索")

    def _excel_search_list_all(self):
        """一键列出搜索路径下所有 Excel 文件名（相对路径，按路径排序）。"""
        root = self._excel_search_search_path
        files = []
        if self._excel_search_recursive.get():
            for dirpath, _dirs, filenames in os.walk(root):
                for fn in filenames:
                    if fn.lower().endswith((".xlsx", ".xls")):
                        files.append(os.path.join(dirpath, fn))
        else:
            try:
                names = os.listdir(root)
            except OSError:
                names = []
            for fn in names:
                if fn.lower().endswith((".xlsx", ".xls")):
                    p = os.path.join(root, fn)
                    if os.path.isfile(p):
                        files.append(p)
        files.sort(key=lambda p: p.lower())
        t = self._excel_result_text
        t.config(state=tk.NORMAL)
        t.delete("1.0", tk.END)
        if not files:
            t.insert(tk.END, "（未找到 Excel 文件）\n", "header")
        else:
            for p in files:
                rel = os.path.relpath(p, root)
                t.insert(tk.END, f"📊 {rel}\n")
        t.config(state=tk.DISABLED)
        self._excel_search_status_var.set(f"共 {len(files)} 个 Excel 文件（路径: {root}）")












# ── 脚本工坊 Text 表格辅助函数 ──────────────────────────────




    def _everything_quick_filter(self, ext_label):
        """点击常用格式快捷按钮（.epub/.pdf/.ppt）：填入ext:语法并直接执行搜索。"""
        ext = ext_label.lstrip(".")
        self._everything_query.set(f"ext:{ext}")
        self._everything_search()

    def _everything_search(self):
        """调用 Everything SDK 搜索文件，将结果填入文件列表。"""
        query = self._everything_query.get().strip()
        if not query:
            return
        self._record_search_history("everything", query)

        if not EverythingSearch.is_available():
            messagebox.showwarning("Everything 未就绪",
                "未找到 Everything SDK (Everything64.dll)。\n"
                "请确认 Everything 已安装。")
            self._everything_info.set("Everything 未就绪")
            return

        self._everything_info.set("正在搜索...")
        self.update_idletasks()

        # 清空本地过滤，避免干扰搜索结果
        self._reader_filter.set("")

        SUPPORTED_EXTS = READER_SUPPORTED_EXTS
        search_query = query
        if "ext:" not in query.lower() and "file:" not in query.lower():
            exts = "|".join(ext.lstrip(".") for ext in SUPPORTED_EXTS)
            search_query = f"{query} ext:{exts}"

        results = EverythingSearch.search(search_query, max_results=2000)

        # 存储为与本地扫描相同的格式 [(full_path, rel_name), ...]
        # 只保留实际存在的支持的文件
        self._reader_pdf_files = []
        for path, name in results:
            if os.path.isfile(path) and any(path.lower().endswith(ext) for ext in SUPPORTED_EXTS):
                self._reader_pdf_files.append((path, name))
        self._reader_pdf_files_all = list(self._reader_pdf_files)

        self._everything_info.set(f"Everything: {len(self._reader_pdf_files)} 个文件")
        self._reader_info.set(f"搜索: {query}")
        self._reader_sort_files()

        # 自动选中第一项并预览
        if self._reader_pdf_files:
            self._reader_lb.selection_set(0)
            self._on_reader_select()

    def _everything_clear(self):
        """清除 Everything 搜索，恢复本地文件夹浏览。"""
        self._everything_query.set("")
        self._everything_info.set("Everything 就绪")
        self._reader_filter.set("")  # 清空本地过滤
        self._reader_scan()  # 恢复本地扫描结果
        if self._reader_pdf_files:
            self._reader_info.set(f"{len(self._reader_pdf_files)} 个文件")
        else:
            self._reader_info.set("0 个文件")
        self._refresh_reader_list()

    def _bind_search_history(self, entry, key, apply_cb):
        """给搜索框绑定历史下拉：点击(ButtonRelease)或按↓弹出。

        用 ButtonRelease 而非 FocusIn：选中记录后焦点回弹到 Entry 且保持聚焦，
        FocusIn 不会再触发（导致「再点不弹」）；ButtonRelease 每次点击都触发，
        且晚于 bind_all 的 <Button-1> 全局关闭，不会刚弹出就被关掉。
        """
        entry.bind(
            "<ButtonRelease-1>",
            lambda e: self._show_search_history(entry, key, apply_cb))
        entry.bind(
            "<Down>",
            lambda e: self._show_search_history(entry, key, apply_cb))

    def _record_search_history(self, key, keyword):
        """记录一次搜索，关键词次数 +1（用于左侧搜索栏历史）。"""
        keyword = (keyword or "").strip()
        if not keyword:
            return
        hist = self._search_histories.setdefault(key, {})
        hist[keyword] = hist.get(keyword, 0) + 1
        # 即时持久化：搜索历史需跨重启保留；搜索是低频动作，全量保存开销可忽略
        try:
            self._save_config()
        except Exception:
            pass

    def _show_search_history(self, entry, key, apply_cb):
        """在搜索框下方弹出历史下拉，按搜索次数降序。"""
        hist = self._search_histories.get(key, {})
        if not hist:
            return
        self._hide_search_history()
        items = sorted(hist.items(), key=lambda kv: (-kv[1], kv[0]))
        items = items[:self._search_history_max]
        win = tk.Toplevel(self)
        win.overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        w = max(entry.winfo_width(), 200)
        h = min(240, 26 * len(items) + 6)
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        win.geometry(f"{w}x{h}+{x}+{y}")
        lb = tk.Listbox(win, bg=CARD, fg=TEXT, selectbackground=ACCENT,
                        selectforeground="white", relief=tk.SOLID, bd=1,
                        font=("Microsoft YaHei UI", 9), highlightthickness=0,
                        activestyle="none")
        lb.pack(fill=tk.BOTH, expand=True)
        self._search_history_items = []
        for kw, cnt in items:
            lb.insert(tk.END, f"{kw}   ({cnt}次)")
            self._search_history_items.append(kw)
        self._search_history_win = win
        self._search_history_lb = lb
        self._search_history_apply = apply_cb

        def _pick(event):
            idx = lb.nearest(event.y)
            if 0 <= idx < len(self._search_history_items):
                kw = self._search_history_items[idx]
                self._hide_search_history()
                apply_cb(kw)

        lb.bind("<Button-1>", _pick)
        lb.bind("<Escape>", lambda e: self._hide_search_history())

    def _hide_search_history(self):
        """关闭搜索历史下拉。"""
        win = getattr(self, "_search_history_win", None)
        if win is not None and win.winfo_exists():
            try:
                win.destroy()
            except Exception:
                pass
        self._search_history_win = None
        self._search_history_lb = None

    def _on_global_click(self, event):
        """全局点击：关闭搜索历史下拉（点击落在下拉内部则忽略）。"""
        win = getattr(self, "_search_history_win", None)
        if win is None or not win.winfo_exists():
            return
        try:
            w = event.widget
            if w is not None and str(w).startswith(str(win)):
                return
        except Exception:
            pass
        self._hide_search_history()

    def _show_reader_dual_menu(self):
        """显示第二、第三PDF阅读会话的打开选项。"""
        menu = tk.Menu(
            self, tearoff=0, bg=PANEL, fg=TEXT,
            activebackground=ACCENT, activeforeground="white",
            font=("Microsoft YaHei UI", 9))
        for slot in (2, 3, 4):
            menu.add_command(
                label=f"阅读区 {slot}：打开当前文件",
                command=lambda s=slot: self._reader_dual_open_same(s))
            menu.add_command(
                label=f"阅读区 {slot}：打开列表选中文件",
                command=lambda s=slot: self._reader_dual_open_selected(s))
            menu.add_command(
                label=f"阅读区 {slot}：选择文件...",
                command=lambda s=slot: self._reader_dual_choose_file(s))
            if self._reader_aux.get(slot, {}).get("visible"):
                menu.add_command(
                    label=f"关闭 阅读区 {slot}",
                    command=lambda s=slot: self._reader_aux_close(s))
            if slot == 3:
                menu.add_separator()
        x = self._reader_dual_btn.winfo_rootx()
        y = self._reader_dual_btn.winfo_rooty() + self._reader_dual_btn.winfo_height()
        menu.post(x, y)

    def _reader_dual_open_same(self, slot=2):
        """复制当前主阅读区的文件到指定阅读区（支持所有文件类型）。"""
        if not self._reader_active_file_path:
            messagebox.showinfo("双开阅读", "请先在主阅读区打开一份文件")
            return
        self._reader_dual_open_path(self._reader_active_file_path, slot)

    def _reader_dual_choose_file(self, slot=2):
        """手动选择一份文件在指定阅读区打开（支持所有支持的文件格式）。"""
        path = filedialog.askopenfilename(
            title=f"选择 PDF {slot} 文件",
            filetypes=[("支持的文件", "*.pdf;*.epub;*.ppt;*.pptx;*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
                       ("PDF文件", "*.pdf"), ("所有文件", "*.*")])
        if path:
            self._reader_dual_open_path(path, slot)

    def _reader_dual_close(self):
        self._reader_aux_close(2)

    def _show_app_theme_menu(self):
        """显示程序全局外观主题菜单。"""
        palette = self._active_app_palette
        menu = tk.Menu(
            self, tearoff=0, bg=palette["PANEL"], fg=palette["TEXT"],
            activebackground=palette["ACCENT"], activeforeground="white",
            selectcolor=palette["CARD"], font=("Microsoft YaHei UI", 9))
        xiaomi_keys = set(getattr(self, "_xiaomi_theme_keys", ()))
        sep_added = False
        for key, theme in self._app_themes.items():
            if key in xiaomi_keys and not sep_added:
                menu.add_separator()
                menu.add_command(label="—— 小米汽车色 ——", state="disabled")
                sep_added = True
            menu.add_radiobutton(
                label=theme["name"], variable=self._app_theme, value=key,
                command=self._apply_app_theme)
        x = self._app_theme_btn.winfo_rootx()
        y = self._app_theme_btn.winfo_rooty() + self._app_theme_btn.winfo_height()
        menu.post(x, y)

    @staticmethod
    def _derive_secondary_colors(palette):
        """根据主题明暗派生弹窗/浮层二级色（硬编码中性色的浅/深两套）。"""
        try:
            r, g, b = _hex_to_rgb(palette["BG"])
            dark = (0.299 * r + 0.587 * g + 0.114 * b) < 128
        except Exception:
            dark = True
        if dark:
            return {"POPUP_BG": "#1a1a1a", "POPUP_CARD": "#2d2d2d",
                    "POPUP_DIM": "#999999"}
        return {"POPUP_BG": "#f5f5f5", "POPUP_CARD": "#ffffff",
                "POPUP_DIM": "#666666"}

    def _apply_app_theme(self, persist=True):
        """动态更新经典Tk控件与ttk控件的全局配色。"""
        theme_key = self._app_theme.get()
        if theme_key not in self._app_themes:
            theme_key = "graphite"
            self._app_theme.set(theme_key)
        old = self._active_app_palette
        new = self._app_themes[theme_key]
        color_map = {
            str(old[key]).lower(): new[key]
            for key in ("BG", "PANEL", "CARD", "ENTRY", "BORDER",
                        "ACCENT", "GREEN", "TEXT", "DIM", "CANVAS")
        }

        def recolor(widget):
            if isinstance(widget, (RoundedButton, SlimScrollbar)):
                widget._refresh_theme(color_map, new)
                return
            try:
                keys = set(widget.keys())
            except Exception:
                keys = set()
            updates = {}
            for option in (
                    "background", "foreground", "activebackground",
                    "activeforeground", "selectcolor", "insertbackground",
                    "highlightbackground", "highlightcolor", "troughcolor"):
                if option not in keys:
                    continue
                try:
                    current = str(widget.cget(option)).lower()
                except Exception:
                    continue
                if current in color_map:
                    updates[option] = color_map[current]
            if updates:
                try:
                    widget.configure(**updates)
                except Exception:
                    pass
            try:
                children = widget.winfo_children()
            except Exception:
                children = ()
            for child in children:
                recolor(child)

        recolor(self)
        globals().update({key: new[key] for key in (
            "BG", "PANEL", "CARD", "ENTRY", "BORDER",
            "ACCENT", "GREEN", "TEXT", "DIM", "CANVAS")})
        # 派生二级色随主题明暗切换（弹窗背景等硬编码中性色跟随主题）
        globals().update(self._derive_secondary_colors(new))
        self._active_app_palette = dict(new)
        self._reader_themes["normal"]["canvas_bg"] = new["CANVAS"]
        self.configure(bg=new["BG"])
        self._app_theme_btn.config(
            text=f"✦ {new['name']}", bg=new["ENTRY"], fg=new["TEXT"])

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background=new["BG"], borderwidth=0)
        style.configure("TEntry", borderwidth=0, relief="flat")
        style.configure("TCombobox", borderwidth=0)
        style.configure(
            "TNotebook.Tab", background=new["PANEL"], foreground=new["TEXT"],
            padding=(12, 6))
        style.map(
            "TNotebook.Tab",
            background=[("selected", new["ACCENT"])],
            foreground=[("selected", "white")])
        style.configure(
            "Treeview", background=new["CARD"], foreground=new["TEXT"],
            fieldbackground=new["CARD"], borderwidth=0)
        style.configure(
            "Treeview.Heading", background=new["PANEL"], foreground=new["TEXT"])
        style.map("Treeview", background=[("selected", new["ACCENT"])])
        style.configure(
            "TScrollbar", background=new["PANEL"], troughcolor=new["BG"],
            arrowcolor=new["TEXT"])
        style.configure(
            "TCombobox", fieldbackground=new["ENTRY"], background=new["ENTRY"],
            foreground=new["TEXT"])
        if persist and hasattr(self, "_config_file"):
            self._save_config()

    def _toggle_reader_linked_search(self):
        """切换联动搜索开关：ON/OFF"""
        new_val = not self._reader_linked_search.get()
        self._reader_linked_search.set(new_val)
        colors = {True: GREEN, False: "#454b63"}
        labels = {True: "🔗 联动搜索:开", False: "🔗 联动搜索:关"}
        self._reader_linked_btn.config(text=labels[new_val], bg=colors[new_val])
        self._save_config()

    def _reader_update_source_slot(self, slot, word, result):
        """更新触发搜索的源 slot 的 UI（高亮 + 匹配面板）。"""
        if not result:
            return

        if slot == 1:
            if result["type"] == "pdf" and self._reader_pdf:
                self._reader_highlight_keywords = [word]
                self._reader_highlight_exact = True
                self._reader_preview_search.set(word)
                occurrences = result["matches"]
                matched_pages = sorted({p for p, _ in occurrences})
                self._reader_match_pages = matched_pages
                self._reader_match_counts = {}
                for p, _ in occurrences:
                    self._reader_match_counts[p] = self._reader_match_counts.get(p, 0) + 1
                self._reader_match_lb.delete(0, tk.END)
                for p in matched_pages:
                    self._reader_match_lb.insert(tk.END, f"  {p + 1}")
                self._reader_match_count_var.set(
                    self._fmt_match_count(len(matched_pages), len(occurrences)))
                self._reader_refresh_main_match_flow()
                self._reader_redraw_highlight_overlay()
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(1)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    self._reader_preview_search.set(word)
                    self._reader_highlight_keywords = [word]
                    self._reader_match_count_var.set(self._fmt_match_count(0, count))
        else:
            state = self._reader_aux.get(slot)
            if not state:
                return
            state["keywords"] = [word]
            state["exact"] = True
            state["search_var"].set(word)
            if result["type"] == "pdf":
                occurrences = result["matches"]
                pages = sorted({p for p, _ in occurrences})
                counts = {}
                for p, _ in occurrences:
                    counts[p] = counts.get(p, 0) + 1
                state["match_pages"] = pages
                state["match_counts"] = counts
                state["match_lb"].delete(0, tk.END)
                for p in pages:
                    state["match_lb"].insert(tk.END, f"第{p + 1}页  {counts[p]}处")
                state["match_count_var"].set(self._fmt_match_count(len(pages), len(occurrences)))
                self._reader_refresh_aux_match_flow(slot)
                self._reader_aux_redraw_highlight_overlay(slot)
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(slot)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    state["match_count_var"].set(self._fmt_match_count(0, count))

    def _reader_update_other_slot_highlight(self, slot, word, result, exact=True):
        """更新其他 slot 的高亮并触发重渲染，使关键词高亮在画布上可见。

        exact=False 用于预览搜索栏的模糊子串传播——让 PDF2/3 的高亮和 PDF1 的
        子串高亮保持一致（点击文本的联动搜索仍走 exact=True 整词匹配）。
        """
        if not result:
            return

        if slot == 1:
            if result["type"] == "pdf" and self._reader_pdf:
                self._reader_highlight_keywords = [word]
                self._reader_highlight_exact = exact
                self._reader_preview_search.set(word)
                occurrences = result["matches"]
                matched_pages = sorted({p for p, _ in occurrences})
                self._reader_match_pages = matched_pages
                self._reader_match_counts = {}
                for p, _ in occurrences:
                    self._reader_match_counts[p] = self._reader_match_counts.get(p, 0) + 1
                self._reader_match_lb.delete(0, tk.END)
                for p in matched_pages:
                    self._reader_match_lb.insert(tk.END, f"  {p + 1}")
                self._reader_match_count_var.set(
                    self._fmt_match_count(len(matched_pages), len(occurrences)))
                self._reader_refresh_main_match_flow()
                # 触发重渲染使关键词高亮可见
                self._reader_redraw_highlight_overlay()
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(1)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    self._reader_highlight_keywords = [word]
                    self._reader_preview_search.set(word)
                    self._reader_match_count_var.set(self._fmt_match_count(0, count))
        else:
            state = self._reader_aux.get(slot)
            if not state:
                return
            state["keywords"] = [word]
            state["exact"] = exact
            state["search_var"].set(word)
            if result["type"] == "pdf":
                occurrences = result["matches"]
                pages = sorted({p for p, _ in occurrences})
                counts = {}
                for p, _ in occurrences:
                    counts[p] = counts.get(p, 0) + 1
                state["match_pages"] = pages
                state["match_counts"] = counts
                state["match_lb"].delete(0, tk.END)
                for p in pages:
                    state["match_lb"].insert(tk.END, f"第{p + 1}页  {counts[p]}处")
                state["match_count_var"].set(self._fmt_match_count(len(pages), len(occurrences)))
                self._reader_refresh_aux_match_flow(slot)
                # 触发重渲染使关键词高亮在画布上可见
                self._reader_aux_redraw_highlight_overlay(slot)
            elif result["type"] in ("epub", "html", "text"):
                text_widget = self._reader_get_text_widget_for_slot(slot)
                if text_widget:
                    count = self._reader_highlight_in_text_widget(text_widget, word)
                    state["match_count_var"].set(self._fmt_match_count(0, count))

    def _reader_linked_search_all(self, word, source_slot=1, quiet=False):
        """联动搜索：在全部已打开的阅读区中搜索关键词，展示统一结果。

        quiet=True 时（如 IPC 后台触发）找不到匹配只静默返回，不弹窗打断用户。
        """
        results = {}  # {slot: {"type": ..., "path": ..., "matches": ...}}

        # 遍历 3 个 slot，分别搜索
        for slot in (1, 2, 3, 4):
            path, file_type, matches = self._reader_search_in_slot(slot, word)
            if path and matches is not None:
                results[slot] = {
                    "type": file_type,
                    "path": path,
                    "matches": matches,
                }

        if not results:
            if not quiet:
                messagebox.showinfo("联动搜索", f"所有已打开文件中未找到: \"{word}\"")
            return

        # 更新源 slot 的高亮和匹配面板
        self._reader_update_source_slot(source_slot, word, results.get(source_slot))

        # 更新其他 slot 的高亮
        for slot in results:
            if slot != source_slot:
                self._reader_update_other_slot_highlight(slot, word, results[slot])

    def _reader_show_linked_results(self, word, results, source_slot):
        """弹窗展示联动搜索结果（所有 slot 的命中汇总）。"""
        win = tk.Toplevel(self)
        win.title(f"🔗 联动搜索: \"{word}\"")
        win.configure(bg=BG)
        # 使用较宽的窗口以便查看结果
        win_w = min(720, self.winfo_screenwidth() - 100)
        win_h = min(560, self.winfo_screenheight() - 100)
        win.geometry(f"{win_w}x{win_h}"
                     f"+{max(0, self.winfo_rootx() + 60)}+{max(0, self.winfo_rooty() + 50)}")
        win.minsize(400, 300)

        # 标题
        header = tk.Frame(win, bg=PANEL)
        header.pack(fill=tk.X)
        tk.Label(header, text=f"联动搜索: \"{word}\" — 共 {len(results)} 个文件有匹配",
                 bg=PANEL, fg=TEXT, anchor="w", padx=12, pady=9,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(
                     side=tk.LEFT, fill=tk.X, expand=True)
        RoundedButton(header, text="✕", bg="#454b63", fg="white",
                  relief=tk.FLAT, font=("Microsoft YaHei UI", 8, "bold"),
                  padx=6, cursor="hand2",
                  command=win.destroy).pack(side=tk.RIGHT, padx=10, pady=7)

        # 结果列表（可滚动的 frame）
        cv = tk.Canvas(win, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(win, command=cv.yview)
        result_frame = tk.Frame(cv, bg=BG)
        result_frame.bind("<Configure>", lambda e: cv.configure(
            scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=result_frame, anchor="nw")
        cv.config(yscrollcommand=scrollbar.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6, pady=6)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮在 result_frame 上也滚动
        def _wheel(event):
            cv.yview_scroll(int(-event.delta / 120), "units")
        cv.bind("<MouseWheel>", _wheel)
        result_frame.bind("<MouseWheel>", _wheel)

        for slot in (1, 2, 3, 4):
            r = results.get(slot)
            if not r:
                continue
            display_name = os.path.basename(r["path"])
            typ = r["type"]
            matches = r["matches"]

            # 每个 slot 一个卡片
            card = tk.Frame(result_frame, bg=CARD, highlightbackground=BORDER,
                            highlightthickness=1)
            card.pack(fill=tk.X, padx=4, pady=3)

            # 标题行
            title_row = tk.Frame(card, bg=CARD)
            title_row.pack(fill=tk.X, padx=8, pady=(6, 2))
            slot_label = {1: "📄 PDF1", 2: "📄 PDF2", 3: "📄 PDF3", 4: "📄 PDF4"}[slot]
            is_source = " ← 来源" if slot == source_slot else ""
            tk.Label(title_row, text=f"{slot_label}: {display_name}{is_source}",
                     bg=CARD, fg=ACCENT if slot == source_slot else TEXT,
                     font=("Microsoft YaHei UI", 9, "bold"), anchor="w").pack(
                         side=tk.LEFT, fill=tk.X, expand=True)

            # 匹配信息
            if typ == "pdf":
                if matches:
                    pages = sorted({p for p, _ in matches})
                    count = len(matches)
                    page_str_list = [str(p + 1) for p in pages[:20]]
                    page_str = ", ".join(page_str_list)
                    if len(pages) > 20:
                        page_str += f" ... 等{len(pages)}页"
                    info = f"匹配 {count} 处，分布在 {len(pages)} 页: {page_str}"
                else:
                    info = "无匹配"
            elif typ in ("epub", "html", "text", "ppt"):
                if matches:
                    count = len(matches)
                    line_list = [str(ln) for ln in matches[:15]]
                    line_str = ", ".join(line_list)
                    if len(matches) > 15:
                        line_str += f" ... 等{len(matches)}行"
                    info = f"匹配 {count} 行: {line_str}"
                else:
                    info = "无匹配"
            else:
                info = f"匹配 {len(matches) if matches else 0} 处"

            info_lbl = tk.Label(card, text=info, bg=CARD, fg=DIM,
                         font=("Microsoft YaHei UI", 8), anchor="w",
                         padx=8, pady=(0, 6),
                         wraplength=win_w - 60, justify=tk.LEFT)
            info_lbl.pack(fill=tk.X)

            # 跳转按钮（非源 slot + 有匹配）
            if slot != source_slot and matches:
                btn_row = tk.Frame(card, bg=CARD)
                btn_row.pack(fill=tk.X, padx=8, pady=(0, 6))
                RoundedButton(btn_row, text=f"跳转到 PDF{slot}", bg=ACCENT,
                          fg="white", relief=tk.FLAT, cursor="hand2",
                          font=("Microsoft YaHei UI", 8), padx=8,
                          command=lambda s=slot, r=r: self._reader_linked_jump_to(s, r)).pack(
                              side=tk.LEFT)

        # 关闭按钮
        RoundedButton(win, text="关闭", bg="#454b63", fg="white",
                  relief=tk.FLAT, font=("Microsoft YaHei UI", 9), padx=14, pady=6,
                  cursor="hand2", command=win.destroy).pack(pady=10)

    def _reader_linked_jump_to(self, slot, result):
        """从联动搜索结果弹窗跳转到指定 slot 的第一个匹配位置。"""
        self._reader_set_active_slot(slot)
        if result["type"] == "pdf":
            first_page = sorted({p for p, _ in result["matches"]})[0]
            if slot == 1:
                self._reader_page = first_page
                self._reader_scroll_to_page(first_page)
                self._reader_update_active_page_info()
                self._reader_highlight_thumb()
            else:
                state = self._reader_aux[slot]
                state["page"] = first_page
                self._reader_aux_render(slot, fit_width=False, target_page=first_page)

    def _toggle_match_col(self):
        """折叠/展开三个PDF预览区(PDF1/2/3)下方的匹配页码块（同一按钮统一控制）。"""
        if self._reader_match_visible:
            self._reader_match_frame.grid_remove()
            for slot in (2, 3, 4):
                st = self._reader_aux.get(slot)
                if st:
                    st["match_frame"].grid_remove()
            self._reader_match_nav_btn.config(bg="#454b63")
            self._reader_match_visible = False
        else:
            self._reader_match_frame.grid()
            for slot in (2, 3, 4):
                st = self._reader_aux.get(slot)
                if st:
                    st["match_frame"].grid()
            self._reader_match_nav_btn.config(bg=ACCENT)
            self._reader_match_visible = True

    def _reader_search_and_show(self, word):
        """搜索整个PDF中所有包含该词的页面，裁剪并顺序显示"""
        if not self._reader_pdf:
            return

        # 尝试使用缓存搜索
        occurrences = self._reader_search_cached(word)
        if occurrences is None:
            # 缓存未就绪，回退到逐页扫描（首次搜索时缓存可能在后台构建中）
            word_lower = word.lower()
            occurrences = []
            for i in range(self._reader_pdf.page_count):
                for w in self._reader_pdf[i].get_text("words"):
                    w_clean = self._clean_word(w[4])
                    if w_clean and w_clean.lower() == word_lower:
                        occurrences.append((i, fitz.Rect(w[0], w[1], w[2], w[3])))
        else:
            # 缓存搜索不需要额外处理
            pass

        if not occurrences:
            messagebox.showinfo("搜索结果", f"整个文档未找到: \"{word}\"")
            return

        # 点击预览文字后，同步高亮PDF内该文字的全部出现位置。
        self._reader_highlight_keywords = [word]
        self._reader_highlight_exact = True
        self._reader_preview_search.set(word)
        matched_pages = sorted({page_idx for page_idx, _ in occurrences})
        self._reader_match_pages = matched_pages
        self._reader_match_counts = {}
        for page_idx, _ in occurrences:
            self._reader_match_counts[page_idx] = (
                self._reader_match_counts.get(page_idx, 0) + 1)
        self._reader_match_lb.delete(0, tk.END)
        for page_idx in matched_pages:
            self._reader_match_lb.insert(tk.END, f"  {page_idx + 1}")
        self._reader_match_count_var.set(
            self._fmt_match_count(len(matched_pages), len(occurrences)))
        self._reader_refresh_main_match_flow()
        if self._reader_page in matched_pages:
            selected = matched_pages.index(self._reader_page)
            self._reader_match_lb.selection_set(selected)
            self._reader_match_lb.see(selected)
        self._reader_redraw_highlight_overlay()
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()




class EverythingSearch:
    """封装 Everything 搜索 SDK，支持全局文件索引搜索。"""
    _dll = None

    @classmethod
    def _init_dll(cls):
        if cls._dll is not None:
            return True
        candidates = [
            r"C:\Program Files\Quicker\Everything64.dll",
            r"C:\Program Files\Everything\Everything64.dll",
            r"C:\Program Files (x86)\Everything\Everything.dll",
            r"C:\Program Files\Everything\Everything.dll",
        ]
        for path in candidates:
            if os.path.exists(path):
                try:
                    cls._dll = ctypes.WinDLL(path)
                    cls._dll.Everything_SetSearchW.argtypes = [wintypes.LPCWSTR]
                    cls._dll.Everything_SetSearchW.restype = None
                    cls._dll.Everything_QueryW.argtypes = [wintypes.BOOL]
                    cls._dll.Everything_QueryW.restype = wintypes.BOOL
                    cls._dll.Everything_GetNumResults.argtypes = []
                    cls._dll.Everything_GetNumResults.restype = wintypes.DWORD
                    cls._dll.Everything_GetResultFullPathNameW.argtypes = [
                        wintypes.DWORD, wintypes.LPWSTR, wintypes.DWORD
                    ]
                    cls._dll.Everything_GetResultFullPathNameW.restype = wintypes.DWORD
                    return True
                except Exception:
                    continue
        return False

    @classmethod
    def search(cls, query: str, max_results: int = 2000) -> list:
        """
        搜索文件，返回 [(full_path, filename), ...]。
        支持多关键词空格分隔（Everything 原生语法）。
        """
        if not cls._init_dll():
            return []
        if not query or not query.strip():
            return []
        cls._dll.Everything_SetSearchW(query.strip())
        if not cls._dll.Everything_QueryW(True):
            return []
        num = cls._dll.Everything_GetNumResults()
        buf = ctypes.create_unicode_buffer(520)
        results = []
        for i in range(min(num, max_results)):
            cls._dll.Everything_GetResultFullPathNameW(i, buf, 520)
            path = buf.value
            if path:
                results.append((path, os.path.basename(path)))
        return results

    @classmethod
    def is_available(cls) -> bool:
        return cls._init_dll()

if __name__ == "__main__":
    import traceback
    try:
        app = ExcelPdfSync()
        app.protocol("WM_DELETE_WINDOW", app.on_close)
        app.mainloop()
    except Exception as _launch_err:
        # 启动失败时弹出错误对话框，不再闪退——把这段错误信息截图发给开发者
        import tkinter as _tk
        import tkinter.messagebox as _mb
        _err_detail = traceback.format_exc()
        try:
            _root = _tk.Tk()
            _root.withdraw()
            _mb.showerror(
                "程序启动失败",
                f"启动时遇到错误，请把以下信息截图反馈：\n\n{_err_detail[:2000]}"
            )
            _root.destroy()
        except Exception:
            # 连tkinter都用不了时，写入桌面日志文件
            import os as _os
            log_path = _os.path.join(_os.path.expanduser("~"), "Desktop",
                                      "excel_pdf_sync_crash.txt")
            with open(log_path, "w", encoding="utf-8") as _f:
                _f.write(_err_detail)
            print(f"崩溃日志已写入: {log_path}")
