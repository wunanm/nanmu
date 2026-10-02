# -*- coding: utf-8 -*-
"""
楠木 Dock —— 仿 BitDock / macOS 风格的桌面 Dock 栏

功能：
  - 鼠标悬停时图标平滑放大（macOS 鱼眼效果），显示名称气泡
  - 左键：启动程序；程序已有窗口时切换过去（再点一次最小化，多窗口时轮流切换）
  - 中键：总是启动新实例
  - 有窗口的程序下方显示小圆点；没固定的正在运行的程序显示在分隔线右侧
  - 把程序 / 快捷方式 / 文件夹 / 网址拖到 Dock 上即可添加
  - 按住图标拖动可调整顺序，拖出 Dock 上方松开即移除
  - 右键菜单：固定、关闭窗口、管理员运行、打开所在位置、重命名、移除、设置
  - 设置：图标大小、放大倍数、自动隐藏、全屏时隐藏、开机自启、
          隐藏 Windows 任务栏、隐藏桌面图标（退出 Dock 时自动恢复，Ctrl+Alt+D 一键切换）
  - 皮肤中心：11 款内置皮肤 + 自定义颜色 / 背景图片
  - 看板娘：一张图片站在 Dock 旁边，点她会说话，开关软件时会提醒
  - 托盘图标，可隐藏/显示 Dock

运行：双击本文件（pythonw），或 `python nanmu_dock.pyw` 查看调试输出。
`nanmu_dock.pyw --restore`：恢复任务栏和桌面图标（程序被强制结束后用）。
配置保存在同目录下的 dock_config.json。
"""
import os
import sys
import json
import math
import time
import random
import shutil
import ctypes
import threading
import traceback
import configparser
import subprocess
import winreg
from ctypes import wintypes

from PySide6.QtCore import Qt, QTimer, QRectF, QPointF, QSize, QFileInfo, QThread, Signal
from PySide6.QtGui import (QPainter, QColor, QPainterPath, QPixmap, QImage, QIcon, QFont, QKeySequence,
                           QFontMetricsF, QCursor, QPen, QBrush, QLinearGradient, QRadialGradient, QTransform,
                           QActionGroup, QGuiApplication)
from PySide6.QtWidgets import (QApplication, QWidget, QMenu, QFileIconProvider, QSystemTrayIcon,
                               QInputDialog, QFileDialog, QStyle, QMessageBox, QDialog, QVBoxLayout,
                               QHBoxLayout, QGridLayout, QLabel, QKeySequenceEdit, QLineEdit,
                               QSlider, QPushButton, QColorDialog, QComboBox, QCheckBox)
from PySide6.QtNetwork import QLocalServer, QLocalSocket

import pythoncom
import win32api
import win32con
import win32gui
import win32process
import win32com.client

APP_NAME = "楠木 Dock"
APP_VERSION = "1.0.0"
INSTANCE_KEY = "NanmuDock_SingleInstance"
APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
CONFIG_PATH = os.path.join(APP_DIR, "dock_config.json")
LINKS_DIR = os.path.join(APP_DIR, "links")
LOG_PATH = os.path.join(APP_DIR, "dock_error.log")
# 随程序发布的资源；打包成 exe 后在 PyInstaller 的解压目录里
ASSETS_DIR = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "assets")
BUILTIN_MASCOT = os.path.join(ASSETS_DIR, "mascot_default.png")

DEFAULT_CONFIG = {
    "icon_size": 52,          # 图标基础大小（逻辑像素）
    "magnify": 1.8,           # 最大放大倍数，1.0 = 关闭放大
    "spread": 2.6,            # 放大影响范围（以图标个数计）
    "skin": "dark",           # 见 SKINS，或 custom_color / custom_image
    "skin_color": "",         # 自定义颜色皮肤的主色
    "skin_image": "",         # 图片皮肤的图片
    "skin_opacity": 1.0,      # 背景不透明度 0.2~1
    "mascot_image": "",       # 看板娘图片（建议透明背景 PNG），站在 Dock 旁边
    "mascot_side": "right",   # left / right
    "mascot_size": 2.6,       # 高度 = 图标大小 × 这个倍数
    "mascot_lines": ["……", "嗯？", "有什么事吗？", "今天也辛苦了。", "（伸了个懒腰）"],
    "mascot_announce": True,  # 打开软件时看板娘报一声
    "mascot_open_lines": ["打开了 {name}。", "……{name}，开好了。", "{name}……要用吗？", "嗯，{name}。"],
    "mascot_announce_close": True,   # 关掉软件时也说一句
    "mascot_close_lines": ["{name}……关掉了。", "{name}，已经关了。", "……{name}，辛苦了。", "嗯，{name}，拜拜。"],
    "auto_hide": False,
    "hide_on_fullscreen": True,
    "show_running": True,     # 显示没固定的正在运行的程序
    "hide_taskbar": False,
    "hide_desktop_icons": False,
    "hotkeys": {},            # 全局快捷键 {功能: 组合键}，见 HOTKEY_ACTIONS，留空 = 不用
    "items": [],
}

MARGIN = 6            # Dock 距屏幕底部
BOUNCE_TIME = 0.9     # 启动时弹跳动画时长（秒）
MASCOT_HOP = 0.45     # 看板娘被点时跳一下的时长
WS_EX_NOACTIVATE = 0x08000000
WS_EX_APPWINDOW = 0x00040000

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
dwmapi = ctypes.windll.dwmapi
user32.PrivateExtractIconsW.argtypes = [wintypes.LPCWSTR, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                        ctypes.POINTER(wintypes.HICON), ctypes.POINTER(wintypes.UINT),
                                        wintypes.UINT, wintypes.UINT]
user32.PrivateExtractIconsW.restype = wintypes.UINT
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
                                                ctypes.POINTER(wintypes.DWORD)]
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]


def log_error(text):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(time.strftime("[%Y-%m-%d %H:%M:%S] ") + text + "\n")
    except OSError:
        pass


def _excepthook(tp, val, tb):
    # pythonw 下没有控制台，未捕获异常写进日志，避免“静默卡死”
    log_error("".join(traceback.format_exception(tp, val, tb)))
    sys.__excepthook__(tp, val, tb)


sys.excepthook = _excepthook


# ---------------------------------------------------------------- 快捷方式 / 图标

_wsh = None


def wsh():
    global _wsh
    if _wsh is None:
        _wsh = win32com.client.Dispatch("WScript.Shell")
    return _wsh


def is_web_url(path):
    return "://" in path and not os.path.exists(path)


def read_lnk(path):
    """返回 (目标路径, 图标文件, 图标序号)"""
    try:
        s = wsh().CreateShortcut(path)
        target = os.path.expandvars(s.TargetPath or "")
        icon_file, _, idx = (s.IconLocation or "").rpartition(",")
        try:
            idx = int(idx)
        except ValueError:
            idx = 0
        return target, os.path.expandvars(icon_file), idx
    except Exception:
        return "", "", 0


def read_url_file(path):
    """读取 .url 网页快捷方式，返回 (图标文件, 图标序号)"""
    try:
        raw = open(path, "rb").read()
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("mbcs", errors="ignore")
        cp = configparser.ConfigParser(interpolation=None, strict=False)
        cp.read_string(text)
        sec = cp["InternetShortcut"]
        return os.path.expandvars(sec.get("iconfile", "")), int(sec.get("iconindex", "0") or 0)
    except Exception:
        return "", 0


def extract_icon(path, idx=0, size=256):
    """用 PrivateExtractIcons 直接提取指定尺寸的高清图标"""
    if not path or not os.path.exists(path):
        return None
    h = wintypes.HICON()
    icon_id = wintypes.UINT()
    n = user32.PrivateExtractIconsW(path, idx, size, size, ctypes.byref(h), ctypes.byref(icon_id), 1, 0)
    if n == 0xFFFFFFFF or n < 1 or not h.value:
        return None
    try:
        img = QImage.fromHICON(h.value)
    finally:
        user32.DestroyIcon(h)
    return None if img.isNull() else img


def window_icon(hwnd):
    """窗口自己的图标（UWP 等没有 exe 图标的程序用），分辨率较低"""
    h = 0
    for kind in (1, 2, 0):      # ICON_BIG, ICON_SMALL2, ICON_SMALL
        try:
            h = win32gui.SendMessageTimeout(hwnd, win32con.WM_GETICON, kind, 0, win32con.SMTO_ABORTIFHUNG, 100)[1]
        except Exception:
            h = 0
        if h:
            break
    if not h:
        try:
            h = win32gui.GetClassLong(hwnd, -14)    # GCL_HICON
        except Exception:
            h = 0
    if not h:
        return None
    img = QImage.fromHICON(h)
    return None if img.isNull() else img


_icon_provider = None


def shell_icon(path):
    global _icon_provider
    if _icon_provider is None:
        _icon_provider = QFileIconProvider()
    pm = _icon_provider.icon(QFileInfo(path)).pixmap(QSize(256, 256))
    return None if pm.isNull() else pm.toImage()


def load_icon_image(path):
    src, idx = path, 0
    ext = os.path.splitext(path)[1].lower()
    if ext == ".lnk":
        target, icon_file, i = read_lnk(path)
        if icon_file and os.path.exists(icon_file):
            src, idx = icon_file, i
        elif target:
            src = target
    elif ext == ".url":
        icon_file, i = read_url_file(path)
        if icon_file and os.path.exists(icon_file):
            src, idx = icon_file, i

    if src and os.path.exists(src):
        if os.path.splitext(src)[1].lower() in (".exe", ".dll", ".ico", ".icl", ".cpl", ".scr"):
            img = extract_icon(src, idx)
            if img is not None:
                return img
        img = shell_icon(src)
        if img is not None:
            return img
    if os.path.exists(path):
        img = shell_icon(path)
        if img is not None:
            return img
    style = QApplication.style()
    sp = QStyle.SP_DriveNetIcon if is_web_url(path) else QStyle.SP_FileIcon
    return style.standardIcon(sp).pixmap(QSize(256, 256)).toImage()


def exe_description(path):
    try:
        lang, cp = win32api.GetFileVersionInfo(path, "\\VarFileInfo\\Translation")[0]
        desc = win32api.GetFileVersionInfo(path, "\\StringFileInfo\\%04x%04x\\FileDescription" % (lang, cp))
        return (desc or "").strip() or None
    except Exception:
        return None


def display_name(path):
    if is_web_url(path):
        return path.split("://", 1)[1].split("/", 1)[0] or path
    p = path.rstrip("\\/")
    if len(p) <= 2 and p.endswith(":"):
        return p + "\\"
    stem, ext = os.path.splitext(os.path.basename(p))
    if ext.lower() == ".exe":
        return exe_description(path) or stem
    if ext.lower() in (".lnk", ".url"):
        return stem
    return os.path.basename(p) or p


def norm(path):
    # realpath：软件被“搬家”工具挪走后，原位置会留一个目录链接，运行时 Windows 报的是真实位置
    try:
        return os.path.normcase(os.path.realpath(path))
    except OSError:
        return os.path.normcase(os.path.abspath(path))


def _env_dir(var, *sub):
    base = os.environ.get(var)
    return norm(os.path.join(base, *sub)) if base else None


# 很多程序共用的目录：不能因为“装在同一个文件夹”就认定是同一个软件
SHARED_DIRS = {d for d in (
    _env_dir("WINDIR"), _env_dir("WINDIR", "System32"), _env_dir("WINDIR", "SysWOW64"),
    _env_dir("ProgramFiles"), _env_dir("ProgramFiles(x86)"), _env_dir("ProgramData"),
    _env_dir("LOCALAPPDATA"), _env_dir("LOCALAPPDATA", "Programs"), _env_dir("APPDATA"),
) if d}
# 这些父进程只是“帮忙启动”，不代表子进程属于它
LAUNCHER_PARENTS = {"explorer.exe", "svchost.exe", "services.exe", "sihost.exe", "runtimebroker.exe",
                    "cmd.exe", "powershell.exe", "pwsh.exe", "conhost.exe", "python.exe", "pythonw.exe",
                    "nanmudock.exe", "userinit.exe", "wininit.exe", "winlogon.exe"}


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_size_t), ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD), ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD), ("szExeFile", ctypes.c_wchar * 260)]


kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
kernel32.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
kernel32.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]


def parent_map():
    """{pid: (父 pid, 进程名小写)}，只在出现新程序时调用一次"""
    snap = kernel32.CreateToolhelp32Snapshot(0x2, 0)       # TH32CS_SNAPPROCESS
    result = {}
    if not snap or snap == wintypes.HANDLE(-1).value:
        return result
    try:
        e = PROCESSENTRY32W()
        e.dwSize = ctypes.sizeof(e)
        ok = kernel32.Process32FirstW(snap, ctypes.byref(e))
        while ok:
            result[e.th32ProcessID] = (e.th32ParentProcessID, e.szExeFile.lower())
            ok = kernel32.Process32NextW(snap, ctypes.byref(e))
    finally:
        kernel32.CloseHandle(snap)
    return result


def find_owner(key, hwnd, pinned):
    """运行中的程序（exe 路径 key）对不上任何固定项时，看它是不是某个固定项的“本体”：
    装在固定项的目录里，并且和启动器在同一个文件夹 / 下一级子文件夹（Oopz、Discord 这类），
    或者就是由固定项启动的（Steam 的 steamwebhelper）。返回固定项的 target 或 None。"""
    if "\\steamapps\\" in key:                 # 从 Steam 启动的游戏是独立的软件
        return None
    folder = os.path.dirname(key)
    candidates = {}
    for target in pinned:
        tdir = os.path.dirname(target)
        if tdir not in SHARED_DIRS and (folder == tdir or folder.startswith(tdir + "\\")):
            candidates[target] = folder == tdir or os.path.dirname(folder) == tdir
    if not candidates:
        return None
    for target, near in candidates.items():
        if near:
            return target
    # 更深的子目录：要求是由这个固定项启动的
    procs = parent_map()
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    for _ in range(4):
        ppid, _ = procs.get(pid, (0, ""))
        if not ppid or ppid == pid or ppid not in procs or procs[ppid][1] in LAUNCHER_PARENTS:
            break
        parent = os.path.normcase(process_path(ppid) or "")
        if parent in candidates:
            return parent
        pid = ppid
    return None


def resolve_exe(path):
    """快捷方式 / 程序最终指向的 exe（规范化），用来和正在运行的窗口对应；其他类型返回 None"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".lnk":
        target = read_lnk(path)[0]
    elif ext == ".exe":
        target = path
    else:
        return None
    return norm(target) if target.lower().endswith(".exe") else None


# ---------------------------------------------------------------- 窗口枚举 / 切换

def process_path(pid):
    h = kernel32.OpenProcess(0x1000, False, pid)    # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return None
    try:
        buf = ctypes.create_unicode_buffer(1024)
        n = wintypes.DWORD(1024)
        return buf.value if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)) else None
    finally:
        kernel32.CloseHandle(h)


_pid_paths = {}
SKIP_CLASSES = {"Progman", "WorkerW", "Shell_TrayWnd", "Shell_SecondaryTrayWnd"}


def pid_path(pid):
    p = _pid_paths.get(pid)
    if p is None:
        p = process_path(pid) or ""
        _pid_paths[pid] = p
    return p


def is_cloaked(hwnd):
    v = ctypes.c_int(0)
    dwmapi.DwmGetWindowAttribute(wintypes.HWND(hwnd), 14, ctypes.byref(v), ctypes.sizeof(v))
    return v.value != 0


def uwp_real_pid(hwnd, host_pid):
    """UWP 应用的窗口属于 ApplicationFrameHost，真正的进程在子窗口里"""
    found = []

    def cb(child, _):
        pid = win32process.GetWindowThreadProcessId(child)[1]
        if pid != host_pid:
            found.append(pid)
            return False
        return True

    try:
        win32gui.EnumChildWindows(hwnd, cb, None)
    except Exception:
        pass
    return found[0] if found else host_pid


def enum_app_windows():
    """按 Z 序返回会出现在任务栏上的窗口：[(hwnd, exe 规范路径, exe 原始路径)]"""
    me = os.getpid()
    result = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        app = ex & WS_EX_APPWINDOW
        if not app and (win32gui.GetWindow(hwnd, win32con.GW_OWNER)
                        or ex & (win32con.WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE)):
            return True
        if not win32gui.GetWindowText(hwnd) or is_cloaked(hwnd):
            return True
        if win32gui.GetClassName(hwnd) in SKIP_CLASSES:
            return True
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        if pid == me:
            return True
        path = pid_path(pid)
        if os.path.basename(path).lower() == "applicationframehost.exe":
            path = pid_path(uwp_real_pid(hwnd, pid)) or path
        if path:
            result.append((hwnd, os.path.normcase(path), path))
        return True

    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return result


def focus_window(hwnd):
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # 前台锁定时，模拟一次 Alt 解锁再试
        try:
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass


class WindowMonitor(QThread):
    """后台每秒枚举一次窗口（只是 EnumWindows，几毫秒），按程序分组"""
    updated = Signal(object)     # [(exe 规范路径, exe 原始路径, [hwnd...])]，按首次出现的 Z 序

    def __init__(self):
        super().__init__()
        self._poke = threading.Event()

    def poke(self):
        self._poke.set()

    def run(self):
        while not self.isInterruptionRequested():
            groups = {}
            try:
                for hwnd, key, path in enum_app_windows():
                    groups.setdefault(key, (path, []))[1].append(hwnd)
                # 进程退出后清掉 pid 缓存，防止 pid 复用
                alive = {win32process.GetWindowThreadProcessId(h)[1] for _, hs in groups.values() for h in hs}
                if len(_pid_paths) > 300:
                    for pid in list(_pid_paths):
                        if pid not in alive:
                            _pid_paths.pop(pid, None)
            except Exception as e:
                log_error("窗口枚举失败: %r" % e)
            self.updated.emit([(k, p, hs) for k, (p, hs) in groups.items()])
            if self._poke.wait(1.0):
                self._poke.clear()
                self.msleep(500)


# ---------------------------------------------------------------- 隐藏任务栏 / 桌面图标

def taskbar_windows():
    wins = [win32gui.FindWindow("Shell_TrayWnd", None)]
    h = 0
    while True:
        h = win32gui.FindWindowEx(0, h, "Shell_SecondaryTrayWnd", None)
        if not h:
            break
        wins.append(h)
    return [w for w in wins if w]


def desktop_icon_views():
    views = []

    def cb(h, _):
        if win32gui.GetClassName(h) in ("Progman", "WorkerW"):
            dv = win32gui.FindWindowEx(h, 0, "SHELLDLL_DefView", None)
            lv = dv and win32gui.FindWindowEx(dv, 0, "SysListView32", None)
            if lv:
                views.append(lv)
        return True

    win32gui.EnumWindows(cb, None)
    return views


def show_windows(hwnds, show):
    for h in hwnds:
        if bool(win32gui.IsWindowVisible(h)) != show:
            win32gui.ShowWindow(h, win32con.SW_SHOW if show else win32con.SW_HIDE)


def restore_shell():
    """把任务栏和桌面图标恢复成正常状态"""
    try:
        show_windows(taskbar_windows(), True)
        show_windows(desktop_icon_views(), True)
    except Exception as e:
        log_error("恢复任务栏/桌面图标失败: %r" % e)


# ---------------------------------------------------------------- 全局快捷键

WM_HOTKEY = 0x0312
# (功能, 显示名, 默认组合键)；注册时用序号 1、2、3… 当热键 id
HOTKEY_ACTIONS = [
    ("desktop_icons", "显示 / 隐藏桌面图标", "Ctrl+Alt+D"),
    ("dock", "显示 / 隐藏 Dock", "Ctrl+Alt+H"),
    ("taskbar", "显示 / 隐藏 Windows 任务栏", ""),
]
MOD_ALT, MOD_CONTROL, MOD_SHIFT, MOD_WIN, MOD_NOREPEAT = 0x1, 0x2, 0x4, 0x8, 0x4000
_SPECIAL_VK = {Qt.Key_Space: 0x20, Qt.Key_PageUp: 0x21, Qt.Key_PageDown: 0x22, Qt.Key_End: 0x23,
               Qt.Key_Home: 0x24, Qt.Key_Left: 0x25, Qt.Key_Up: 0x26, Qt.Key_Right: 0x27, Qt.Key_Down: 0x28,
               Qt.Key_Insert: 0x2D, Qt.Key_Delete: 0x2E, Qt.Key_Pause: 0x13}


def parse_hotkey(text):
    """"Ctrl+Alt+D" → (Win32 修饰键, 虚拟键码)；无效或没带修饰键时返回 None"""
    seq = QKeySequence(text)
    if seq.isEmpty():
        return None
    combo = seq[0]
    qmods, key = combo.keyboardModifiers(), combo.key()
    mods = 0
    if qmods & Qt.ControlModifier:
        mods |= MOD_CONTROL
    if qmods & Qt.AltModifier:
        mods |= MOD_ALT
    if qmods & Qt.ShiftModifier:
        mods |= MOD_SHIFT
    if qmods & Qt.MetaModifier:
        mods |= MOD_WIN
    k = key.value if hasattr(key, "value") else int(key)
    if Qt.Key_A.value <= k <= Qt.Key_Z.value or Qt.Key_0.value <= k <= Qt.Key_9.value:
        vk = k                                       # 字母、数字的 Qt 键值就是 ASCII
    elif Qt.Key_F1.value <= k <= Qt.Key_F24.value:
        vk = 0x70 + k - Qt.Key_F1.value
    else:
        vk = next((v for q, v in _SPECIAL_VK.items() if q.value == k), None)
    if not mods or vk is None:
        return None
    return mods, vk


user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]


# ---------------------------------------------------------------- 数据

class DockItem:
    MIP_SIZES = (48, 96, 160, 256)

    def __init__(self, path, name=None, args="", pinned=True, image=None):
        self.path = path
        self.name = name or display_name(path)
        self.args = args
        self.pinned = pinned
        self.target = resolve_exe(path)
        self.image = image if image is not None else load_icon_image(path)
        self.hwnds = []               # 当前属于它的窗口
        self.bounce_start = -10.0
        self._mips = {}
        self._shadow = None
        self._reflection = None

    def _base96(self):
        return self.image.scaled(96, 96, Qt.KeepAspectRatio, Qt.SmoothTransformation) \
            .convertToFormat(QImage.Format_ARGB32_Premultiplied)

    def shadow_pixmap(self):
        """柔和投影：剪影缩小再放大 = 便宜的模糊，只算一次"""
        if self._shadow is None:
            canvas = QImage(128, 128, QImage.Format_ARGB32_Premultiplied)
            canvas.fill(Qt.transparent)
            q = QPainter(canvas)
            q.drawImage(16, 16, self._base96())
            q.setCompositionMode(QPainter.CompositionMode_SourceIn)
            q.fillRect(canvas.rect(), QColor(0, 0, 0, 150))
            q.end()
            blurred = canvas.scaled(18, 18, Qt.IgnoreAspectRatio, Qt.SmoothTransformation) \
                .scaled(128, 128, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
            self._shadow = QPixmap.fromImage(blurred)
        return self._shadow

    def reflection_pixmap(self):
        """倒影：上下翻转后用渐变把下半部分淡出"""
        if self._reflection is None:
            img = self._base96().transformed(QTransform().scale(1, -1))
            q = QPainter(img)
            q.setCompositionMode(QPainter.CompositionMode_DestinationIn)
            g = QLinearGradient(0, 0, 0, img.height())
            g.setColorAt(0, QColor(0, 0, 0, 150))
            g.setColorAt(0.45, QColor(0, 0, 0, 0))
            q.fillRect(img.rect(), g)
            q.end()
            self._reflection = QPixmap.fromImage(img)
        return self._reflection

    @property
    def running(self):
        return bool(self.hwnds)

    def pixmap_for(self, px):
        size = next((s for s in self.MIP_SIZES if s >= px), self.MIP_SIZES[-1])
        pm = self._mips.get(size)
        if pm is None:
            pm = QPixmap.fromImage(self.image.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self._mips[size] = pm
        return pm

    def to_dict(self):
        return {"name": self.name, "path": self.path, "args": self.args}


# ---------------------------------------------------------------- 皮肤

SKIN_BASE = dict(
    name="", bg="glass",                      # glass 圆角玻璃 / shelf 3D 玻璃台 / none 无背景 / image 图片
    bg_top=(60, 60, 68, 165), bg_bottom=(22, 22, 26, 175),
    border=(255, 255, 255, 50), border2=None, border_w=1.0,   # border2：边框做横向渐变
    highlight=(255, 255, 255, 30),            # 顶部内高光
    radius=0.3,                               # 圆角 = 高度 × radius
    indicator="dot", accent=(255, 255, 255, 230),             # 运行指示：dot / bar / glow
    sep=(255, 255, 255, 70),
    label_bg=(28, 28, 32, 225), label_fg=(255, 255, 255), label_border=None,
    shadow=False, reflection=False, glow=False, decor=None,   # decor：stars / sakura / bubbles / grid
    text=(235, 235, 235),
)

SKINS = {
    "dark": dict(name="深色毛玻璃"),
    "light": dict(name="浅色毛玻璃", bg_top=(255, 255, 255, 150), bg_bottom=(225, 228, 235, 130),
                  border=(255, 255, 255, 190), highlight=(255, 255, 255, 120), accent=(30, 30, 30, 210),
                  sep=(0, 0, 0, 55), text=(40, 40, 40), shadow=True),
    "neon": dict(name="赛博霓虹", bg_top=(30, 12, 56, 205), bg_bottom=(8, 6, 26, 220),
                 border=(255, 60, 200, 235), border2=(40, 230, 255, 235), border_w=1.6, highlight=None,
                 radius=0.32, indicator="glow", accent=(40, 230, 255, 255), sep=(255, 60, 200, 130),
                 label_bg=(20, 8, 40, 235), label_fg=(130, 240, 255), label_border=(255, 60, 200, 210),
                 glow=True, decor="grid"),
    "sakura": dict(name="樱花粉", bg_top=(255, 220, 235, 200), bg_bottom=(255, 178, 205, 185),
                   border=(255, 255, 255, 220), highlight=(255, 255, 255, 160), radius=0.5, indicator="bar",
                   accent=(226, 64, 128, 240), sep=(226, 64, 128, 90), label_bg=(226, 64, 128, 235),
                   shadow=True, decor="sakura", text=(150, 40, 90)),
    "ocean": dict(name="深海蓝", bg_top=(20, 150, 210, 175), bg_bottom=(4, 60, 130, 200),
                  border=(170, 235, 255, 170), highlight=(255, 255, 255, 80), accent=(200, 250, 255, 245),
                  sep=(200, 250, 255, 90), label_bg=(4, 50, 110, 235), reflection=True, decor="bubbles"),
    "gold": dict(name="黑金", bg_top=(34, 30, 22, 225), bg_bottom=(10, 9, 7, 235),
                 border=(190, 150, 50, 235), border2=(255, 228, 150, 235), border_w=1.4,
                 highlight=(255, 220, 140, 45), radius=0.22, indicator="bar", accent=(235, 195, 90, 255),
                 sep=(212, 175, 55, 120), label_bg=(20, 17, 12, 240), label_fg=(240, 205, 110),
                 label_border=(212, 175, 55, 200), shadow=True),
    "shelf": dict(name="3D 玻璃台", bg="shelf", bg_top=(235, 240, 250, 140), bg_bottom=(150, 160, 185, 170),
                  border=(255, 255, 255, 220), highlight=(255, 255, 255, 200), indicator="glow",
                  accent=(140, 210, 255, 255), sep=(255, 255, 255, 120), reflection=True),
    "starry": dict(name="星空", bg_top=(28, 32, 78, 205), bg_bottom=(6, 7, 24, 220),
                   border=(140, 160, 255, 120), highlight=(255, 255, 255, 35), indicator="glow",
                   accent=(255, 236, 170, 255), sep=(170, 180, 255, 90), label_bg=(14, 16, 44, 235),
                   label_fg=(255, 236, 170), decor="stars"),
    "minimal": dict(name="极简透明", bg="none", shadow=True, sep=(255, 255, 255, 110),
                    accent=(255, 255, 255, 240)),
    # 若叶睦：代表色鼠尾草绿 #779977 + 奶油色 + 琥珀金（眼睛的颜色），背景是她最爱的黄瓜片
    "mutsumi": dict(name="若叶睦", bg_top=(176, 202, 166, 205), bg_bottom=(119, 153, 119, 215),
                    border=(244, 238, 214, 225), highlight=(255, 255, 255, 130), radius=0.42, indicator="glow",
                    accent=(236, 192, 88, 255), sep=(244, 238, 214, 130), label_bg=(86, 114, 86, 240),
                    label_fg=(255, 248, 225), label_border=(236, 192, 88, 190), shadow=True,
                    decor="cucumber", text=(40, 62, 40)),
    # Mortis：Ave Mujica 舞台上的另一面，哥特墨绿 + 金色
    "mortis": dict(name="Mortis", bg_top=(38, 52, 42, 228), bg_bottom=(9, 13, 11, 238),
                   border=(140, 112, 56, 235), border2=(228, 198, 122, 240), border_w=1.4,
                   highlight=(200, 230, 200, 28), radius=0.2, indicator="bar", accent=(214, 182, 100, 255),
                   sep=(150, 120, 60, 130), label_bg=(14, 20, 16, 240), label_fg=(228, 202, 132),
                   label_border=(150, 120, 60, 210), glow=True, decor="gothic"),
}
CUSTOM_COLOR, CUSTOM_IMAGE = "custom_color", "custom_image"

_rng = random.Random(20261002)     # 固定种子：装饰位置每次都一样，放大时不会乱跳
DECOR_STARS = [(_rng.random(), _rng.random(), _rng.uniform(0.5, 1.4), _rng.randint(60, 220)) for _ in range(70)]
DECOR_BUBBLES = [(_rng.random(), _rng.uniform(0.2, 1.0), _rng.uniform(0.4, 1.0)) for _ in range(18)]
DECOR_FLOWERS = [(_rng.random(), _rng.uniform(0.1, 0.95), _rng.uniform(0.5, 1.0), _rng.uniform(0, 72))
                 for _ in range(14)]
DECOR_SLICES = [(_rng.random(), _rng.uniform(0.15, 0.9), _rng.uniform(0.6, 1.0), _rng.uniform(0, 360))
                for _ in range(12)]


def draw_cucumber_slice(p, center, r, rot):
    p.setPen(QPen(QColor(78, 128, 60, 170), max(1.0, r * 0.16)))
    p.setBrush(QColor(206, 234, 176, 150))
    p.drawEllipse(center, r, r)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(246, 252, 228, 170))
    for i in range(6):          # 一圈籽
        a = math.radians(rot + 60 * i)
        p.drawEllipse(QPointF(center.x() + math.cos(a) * r * 0.45, center.y() + math.sin(a) * r * 0.45),
                      r * 0.11, r * 0.17)


def build_skin(skin_id, cfg):
    th = dict(SKIN_BASE)
    if skin_id == CUSTOM_COLOR:
        c = QColor(cfg.get("skin_color") or "#6a5acd")
        light = c.lightnessF() > 0.62
        top, bot, edge, deep = c.lighter(125), c.darker(150), c.lighter(165), c.darker(240)
        th.update(name="自定义颜色", bg_top=(top.red(), top.green(), top.blue(), 190),
                  bg_bottom=(bot.red(), bot.green(), bot.blue(), 205),
                  border=(edge.red(), edge.green(), edge.blue(), 180), highlight=(255, 255, 255, 70),
                  accent=(30, 30, 30, 220) if light else (255, 255, 255, 235),
                  sep=(0, 0, 0, 60) if light else (255, 255, 255, 80),
                  label_bg=(deep.red(), deep.green(), deep.blue(), 235),
                  text=(40, 40, 40) if light else (240, 240, 240), shadow=True)
    elif skin_id == CUSTOM_IMAGE:
        th.update(name="图片皮肤", bg="image", bg_top=(0, 0, 0, 30), bg_bottom=(0, 0, 0, 90),
                  border=(255, 255, 255, 140), highlight=(255, 255, 255, 60), shadow=True)
    else:
        th.update(SKINS.get(skin_id, SKINS["dark"]))
    th["opacity"] = min(1.0, max(0.2, float(cfg.get("skin_opacity", 1.0))))
    return th


def qc(rgba, alpha=None):
    c = QColor(*rgba)
    if alpha is not None:
        c.setAlpha(alpha)
    return c


def bar_path(th, bar):
    path = QPainterPath()
    if th["bg"] == "shelf":
        h = bar.height() * 0.55
        top, inset = bar.bottom() - h, h * 0.8
        path.moveTo(bar.left() + inset, top)
        path.lineTo(bar.right() - inset, top)
        path.lineTo(bar.right(), bar.bottom())
        path.lineTo(bar.left(), bar.bottom())
        path.closeSubpath()
    else:
        r = bar.height() * th["radius"]
        path.addRoundedRect(bar, r, r)
    return path


def icon_base_y(th, bar, pad):
    """图标底边的 y 坐标"""
    return bar.bottom() - (bar.height() * 0.28 if th["bg"] == "shelf" else pad)


def indicator_y(th, bar, pad):
    return bar.bottom() - (bar.height() * 0.1 if th["bg"] == "shelf" else pad * 0.42)


def draw_flower(p, center, size, rot, color):
    p.setBrush(color)
    for i in range(5):
        p.save()
        p.translate(center)
        p.rotate(rot + 72 * i)
        p.drawEllipse(QRectF(-size * 0.28, -size, size * 0.56, size * 0.9))
        p.restore()
    p.setBrush(QColor(255, 240, 160, color.alpha()))
    p.drawEllipse(center, size * 0.14, size * 0.14)


def paint_decor(p, th, bar, path):
    d = th["decor"]
    if not d:
        return
    p.save()
    p.setClipPath(path)
    p.setPen(Qt.NoPen)
    w, h = bar.width(), bar.height()
    if d == "stars":
        for fx, fy, r, a in DECOR_STARS:
            p.setBrush(QColor(255, 255, 255, a))
            p.drawEllipse(QPointF(bar.left() + fx * w, bar.top() + fy * h), r, r)
    elif d == "bubbles":
        p.setBrush(Qt.NoBrush)
        for fx, fy, r in DECOR_BUBBLES:
            p.setPen(QPen(QColor(255, 255, 255, 55), 1))
            p.drawEllipse(QPointF(bar.left() + fx * w, bar.top() + fy * h), r * h * 0.12, r * h * 0.12)
    elif d == "sakura":
        for fx, fy, s, rot in DECOR_FLOWERS:
            draw_flower(p, QPointF(bar.left() + fx * w, bar.top() + fy * h), s * h * 0.16, rot,
                        QColor(255, 255, 255, 120))
    elif d == "cucumber":
        for fx, fy, s, rot in DECOR_SLICES:
            draw_cucumber_slice(p, QPointF(bar.left() + fx * w, bar.top() + fy * h), s * h * 0.13, rot)
    elif d == "gothic":
        # 一排金色菱形纹 + 细竖纹
        gold = qc(th["accent"], 40)
        p.setPen(QPen(qc(th["accent"], 14), 1))
        step = h * 0.16
        x = bar.left()
        while x < bar.right():
            p.drawLine(QPointF(x, bar.top()), QPointF(x, bar.bottom()))
            x += step
        p.setPen(QPen(gold, 1))
        p.setBrush(Qt.NoBrush)
        s, cy = h * 0.11, bar.center().y()
        x = bar.left() + h * 0.4
        while x < bar.right():
            diamond = QPainterPath()
            diamond.moveTo(x, cy - s)
            diamond.lineTo(x + s * 0.7, cy)
            diamond.lineTo(x, cy + s)
            diamond.lineTo(x - s * 0.7, cy)
            diamond.closeSubpath()
            p.drawPath(diamond)
            p.drawEllipse(QPointF(x + h * 0.4, cy), 1.2, 1.2)
            x += h * 0.8
    elif d == "grid":
        p.setPen(QPen(qc(th["accent"], 26), 1))
        step = h * 0.28
        x = bar.left() + step / 2
        while x < bar.right():
            p.drawLine(QPointF(x, bar.top()), QPointF(x, bar.bottom()))
            x += step
        p.drawLine(QPointF(bar.left(), bar.center().y()), QPointF(bar.right(), bar.center().y()))
    p.restore()


def paint_background(p, th, bar, image=None):
    if th["bg"] == "none":
        return
    path = bar_path(th, bar)
    pr = path.boundingRect()
    p.save()
    p.setOpacity(p.opacity() * th["opacity"])
    if th["bg"] == "image" and image is not None and not image.isNull():
        p.setClipPath(path)
        scale = max(pr.width() / image.width(), pr.height() / image.height())
        w, h = image.width() * scale, image.height() * scale
        p.drawPixmap(QRectF(pr.center().x() - w / 2, pr.center().y() - h / 2, w, h), image, QRectF(image.rect()))
        p.setClipping(False)
    grad = QLinearGradient(pr.topLeft(), pr.bottomLeft())
    grad.setColorAt(0, qc(th["bg_top"]))
    grad.setColorAt(1, qc(th["bg_bottom"]))
    p.fillPath(path, grad)
    paint_decor(p, th, bar, path)

    if th["highlight"]:
        hl = qc(th["highlight"])
        p.save()
        p.setClipPath(path)
        if th["bg"] == "shelf":
            p.fillRect(QRectF(pr.left(), pr.bottom() - 3, pr.width(), 3), hl)      # 玻璃台前沿
        p.strokePath(path.translated(0, 1.2), QPen(hl, 1))
        p.restore()
    if th["border2"]:
        g = QLinearGradient(pr.topLeft(), pr.topRight())
        g.setColorAt(0, qc(th["border"]))
        g.setColorAt(0.5, qc(th["border2"]))
        g.setColorAt(1, qc(th["border"]))
        pen = QPen(QBrush(g), th["border_w"])
    else:
        pen = QPen(qc(th["border"]), th["border_w"])
    p.strokePath(path, pen)
    p.restore()


def paint_indicator(p, th, cx, y, B):
    acc = qc(th["accent"])
    p.setPen(Qt.NoPen)
    if th["indicator"] == "bar":
        w, h = B * 0.34, max(2.5, B / 17)
        p.setBrush(acc)
        p.drawRoundedRect(QRectF(cx - w / 2, y - h / 2, w, h), h / 2, h / 2)
        return
    r = max(2.0, B / 22)
    if th["indicator"] == "glow":
        g = QRadialGradient(QPointF(cx, y), r * 4)
        g.setColorAt(0, qc(th["accent"], 160))
        g.setColorAt(1, qc(th["accent"], 0))
        p.setBrush(g)
        p.drawEllipse(QPointF(cx, y), r * 4, r * 4)
    p.setBrush(acc)
    p.drawEllipse(QPointF(cx, y), r, r)


def paint_separator(p, th, cx, bar, pad):
    p.setPen(QPen(qc(th["sep"]), 1))
    cx = round(cx) + 0.5
    if th["bg"] == "shelf":
        y1, y2 = bar.bottom() - bar.height() * 0.48, bar.bottom() - 4
    else:
        y1, y2 = bar.top() + pad * 0.7, bar.bottom() - pad * 0.7
    p.drawLine(QPointF(cx, y1), QPointF(cx, y2))


def paint_item(p, th, item, rect, base_y, running, hovered, path, ind_y, B, dpr):
    w = rect.width()
    if hovered and th["glow"]:
        g = QRadialGradient(rect.center(), w * 0.78)
        g.setColorAt(0, qc(th["accent"], 120))
        g.setColorAt(1, qc(th["accent"], 0))
        p.setPen(Qt.NoPen)
        p.setBrush(g)
        p.drawEllipse(rect.center(), w * 0.78, w * 0.78)
    if th["shadow"]:
        s = item.shadow_pixmap()
        k = w * 128 / 96
        p.drawPixmap(QRectF(rect.center().x() - k / 2, rect.center().y() - k / 2 + w * 0.06, k, k), s, QRectF(s.rect()))
    if th["reflection"]:
        rp = item.reflection_pixmap()
        p.save()
        p.setClipPath(path)
        p.setOpacity(p.opacity() * 0.6)
        # 倒影跟弹跳方向相反
        p.drawPixmap(QRectF(rect.left(), 2 * base_y - rect.bottom() + 1, w, w), rp, QRectF(rp.rect()))
        p.restore()
    pm = item.pixmap_for(w * dpr)
    p.drawPixmap(rect, pm, QRectF(pm.rect()))
    if running:
        paint_indicator(p, th, rect.center().x(), ind_y, B)


def load_config():
    cfg, saved, first_run = dict(DEFAULT_CONFIG), {}, False
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            saved = json.load(f)
        cfg.update(saved)
        if "skin" not in saved:                  # 旧版本只有 dark / light 两种外观
            cfg["skin"] = saved.get("style", "dark")
        cfg.pop("style", None)
    except FileNotFoundError:
        first_run = True
    except Exception as e:
        log_error("配置文件损坏，已使用默认配置: %r" % e)
        try:
            shutil.copy2(CONFIG_PATH, CONFIG_PATH + ".bak")
        except OSError:
            pass
        first_run = True
    # 快捷键：没设置过的功能用默认值；旧版本只有一个 hotkey_desktop_icons
    hotkeys = {action: default for action, _, default in HOTKEY_ACTIONS}
    if "hotkey_desktop_icons" in saved:
        hotkeys["desktop_icons"] = saved["hotkey_desktop_icons"]
    hotkeys.update(saved.get("hotkeys") or {})
    cfg["hotkeys"] = hotkeys
    cfg.pop("hotkey_desktop_icons", None)
    return cfg, first_run


def write_config(cfg):
    tmp = CONFIG_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CONFIG_PATH)
    except OSError as e:
        log_error("保存配置失败: %r" % e)


def default_item_paths():
    """首次运行：导入 BitDock 的图标（复制一份到 links，卸载 BitDock 也不受影响），再补几个常用程序"""
    paths = []
    bitdock_links = os.path.join(os.path.expanduser("~"), "Documents", "BitSoft Files", "BitDock", "link")
    if os.path.isdir(bitdock_links):
        os.makedirs(LINKS_DIR, exist_ok=True)
        for name in sorted(os.listdir(bitdock_links)):
            if os.path.splitext(name)[1].lower() in (".lnk", ".url", ".exe"):
                dst = os.path.join(LINKS_DIR, name)
                try:
                    shutil.copy2(os.path.join(bitdock_links, name), dst)
                    paths.append(dst)
                except OSError:
                    pass
    win = os.environ.get("WINDIR", r"C:\Windows")
    candidates = [os.path.join(win, "explorer.exe")]
    if not paths:
        candidates += [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            os.path.join(win, "System32", "notepad.exe"),
            os.path.join(win, "System32", "calc.exe"),
            os.path.join(win, "System32", "Taskmgr.exe"),
        ]
    return [p for p in candidates if os.path.exists(p)] + paths


# ---------------------------------------------------------------- Dock 窗口

MENU_QSS = """
QMenu { background: #2a2a30; color: #eee; border: 1px solid #48484f; padding: 5px; }
QMenu::item { padding: 6px 26px 6px 24px; border-radius: 5px; }
QMenu::item:selected { background: #3b6cf0; }
QMenu::item:disabled { color: #888; }
QMenu::separator { height: 1px; background: #48484f; margin: 5px 8px; }
QMenu::indicator { width: 13px; height: 13px; left: 6px; }
"""


class Dock(QWidget):
    launch_failed = Signal(str, str)

    def __init__(self, cfg, tray_message):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.NoDropShadowWindowHint)
        self.setWindowTitle(APP_NAME)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setMouseTracking(True)
        self.setAcceptDrops(True)

        self.cfg = cfg
        self.theme = build_skin(cfg["skin"], cfg)
        self._skin_pm = (None, None)  # (缓存键, 缩放好的背景图)
        self._mascot_pm = (None, None)
        self.mascot_h = 0
        self.mascot_bounce = -10.0
        self.mascot_line = ""
        self.mascot_say_until = 0.0
        self._seen_apps = None        # {exe: 最后一次看到它有窗口的时间}，用来判断“新打开”
        self._last_state = None
        self._owner_sig, self._owners = None, {}   # 运行中程序 → 所属固定项 的缓存
        self._open_apps = {}          # {exe: 显示名}，当前开着窗口的程序
        self._gone_since = {}         # {exe: 窗口消失的时间}，用来确认是真的关了
        self._say_queue = []
        self.tray_message = tray_message
        self.items = [DockItem(d["path"], d.get("name"), d.get("args", "")) for d in cfg["items"] if d.get("path")]
        self.running_items = []       # 没固定、正在运行的程序

        # 交互状态
        self.hovered = False
        self.mouse_pos = None
        self.hover_amt = 0.0          # 0..1 放大程度（带动画）
        self.reveal = 1.0             # 0..1 自动隐藏的显示程度
        self.hide_after = 0.0
        self.press_item = None
        self.press_pos = None
        self.dragging = False
        self.drag_idx = None
        self.ext_drag = False
        self.drop_idx = None
        self.menu_open = False
        self.user_hidden = False
        self.fs_hidden = False
        self.taskbar_hidden = False
        self.icons_hidden = False

        self.apply_metrics()

        self.anim = QTimer(self, interval=16, timeout=self.tick)
        self._last_tick = time.monotonic()
        self.hover_check = QTimer(self, interval=120, timeout=self.check_hover)
        self.shell_timer = QTimer(self, interval=1500, timeout=self.periodic)
        self.shell_timer.start()

        self.monitor = WindowMonitor()
        self.monitor.updated.connect(self.on_windows_update)
        self.monitor.start()

        self.launch_failed.connect(lambda n, err: self.tray_message("启动失败：" + n, err))
        screen = QGuiApplication.primaryScreen()
        screen.availableGeometryChanged.connect(lambda _: self.relayout_window())

        self.apply_shell()
        failed = self.register_hotkeys()
        if failed:
            self.tray_message("快捷键不可用", "、".join("%s（%s）" % f for f in failed)
                              + " 已被其他程序占用，请在 设置 → 快捷键设置 里换一个")

    # ---------- 全局快捷键

    def register_hotkeys(self):
        """注册全部快捷键，返回注册失败的 [(功能名, 组合键)]"""
        hwnd = int(self.winId())
        failed = []
        for i, (action, title, _) in enumerate(HOTKEY_ACTIONS, 1):
            user32.UnregisterHotKey(hwnd, i)
            text = self.cfg["hotkeys"].get(action) or ""
            if not text:
                continue
            parsed = parse_hotkey(text)
            if parsed is None or not user32.RegisterHotKey(hwnd, i, parsed[0] | MOD_NOREPEAT, parsed[1]):
                failed.append((title, text))
        return failed

    def nativeEvent(self, event_type, message):
        if event_type == b"windows_generic_MSG":
            msg = wintypes.MSG.from_address(int(message))
            if msg.message == WM_HOTKEY and 1 <= msg.wParam <= len(HOTKEY_ACTIONS):
                self.run_hotkey(HOTKEY_ACTIONS[msg.wParam - 1][0])
                return True, 0
        return False, 0

    def run_hotkey(self, action):
        if action == "desktop_icons":
            self.set_option("hide_desktop_icons", not self.cfg["hide_desktop_icons"])
        elif action == "dock":
            self.set_user_hidden(not self.user_hidden)
        elif action == "taskbar":
            self.set_option("hide_taskbar", not self.cfg["hide_taskbar"])

    def edit_hotkeys(self):
        dlg = QDialog()
        dlg.setWindowTitle("快捷键设置 - " + APP_NAME)
        dlg.setWindowFlag(Qt.WindowStaysOnTopHint)
        dlg.setStyleSheet(SKIN_CENTER_QSS)
        dlg.setMinimumWidth(560)
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(12)
        title = QLabel("快捷键设置")
        title.setObjectName("title")
        lay.addWidget(title)
        lay.addWidget(QLabel("点输入框后按下组合键（要带 Ctrl / Alt / Shift / Win），在任何程序里都能用。"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        edits = {}
        for row, (action, name, _) in enumerate(HOTKEY_ACTIONS):
            grid.addWidget(QLabel(name), row, 0)
            edit = QKeySequenceEdit(QKeySequence(self.cfg["hotkeys"].get(action) or ""))
            edit.setMaximumSequenceLength(1)
            edit.setMinimumWidth(180)
            line = edit.findChild(QLineEdit)
            if line is not None:
                line.setPlaceholderText("未设置")
            grid.addWidget(edit, row, 1)
            clear = QPushButton("不用")
            clear.clicked.connect(edit.clear)
            grid.addWidget(clear, row, 2)
            edits[action] = edit
        lay.addLayout(grid)
        buttons = QHBoxLayout()
        reset = QPushButton("恢复默认")
        reset.clicked.connect(lambda: [edits[a].setKeySequence(QKeySequence(d)) for a, _, d in HOTKEY_ACTIONS])
        ok, cancel = QPushButton("确定"), QPushButton("取消")
        ok.clicked.connect(dlg.accept)
        cancel.clicked.connect(dlg.reject)
        buttons.addWidget(reset)
        buttons.addStretch(1)
        buttons.addWidget(cancel)
        buttons.addWidget(ok)
        lay.addLayout(buttons)
        dlg.show()
        dlg.activateWindow()
        if not dlg.exec():
            return

        new = {a: e.keySequence().toString(QKeySequence.PortableText) for a, e in edits.items()}
        names = {a: n for a, n, _ in HOTKEY_ACTIONS}
        bad = [names[a] + "（%s）" % t for a, t in new.items() if t and parse_hotkey(t) is None]
        if bad:
            QMessageBox.warning(None, APP_NAME, "这些组合键不能用作全局快捷键（要带 Ctrl/Alt/Shift/Win，"
                                                "主键用字母、数字或 F1~F24 等）：\n" + "\n".join(bad))
            return
        used = [t for t in new.values() if t]
        if len(used) != len(set(used)):
            QMessageBox.warning(None, APP_NAME, "有两个功能用了同一个组合键，请改成不同的。")
            return
        old = dict(self.cfg["hotkeys"])
        self.cfg["hotkeys"] = new
        failed = self.register_hotkeys()
        if failed:
            # 被占用的那几个退回原来的设置
            for action, name, _ in HOTKEY_ACTIONS:
                if any(f[0] == name for f in failed):
                    self.cfg["hotkeys"][action] = old.get(action, "")
            self.register_hotkeys()
            QMessageBox.warning(None, APP_NAME, "这些组合键已被其他程序占用，已保留原来的设置：\n"
                                + "\n".join("%s（%s）" % f for f in failed))
        self.save_config()

    # ---------- 尺寸 / 配置

    def apply_metrics(self):
        self.B = int(self.cfg["icon_size"])
        self.M = max(1.0, float(self.cfg["magnify"]))
        self.spread = max(1.0, float(self.cfg["spread"]))
        self.gap = round(self.B * 0.16)
        self.pad = round(self.B * 0.2)
        self.sep_w = max(8, round(self.B * 0.22))
        self.bar_h = self.B + 2 * self.pad
        self.relayout_window()

    def relayout_window(self):
        screen = QGuiApplication.primaryScreen()
        # 任务栏被我们藏起来时，系统仍给它留着位置，Dock 直接贴到屏幕最底下占用那块地方
        scr = screen.geometry() if self.taskbar_hidden else screen.availableGeometry()
        B, M = self.B, self.M
        n = len(self.items) + len(self.running_items) + 3
        above = max(B * (M - 1) + B * 0.4 + 46, B * 2.1)
        height = int(MARGIN + self.bar_h + above)
        self.mascot_h = min(B * float(self.cfg.get("mascot_size", 2.6)), height - MARGIN - 6)
        mascot = self.mascot_pixmap()
        mascot_w = mascot.width() / mascot.height() * self.mascot_h if mascot else 0
        width = n * (B + self.gap) + 2 * self.pad + B * (M - 1) * self.spread * 1.3 + B * 3
        width += 2 * (mascot_w + self.gap * 2)       # Dock 居中，两边都留出看板娘的位置
        width = min(int(width), scr.width())
        self.setGeometry(scr.x() + (scr.width() - width) // 2, scr.y() + scr.height() - height, width, height)
        self.update()

    def save_config(self):
        self.cfg["items"] = [it.to_dict() for it in self.items]
        write_config(self.cfg)

    def set_option(self, key, value):
        self.cfg[key] = value
        self.save_config()
        if key in ("icon_size", "magnify", "spread"):
            self.apply_metrics()
        if key == "hide_on_fullscreen" and not value and self.fs_hidden:
            self.fs_hidden = False
            self.sync_visible()
        if key in ("hide_taskbar", "hide_desktop_icons"):
            self.apply_shell()
        if key == "show_running":
            self.monitor.poke()
            if not value:
                self.running_items = []
                self.relayout_window()
        if key.startswith("skin"):
            self.theme = build_skin(self.cfg["skin"], self.cfg)
            self._skin_pm = (None, None)      # 换了图（可能还是同一个文件名）就重新读
        if key.startswith("mascot"):
            self._mascot_pm = (None, None)
            self.relayout_window()
        self.kick()

    # 图片缓存只认 (路径, 尺寸)，不在每帧去碰硬盘；换图时由 set_option 清掉缓存

    def skin_image(self):
        """图片皮肤的背景图，预先缩到够用的尺寸，避免每帧缩放大图"""
        path = self.cfg.get("skin_image") or ""
        if self.cfg["skin"] != CUSTOM_IMAGE or not path:
            return None
        height = int(self.bar_h * 2 * self.devicePixelRatioF())
        if self._skin_pm[0] != (path, height):
            img = QImage(path)
            pm = None if img.isNull() else QPixmap.fromImage(
                img.scaledToHeight(min(height, img.height()), Qt.SmoothTransformation))
            self._skin_pm = ((path, height), pm)
        return self._skin_pm[1]

    # ---------- 看板娘

    def mascot_pixmap(self):
        path = self.cfg.get("mascot_image") or ""
        if not path or self.mascot_h <= 0:
            return None
        height = int(self.mascot_h * self.devicePixelRatioF())
        if self._mascot_pm[0] != (path, height):
            img = QImage(path)
            pm = None if img.isNull() else QPixmap.fromImage(img.scaledToHeight(height, Qt.SmoothTransformation))
            self._mascot_pm = ((path, height), pm)
        return self._mascot_pm[1]

    def mascot_rect(self, bar):
        pm = self.mascot_pixmap()
        if pm is None:
            return None
        h = self.mascot_h
        w = pm.width() / pm.height() * h
        t = time.monotonic() - self.mascot_bounce
        hop = math.sin(math.pi * t / MASCOT_HOP) * self.B * 0.28 if 0 <= t < MASCOT_HOP else 0.0
        x = bar.right() + self.gap * 1.5 if self.cfg.get("mascot_side") != "left" else bar.left() - self.gap * 1.5 - w
        return QRectF(x, bar.bottom() - h - hop, w, h)

    def poke_mascot(self):
        lines = [s for s in self.cfg.get("mascot_lines") or [] if s] or ["……"]
        choices = [s for s in lines if s != self.mascot_line] or lines     # 不连着说同一句
        self.mascot_say(random.choice(choices))

    def mascot_say(self, text, seconds=2.8):
        self.mascot_line = text
        now = time.monotonic()
        self.mascot_bounce = now
        self.mascot_say_until = now + seconds
        QTimer.singleShot(int(seconds * 1000) + 100, self._say_done)
        self.kick()

    def _say_done(self):
        self.update()
        # 旧的定时器也会回调，只有当前这句真的说完了才接着说下一句
        if self._say_queue and time.monotonic() >= self.mascot_say_until - 0.05:
            self.mascot_say(*self._say_queue.pop(0))

    def queue_say(self, text, seconds=3.2):
        """正在说话就排队，说完再说；队列最多攒 3 句，免得一下开关一堆程序时念个没完"""
        if time.monotonic() < self.mascot_say_until:
            if len(self._say_queue) < 3:
                self._say_queue.append((text, seconds))
        else:
            self.mascot_say(text, seconds)

    def announce(self, names, enabled_key, lines_key, fallback):
        """看板娘报一句“打开了 xxx” / “xxx……关掉了。”"""
        if not names or not self.cfg.get(enabled_key, True) or self.mascot_pixmap() is None:
            return
        names = list(dict.fromkeys(names))     # 同一个软件的几个进程只报一次
        shown = "、".join(names[:3]) + ("等 %d 个" % len(names) if len(names) > 3 else "")
        templates = [s for s in self.cfg.get(lines_key) or [] if "{name}" in s] or [fallback]
        self.queue_say(random.choice(templates).replace("{name}", shown))

    # ---------- 任务栏 / 桌面图标

    def apply_shell(self):
        # Dock 被用户隐藏时把任务栏还回去，免得托盘图标也跟着看不见
        want_taskbar = self.cfg["hide_taskbar"] and not self.user_hidden
        want_icons = self.cfg["hide_desktop_icons"]
        try:
            # 只隐藏任务栏窗口，不改系统的“自动隐藏”设置（改了 explorer 会反复把任务栏弹回来）
            if want_taskbar:
                show_windows(taskbar_windows(), False)
                if not self.taskbar_hidden:
                    self.taskbar_hidden = True
                    self.relayout_window()
            elif self.taskbar_hidden:
                show_windows(taskbar_windows(), True)
                self.taskbar_hidden = False
                self.relayout_window()

            if want_icons:
                show_windows(desktop_icon_views(), False)
                self.icons_hidden = True
            elif self.icons_hidden:
                show_windows(desktop_icon_views(), True)
                self.icons_hidden = False
        except Exception as e:
            log_error("隐藏/恢复任务栏失败: %r" % e)

    def restore_shell_options(self):
        """外部 --restore 调用：关掉两个隐藏选项并恢复"""
        self.cfg["hide_taskbar"] = False
        self.cfg["hide_desktop_icons"] = False
        self.apply_shell()
        restore_shell()
        self.save_config()

    def periodic(self):
        self.check_fullscreen()
        # explorer 偶尔会自己把任务栏/图标显示回来，定时再藏一下
        if self.taskbar_hidden or self.icons_hidden:
            self.apply_shell()

    # ---------- 布局计算

    def visible_pinned(self):
        if self.dragging and self.drag_idx is not None:
            return [it for i, it in enumerate(self.items) if i != self.drag_idx]
        return list(self.items)

    def build_slots(self, with_gap=True):
        """[(类型, item)]，类型：item / gap（拖放占位）/ sep（分隔线）"""
        slots = [("item", it) for it in self.visible_pinned()]
        if with_gap and self.drop_idx is not None:
            slots.insert(min(self.drop_idx, len(slots)), ("gap", None))
        if self.running_items:
            if slots:
                slots.append(("sep", None))
            slots.extend(("item", it) for it in self.running_items)
        return slots

    def rest_layout(self, slots):
        widths = [self.sep_w if kind == "sep" else self.B for kind, _ in slots]
        total = sum(widths) + self.gap * max(len(slots) - 1, 0)
        x = (self.width() - total) / 2
        lefts = []
        for w in widths:
            lefts.append(x)
            x += w + self.gap
        return lefts, widths

    def layout(self):
        """返回 slots[(类型, item, x, 宽)]、bar 矩形、悬停区域"""
        B, gap, pad = self.B, self.gap, self.pad
        slots = self.build_slots()
        lefts, rws = self.rest_layout(slots)
        mx = self.mouse_pos.x() if self.mouse_pos is not None else None

        sizes, extra_left = [], 0.0
        for (kind, _), left, rw in zip(slots, lefts, rws):
            w = rw
            if kind != "sep" and mx is not None and self.hover_amt > 0 and self.M > 1:
                d = abs(mx - (left + rw / 2)) / (B + gap)
                if d < self.spread:
                    w = rw * (1 + (self.M - 1) * (math.cos(math.pi * d / self.spread) + 1) / 2 * self.hover_amt)
            sizes.append(w)
            if mx is not None:
                # 光标左侧的增量向左长，右侧的向右长 —— 保证光标下的图标不“跑掉”
                extra_left += (w - rw) * min(max((mx - left) / rw, 0.0), 1.0)

        result = []
        x = (lefts[0] if lefts else self.width() / 2) - extra_left
        for (kind, item), w in zip(slots, sizes):
            result.append((kind, item, x, w))
            x += w + gap

        hide_off = (1 - self.reveal) * (self.bar_h + MARGIN + 4)
        bottom = self.height() - MARGIN + hide_off
        if result:
            left, right = result[0][2] - pad, result[-1][2] + result[-1][3] + pad
        else:
            left, right = self.width() / 2 - B * 1.6, self.width() / 2 + B * 1.6
        bar = QRectF(left, bottom - self.bar_h, right - left, self.bar_h)
        zone_top = self.height() - MARGIN - pad - B * self.M - 10
        zone = QRectF(left, zone_top, right - left, self.height() - zone_top)
        return result, bar, zone

    def item_at(self, pos):
        slots, bar, zone = self.layout()
        if not zone.contains(pos) and not bar.contains(pos):
            return None
        for kind, item, x, w in slots:
            if kind == "item" and x - self.gap / 2 <= pos.x() < x + w + self.gap / 2:
                return item
        return None

    def calc_drop_idx(self, x):
        slots = self.build_slots(with_gap=False)
        lefts, rws = self.rest_layout(slots)
        return sum(1 for (kind, it), left, rw in zip(slots, lefts, rws)
                   if kind == "item" and it.pinned and x > left + rw / 2)

    def in_remove_zone(self, pos):
        bar_top = self.height() - MARGIN - self.bar_h
        return pos.y() < bar_top - self.B * 1.3

    # ---------- 动画

    def kick(self):
        if not self.anim.isActive():
            self._last_tick = time.monotonic()
            self.anim.start()
        self.update()

    def tick(self):
        now = time.monotonic()
        dt = min(now - self._last_tick, 0.05)
        self._last_tick = now
        k = min(1.0, dt * 14)

        active = self.hovered or self.dragging or self.ext_drag
        dragging_out = self.dragging and self.mouse_pos is not None and self.in_remove_zone(self.mouse_pos)
        target_h = 1.0 if active and not dragging_out else 0.0
        self.hover_amt += (target_h - self.hover_amt) * k
        if abs(self.hover_amt - target_h) < 0.004:
            self.hover_amt = target_h

        if self.cfg["auto_hide"] and not (active or self.menu_open) and now >= self.hide_after:
            target_r = 0.0
        else:
            target_r = 1.0
        self.reveal += (target_r - self.reveal) * min(1.0, dt * 11)
        if abs(self.reveal - target_r) < 0.004:
            self.reveal = target_r

        bouncing = any(now - it.bounce_start < BOUNCE_TIME for it in self.items + self.running_items) \
            or now - self.mascot_bounce < MASCOT_HOP
        waiting_hide = self.cfg["auto_hide"] and not active and now < self.hide_after
        self.update()
        if self.hover_amt == target_h and self.reveal == target_r and not bouncing and not waiting_hide:
            self.anim.stop()

    def set_hovered(self, value):
        if value == self.hovered:
            return
        self.hovered = value
        if value:
            self.hover_check.start()
        else:
            self.hover_check.stop()
            self.hide_after = time.monotonic() + 0.6
        self.kick()

    def check_hover(self):
        # 兜底：鼠标快速划出时 leaveEvent 偶尔收不到
        if self.dragging or self.menu_open:
            return
        pos = QPointF(self.mapFromGlobal(QCursor.pos()))
        _, bar, zone = self.layout()
        if zone.contains(pos) or bar.contains(pos):
            if self.mouse_pos != pos:
                self.mouse_pos = pos
                self.update()
        else:
            self.set_hovered(False)

    # ---------- 绘制

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        slots, bar, zone = self.layout()
        th = self.theme
        now = time.monotonic()
        B = self.B

        # 几乎透明的命中区域：分层窗口里 alpha=0 的像素会被鼠标穿透
        if self.hovered or self.dragging or self.ext_drag:
            p.fillRect(zone, QColor(0, 0, 0, 1))
        if self.cfg["auto_hide"] and self.reveal < 0.5:
            p.fillRect(QRectF(bar.left(), self.height() - 3, bar.width(), 3), QColor(0, 0, 0, 1))
        if self.reveal > 0.5:
            p.fillRect(bar, QColor(0, 0, 0, 1))     # 透明皮肤 / 低不透明度时也能接住鼠标

        hovered_item = None
        if self.reveal > 0.01:
            p.setOpacity(min(1.0, self.reveal * 1.5))
            path = bar_path(th, bar)
            paint_background(p, th, bar, self.skin_image())

            if not slots:
                p.setPen(qc(th["text"]))
                p.setFont(QFont("Microsoft YaHei UI", max(9, B // 6)))
                p.drawText(bar, Qt.AlignCenter, "把程序拖到这里")

            dpr = self.devicePixelRatioF()
            base_y = icon_base_y(th, bar, self.pad)
            ind_y = indicator_y(th, bar, self.pad)
            for kind, item, x, w in slots:
                if kind == "sep":
                    paint_separator(p, th, x + w / 2, bar, self.pad)
                    continue
                if item is None:
                    continue
                t = now - item.bounce_start
                bounce = abs(math.sin(math.pi * t / (BOUNCE_TIME / 2))) * B * 0.38 * (1 - t / BOUNCE_TIME) \
                    if 0 <= t < BOUNCE_TIME else 0.0
                rect = QRectF(x, base_y - w - bounce, w, w)
                hovered = (self.mouse_pos is not None and self.hovered and not self.dragging
                           and x - self.gap / 2 <= self.mouse_pos.x() < x + w + self.gap / 2)
                paint_item(p, th, item, rect, base_y, item.running, hovered and self.hover_amt > 0.5,
                           path, ind_y, B, dpr)
                if hovered:
                    hovered_item = (item, rect)

            mrect = self.mascot_rect(bar)
            if mrect is not None:
                pm = self.mascot_pixmap()
                p.drawPixmap(mrect, pm, QRectF(pm.rect()))
                if now < self.mascot_say_until:
                    # 她个子高、头顶没地方，气泡放在外侧肩膀旁边
                    outer = "left" if self.cfg.get("mascot_side") == "left" else "right"
                    self.draw_label(p, self.mascot_line, mrect, beside=outer)

            if hovered_item and self.hover_amt > 0.5:
                self.draw_label(p, hovered_item[0].name, hovered_item[1])

        # 正在拖动的图标
        if self.dragging and self.mouse_pos is not None:
            item = self.items[self.drag_idx]
            size = B * 1.1
            c = self.mouse_pos
            cy = max(size / 2, c.y())
            removing = self.in_remove_zone(c)
            p.setOpacity(0.55 if removing else 0.9)
            pm = item.pixmap_for(size * self.devicePixelRatioF())
            p.drawPixmap(QRectF(c.x() - size / 2, cy - size / 2, size, size), pm, QRectF(pm.rect()))
            p.setOpacity(1.0)
            if removing:
                self.draw_label(p, "松开移除", QRectF(c.x() - size / 2, cy - size / 2, size, size), red=True)
        p.end()

    def draw_label(self, p, text, icon_rect, red=False, beside=None):
        font = QFont("Microsoft YaHei UI")
        font.setPixelSize(max(12, round(self.B * 0.25)))
        p.setFont(font)
        fm = QFontMetricsF(font)
        tw, th = fm.horizontalAdvance(text) + 20, fm.height() + 8
        x = min(max(icon_rect.center().x() - tw / 2, 2), self.width() - tw - 2)
        y = icon_rect.top() - th - 8
        if beside:
            x = icon_rect.right() + 6 if beside == "right" else icon_rect.left() - tw - 6
            x = min(max(x, 2), self.width() - tw - 2)
            y = icon_rect.top() + icon_rect.height() * 0.12
        elif y < 2:    # 上方放不下就放到图标下面
            y = icon_rect.bottom() + 6
        r = QRectF(x, y, tw, th)
        skin = self.theme
        border = skin["label_border"]
        p.setPen(QPen(qc(border), 1) if border and not red else Qt.NoPen)
        p.setBrush(QColor(200, 45, 45, 230) if red else qc(skin["label_bg"]))
        p.drawRoundedRect(r, th / 2, th / 2)
        p.setPen(QColor(255, 255, 255) if red else qc(skin["label_fg"]))
        p.drawText(r, Qt.AlignCenter, text)

    # ---------- 鼠标

    def enterEvent(self, e):
        pos = e.position()
        _, bar, zone = self.layout()
        if bar.contains(pos) or (self.cfg["auto_hide"] and pos.y() >= self.height() - 4):
            self.mouse_pos = pos
            self.set_hovered(True)

    def leaveEvent(self, _):
        if not self.dragging and not self.menu_open:
            self.set_hovered(False)

    def mouseMoveEvent(self, e):
        pos = e.position()
        self.mouse_pos = pos
        if (self.press_item is not None and self.press_item.pinned and not self.dragging
                and (e.buttons() & Qt.LeftButton) and (pos - self.press_pos).manhattanLength() > 6):
            self.dragging = True
            self.drag_idx = self.items.index(self.press_item)
        if self.dragging:
            self.drop_idx = None if self.in_remove_zone(pos) else self.calc_drop_idx(pos.x())
        else:
            _, bar, zone = self.layout()
            self.set_hovered(bar.contains(pos) or (self.hovered and zone.contains(pos)))
        self.kick()

    def mousePressEvent(self, e):
        pos = e.position()
        mrect = self.mascot_rect(self.layout()[1])
        if e.button() == Qt.LeftButton and mrect is not None and mrect.contains(pos):
            self.poke_mascot()
            return
        item = self.item_at(pos)
        if e.button() == Qt.LeftButton:
            self.press_item, self.press_pos = item, pos
        elif e.button() == Qt.MiddleButton and item is not None:
            self.launch(item)
        elif e.button() == Qt.RightButton:
            self.show_menu(e.globalPosition().toPoint(), item)

    def mouseReleaseEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        pos = e.position()
        if self.dragging:
            item = self.items.pop(self.drag_idx)
            if not self.in_remove_zone(pos):
                self.items.insert(min(self.drop_idx or 0, len(self.items)), item)
            elif item.running and self.cfg["show_running"]:
                # 移除的程序还开着，就回到“正在运行”区
                item.pinned = False
                self.running_items.append(item)
            self.save_config()
            self.relayout_window()
        elif self.press_item is not None and self.item_at(pos) is self.press_item:
            self.activate(self.press_item)
        self.press_item = self.press_pos = None
        self.dragging = False
        self.drag_idx = self.drop_idx = None
        self.check_hover()
        self.kick()

    # ---------- 外部拖入

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.ext_drag = True
            self.mouse_pos = e.position()
            self.drop_idx = self.calc_drop_idx(e.position().x())
            self.kick()

    def dragMoveEvent(self, e):
        e.acceptProposedAction()
        self.mouse_pos = e.position()
        self.drop_idx = self.calc_drop_idx(e.position().x())
        self.kick()

    def dragLeaveEvent(self, _):
        self.ext_drag = False
        self.drop_idx = None
        self.hide_after = time.monotonic() + 0.6
        self.kick()

    def dropEvent(self, e):
        idx = self.drop_idx if self.drop_idx is not None else len(self.items)
        paths = []
        for url in e.mimeData().urls():
            paths.append(os.path.normpath(url.toLocalFile()) if url.isLocalFile() else url.toString())
        self.ext_drag = False
        self.drop_idx = None
        self.add_paths(paths, idx)
        e.acceptProposedAction()

    def add_paths(self, paths, index=None):
        existing = {os.path.normcase(it.path) for it in self.items}
        index = len(self.items) if index is None else index
        for path in paths:
            if not path or os.path.normcase(path) in existing:
                continue
            try:
                self.items.insert(index, DockItem(path))
                existing.add(os.path.normcase(path))
                index += 1
            except Exception as ex:
                log_error("添加失败 %s: %r" % (path, ex))
        self.save_config()
        self.monitor.poke()
        self.relayout_window()
        self.kick()

    # ---------- 正在运行的程序

    def owner_of(self, key, hwnd, pinned):
        """运行中的程序归哪个固定项（结果缓存，固定项变了就重新算）"""
        if key in pinned:
            return pinned[key]
        sig = tuple(sorted(pinned))
        if sig != self._owner_sig:
            self._owner_sig, self._owners = sig, {}
        if key not in self._owners:
            self._owners[key] = find_owner(key, hwnd, pinned)
        return pinned.get(self._owners[key])

    def on_windows_update(self, groups):
        pinned = {it.target: it for it in self.items if it.target}
        old_running = {it.target: it for it in self.running_items}
        new_running = []
        for it in pinned.values():
            it.hwnds = []
        owner_names = {}
        for key, path, hwnds in groups:
            owner = self.owner_of(key, hwnds[0], pinned)
            if owner is not None:
                owner.hwnds = owner.hwnds + hwnds
                owner_names[key] = owner.name
            elif self.cfg["show_running"]:
                it = old_running.get(key)
                if it is None:
                    image = extract_icon(path)
                    if image is None:
                        image = window_icon(hwnds[0])
                    it = DockItem(path, pinned=False, image=image)
                it.hwnds = hwnds
                new_running.append(it)
        # 已有的保持原来顺序，新开的排到最后
        order = {it.target: i for i, it in enumerate(self.running_items)}
        new_running.sort(key=lambda it: order.get(it.target, len(order)))
        count_changed = len(new_running) != len(self.running_items)
        self.running_items = new_running
        if count_changed:
            self.relayout_window()
        # 只有运行状态真的变了才重画（每秒重画整个分层窗口很费 CPU）
        state = (tuple(it.running for it in self.items), tuple(it.target for it in new_running))
        if state != self._last_state:
            self._last_state = state
            self.update()

        now = time.monotonic()
        names = {it.target: it.name for it in new_running}
        names.update({t: it.name for t, it in pinned.items()})
        names.update(owner_names)
        current = {key: names.get(key) or display_name(path) for key, path, _ in groups}
        if self._seen_apps is None:            # 启动时已经开着的不算“新打开”
            self._seen_apps = dict.fromkeys(current, now)
            self._open_apps = dict(current)
            return

        # 新打开：关掉 5 秒内又打开的不重复报
        opened = [name for key, name in current.items()
                  if key not in self._seen_apps or now - self._seen_apps[key] > 5]
        self._seen_apps.update(dict.fromkeys(current, now))

        # 关掉：窗口全部消失超过 2 秒才算（软件重启、切换窗口时会短暂没有窗口）
        closed = []
        for key in list(self._open_apps):
            if key in current:
                self._gone_since.pop(key, None)
            elif now - self._gone_since.setdefault(key, now) >= 2:
                closed.append(self._open_apps.pop(key))
                self._gone_since.pop(key, None)
        self._open_apps.update(current)

        self.announce(opened, "mascot_announce", "mascot_open_lines", "打开了 {name}。")
        self.announce(closed, "mascot_announce_close", "mascot_close_lines", "{name}……关掉了。")

    def activate(self, item):
        pinned = {it.target: it for it in self.items if it.target}
        wins = [h for h, key, _ in enum_app_windows()
                if item.target and (key == item.target or self.owner_of(key, h, pinned) is item)]
        if wins:
            fg = win32gui.GetForegroundWindow()
            if fg in wins:
                if len(wins) == 1:
                    win32gui.ShowWindow(fg, win32con.SW_MINIMIZE)
                    return
                focus_window(wins[(wins.index(fg) + 1) % len(wins)])
            else:
                focus_window(wins[0])
            return
        self.launch(item)

    def launch(self, item, admin=False):
        item.bounce_start = time.monotonic()
        self.kick()
        path, args, name = item.path, item.args, item.name
        cwd = os.path.dirname(path) if path.lower().endswith(".exe") else None

        def run():
            pythoncom.CoInitialize()
            try:
                if admin:
                    r = ctypes.windll.shell32.ShellExecuteW(None, "runas", path, args or None, cwd, 1)
                    if r <= 32 and r != 5:      # 5 = 用户取消了 UAC
                        raise OSError("ShellExecute 错误码 %d" % r)
                elif not is_web_url(path) and not os.path.exists(path):
                    raise FileNotFoundError("文件不存在：" + path)
                else:
                    os.startfile(path, arguments=args, cwd=cwd)
            except Exception as ex:
                self.launch_failed.emit(name, str(ex))
            finally:
                pythoncom.CoUninitialize()
            self.monitor.poke()

        threading.Thread(target=run, daemon=True).start()

    def close_windows(self, item):
        for h in item.hwnds:
            try:
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
        self.monitor.poke()

    def pin(self, item):
        if item in self.running_items:
            self.running_items.remove(item)
        item.pinned = True
        self.items.append(item)
        self.save_config()
        self.relayout_window()

    # ---------- 右键菜单

    def show_menu(self, gpos, item):
        self.menu_open = True
        m = QMenu()
        m.setStyleSheet(MENU_QSS)
        if item is not None:
            m.addAction("打开" if not item.running else "切换到窗口", lambda: self.activate(item))
            m.addAction("启动新实例", lambda: self.launch(item))
            if item.target or item.path.lower().endswith((".exe", ".lnk", ".bat", ".cmd")):
                m.addAction("以管理员身份运行", lambda: self.launch(item, admin=True))
            loc = m.addAction("打开文件所在位置", lambda: self.open_location(item))
            loc.setEnabled(not is_web_url(item.path))
            if item.running:
                m.addAction("关闭窗口" if len(item.hwnds) == 1 else "关闭全部 %d 个窗口" % len(item.hwnds),
                            lambda: self.close_windows(item))
            m.addSeparator()
            if item.pinned:
                m.addAction("重命名…", lambda: self.rename(item))
                m.addAction("从 Dock 移除", lambda: self.remove(item))
            else:
                m.addAction("固定到 Dock", lambda: self.pin(item))
            m.addSeparator()
        m.addAction("添加程序 / 文件…", self.add_file_dialog)
        m.addAction("添加文件夹…", self.add_folder_dialog)
        self.build_settings_menu(m.addMenu("设置"))
        m.addSeparator()
        m.addAction("隐藏 Dock（托盘可恢复）", lambda: self.set_user_hidden(True))
        m.addAction("退出", QApplication.quit)
        m.exec(gpos)
        self.menu_open = False
        self.check_hover()
        self.kick()

    def build_settings_menu(self, sm):
        sm.setStyleSheet(MENU_QSS)

        def radio(title, key, options):
            sub = sm.addMenu(title)
            sub.setStyleSheet(MENU_QSS)
            group = QActionGroup(sub)
            for label, value in options:
                a = sub.addAction(label)
                a.setCheckable(True)
                a.setChecked(self.cfg[key] == value)
                group.addAction(a)
                a.triggered.connect(lambda _=False, v=value: self.set_option(key, v))

        def toggle(title, checked, fn):
            a = sm.addAction(title)
            a.setCheckable(True)
            a.setChecked(checked)
            a.triggered.connect(fn)

        def option(title, key):
            toggle(title, self.cfg[key], lambda c: self.set_option(key, c))

        radio("图标大小", "icon_size", [("小", 40), ("中", 52), ("大", 64), ("特大", 80)])
        radio("放大效果", "magnify", [("关闭", 1.0), ("轻微", 1.4), ("标准", 1.8), ("夸张", 2.3)])
        skins = [(s["name"], sid) for sid, s in SKINS.items()]
        if self.cfg.get("skin_color"):
            skins.append(("自定义颜色", CUSTOM_COLOR))
        if self.cfg.get("skin_image"):
            skins.append(("图片皮肤", CUSTOM_IMAGE))
        radio("切换皮肤", "skin", skins)
        sm.addAction("皮肤中心…", self.open_skin_center)
        sm.addSeparator()
        option("显示正在运行的程序", "show_running")
        option("隐藏 Windows 任务栏", "hide_taskbar")
        hk = self.cfg["hotkeys"].get("desktop_icons")
        option("隐藏桌面图标" + ("\t" + hk if hk else ""), "hide_desktop_icons")
        sm.addAction("快捷键设置…", self.edit_hotkeys)
        sm.addSeparator()
        option("自动隐藏 Dock", "auto_hide")
        option("全屏程序时隐藏 Dock", "hide_on_fullscreen")
        toggle("开机自启", autostart_enabled(), set_autostart)

    def open_skin_center(self):
        dlg = SkinCenter(self)
        dlg.show()
        dlg.activateWindow()
        dlg.exec()

    def open_location(self, item):
        path = item.path
        if path.lower().endswith(".lnk"):
            target = read_lnk(path)[0]
            if target and os.path.exists(target):
                path = target
        if os.path.exists(path):
            subprocess.Popen(["explorer", "/select,", path])

    def rename(self, item):
        dlg = QInputDialog()
        dlg.setWindowFlag(Qt.WindowStaysOnTopHint)
        dlg.setWindowTitle("重命名")
        dlg.setLabelText("显示名称：")
        dlg.setTextValue(item.name)
        dlg.show()
        dlg.activateWindow()
        if dlg.exec() and dlg.textValue().strip():
            item.name = dlg.textValue().strip()
            self.save_config()

    def remove(self, item):
        if item in self.items:
            self.items.remove(item)
            self.save_config()
            self.monitor.poke()
            self.relayout_window()

    def add_file_dialog(self):
        start = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs")
        files, _ = QFileDialog.getOpenFileNames(None, "选择要添加到 Dock 的程序或文件", start,
                                                "程序和快捷方式 (*.exe *.lnk *.url *.bat *.cmd);;所有文件 (*)")
        if files:
            self.add_paths([os.path.normpath(f) for f in files])

    def add_folder_dialog(self):
        d = QFileDialog.getExistingDirectory(None, "选择要添加到 Dock 的文件夹")
        if d:
            self.add_paths([os.path.normpath(d)])

    # ---------- 显示 / 隐藏

    def set_user_hidden(self, hidden):
        self.user_hidden = hidden
        self.sync_visible()
        self.apply_shell()

    def sync_visible(self):
        should_show = not self.user_hidden and not self.fs_hidden
        if should_show and not self.isVisible():
            self.show()
        elif not should_show and self.isVisible():
            self.hide()

    def check_fullscreen(self):
        fs = False
        if self.cfg["hide_on_fullscreen"]:
            try:
                hwnd = win32gui.GetForegroundWindow()
                if hwnd and win32process.GetWindowThreadProcessId(hwnd)[1] != os.getpid() \
                        and win32gui.GetClassName(hwnd) not in SKIP_CLASSES:
                    l, t, r, b = win32gui.GetWindowRect(hwnd)
                    mon = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
                    ml, mt, mr, mb = win32api.GetMonitorInfo(mon)["Monitor"]
                    mine = win32api.MonitorFromWindow(int(self.winId()), win32con.MONITOR_DEFAULTTONEAREST)
                    fs = mon == mine and l <= ml and t <= mt and r >= mr and b >= mb
            except Exception:
                fs = False
        if fs != self.fs_hidden:
            self.fs_hidden = fs
            self.sync_visible()

    def showEvent(self, e):
        super().showEvent(e)
        # 点击 Dock 不抢焦点，这样“再点一次最小化”才能判断出前台窗口
        hwnd = int(self.winId())
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex | WS_EX_NOACTIVATE)

    def shutdown(self):
        # 退出时一定把任务栏和桌面图标还回去
        restore_shell()
        self.save_config()
        for i in range(1, len(HOTKEY_ACTIONS) + 1):
            user32.UnregisterHotKey(int(self.winId()), i)
        self.monitor.requestInterruption()
        self.monitor.poke()
        self.monitor.wait(3000)


# ---------------------------------------------------------------- 皮肤中心

SKIN_CENTER_QSS = """
QDialog { background: #1d1d22; }
QLabel { color: #d8d8de; font-size: 13px; }
QLabel#title { color: #ffffff; font-size: 18px; font-weight: bold; }
QPushButton { background: #303038; color: #eee; border: 1px solid #45454e; border-radius: 7px;
              padding: 7px 16px; font-size: 13px; }
QPushButton:hover { background: #3b6cf0; border-color: #3b6cf0; }
QCheckBox { color: #d8d8de; font-size: 13px; }
QLineEdit { background: #303038; color: #fff; border: 1px solid #45454e; border-radius: 6px; padding: 6px 8px;
            font-size: 13px; }
QLineEdit:focus { border-color: #3b6cf0; }
QComboBox { background: #303038; color: #eee; border: 1px solid #45454e; border-radius: 7px;
            padding: 6px 10px; font-size: 13px; }
QComboBox QAbstractItemView { background: #2a2a30; color: #eee; selection-background-color: #3b6cf0; }
QSlider::groove:horizontal { height: 4px; background: #44444c; border-radius: 2px; }
QSlider::sub-page:horizontal { background: #3b6cf0; border-radius: 2px; }
QSlider::handle:horizontal { width: 16px; margin: -6px 0; border-radius: 8px; background: #ffffff; }
"""


class SkinCard(QWidget):
    """一张皮肤预览卡：示例壁纸 + 用真实图标画出的小 Dock"""
    clicked = Signal(str)

    def __init__(self, center, skin_id):
        super().__init__()
        self.center = center
        self.skin_id = skin_id
        self.hover = False
        self.setFixedSize(250, 120)
        self.setCursor(Qt.PointingHandCursor)

    def enterEvent(self, _):
        self.hover = True
        self.update()

    def leaveEvent(self, _):
        self.hover = False
        self.update()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit(self.skin_id)

    def paintEvent(self, _):
        dock = self.center.dock
        cfg = dock.cfg
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        card = QRectF(self.rect()).adjusted(3, 3, -3, -3)
        clip = QPainterPath()
        clip.addRoundedRect(card, 12, 12)
        wall = QLinearGradient(card.topLeft(), card.bottomRight())
        wall.setColorAt(0, QColor("#3f5fa8"))
        wall.setColorAt(0.55, QColor("#9c78b8"))
        wall.setColorAt(1, QColor("#f2a66a"))
        p.fillPath(clip, wall)
        p.setClipPath(clip)

        th = build_skin(self.skin_id, cfg)
        B, gap = 34, 6
        pad = round(B * 0.2)
        items = self.center.samples
        sizes = [B * 1.3 if i == 2 else B for i in range(len(items))]
        total = sum(sizes) + gap * max(len(items) - 1, 0) + 2 * pad
        bar = QRectF(card.center().x() - total / 2, card.bottom() - 30 - (B + 2 * pad), total, B + 2 * pad)

        image = None
        missing = False
        if self.skin_id == CUSTOM_IMAGE:
            path = cfg.get("skin_image") or ""
            img = QImage(path) if os.path.exists(path) else QImage()
            missing = img.isNull()
            if not missing:
                image = QPixmap.fromImage(img.scaledToHeight(min(160, img.height()), Qt.SmoothTransformation))
        paint_background(p, th, bar, image)
        path = bar_path(th, bar)
        base_y, ind_y = icon_base_y(th, bar, pad), indicator_y(th, bar, pad)
        x = bar.left() + pad
        dpr = self.devicePixelRatioF()
        for i, (item, w) in enumerate(zip(items, sizes)):
            paint_item(p, th, item, QRectF(x, base_y - w, w, w), base_y, i in (0, 2), i == 2, path, ind_y, B, dpr)
            x += w + gap

        # 底部名称条
        strip = QRectF(card.left(), card.bottom() - 24, card.width(), 24)
        p.fillRect(strip, QColor(0, 0, 0, 110))
        font = QFont("Microsoft YaHei UI")
        font.setPixelSize(13)
        p.setFont(font)
        p.setPen(QColor(255, 255, 255))
        name = th["name"]
        if self.skin_id == CUSTOM_COLOR and not cfg.get("skin_color"):
            name += "（点击选色）"
        if missing:
            name += "（点击选择图片）"
        selected = cfg["skin"] == self.skin_id
        p.drawText(strip.adjusted(10, 0, -10, 0), Qt.AlignVCenter | Qt.AlignLeft, ("✓  " if selected else "") + name)
        p.setClipping(False)
        if selected or self.hover:
            p.setPen(QPen(QColor("#3b6cf0") if selected else QColor(255, 255, 255, 120), 2.5 if selected else 1.5))
            p.setBrush(Qt.NoBrush)
            p.drawRoundedRect(card.adjusted(1, 1, -1, -1), 11, 11)
        p.end()


class SkinCenter(QDialog):
    def __init__(self, dock):
        super().__init__()
        self.dock = dock
        self.setWindowTitle("皮肤中心 - " + APP_NAME)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.setStyleSheet(SKIN_CENTER_QSS)

        self.samples = (dock.items + dock.running_items)[:5]
        if len(self.samples) < 5:
            win = os.environ.get("WINDIR", r"C:\Windows")
            for exe in ("explorer.exe", r"System32\notepad.exe", r"System32\calc.exe", r"System32\mspaint.exe",
                        r"System32\Taskmgr.exe"):
                if len(self.samples) >= 5:
                    break
                if os.path.exists(os.path.join(win, exe)):
                    self.samples.append(DockItem(os.path.join(win, exe)))

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(12)
        title = QLabel("皮肤中心")
        title.setObjectName("title")
        lay.addWidget(title)
        lay.addWidget(QLabel("点一下卡片立即换上，Dock 会实时变化。"))

        grid = QGridLayout()
        grid.setSpacing(10)
        self.cards = []
        for i, sid in enumerate(list(SKINS) + [CUSTOM_COLOR, CUSTOM_IMAGE]):
            card = SkinCard(self, sid)
            card.clicked.connect(self.choose)
            grid.addWidget(card, i // 4, i % 4)
            self.cards.append(card)
        lay.addLayout(grid)

        row = QHBoxLayout()
        row.addWidget(QLabel("背景不透明度"))
        self.slider = QSlider(Qt.Horizontal)
        self.slider.setRange(20, 100)
        self.slider.setValue(int(dock.cfg.get("skin_opacity", 1.0) * 100))
        self.value_label = QLabel("%d%%" % self.slider.value())
        self.value_label.setFixedWidth(40)
        self.slider.valueChanged.connect(self.on_opacity)
        row.addWidget(self.slider, 1)
        row.addWidget(self.value_label)
        lay.addLayout(row)

        # 看板娘
        cfg = dock.cfg
        mrow = QHBoxLayout()
        mrow.addWidget(QLabel("看板娘"))
        b_mascot = QPushButton("选择图片…")
        b_mascot.clicked.connect(self.pick_mascot)
        b_builtin = QPushButton("内置小人")
        b_builtin.setToolTip("原创 Q 版小人，抱着黄瓜")
        b_builtin.clicked.connect(self.use_builtin_mascot)
        b_builtin.setEnabled(os.path.exists(BUILTIN_MASCOT))
        b_lines = QPushButton("编辑台词…")
        b_lines.clicked.connect(self.edit_lines)
        b_remove = QPushButton("移除")
        b_remove.clicked.connect(lambda: self.dock.set_option("mascot_image", ""))
        side = QComboBox()
        side.addItems(["站在右边", "站在左边"])
        side.setCurrentIndex(1 if cfg.get("mascot_side") == "left" else 0)
        side.currentIndexChanged.connect(lambda i: self.dock.set_option("mascot_side", "left" if i else "right"))
        size = QComboBox()
        sizes = [("小", 2.0), ("中", 2.6), ("大", 3.2)]
        size.addItems([s[0] for s in sizes])
        size.setCurrentIndex(min(range(3), key=lambda i: abs(sizes[i][1] - float(cfg.get("mascot_size", 2.6)))))
        size.currentIndexChanged.connect(lambda i: self.dock.set_option("mascot_size", sizes[i][1]))
        for wdg in (b_builtin, b_mascot, b_lines, side, size, b_remove):
            mrow.addWidget(wdg)
        mrow.addStretch(1)
        lay.addLayout(mrow)
        arow = QHBoxLayout()
        arow.addSpacing(52)
        for title, key, lines_key, what in (("打开软件时报一声", "mascot_announce", "mascot_open_lines", "打开"),
                                            ("关掉软件时也说", "mascot_announce_close", "mascot_close_lines", "关掉")):
            box = QCheckBox(title)
            box.setChecked(cfg.get(key, True))
            box.toggled.connect(lambda c, k=key: self.dock.set_option(k, c))
            btn = QPushButton("编辑…")
            btn.clicked.connect(lambda _=False, k=lines_key, w=what: self.edit_announce_lines(k, w))
            arow.addWidget(box)
            arow.addWidget(btn)
            arow.addSpacing(18)
        arow.addStretch(1)
        lay.addLayout(arow)
        tip = QLabel("建议用透明背景的 PNG 立绘；点 Dock 旁边的她会跳一下并说一句台词，开关软件时她也会告诉你。")
        tip.setStyleSheet("color: #8a8a94; font-size: 12px;")
        lay.addWidget(tip)

        buttons = QHBoxLayout()
        b_color = QPushButton("自定义颜色…")
        b_color.clicked.connect(self.pick_color)
        b_image = QPushButton("选择背景图片…")
        b_image.clicked.connect(self.pick_image)
        b_close = QPushButton("完成")
        b_close.clicked.connect(self.accept)
        buttons.addWidget(b_color)
        buttons.addWidget(b_image)
        buttons.addStretch(1)
        buttons.addWidget(b_close)
        lay.addLayout(buttons)

    def refresh(self):
        for c in self.cards:
            c.update()

    def choose(self, sid):
        if sid == CUSTOM_COLOR and not self.dock.cfg.get("skin_color"):
            return self.pick_color()
        if sid == CUSTOM_IMAGE and not os.path.exists(self.dock.cfg.get("skin_image") or ""):
            return self.pick_image()
        self.dock.set_option("skin", sid)
        self.refresh()

    def on_opacity(self, v):
        self.value_label.setText("%d%%" % v)
        self.dock.cfg["skin_opacity"] = v / 100
        self.dock.theme = build_skin(self.dock.cfg["skin"], self.dock.cfg)
        self.dock.update()
        self.refresh()
        self._save_later()

    def _save_later(self):
        # 拖滑块时不要每一格都写一次文件
        if not hasattr(self, "_save_timer"):
            self._save_timer = QTimer(self, singleShot=True, interval=400, timeout=self.dock.save_config)
        self._save_timer.start()

    def pick_color(self):
        old = QColor(self.dock.cfg.get("skin_color") or "#6a5acd")
        c = QColorDialog.getColor(old, self, "选择 Dock 主色")
        if c.isValid():
            self.dock.cfg["skin_color"] = c.name()
            self.dock.set_option("skin", CUSTOM_COLOR)
            self.refresh()

    def pick_image(self):
        start = os.path.join(os.path.expanduser("~"), "Pictures")
        f, _ = QFileDialog.getOpenFileName(self, "选择 Dock 背景图片", start,
                                           "图片 (*.png *.jpg *.jpeg *.bmp *.webp *.gif)")
        if not f:
            return
        if QImage(f).isNull():
            QMessageBox.warning(self, APP_NAME, "这张图片打不开，换一张试试。")
            return
        self.dock.cfg["skin_image"] = copy_to_skins(f, "background")
        self.dock.set_option("skin", CUSTOM_IMAGE)
        self.refresh()

    def pick_mascot(self):
        start = os.path.join(os.path.expanduser("~"), "Pictures")
        f, _ = QFileDialog.getOpenFileName(self, "选择看板娘图片（建议透明背景 PNG）", start,
                                           "图片 (*.png *.webp *.gif *.jpg *.jpeg *.bmp)")
        if not f:
            return
        if QImage(f).isNull():
            QMessageBox.warning(self, APP_NAME, "这张图片打不开，换一张试试。")
            return
        self.dock.set_option("mascot_image", copy_to_skins(f, "mascot"))
        self.dock.poke_mascot()

    def use_builtin_mascot(self):
        # 复制出来再用：exe 的解压目录每次启动可能不一样
        self.dock.set_option("mascot_image", copy_to_skins(BUILTIN_MASCOT, "mascot"))
        self.dock.poke_mascot()

    def edit_lines(self):
        text, ok = QInputDialog.getMultiLineText(self, "编辑台词", "点看板娘时随机说一句，每行一句：",
                                                 "\n".join(self.dock.cfg.get("mascot_lines") or []))
        if ok:
            self.dock.set_option("mascot_lines", [s.strip() for s in text.splitlines() if s.strip()])

    def edit_announce_lines(self, key, what):
        text, ok = QInputDialog.getMultiLineText(
            self, "编辑提醒语", "%s软件时随机说一句，每行一句，{name} 会换成软件名：" % what,
            "\n".join(self.dock.cfg.get(key) or []))
        if not ok:
            return
        lines = [s.strip() for s in text.splitlines() if s.strip()]
        bad = [s for s in lines if "{name}" not in s]
        if bad:
            QMessageBox.information(self, APP_NAME, "这几句没有 {name}，不知道该把软件名放哪，已跳过：\n" + "\n".join(bad))
        self.dock.set_option(key, [s for s in lines if "{name}" in s])


def copy_to_skins(src, basename):
    """复制一份到 skins 目录，原图删了/挪了皮肤也不会丢"""
    skins_dir = os.path.join(APP_DIR, "skins")
    os.makedirs(skins_dir, exist_ok=True)
    dst = os.path.join(skins_dir, basename + os.path.splitext(src)[1].lower())
    try:
        if os.path.normcase(os.path.abspath(src)) != os.path.normcase(dst):
            for old in os.listdir(skins_dir):
                if old.startswith(basename + "."):
                    os.remove(os.path.join(skins_dir, old))
            shutil.copy2(src, dst)
        return dst
    except OSError as e:
        log_error("复制图片失败: %r" % e)
        return src


# ---------------------------------------------------------------- 开机自启

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_NAME = "NanmuDock"


def autostart_command():
    if getattr(sys, "frozen", False):
        return '"%s"' % sys.executable
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return '"%s" "%s"' % (pythonw, os.path.abspath(__file__))


def autostart_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, RUN_NAME)
            return True
    except OSError:
        return False


def set_autostart(enable):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            if enable:
                winreg.SetValueEx(k, RUN_NAME, 0, winreg.REG_SZ, autostart_command())
            else:
                try:
                    winreg.DeleteValue(k, RUN_NAME)
                except FileNotFoundError:
                    pass
    except OSError as e:
        log_error("设置开机自启失败: %r" % e)


# ---------------------------------------------------------------- 入口

def make_app_icon(size=64):
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 64, size / 64)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(40, 40, 48, 235))
    p.drawRoundedRect(QRectF(2, 30, 60, 26), 9, 9)
    for i, (c, h) in enumerate([("#3b82f6", 16), ("#22c55e", 24), ("#f59e0b", 16)]):
        p.setBrush(QColor(c))
        p.drawRoundedRect(QRectF(8 + i * 17, 52 - h - 2, 14, h), 4, 4)
    p.end()
    return QIcon(pm)


def send_to_running(msg):
    sock = QLocalSocket()
    sock.connectToServer(INSTANCE_KEY)
    if not sock.waitForConnected(300):
        return False
    sock.write(msg)
    sock.waitForBytesWritten(300)
    sock.disconnectFromServer()
    return True


def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setApplicationName(APP_NAME)

    if "--restore" in sys.argv:
        if not send_to_running(b"restore"):
            cfg, _ = load_config()
            cfg["hide_taskbar"] = False
            cfg["hide_desktop_icons"] = False
            restore_shell()
            write_config(cfg)
        else:
            time.sleep(0.5)
            restore_shell()
        QMessageBox.information(None, APP_NAME, "任务栏和桌面图标已恢复。")
        return 0

    # 单实例：已在运行就通知它显示出来
    if send_to_running(b"show"):
        return 0
    QLocalServer.removeServer(INSTANCE_KEY)
    server = QLocalServer()
    server.listen(INSTANCE_KEY)

    icon = make_app_icon()
    app.setWindowIcon(icon)
    tray = QSystemTrayIcon(icon)
    tray.setToolTip(APP_NAME)

    def tray_message(title, text):
        tray.showMessage(title, text, QSystemTrayIcon.Warning, 4000)

    cfg, first_run = load_config()
    if first_run:
        cfg["items"] = [{"path": p} for p in default_item_paths()]
    dock = Dock(cfg, tray_message)
    if first_run:
        dock.save_config()

    def on_connection():
        conn = server.nextPendingConnection()

        def read():
            if b"restore" in bytes(conn.readAll()):
                dock.restore_shell_options()
            else:
                dock.set_user_hidden(False)
        conn.readyRead.connect(read)

    server.newConnection.connect(on_connection)

    tmenu = QMenu()
    tmenu.setStyleSheet(MENU_QSS)
    tmenu.addAction("显示 Dock", lambda: dock.set_user_hidden(False))
    tmenu.addAction("隐藏 Dock", lambda: dock.set_user_hidden(True))
    tmenu.addAction("添加程序 / 文件…", dock.add_file_dialog)
    tmenu.addAction("皮肤中心…", dock.open_skin_center)
    tmenu.addAction("快捷键设置…", dock.edit_hotkeys)
    tmenu.addSeparator()
    tmenu.addAction("退出", app.quit)
    tray.setContextMenu(tmenu)
    tray.activated.connect(lambda reason: dock.set_user_hidden(dock.isVisible())
                           if reason == QSystemTrayIcon.Trigger else None)
    tray.show()

    dock.show()
    app.aboutToQuit.connect(dock.shutdown)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
