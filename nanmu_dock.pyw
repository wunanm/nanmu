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
import gc
import re
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

from PySide6.QtCore import (Qt, QTimer, QRectF, QPointF, QRect, QSize, QFileInfo, QThread, Signal,
                            QVariantAnimation)
from PySide6.QtGui import (QPainter, QColor, QPainterPath, QPixmap, QImage, QIcon, QFont, QKeySequence,
                           QFontMetricsF, QCursor, QPen, QBrush, QLinearGradient, QRadialGradient, QTransform,
                           QActionGroup, QGuiApplication, QFontDatabase)
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
from win32com.shell import shell, shellcon

# “正在播放”用 Windows 的媒体控制接口（pip install winrt-Windows.Media.Control 等），没装就不显示这个功能
try:
    import asyncio
    from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager as MediaManager
    from winrt.windows.storage.streams import DataReader
    HAVE_MEDIA = True
except Exception:
    HAVE_MEDIA = False

APP_NAME = "楠木 Dock"
APP_VERSION = "1.4.1"
INSTANCE_KEY = "NanmuDock_SingleInstance"
APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
CONFIG_PATH = os.path.join(APP_DIR, "dock_config.json")
LINKS_DIR = os.path.join(APP_DIR, "links")
LOG_PATH = os.path.join(APP_DIR, "dock_error.log")
# 随程序发布的资源；打包成 exe 后在 PyInstaller 的解压目录里
ASSETS_DIR = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "assets")
BUILTIN_MASCOT = os.path.join(ASSETS_DIR, "mascot_default.png")

DEFAULT_CONFIG = {
    "position": "bottom",     # Dock 位置：bottom / top / left / right
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
    "notify_flash": True,     # 程序闪烁（微信 / QQ 来消息）时图标跳一下、挂红点
    "mascot_notify": True,    # 来消息时看板娘也说一句
    "mascot_notify_lines": ["{name}……有新消息。", "……{name}在叫你。", "{name}，有人找你。"],
    "mascot_breathe": True,   # 看板娘待机呼吸动画
    "mascot_greet": True,     # 按时间段问候
    "remind_water": 0,        # 喝水提醒间隔（分钟），0 = 关
    "remind_game": 120,       # 连续全屏（玩游戏 / 看视频）多久提醒休息（分钟），0 = 关
    "show_media": True,       # Dock 右端的“正在播放”
    "show_clock": True,       # Dock 右端的时钟
    "show_power": True,       # 关机键（锁定 / 睡眠 / 注销 / 重启 / 关机）
    "show_volume": True,      # 音量按钮
    "show_tray_button": True,  # 托盘按钮（隐藏任务栏时才显示）
    "auto_hide": False,
    "hide_on_fullscreen": True,
    "show_running": True,     # 显示没固定的正在运行的程序
    "window_preview": "multi",  # 鼠标停在图标上时预览窗口：off / multi（多个窗口时）/ all
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


# ---------------------------------------------------------------- 系统项目（回收站、此电脑…）
# 这些是 Windows 的“虚拟文件夹”，没有真实路径，用 shell:xxx 或 ::{CLSID} 表示

RECYCLE_BIN = "shell:RecycleBinFolder"
RECYCLE_BIN_CLSID = "::{645FF040-5081-101B-9F08-00AA002F954E}"
SYSTEM_ITEMS = [("回收站", RECYCLE_BIN), ("此电脑", "shell:MyComputerFolder"),
                ("控制面板", "::{26EE0668-A00A-44D7-9371-BEB064C98683}"), ("网络", "shell:NetworkPlacesFolder"),
                ("下载", "shell:Downloads")]
IMAGERES = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "imageres.dll")


def is_shell_path(path):
    return path.startswith("::") or path.lower().startswith("shell:")


def launch_target(path):
    return "shell:" + path if path.startswith("::") else path


def shell_pidl(path):
    try:
        return shell.SHGetDesktopFolder().ParseDisplayName(0, None, launch_target(path), 0)[1]
    except Exception:
        try:      # ::{CLSID} 形式不带 shell: 前缀也试一次
            return shell.SHGetDesktopFolder().ParseDisplayName(0, None, path, 0)[1]
        except Exception:
            return None


def shell_name(path, kind=None):
    pidl = shell_pidl(path)
    if pidl is None:
        return None
    try:
        return shell.SHGetNameFromIDList(pidl, kind if kind is not None else shellcon.SIGDN_NORMALDISPLAY)
    except Exception:
        return None


def is_recycle_bin(path):
    return is_shell_path(path) and (path.lower() == RECYCLE_BIN.lower()
                                    or shell_name(path, shellcon.SIGDN_DESKTOPABSOLUTEPARSING) == RECYCLE_BIN_CLSID)


def recycle_bin_full():
    try:
        return shell.SHQueryRecycleBin(None)[1] > 0
    except Exception:
        return False


def shell_icon_image(path):
    """系统项目的图标：先查图标所在的文件和序号取高清版，不行就要一个普通大小的"""
    if is_recycle_bin(path):
        return extract_icon(IMAGERES, -54 if recycle_bin_full() else -55)
    pidl = shell_pidl(path)
    if pidl is None:
        return None
    try:
        _, (_, idx, _, icon_file, _) = shell.SHGetFileInfo(pidl, 0, shellcon.SHGFI_PIDL | shellcon.SHGFI_ICONLOCATION)
        img = extract_icon(os.path.expandvars(icon_file), idx) if icon_file else None
        if img is not None:
            return img
        _, (hicon, _, _, _, _) = shell.SHGetFileInfo(pidl, 0, shellcon.SHGFI_PIDL | shellcon.SHGFI_ICON |
                                                     shellcon.SHGFI_LARGEICON)
        if hicon:
            try:
                img = QImage.fromHICON(hicon)
            finally:
                user32.DestroyIcon(hicon)
            return None if img.isNull() else img
    except Exception:
        pass
    return None


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
    if is_shell_path(path):
        img = shell_icon_image(path)
        if img is not None:
            return img
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
    if is_shell_path(path):
        return shell_name(path) or path
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


def find_app_root(key, path, hwnd):
    """窗口属于某个程序在自己安装目录里启动的子组件时（WeGame 的 browser.exe），
    返回主程序 (规范路径, 原始路径)，这样 Dock 上显示的是 WeGame 本体；否则原样返回"""
    if "\\steamapps\\" in key:
        return key, path
    procs = parent_map()
    pid = win32process.GetWindowThreadProcessId(hwnd)[1]
    root, root_path = key, path
    for _ in range(3):
        ppid, _ = procs.get(pid, (0, ""))
        if not ppid or ppid == pid or ppid not in procs or procs[ppid][1] in LAUNCHER_PARENTS:
            break
        parent_path = process_path(ppid) or ""
        parent, pdir = os.path.normcase(parent_path), os.path.dirname(os.path.normcase(parent_path))
        folder = os.path.dirname(root)
        if not parent_path or pdir in SHARED_DIRS or not (folder == pdir or folder.startswith(pdir + "\\")):
            break
        root, root_path, pid = parent, parent_path, ppid
    return root, root_path


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


def minimize_window(hwnd):
    # 用“标题栏最小化按钮”的系统命令：有些程序（WeGame 等 CEF 窗口）不理 ShowWindow
    win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_MINIMIZE, 0)


def focus_window(hwnd):
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            if win32gui.IsIconic(hwnd):     # 不理 ShowWindow 的程序，改发“还原按钮”的系统命令
                win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
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
    recycle_state = Signal(bool)  # 回收站 空 / 有东西 变了

    def __init__(self):
        super().__init__()
        self._poke = threading.Event()
        self.watch_recycle = lambda: False      # Dock 上有回收站图标时才查
        self._recycle_full = None

    def poke(self):
        self._poke.set()

    def run(self):
        tick, poked = 0, False
        while not self.isInterruptionRequested():
            tick += 1
            # 回收站状态每 3 秒查一次（被 poke 时立刻查，比如刚删了文件）
            if self.watch_recycle() and (tick % 3 == 0 or self._recycle_full is None or poked):
                full = recycle_bin_full()
                if full != self._recycle_full:
                    self._recycle_full = full
                    self.recycle_state.emit(full)
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
            poked = self._poke.wait(1.0)
            if poked:
                self._poke.clear()
                self.msleep(500)


class MediaWatcher(QThread):
    """后台每秒读一次“正在播放”（网易云、QQ 音乐、B 站、浏览器…都会报给 Windows），有变化才通知"""
    changed = Signal(object)     # None 或 {app, title, artist, playing, thumb(bytes|None)}

    def __init__(self):
        super().__init__()
        self.loop = None
        self.mgr = None
        self._stop = False
        self._last = None
        self._thumb_key = self._thumb = None

    def run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_until_complete(self.main())
        except Exception as e:
            log_error("读取正在播放失败: %r" % e)

    async def main(self):
        self.mgr = await MediaManager.request_async()
        while not self._stop:
            try:
                info = await self.read()
            except Exception:
                info = None
            if info != self._last:
                self._last = info
                self.changed.emit(info)
            for _ in range(10):
                if self._stop:
                    return
                await asyncio.sleep(0.1)

    async def read(self):
        s = self.mgr.get_current_session()
        if s is None:
            return None
        props = await s.try_get_media_properties_async()
        title = (props.title or "").strip()
        if not title:
            return None
        artist = (props.artist or "").strip()
        key = (s.source_app_user_model_id, title, artist)
        if key != self._thumb_key:          # 换歌了才重新读封面
            self._thumb_key = key
            self._thumb = await self.read_thumb(props.thumbnail)
        return {"app": s.source_app_user_model_id or "", "title": title, "artist": artist,
                "playing": int(s.get_playback_info().playback_status) == 4, "thumb": self._thumb}

    @staticmethod
    async def read_thumb(ref):
        if ref is None:
            return None
        try:
            stream = await ref.open_read_async()
            size = int(stream.size)
            if not size or size > 8 * 2**20:
                return None
            reader = DataReader(stream.get_input_stream_at(0))
            await reader.load_async(size)
            return bytes(reader.read_buffer(size))
        except Exception:
            return None

    def command(self, name):
        if self.loop is not None and self.mgr is not None:
            asyncio.run_coroutine_threadsafe(self._command(name), self.loop)

    async def _command(self, name):
        s = self.mgr.get_current_session()
        if s is None:
            return
        try:
            if name == "toggle":
                await s.try_toggle_play_pause_async()
            elif name == "next":
                await s.try_skip_next_async()
            elif name == "prev":
                await s.try_skip_previous_async()
        except Exception:
            pass
        await asyncio.sleep(0.25)
        self._last = None                     # 马上刷新一次状态

    def stop(self):
        self._stop = True
        self.wait(2000)


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


_icon_views = []


def desktop_icon_views():
    """桌面图标的列表窗口。只在 Progman / WorkerW 里找（不枚举全部顶层窗口，开的程序多时有上千个），
    找到后缓存，explorer 重启导致句柄失效时再重新找"""
    global _icon_views
    if _icon_views and all(win32gui.IsWindow(h) for h in _icon_views):
        return _icon_views
    views = []
    for cls in ("Progman", "WorkerW"):
        h = 0
        while True:
            h = win32gui.FindWindowEx(0, h, cls, None)
            if not h:
                break
            dv = win32gui.FindWindowEx(h, 0, "SHELLDLL_DefView", None)
            lv = dv and win32gui.FindWindowEx(dv, 0, "SysListView32", None)
            if lv:
                views.append(lv)
    _icon_views = views
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
HSHELL_WINDOWACTIVATED, HSHELL_RUDEAPPACTIVATED, HSHELL_FLASH = 4, 0x8004, 0x8006
WIDGET_KINDS = ("media", "volume", "tray", "clock", "power")     # Dock 右端的小组件
WIDGET_GLYPHS = {"volume": "", "tray": "", "power": ""}   # Segoe 图标字体里的字符
VK_VOLUME_MUTE, VK_VOLUME_DOWN, VK_VOLUME_UP = 0xAD, 0xAE, 0xAF
WEEKDAYS = "一二三四五六日"


def press_key(vk):
    win32api.keybd_event(vk, 0, 0, 0)
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]


def idle_seconds():
    """多久没碰键盘鼠标了（判断人在不在电脑前）"""
    info = LASTINPUTINFO(cbSize=ctypes.sizeof(LASTINPUTINFO))
    if not user32.GetLastInputInfo(ctypes.byref(info)):
        return 0
    return (kernel32.GetTickCount() - info.dwTime) / 1000.0


class SYSTEM_POWER_STATUS(ctypes.Structure):
    _fields_ = [("ACLineStatus", ctypes.c_ubyte), ("BatteryFlag", ctypes.c_ubyte),
                ("BatteryLifePercent", ctypes.c_ubyte), ("SystemStatusFlag", ctypes.c_ubyte),
                ("BatteryLifeTime", wintypes.DWORD), ("BatteryFullLifeTime", wintypes.DWORD)]


def battery_percent():
    """笔记本返回电量百分比，台式机 / 读不到返回 None"""
    s = SYSTEM_POWER_STATUS()
    if not kernel32.GetSystemPowerStatus(ctypes.byref(s)) or s.BatteryFlag & 128 or s.BatteryLifePercent > 100:
        return None
    return s.BatteryLifePercent


_icon_font = None


def icon_font():
    """Windows 自带的图标字体（Win11：Segoe Fluent Icons，Win10：Segoe MDL2 Assets）"""
    global _icon_font
    if _icon_font is None:
        families = set(QFontDatabase.families())
        _icon_font = next((f for f in ("Segoe Fluent Icons", "Segoe MDL2 Assets") if f in families), "Segoe UI Symbol")
    return _icon_font
user32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]
user32.RegisterWindowMessageW.restype = wintypes.UINT
user32.RegisterShellHookWindow.argtypes = [wintypes.HWND]
user32.DeregisterShellHookWindow.argtypes = [wintypes.HWND]

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

ICON_STORE = 128     # 图标原图只留这么大（默认设置下最大显示 ~94px）；Dock.apply_metrics 会按需调大


def shrink_icon(img):
    if max(img.width(), img.height()) > ICON_STORE:
        img = img.scaled(ICON_STORE, ICON_STORE, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    return img.convertToFormat(QImage.Format_ARGB32_Premultiplied)


class DockItem:
    MIP_SIZES = (48, 96, 128, 160, 256)

    def __init__(self, path, name=None, args="", pinned=True, image=None):
        self.path = path
        self.name = name or display_name(path)
        self.args = args
        self.pinned = pinned
        self.target = resolve_exe(path)
        self.is_recycle = is_recycle_bin(path)
        self.recycle_full = recycle_bin_full() if self.is_recycle else False
        img = image if image is not None else load_icon_image(path)
        self._full_res = max(img.width(), img.height())     # 原图有多大，调大图标时用来判断要不要重新读
        self.image = shrink_icon(img)
        self.hwnds = []               # 当前属于它的窗口
        self.bounce_start = -10.0
        self.attention = False        # 窗口在闪烁（来消息了），显示红点
        self.children = None          # 分组才有
        self._mips = {}
        self._shadow = None
        self._reflection = None

    def ensure_resolution(self):
        """图标调大后，存的原图不够清晰就从文件重新读一份"""
        have = max(self.image.width(), self.image.height())
        if have < min(ICON_STORE, self._full_res):
            self.image = shrink_icon(extract_icon(self.path) or load_icon_image(self.path))
            self._mips, self._shadow, self._reflection = {}, None, None

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

    def exact_pixmaps(self, w, dpr):
        """没被放大的图标尺寸固定：预先缩好图标和投影，每帧直接贴图，省掉现场平滑缩放"""
        key = ("exact", round(w * dpr))
        cached = self._mips.get(key)
        if cached is None:
            px = round(w * dpr)
            icon = QPixmap.fromImage(self.image.scaled(px, px, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            icon.setDevicePixelRatio(dpr)
            sk = round(px * 128 / 96)
            shadow = self.shadow_pixmap().scaled(sk, sk, Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
            shadow.setDevicePixelRatio(dpr)
            cached = self._mips[key] = (icon, shadow)
        return cached

    def pixmap_for(self, px):
        size = next((s for s in self.MIP_SIZES if s >= px), self.MIP_SIZES[-1])
        size = min(size, max(self.image.width(), self.image.height()))   # 不比存的原图更大
        pm = self._mips.get(size)
        if pm is None:
            pm = QPixmap.fromImage(self.image.scaled(size, size, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self._mips[size] = pm
        return pm

    def to_dict(self):
        return {"name": self.name, "path": self.path, "args": self.args}


class GroupItem(DockItem):
    """图标分组：Dock 上显示成一个格子图标，点开弹出里面的图标"""

    def __init__(self, name, children):
        self.path = ""
        self.name = name or "分组"
        self.args = ""
        self.pinned = True
        self.target = None
        self.is_recycle = False
        self.recycle_full = False
        self.children = list(children)
        for c in self.children:
            c.pinned = True
        self.bounce_start = -10.0
        self.recompose()

    def recompose(self):
        """图标 = 半透明圆角方块里 2×2 排着前 4 个图标"""
        size = 256
        img = QImage(size, size, QImage.Format_ARGB32_Premultiplied)
        img.fill(Qt.transparent)
        q = QPainter(img)
        q.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        q.setPen(QPen(QColor(255, 255, 255, 120), 4))
        q.setBrush(QColor(255, 255, 255, 70))
        q.drawRoundedRect(QRectF(10, 10, size - 20, size - 20), 52, 52)
        cell, pad = 96, 24
        for i, c in enumerate(self.children[:4]):
            r = QRectF(pad + (i % 2) * (cell + 16), pad + (i // 2) * (cell + 16), cell, cell)
            q.drawImage(r, c.image.scaled(cell, cell, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        q.end()
        self._full_res = size
        self.image = shrink_icon(img)
        self._mips, self._shadow, self._reflection = {}, None, None

    def ensure_resolution(self):
        for c in self.children:
            c.ensure_resolution()
        self.recompose()

    @property
    def running(self):
        return any(c.running for c in self.children)

    @property
    def hwnds(self):
        return [h for c in self.children for h in c.hwnds]

    @hwnds.setter
    def hwnds(self, _):
        pass

    @property
    def attention(self):
        return any(c.attention for c in self.children)

    @attention.setter
    def attention(self, value):
        if not value:
            for c in self.children:
                c.attention = False

    def to_dict(self):
        return {"name": self.name, "group": [c.to_dict() for c in self.children]}


def make_item(d):
    """配置里的一条 → DockItem / GroupItem"""
    if "group" in d:
        children = [DockItem(c["path"], c.get("name"), c.get("args", "")) for c in d["group"] if c.get("path")]
        return GroupItem(d.get("name"), children) if children else None
    return DockItem(d["path"], d.get("name"), d.get("args", "")) if d.get("path") else None


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
        r = min(bar.width(), bar.height()) * th["radius"]     # 按短边算：竖着放时也是正常圆角
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
    if bar.height() > bar.width():
        # 竖着的 Dock：把画布转 90°，装饰照横条的样子画，大小和排布都不变
        p.translate(bar.left(), bar.bottom())
        p.rotate(-90)
        bar = QRectF(0, 0, bar.height(), bar.width())
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


_decor_cache = {}


def decor_pixmap(th, w, h, dpr):
    """装饰图案（星星、樱花、黄瓜片…）每帧画几十个小图形很费，按宽度每 16px 一档缓存成图片"""
    # 长边（会随放大变化）按 16px 分档，短边（Dock 厚度）固定
    vertical = h > w
    bucket = int((h if vertical else w) // 16)
    short = int(w if vertical else h)
    key = (th["name"], th["decor"], th["accent"], bucket, short, vertical, round(dpr, 2))
    pm = _decor_cache.get(key)
    if pm is None:
        if len(_decor_cache) > 24:
            _decor_cache.clear()
        long_side = (bucket + 1) * 16
        bw, bh = (short, long_side) if vertical else (long_side, short)
        pm = QPixmap(max(1, int(bw * dpr)), max(1, int(bh * dpr)))
        pm.setDevicePixelRatio(dpr)
        pm.fill(Qt.transparent)
        q = QPainter(pm)
        q.setRenderHint(QPainter.Antialiasing)
        r = QRectF(0, 0, bw, bh)
        paint_decor(q, th, r, bar_path(th, r))
        q.end()
        _decor_cache[key] = pm
    return pm


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
    vertical = pr.height() > pr.width()
    # 渐变沿 Dock 的厚度方向（竖着放时是左右方向）
    grad = QLinearGradient(pr.topLeft(), pr.topRight() if vertical else pr.bottomLeft())
    grad.setColorAt(0, qc(th["bg_top"]))
    grad.setColorAt(1, qc(th["bg_bottom"]))
    p.fillPath(path, grad)
    if th["decor"]:
        # 缓存图已经按 Dock 形状裁好，拉伸不到 2%：不再裁边、也不用平滑缩放（这两样每帧都很贵）
        pm = decor_pixmap(th, bar.width(), bar.height(), p.device().devicePixelRatioF())
        p.save()
        p.setRenderHint(QPainter.SmoothPixmapTransform, False)
        p.drawPixmap(bar, pm, QRectF(pm.rect()))
        p.restore()

    if th["highlight"]:
        # 顶部内高光：直接画一条线，不再按形状裁剪（裁剪路径每帧都很贵）
        hl = qc(th["highlight"])
        p.setPen(QPen(hl, 1))
        y = pr.top() + 1.5
        if th["bg"] == "shelf":
            p.fillRect(QRectF(pr.left() + 2, pr.bottom() - 3, pr.width() - 4, 2), hl)     # 玻璃台前沿
            inset = pr.height() * 0.8
            p.drawLine(QPointF(pr.left() + inset + 2, y), QPointF(pr.right() - inset - 2, y))
        elif vertical:
            r = pr.width() * th["radius"]
            x = pr.left() + 1.5
            p.drawLine(QPointF(x, pr.top() + r * 0.8), QPointF(x, pr.bottom() - r * 0.8))
        else:
            r = pr.height() * th["radius"]
            p.drawLine(QPointF(pr.left() + r * 0.8, y), QPointF(pr.right() - r * 0.8, y))
    if th["border2"]:
        g = QLinearGradient(pr.topLeft(), pr.bottomLeft() if vertical else pr.topRight())
        g.setColorAt(0, qc(th["border"]))
        g.setColorAt(0.5, qc(th["border2"]))
        g.setColorAt(1, qc(th["border"]))
        pen = QPen(QBrush(g), th["border_w"])
    else:
        pen = QPen(qc(th["border"]), th["border_w"])
    p.strokePath(path, pen)
    p.restore()


_indicator_cache = {}


def indicator_sprite(th, B, dpr):
    """运行指示（圆点 / 短横 / 发光点）画一次缓存成小图，每帧直接贴"""
    key = (th["indicator"], th["accent"], B, round(dpr, 2))
    pm = _indicator_cache.get(key)
    if pm is None:
        r = max(2.0, B / 22)
        half = max(B * 0.17, r * 4) + 1          # 小图半宽，够装下发光圈 / 短横
        size = int(half * 2 + 1)
        pm = QPixmap(int(size * dpr), int(size * dpr))
        pm.setDevicePixelRatio(dpr)
        pm.fill(Qt.transparent)
        q = QPainter(pm)
        q.setRenderHint(QPainter.Antialiasing)
        q.setPen(Qt.NoPen)
        c = QPointF(size / 2, size / 2)
        if th["indicator"] == "bar":
            w, h = B * 0.34, max(2.5, B / 17)
            q.setBrush(qc(th["accent"]))
            q.drawRoundedRect(QRectF(c.x() - w / 2, c.y() - h / 2, w, h), h / 2, h / 2)
        else:
            if th["indicator"] == "glow":
                g = QRadialGradient(c, r * 4)
                g.setColorAt(0, qc(th["accent"], 160))
                g.setColorAt(1, qc(th["accent"], 0))
                q.setBrush(g)
                q.drawEllipse(c, r * 4, r * 4)
            q.setBrush(qc(th["accent"]))
            q.drawEllipse(c, r, r)
        q.end()
        if len(_indicator_cache) > 16:
            _indicator_cache.clear()
        _indicator_cache[key] = pm
    return pm


def paint_indicator(p, th, cx, y, B):
    pm = indicator_sprite(th, B, p.device().devicePixelRatioF())
    half = pm.width() / pm.devicePixelRatio() / 2
    p.drawPixmap(QPointF(cx - half, y - half), pm)


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
    exact = abs(w - B) < 0.01          # 没被放大：贴预先缩好的图，不用现场缩放
    if exact:
        icon_pm, shadow_pm = item.exact_pixmaps(w, dpr)
    if th["shadow"]:
        k = w * 128 / 96
        if exact:
            p.drawPixmap(QPointF(round(rect.center().x() - k / 2), round(rect.center().y() - k / 2 + w * 0.06)),
                         shadow_pm)
        else:
            s = item.shadow_pixmap()
            p.drawPixmap(QRectF(rect.center().x() - k / 2, rect.center().y() - k / 2 + w * 0.06, k, k), s,
                         QRectF(s.rect()))
    if th["reflection"]:
        rp = item.reflection_pixmap()
        p.save()
        p.setClipPath(path)
        p.setOpacity(p.opacity() * 0.6)
        # 倒影跟弹跳方向相反
        p.drawPixmap(QRectF(rect.left(), 2 * base_y - rect.bottom() + 1, w, w), rp, QRectF(rp.rect()))
        p.restore()
    if exact:
        p.drawPixmap(QPointF(round(rect.left()), round(rect.top())), icon_pm)
    else:
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


kernel32.GetCurrentProcess.restype = wintypes.HANDLE
kernel32.SetProcessWorkingSetSize.argtypes = [wintypes.HANDLE, ctypes.c_size_t, ctypes.c_size_t]


def trim_memory():
    """把启动、打开对话框时用过、之后不再需要的内存页还给系统（任务管理器里的“内存”会明显下降）。
    真要用到时系统会自动换回来，只在空闲时调用"""
    gc.collect()
    kernel32.SetProcessWorkingSetSize(kernel32.GetCurrentProcess(), ctypes.c_size_t(-1).value,
                                      ctypes.c_size_t(-1).value)


USAGE_PATH = os.path.join(APP_DIR, "usage.json")


def fmt_duration(seconds):
    m = int(seconds // 60)
    if m < 1:
        return "不到 1 分钟"
    if m < 60:
        return "%d 分钟" % m
    return "%d 小时 %d 分" % (m // 60, m % 60) if m % 60 else "%d 小时" % (m // 60)


class UsageStats:
    """每天每个程序在前台用了多久。只存在本机 usage.json，保留最近 35 天"""

    def __init__(self):
        self.days, self.names = {}, {}
        try:
            with open(USAGE_PATH, encoding="utf-8") as f:
                data = json.load(f)
            self.days, self.names = data.get("days", {}), data.get("names", {})
        except (OSError, ValueError):
            pass
        self.dirty = False

    @staticmethod
    def today_key(offset=0):
        return time.strftime("%Y-%m-%d", time.localtime(time.time() - offset * 86400))

    def add(self, key, name, seconds):
        day = self.days.setdefault(self.today_key(), {})
        day[key] = day.get(key, 0) + seconds
        if name:
            self.names[key] = name
        self.dirty = True

    def today(self, key):
        return self.days.get(self.today_key(), {}).get(key, 0)

    def totals(self, days=1):
        """最近 days 天合计：[(key, 秒)] 从多到少"""
        total = {}
        for i in range(days):
            for k, s in self.days.get(self.today_key(i), {}).items():
                total[k] = total.get(k, 0) + s
        return sorted(total.items(), key=lambda kv: -kv[1])

    def save(self):
        if not self.dirty:
            return
        keep = {self.today_key(i) for i in range(35)}
        self.days = {d: v for d, v in self.days.items() if d in keep}
        used = {k for v in self.days.values() for k in v}
        self.names = {k: n for k, n in self.names.items() if k in used}
        tmp = USAGE_PATH + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"days": self.days, "names": self.names}, f, ensure_ascii=False)
            os.replace(tmp, USAGE_PATH)
            self.dirty = False
        except OSError as e:
            log_error("保存使用统计失败: %r" % e)


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
        self._roots = {}              # 运行中程序 → (主程序规范路径, 原始路径) 的缓存
        self.preview = WindowPreview(self)
        self.group_popup = GroupPopup(self)
        self.press_widget = None
        self.peek_timer = QTimer(self, interval=400, timeout=self.check_peek)
        self._preview_item = None     # 鼠标正停在哪个图标上（用来决定弹哪个预览）
        self.preview_timer = QTimer(self, singleShot=True, interval=400, timeout=self.show_preview)
        QTimer.singleShot(2000, self.preview.warm_up)
        QTimer.singleShot(10000, trim_memory)      # 启动时加载图标等用过的内存，稳定后还回去
        self._last_dirty = QRect()    # 上一帧重画的区域（缩小时要把旧内容擦掉）
        self._open_apps = {}          # {exe: 显示名}，当前开着窗口的程序
        self._gone_since = {}         # {exe: 窗口消失的时间}，用来确认是真的关了
        self._say_queue = []
        self.tray_message = tray_message
        self.items = [it for it in map(make_item, cfg["items"]) if it is not None]
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
        self.drop_on = None           # 文件正拖在回收站图标上
        self.menu_open = False
        self.user_hidden = False
        self.fs_hidden = False
        self.taskbar_hidden = False
        self.icons_hidden = False
        self.attention_until = 0.0    # 来消息时让自动隐藏的 Dock 冒出来一会儿
        self.taskbar_peek = False     # 点托盘按钮临时显示任务栏
        self.peek_started = self.peek_last_over = 0.0
        self.mascot_press = None      # 按住看板娘的位置（拖动换边用）
        self.mascot_dragging = False
        self.group_drop = None        # 拖动图标时，正要放进去的那个图标（建分组）
        # 看板娘提醒
        self.last_water = time.monotonic()
        self.fs_since = None          # 这次全屏从什么时候开始
        self.game_reminded = 0        # 这次全屏已经提醒过几次
        self.pending_say = None       # 全屏时看不到气泡，攒着等退出全屏再说
        self.greeted = set()          # 今天已经问候过的时段
        self.usage = UsageStats()
        self._usage_last = time.monotonic()
        self.summarized_day = None    # 今晚总结过了没有
        self.mouse_win = None         # 鼠标的窗口坐标（看板娘的点击 / 拖动、媒体按钮用）
        self.update_orient()

        self.apply_metrics()

        self.anim = QTimer(self, interval=16, timeout=self.tick)
        self._last_tick = time.monotonic()
        self.hover_check = QTimer(self, interval=120, timeout=self.check_hover)
        self.shell_timer = QTimer(self, interval=1500, timeout=self.periodic)
        self.shell_timer.start()

        self.monitor = WindowMonitor()
        self.monitor.updated.connect(self.on_windows_update)
        self.monitor.recycle_state.connect(self.on_recycle_state)
        self.monitor.watch_recycle = lambda: any(it.is_recycle for it in self.pinned_flat())
        self.monitor.start()

        self.media_info = None
        self._media_thumb = (None, None)          # (封面原始数据, 缩好的图)
        self.media = None
        if HAVE_MEDIA:
            self.media = MediaWatcher()
            self.media.changed.connect(self.on_media)
            self.media.start()

        self.launch_failed.connect(lambda n, err: self.tray_message("启动失败：" + n, err))
        screen = QGuiApplication.primaryScreen()
        screen.availableGeometryChanged.connect(lambda _: self.relayout_window())

        self.apply_shell()
        failed = self.register_hotkeys()
        if failed:
            self.tray_message("快捷键不可用", "、".join("%s（%s）" % f for f in failed)
                              + " 已被其他程序占用，请在 设置 → 快捷键设置 里换一个")

        # 接收 Shell 消息：窗口闪烁（来消息）、窗口被激活。任务栏藏起来也照样收得到
        self.WM_SHELLHOOK = user32.RegisterWindowMessageW("SHELLHOOK")
        user32.RegisterShellHookWindow(int(self.winId()))

        # 时钟：每分钟整点刷新一次，顺便检查问候和提醒
        self.minute_timer = QTimer(self, singleShot=True, timeout=self.on_minute)
        self.schedule_minute()
        # 看板娘呼吸：只刷新她那一小块，约 12 帧/秒
        self.breath_timer = QTimer(self, interval=80, timeout=self.breathe)
        self.update_breath_timer()
        QTimer.singleShot(3000, self.greet)

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
            if msg.message == getattr(self, "WM_SHELLHOOK", None):
                code, hwnd = msg.wParam, msg.lParam
                # 不在系统回调里干重活，交给事件循环
                if code == HSHELL_FLASH:
                    QTimer.singleShot(0, lambda h=hwnd: self.on_window_flash(h))
                elif code in (HSHELL_WINDOWACTIVATED, HSHELL_RUDEAPPACTIVATED):
                    QTimer.singleShot(0, lambda h=hwnd: self.on_window_activated(h))
        return False, 0

    # ---------- 消息提醒（窗口闪烁）

    def item_for_hwnd(self, hwnd):
        """这个窗口属于 Dock 上的哪个图标（固定的、分组里的、或运行区的）"""
        try:
            pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        except Exception:
            return None
        path = pid_path(pid)
        if not path:
            return None
        key = os.path.normcase(path)
        pinned = self.pinned_targets()
        owner = self.owner_of(key, hwnd, pinned)
        if owner is not None:
            return owner
        if key not in self._roots:
            self._roots[key] = find_app_root(key, path, hwnd)
        root = self._roots[key][0]
        return next((it for it in self.running_items if it.target == root), None)

    def on_window_flash(self, hwnd):
        if not self.cfg.get("notify_flash", True) or hwnd == win32gui.GetForegroundWindow():
            return
        item = self.item_for_hwnd(hwnd)
        if item is None or item.attention:
            return                        # 已经在提醒了（闪烁会连续来好几次）
        item.attention = True
        item.bounce_start = time.monotonic()
        group = self.group_of(item)
        if group is not None:
            group.bounce_start = item.bounce_start
        self.attention_until = time.monotonic() + 4      # 自动隐藏模式下冒出来 4 秒
        self.kick()
        self.announce([item.name], "mascot_notify", "mascot_notify_lines", "{name}……有新消息。")

    def on_window_activated(self, hwnd):
        item = self.item_for_hwnd(hwnd) if hwnd else None
        if item is not None and item.attention:
            item.attention = False
            self.update()

    def clear_seen_attention(self):
        """前台已经是那个程序的窗口了，就算看过了"""
        fg = win32gui.GetForegroundWindow()
        for it in self.all_items():
            if it.attention and fg in it.hwnds:
                it.attention = False
                self.update()

    def paint_badge(self, p, rect):
        """来消息的红点，挂在图标右上角"""
        r = max(5.0, rect.width() * 0.13)
        c = QPointF(rect.right() - r * 0.6, rect.top() + r * 0.6)
        p.setPen(QPen(QColor(255, 255, 255), max(1.5, r * 0.28)))
        p.setBrush(QColor(235, 64, 52))
        p.drawEllipse(c, r, r)

    # ---------- 时钟 / 音量 / 托盘按钮

    def paint_widget(self, p, kind, rect, hovered=False):
        """rect 是窗口坐标；竖着放时 rect 是窄而高的"""
        th = self.theme
        color = qc(th["text"])
        if hovered and kind != "media":
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(color.red(), color.green(), color.blue(), 32))
            m = self.pad * 0.5
            p.drawRoundedRect(rect.adjusted(m, 1, -m, -1) if self.vertical else rect.adjusted(1, m, -1, -m), 8, 8)
        p.setPen(color)
        if kind == "clock":
            now = time.localtime()
            f1 = QFont("Microsoft YaHei UI")
            f1.setPixelSize(max(12, round(self.B * (0.3 if self.vertical else 0.33))))
            f1.setBold(True)
            f2 = QFont("Microsoft YaHei UI")
            f2.setPixelSize(max(10, round(self.B * (0.18 if self.vertical else 0.2))))
            date = "%d/%d 周%s" % (now.tm_mon, now.tm_mday, WEEKDAYS[now.tm_wday])
            bat = battery_percent()
            if bat is not None and not self.vertical:
                date += " %d%%" % bat
            mid = rect.center().y()
            p.setFont(f1)
            p.drawText(QRectF(rect.left(), mid - self.B * 0.42, rect.width(), self.B * 0.46),
                       Qt.AlignHCenter | Qt.AlignBottom, time.strftime("%H:%M", now))
            p.setFont(f2)
            p.drawText(QRectF(rect.left() - 6, mid + self.B * 0.04, rect.width() + 12, self.B * 0.34),
                       Qt.AlignHCenter | Qt.AlignTop, date)
        elif kind == "media":
            self.paint_media(p, rect, color, hovered)
        else:
            f = QFont(icon_font())
            f.setPixelSize(max(12, round(self.B * 0.36)))
            p.setFont(f)
            p.drawText(rect, Qt.AlignCenter, WIDGET_GLYPHS[kind])

    def widget_rect(self, kind):
        slots, bar, _ = self.layout()
        for k, _, x, w in slots:
            if k == kind:
                return self.win_rect(QRectF(x, bar.top(), w, bar.height()))
        return None

    def widget_tip(self, kind):
        if kind == "clock":
            now = time.localtime()
            return "%d年%d月%d日 星期%s · 点击打开日历和通知" % (now.tm_year, now.tm_mon, now.tm_mday,
                                                    WEEKDAYS[now.tm_wday])
        if kind == "volume":
            return "滚轮调音量 · 点击静音"
        if kind == "media" and self.media_info:
            info = self.media_info
            app = self.media_app_item()
            text = info["title"] + (" — " + info["artist"] if info["artist"] else "")
            return text + (" · " + app.name if app else "")
        if kind == "power":
            return "电源：锁定 / 睡眠 / 注销 / 重启 / 关机"
        return "显示托盘（临时显示任务栏）"

    # ---------- 正在播放

    def on_media(self, info):
        had = bool(self.media_info)
        self.media_info = info
        if bool(info) != had:
            self.relayout_window()        # 出现 / 消失：Dock 长度要变
        self.update()

    def media_app_item(self):
        """正在播放的是 Dock 上哪个程序：用播放器报的 ID（如 com.bilibili.bilibiliPC、cloudmusic.exe）去匹配路径"""
        if not self.media_info:
            return None
        aumid = self.media_info["app"].lower()
        tokens = [t for t in re.split(r"[.!_\\\-\s]+", aumid)
                  if len(t) >= 4 and t not in ("exe", "desktop", "microsoft", "windows", "com")]
        for it in self.all_items():
            if it.target and any(t in it.target for t in tokens):
                return it
        return None

    def media_thumb(self):
        data = self.media_info.get("thumb") if self.media_info else None
        if data and self._media_thumb[0] is not data:
            img = QImage.fromData(data)
            pm = None if img.isNull() else QPixmap.fromImage(img.scaled(128, 128, Qt.KeepAspectRatioByExpanding,
                                                                         Qt.SmoothTransformation))
            self._media_thumb = (data, pm)
        if data and self._media_thumb[1] is not None:
            return self._media_thumb[1]
        app = self.media_app_item()
        return app.pixmap_for(96) if app else None

    def media_parts(self, rect):
        """正在播放组件里各部分的位置（窗口坐标）"""
        B, pad = self.B, self.pad
        if self.vertical:
            t = min(rect.width() - pad, B * 0.8)
            thumb = QRectF(rect.center().x() - t / 2, rect.top() + 4, t, t)
            return {"thumb": thumb, "toggle": QRectF(rect.left(), thumb.bottom() + 2, rect.width(),
                                                      rect.bottom() - thumb.bottom() - 2)}
        t = B * 0.8
        bw = B * 0.44
        nxt = QRectF(rect.right() - bw, rect.top(), bw, rect.height())
        tog = nxt.translated(-bw, 0)
        prv = tog.translated(-bw, 0)
        thumb = QRectF(rect.left() + pad * 0.3, rect.center().y() - t / 2, t, t)
        text = QRectF(thumb.right() + 8, rect.top(), prv.left() - thumb.right() - 12, rect.height())
        return {"thumb": thumb, "text": text, "prev": prv, "toggle": tog, "next": nxt}

    def paint_media(self, p, rect, color, hovered):
        info = self.media_info
        if not info:
            return
        parts = self.media_parts(rect)
        mw = getattr(self, "mouse_win", None)
        tr = parts["thumb"]
        pm = self.media_thumb()
        clip = QPainterPath()
        clip.addRoundedRect(tr, 6, 6)
        p.save()
        p.setClipPath(clip)
        if pm is not None:
            p.drawPixmap(tr, pm, QRectF(pm.rect()))
        else:
            p.fillRect(tr, QColor(color.red(), color.green(), color.blue(), 40))
            f = QFont(icon_font())
            f.setPixelSize(round(tr.height() * 0.5))
            p.setFont(f)
            p.setPen(color)
            p.drawText(tr, Qt.AlignCenter, "")       # 音符
        p.restore()
        if "text" in parts:
            r = parts["text"]
            f1 = QFont("Microsoft YaHei UI")
            f1.setPixelSize(max(11, round(self.B * 0.23)))
            f1.setBold(True)
            f2 = QFont("Microsoft YaHei UI")
            f2.setPixelSize(max(10, round(self.B * 0.19)))
            p.setPen(color)
            p.setFont(f1)
            p.drawText(QRectF(r.left(), r.top(), r.width(), r.height() / 2 + 2), Qt.AlignLeft | Qt.AlignBottom,
                       QFontMetricsF(f1).elidedText(info["title"], Qt.ElideRight, r.width()))
            p.setFont(f2)
            p.setPen(QColor(color.red(), color.green(), color.blue(), 170))
            sub = info["artist"] or (self.media_app_item().name if self.media_app_item() else "")
            p.drawText(QRectF(r.left(), r.center().y() + 3, r.width(), r.height() / 2 - 3), Qt.AlignLeft | Qt.AlignTop,
                       QFontMetricsF(f2).elidedText(sub, Qt.ElideRight, r.width()))
        f = QFont(icon_font())
        f.setPixelSize(max(11, round(self.B * 0.3)))
        p.setFont(f)
        glyphs = {"prev": "", "toggle": "" if info["playing"] else "", "next": ""}
        for name, glyph in glyphs.items():
            if name not in parts:
                continue
            br = parts[name]
            if hovered and mw is not None and br.contains(mw):
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(color.red(), color.green(), color.blue(), 36))
                p.drawEllipse(br.center(), min(br.width(), br.height()) * 0.42, min(br.width(), br.height()) * 0.42)
            p.setPen(color)
            p.drawText(br, Qt.AlignCenter, glyph)

    def run_widget(self, kind, pos=None):
        if kind == "media":
            rect = self.widget_rect("media")
            if rect is None or self.media is None:
                return
            parts = self.media_parts(rect)
            hit = next((n for n in ("prev", "toggle", "next") if n in parts and pos is not None
                        and parts[n].contains(pos)), None)
            if hit:
                self.media.command(hit)
            else:
                app = self.media_app_item()          # 点封面 / 歌名：切到播放器
                if app is not None:
                    self.activate(app)
        elif kind == "volume":
            press_key(VK_VOLUME_MUTE)
        elif kind == "tray":
            self.peek_taskbar()
        elif kind == "power":
            self.show_power_menu()
        elif kind == "clock":
            try:
                os.startfile("ms-actioncenter:")          # Win11：通知中心 + 日历
            except OSError:
                win32api.keybd_event(win32con.VK_LWIN, 0, 0, 0)
                press_key(ord("N"))
                win32api.keybd_event(win32con.VK_LWIN, 0, win32con.KEYEVENTF_KEYUP, 0)

    # ---------- 关机键

    def show_power_menu(self):
        rect = self.widget_rect("power")
        if rect is None:
            return
        m = QMenu()
        m.setStyleSheet(MENU_QSS)
        for title, action in (("锁定", "lock"), ("睡眠", "sleep"), (None, None), ("注销", "logoff"),
                              ("重启", "restart"), ("关机", "shutdown")):
            if title is None:
                m.addSeparator()
            else:
                m.addAction(title, lambda a=action: self.power_action(a))
        # 菜单朝屏幕里面弹出（Dock 在底部就往上弹）
        size = m.sizeHint()
        edge = {"up": rect.top(), "down": rect.bottom(), "left": rect.left(), "right": rect.right()}[self.inward()]
        if self.vertical:
            anchor = self.mapToGlobal(QPointF(edge, rect.center().y()).toPoint())
        else:
            anchor = self.mapToGlobal(QPointF(rect.center().x(), edge).toPoint())
        self.menu_open = True
        m.exec(place_popup(size.width(), size.height(), anchor, self.inward()).topLeft())
        m.deleteLater()
        self.menu_open = False
        self.check_hover()
        self.kick()

    def confirm_power(self, name):
        """注销 / 重启 / 关机会关掉所有程序，先确认一下，免得手滑丢了没保存的东西"""
        box = QMessageBox(QMessageBox.Question, APP_NAME, "确定要%s吗？\n没保存的内容可能会丢失。" % name,
                          QMessageBox.Yes | QMessageBox.No)
        box.setDefaultButton(QMessageBox.No)
        box.button(QMessageBox.Yes).setText(name)
        box.button(QMessageBox.No).setText("取消")
        box.setWindowFlag(Qt.WindowStaysOnTopHint)
        box.show()
        box.activateWindow()
        return box.exec() == QMessageBox.Yes

    def power_action(self, action):
        names = {"logoff": "注销", "restart": "重启", "shutdown": "关机"}
        if action in names:
            if not self.confirm_power(names[action]):
                return
            # 系统一关，Dock 来不及走正常退出，先把设置和统计存好
            self.save_config()
            self.usage.save()
            flag = {"logoff": "/l", "restart": "/r", "shutdown": "/s"}[action]
            args = ["shutdown", flag] if action == "logoff" else ["shutdown", flag, "/t", "0"]
            subprocess.Popen(args, creationflags=subprocess.CREATE_NO_WINDOW)
        elif action == "lock":
            user32.LockWorkStation()
        elif action == "sleep":
            self.usage.save()
            ctypes.windll.powrprof.SetSuspendState(False, False, False)

    def peek_taskbar(self):
        """隐藏任务栏时，临时把它显示出来用托盘；鼠标离开任务栏和托盘弹窗 1.5 秒后自动藏回去"""
        if not self.cfg["hide_taskbar"]:
            return
        self.taskbar_peek = True
        self.peek_started = self.peek_last_over = time.monotonic()
        bars = taskbar_windows()
        for h in bars:
            win32gui.ShowWindow(h, win32con.SW_SHOWNOACTIVATE)
            win32gui.SetWindowPos(h, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
        if bars:
            focus_window(bars[0])
        self.peek_timer.start()

    def check_peek(self):
        now = time.monotonic()
        cur = win32api.GetCursorPos()
        over = False
        for h in taskbar_windows():
            l, t, r, b = win32gui.GetWindowRect(h)
            over = over or (l <= cur[0] < r and t <= cur[1] < b)
        # 托盘溢出区、音量 / 网络面板这些弹窗开着时也不要藏
        try:
            fg = win32gui.GetForegroundWindow()
            proc = os.path.basename(pid_path(win32process.GetWindowThreadProcessId(fg)[1])).lower()
            shell_ui = proc in ("explorer.exe", "shellexperiencehost.exe", "shellhost.exe",
                                "startmenuexperiencehost.exe", "searchhost.exe") \
                and win32gui.GetClassName(fg) not in ("CabinetWClass", "Progman", "WorkerW")
        except Exception:
            shell_ui = False
        if over or shell_ui:
            self.peek_last_over = now
        if now - self.peek_started > 3 and now - self.peek_last_over > 1.5:
            self.taskbar_peek = False
            self.peek_timer.stop()
            self.apply_shell()
            self.raise_()

    # ---------- 分组

    def merge_into_group(self, item, target):
        """把 item 放进 target：target 是分组就加进去，否则两个合成一个新分组"""
        if target.children is not None:
            target.children.append(item)
            item.pinned = True
            target.recompose()
        elif target in self.items:
            self.items[self.items.index(target)] = GroupItem("分组", [target, item])

    def open_group(self, group):
        if self.group_popup.isVisible() and self.group_popup.group is group:
            self.group_popup.hide()
            return
        anchor = self.popup_anchor(group)
        if anchor is not None:
            self.group_popup.show_for(group, anchor, self.inward())
            self.kick()

    def take_out_of_group(self, group, child):
        """从分组里拿出来，放回 Dock 上分组的后面；分组只剩一个就自动解散"""
        if child not in group.children or group not in self.items:
            return
        group.children.remove(child)
        idx = self.items.index(group)
        self.items.insert(idx + 1, child)
        if len(group.children) == 1:
            self.items[idx] = group.children[0]
        elif not group.children:
            self.items.remove(group)
        else:
            group.recompose()
        self.save_config()
        self.relayout_window()

    def delete_from_group(self, group, child):
        if child in group.children:
            group.children.remove(child)
            if len(group.children) == 1 and group in self.items:
                self.items[self.items.index(group)] = group.children[0]
            elif not group.children and group in self.items:
                self.items.remove(group)
            else:
                group.recompose()
            self.save_config()
            self.relayout_window()

    def dissolve_group(self, group):
        if group in self.items:
            idx = self.items.index(group)
            self.items[idx:idx + 1] = group.children
            self.group_popup.hide()
            self.save_config()
            self.relayout_window()

    # ---------- 看板娘：问候、提醒、呼吸

    def schedule_minute(self):
        now = time.time()
        self.minute_timer.start(int((60 - now % 60) * 1000) + 50)    # 对齐到整分

    def on_minute(self):
        self.schedule_minute()
        if self.widget_kinds():
            self.update()                 # 时钟走字
        self.usage.save()
        self.greet()
        self.evening_summary()
        mono = time.monotonic()
        idle = idle_seconds()
        # 喝水：人离开电脑超过 5 分钟就重新计时
        water = int(self.cfg.get("remind_water", 0) or 0)
        if idle > 300:
            self.last_water = mono
        elif water > 0 and mono - self.last_water >= water * 60:
            self.last_water = mono
            self.remind(random.choice(["……喝点水吧。", "该喝水了。", "喝口水，休息一下眼睛。"]))
        # 连续全屏（玩游戏 / 看视频）
        game = int(self.cfg.get("remind_game", 0) or 0)
        if game > 0 and self.fs_since is not None:
            n = int((mono - self.fs_since) // (game * 60))
            if n > self.game_reminded:
                self.game_reminded = n
                minutes = n * game
                span = ("%d 小时" % (minutes // 60)) if minutes % 60 == 0 else ("%d 分钟" % minutes)
                self.remind("已经连续玩了 %s 了……休息一下吧。" % span)

    def remind(self, text):
        """全屏时看不到 Dock，就弹系统通知，并攒着等退出全屏再让看板娘说"""
        if self.isVisible() and not self.fs_hidden and self.mascot_pixmap() is not None:
            self.queue_say(text, 4.0)
        else:
            self.tray_message(APP_NAME, text)
            if self.mascot_pixmap() is not None:
                self.pending_say = text

    def greet(self):
        if not self.cfg.get("mascot_greet", True) or self.mascot_pixmap() is None or self.fs_hidden:
            return
        if idle_seconds() > 300:
            return                        # 人不在，问候给谁听
        now = time.localtime()
        h = now.tm_hour
        slots = [(5, 11, "早上好……今天也加油吧。"), (11, 14, "中午了……吃饭了吗？"),
                 (14, 18, "下午好。"), (18, 23, "晚上好。"), (23, 24, "……很晚了，还不睡吗？"),
                 (0, 5, "……这么晚了，早点睡吧。")]
        for start, end, text in slots:
            if start <= h < end:
                key = (now.tm_yday, start)
                if key not in self.greeted:
                    self.greeted.add(key)
                    self.queue_say(text, 3.5)
                return

    def update_breath_timer(self):
        on = (self.cfg.get("mascot_breathe", True) and bool(self.cfg.get("mascot_image"))
              and self.isVisible() and not self.fs_hidden)
        if on and not self.breath_timer.isActive():
            self.breath_timer.start()
        elif not on and self.breath_timer.isActive():
            self.breath_timer.stop()

    def breathe(self):
        if self.mascot_dragging:
            return
        mrect = self.mascot_rect(self.layout()[1])
        if mrect is not None:
            self.update(mrect.adjusted(-4, -10, 4, 2).toAlignedRect())

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
        accepted = dlg.exec()
        QTimer.singleShot(1500, trim_memory)
        if not accepted:
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
        self.M = max(1.0, float(self.cfg["magnify"]))
        self.spread = max(1.0, float(self.cfg["spread"]))
        self.set_metrics(int(self.cfg["icon_size"]))
        # 图标最大会显示多大，原图就留多大（多留的只是白占内存）
        global ICON_STORE
        dpr = QGuiApplication.primaryScreen().devicePixelRatio()
        ICON_STORE = 128 if int(self.cfg["icon_size"]) * self.M * dpr <= 128 else 256
        for it in self.items + self.running_items:
            it.ensure_resolution()
        self.relayout_window()

    def set_metrics(self, B):
        """图标实际大小（设定值，或者屏幕放不下时自动缩小后的值）及跟着它走的尺寸"""
        self.B = B
        self.gap = round(B * 0.16)
        self.pad = round(B * 0.2)
        self.sep_w = max(6, round(B * 0.22))
        self.bar_h = B + 2 * self.pad

    def mascot_len(self):
        """看板娘沿 Dock 方向占多长（横着放是她的宽度，竖着放是她的身高）"""
        src = self.mascot_pixmap_raw()
        if src is None:
            return 0
        h = self.B * float(self.cfg.get("mascot_size", 2.6))
        return h if self.vertical else src.width() / src.height() * h

    def needed_main(self):
        """按现在的图标、小组件、看板娘，Dock 窗口沿长边需要多长"""
        slots = self.build_slots(with_gap=False)
        rest = sum(self.slot_width(k) for k, _ in slots) + self.gap * max(len(slots) - 1, 0)
        magnify = self.B * (self.M - 1) * self.spread          # 放大时整条会变长这么多
        return rest + 2 * self.pad + magnify + 2 * (self.mascot_len() + self.gap * 2) + self.B

    # ---------- Dock 位置（底 / 顶 / 左 / 右）
    # 布局全部按“横着贴在底边”来算（抽象坐标：x 沿着 Dock，y 往屏幕边缘增大），
    # 画的时候再换算成窗口坐标。图标本身始终正着画，不旋转。

    def update_orient(self):
        self.orient = self.cfg.get("position", "bottom")
        if self.orient not in ("bottom", "top", "left", "right"):
            self.orient = "bottom"
        self.vertical = self.orient in ("left", "right")

    @property
    def L(self):
        """沿着 Dock 方向的窗口长度"""
        return self.height() if self.vertical else self.width()

    @property
    def C(self):
        """垂直于 Dock 方向的窗口长度"""
        return self.width() if self.vertical else self.height()

    def win_rect(self, r):
        o, C = self.orient, self.C
        if o == "bottom":
            return QRectF(r)
        if o == "top":
            return QRectF(r.x(), C - r.y() - r.height(), r.width(), r.height())
        if o == "left":
            return QRectF(C - r.y() - r.height(), r.x(), r.height(), r.width())
        return QRectF(r.y(), r.x(), r.height(), r.width())

    def win_pt(self, pt):
        o, C = self.orient, self.C
        if o == "bottom":
            return QPointF(pt)
        if o == "top":
            return QPointF(pt.x(), C - pt.y())
        if o == "left":
            return QPointF(C - pt.y(), pt.x())
        return QPointF(pt.y(), pt.x())

    def abs_pt(self, pt):
        o, C = self.orient, self.C
        if o == "bottom":
            return QPointF(pt)
        if o == "top":
            return QPointF(pt.x(), C - pt.y())
        if o == "left":
            return QPointF(pt.y(), C - pt.x())
        return QPointF(pt.y(), pt.x())

    def inward(self):
        """从屏幕边缘往里的方向：弹窗、名称气泡朝这边放"""
        return {"bottom": "up", "top": "down", "left": "right", "right": "left"}[self.orient]

    def popup_anchor(self, item):
        """图标朝里那一侧边缘的中点（全局坐标），弹窗贴在这里"""
        slots, bar, _ = self.layout()
        for kind, it, x, w in slots:
            if it is item:
                top = icon_base_y(self.paint_theme(), bar, self.pad) - w - 8
                return self.mapToGlobal(self.win_pt(QPointF(x + w / 2, top)).toPoint())
        return None

    def relayout_window(self):
        screen = QGuiApplication.primaryScreen()
        # 任务栏被我们藏起来时，系统仍给它留着位置，Dock 直接贴到屏幕边上占用那块地方
        scr = screen.geometry() if self.taskbar_hidden else screen.availableGeometry()
        avail = scr.height() if self.vertical else scr.width()
        # 屏幕放不下就按比例缩小图标（竖着放、图标多时常见）；有空间了再长回设定的大小
        for _ in range(3):
            target = max(28, min(int(self.cfg["icon_size"]), int(avail * self.B / self.needed_main())))
            if target == self.B:
                break
            self.set_metrics(target)
        B, M = self.B, self.M
        above = max(B * (M - 1) + B * 0.4 + 46, B * 2.1)
        cross = int(MARGIN + self.bar_h + above)
        mascot = self.mascot_pixmap_raw()
        if self.vertical:
            # 竖着放时看板娘站在 Dock 的一头，高度不受窗口厚度限制，但窗口要够宽装下她
            self.mascot_h = B * float(self.cfg.get("mascot_size", 2.6))
            mascot_w = mascot.width() / mascot.height() * self.mascot_h if mascot else 0
            # 名称气泡在图标旁边，窗口要留出放气泡的宽度
            cross = max(cross, int(mascot_w + MARGIN + 12), int(MARGIN + self.bar_h + 230))
        else:
            self.mascot_h = min(B * float(self.cfg.get("mascot_size", 2.6)), cross - MARGIN - 6)
        main = self.needed_main()
        o = self.orient
        if self.vertical:
            main = min(int(main), scr.height())
            x = scr.x() if o == "left" else scr.x() + scr.width() - cross
            self.setGeometry(x, scr.y() + (scr.height() - main) // 2, cross, main)
        else:
            main = min(int(main), scr.width())
            y = scr.y() if o == "top" else scr.y() + scr.height() - cross
            self.setGeometry(scr.x() + (scr.width() - main) // 2, y, main, cross)
        self._last_dirty = self.rect()    # 窗口变了，下一帧整个重画
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
            self._mascot_src = (None, None)   # 换了图（可能同名）要重新读
            self.relayout_window()
            self.update_breath_timer()
        if key in ("show_clock", "show_volume", "show_tray_button", "show_media", "show_power", "hide_taskbar"):
            self.relayout_window()        # 右端小组件变了，Dock 宽度跟着变
        if key == "position":
            self.update_orient()
            self.preview.hide()
            self.group_popup.hide()
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

    def mascot_pixmap_raw(self):
        """看板娘原图（只用来算宽高比），按路径缓存"""
        path = self.cfg.get("mascot_image") or ""
        if not path:
            return None
        if getattr(self, "_mascot_src", (None,))[0] != path:
            img = QImage(path)
            self._mascot_src = (path, None if img.isNull() else img)
        return self._mascot_src[1]

    def mascot_pixmap(self):
        src = self.mascot_pixmap_raw()
        if src is None or self.mascot_h <= 0:
            return None
        height = int(self.mascot_h * self.devicePixelRatioF())
        key = (self.cfg.get("mascot_image"), height)
        if self._mascot_pm[0] != key:
            self._mascot_pm = (key, QPixmap.fromImage(src.scaledToHeight(height, Qt.SmoothTransformation)))
        return self._mascot_pm[1]

    def mascot_rect(self, bar):
        pm = self.mascot_pixmap()
        if pm is None:
            return None
        """返回窗口坐标（她总是正着站，不跟着 Dock 方向转）"""
        h = self.mascot_h
        w = pm.width() / pm.height() * h
        mw = getattr(self, "mouse_win", None)
        if self.mascot_dragging and mw is not None:
            # 拖着她走：跟着鼠标，不出窗口
            x = min(max(mw.x() - w / 2, 0), self.width() - w)
            y = min(max(mw.y() - h * 0.4, 0), self.height() - h)
            return QRectF(x, y, w, h)
        now = time.monotonic()
        t = now - self.mascot_bounce
        hop = math.sin(math.pi * t / MASCOT_HOP) * self.B * 0.28 if 0 <= t < MASCOT_HOP else 0.0
        breath = getattr(self, "breath_timer", None)
        if breath is not None and breath.isActive() and hop == 0:
            # 呼吸：以脚底为基准轻轻伸缩，约 3.2 秒一次
            s = 1 + 0.018 * math.sin(now * 2 * math.pi / 3.2)
            w, h = w * (1 - (s - 1) * 0.5), h * s
        at_start = self.cfg.get("mascot_side") == "left"      # left = Dock 开头那一端（竖着时是上端）
        if self.vertical:
            bw = self.win_rect(bar)
            y = bw.top() - self.gap * 1.5 - h if at_start else bw.bottom() + self.gap * 1.5
            x = bw.left() + hop if self.orient == "left" else bw.right() - w - hop
            return QRectF(x, y, w, h)
        x = bar.left() - self.gap * 1.5 - w if at_start else bar.right() + self.gap * 1.5
        return self.win_rect(QRectF(x, bar.bottom() - h - hop, w, h))

    def paint_theme(self):
        """3D 玻璃台和倒影只适合贴在底边，换到别的位置时改成普通玻璃"""
        th = self.theme
        if self.orient != "bottom" and (th["bg"] == "shelf" or th["reflection"]):
            th = dict(th, reflection=False, bg="glass" if th["bg"] == "shelf" else th["bg"])
        return th

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
                if not self.taskbar_peek:         # 点了托盘按钮、正在临时显示时别藏
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
        self.sample_usage()
        # explorer 偶尔会自己把任务栏/图标显示回来，定时再藏一下
        if self.taskbar_hidden or self.icons_hidden:
            self.apply_shell()

    # ---------- 使用时长

    def sample_usage(self):
        """每 1.5 秒看一眼前台是谁，把这段时间记到它头上（人离开超过 2 分钟不算）"""
        now = time.monotonic()
        dt, self._usage_last = now - self._usage_last, now
        if dt > 10 or idle_seconds() > 120:
            return                        # 睡眠唤醒 / 人不在电脑前
        try:
            fg = win32gui.GetForegroundWindow()
            if not fg or win32gui.GetClassName(fg) in SKIP_CLASSES:
                return
            pid = win32process.GetWindowThreadProcessId(fg)[1]
            if pid == os.getpid():
                return
            path = pid_path(pid)
            if os.path.basename(path).lower() == "applicationframehost.exe":
                path = pid_path(uwp_real_pid(fg, pid)) or path
        except Exception:
            return
        if not path:
            return
        key = os.path.normcase(path)
        owner = self.owner_of(key, fg, self.pinned_targets())
        if owner is not None:
            key, name = owner.target, owner.name
        else:
            if key not in self._roots:
                self._roots[key] = find_app_root(key, path, fg)
            key = self._roots[key][0]
            running = next((it for it in self.running_items if it.target == key), None)
            name = running.name if running else self.usage.names.get(key) or display_name(self._roots[key][1])
        self.usage.add(key, name, dt)

    def usage_today(self, item):
        if item.children is not None:
            return sum(self.usage_today(c) for c in item.children)
        return self.usage.today(item.target) if item.target else 0

    def evening_summary(self):
        """晚上 10 点后，看板娘总结一下今天"""
        day = UsageStats.today_key()
        if self.summarized_day == day or time.localtime().tm_hour < 22 or self.mascot_pixmap() is None:
            return
        top = self.usage.totals(1)
        total = sum(s for _, s in top)
        self.summarized_day = day
        if total < 1800 or not top:
            return
        key, secs = top[0]
        name = self.usage.names.get(key) or os.path.basename(key)
        self.remind("今天用了 %s 电脑，最久的是 %s（%s）……早点休息哦。" % (fmt_duration(total), name, fmt_duration(secs)))

    def open_usage(self):
        dlg = UsageWindow(self)
        dlg.show()
        dlg.activateWindow()
        dlg.exec()
        dlg.deleteLater()
        QTimer.singleShot(1500, trim_memory)

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
        widgets = self.widget_kinds()
        if widgets:
            if slots:
                slots.append(("sep", None))
            slots.extend((k, None) for k in widgets)
        return slots

    def widget_kinds(self):
        kinds = []
        if self.cfg.get("show_media", True) and getattr(self, "media_info", None):
            kinds.append("media")
        if self.cfg.get("show_volume", True):
            kinds.append("volume")
        if self.cfg.get("show_tray_button", True) and self.cfg["hide_taskbar"]:
            kinds.append("tray")      # 任务栏没藏时不需要
        if self.cfg.get("show_clock", True):
            kinds.append("clock")
        if self.cfg.get("show_power", True):
            kinds.append("power")
        return kinds

    def slot_width(self, kind):
        if kind == "sep":
            return self.sep_w
        if kind == "clock":
            return round(self.B * (1.25 if self.vertical else 1.5))
        if kind == "media":
            return round(self.B * (1.9 if self.vertical else 4.8))
        if kind in WIDGET_KINDS:
            return round(self.B * 0.62)
        return self.B

    def rest_layout(self, slots):
        widths = [self.slot_width(kind) for kind, _ in slots]
        total = sum(widths) + self.gap * max(len(slots) - 1, 0)
        x = (self.L - total) / 2
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
            if kind in ("item", "gap") and mx is not None and self.hover_amt > 0 and self.M > 1:
                d = abs(mx - (left + rw / 2)) / (B + gap)
                if d < self.spread:
                    w = rw * (1 + (self.M - 1) * (math.cos(math.pi * d / self.spread) + 1) / 2 * self.hover_amt)
            sizes.append(w)
            if mx is not None:
                # 光标左侧的增量向左长，右侧的向右长 —— 保证光标下的图标不“跑掉”
                extra_left += (w - rw) * min(max((mx - left) / rw, 0.0), 1.0)

        result = []
        L, C = self.L, self.C
        x = (lefts[0] if lefts else L / 2) - extra_left
        for (kind, item), w in zip(slots, sizes):
            result.append((kind, item, x, w))
            x += w + gap

        hide_off = (1 - self.reveal) * (self.bar_h + MARGIN + 4)
        bottom = C - MARGIN + hide_off
        if result:
            left, right = result[0][2] - pad, result[-1][2] + result[-1][3] + pad
        else:
            left, right = L / 2 - B * 1.6, L / 2 + B * 1.6
        bar = QRectF(left, bottom - self.bar_h, right - left, self.bar_h)
        zone_top = C - MARGIN - pad - B * self.M - 10
        zone = QRectF(left, zone_top, right - left, C - zone_top)
        return result, bar, zone

    def item_at(self, pos):
        slots, bar, zone = self.layout()
        if not zone.contains(pos) and not bar.contains(pos):
            return None
        for kind, item, x, w in slots:
            if kind == "item" and x - self.gap / 2 <= pos.x() < x + w + self.gap / 2:
                return item
        return None

    def widget_at(self, pos):
        slots, bar, zone = self.layout()
        if not zone.contains(pos) and not bar.contains(pos):
            return None
        for kind, _, x, w in slots:
            if kind in WIDGET_KINDS and x - self.gap / 2 <= pos.x() < x + w + self.gap / 2:
                return kind
        return None

    def calc_drop_idx(self, x):
        slots = self.build_slots(with_gap=False)
        lefts, rws = self.rest_layout(slots)
        return sum(1 for (kind, it), left, rw in zip(slots, lefts, rws)
                   if kind == "item" and it.pinned and x > left + rw / 2)

    def group_target_at(self, x):
        """拖动图标时，光标正对着另一个固定图标的正中间 → 松手就放进去（建分组）"""
        dragged = self.items[self.drag_idx] if self.drag_idx is not None else None
        if dragged is None or dragged.children is not None:
            return None                   # 分组不能再套分组
        slots = self.build_slots(with_gap=False)
        lefts, rws = self.rest_layout(slots)
        for (kind, it), left, rw in zip(slots, lefts, rws):
            if kind == "item" and it.pinned and abs(x - (left + rw / 2)) < rw * 0.22:
                return it
        return None

    def in_remove_zone(self, pos):
        bar_top = self.C - MARGIN - self.bar_h
        return pos.y() < bar_top - self.B * 1.3

    # ---------- 动画

    def kick(self):
        if not self.anim.isActive():
            self._last_tick = time.monotonic()
            self.anim.start()
        self.update_dock()

    def dirty_rect(self):
        """这一帧可能变化的区域：Dock 条 + 名称气泡 + 看板娘和她的台词气泡"""
        if self.dragging or self.ext_drag or self.vertical or self.mascot_dragging:
            return self.rect()            # 拖动的东西到处跑；竖着放时窗口本来就窄，整个刷新
        _, bar, _ = self.layout()
        margin = max(self.B, 120)         # 两头图标的名称气泡可能伸出 Dock 条
        r = QRect(int(bar.left() - margin), 0, int(bar.width() + 2 * margin), self.height())
        mrect = self.mascot_rect(bar)
        if mrect is not None:
            left = self.cfg.get("mascot_side") == "left"
            extra = 280                   # 台词气泡在外侧
            r = r.united(QRect(int(mrect.left() - (extra if left else 10)), 0,
                               int(mrect.width() + extra + 10), self.height()))
        return r & self.rect()

    def update_dock(self):
        # 分层窗口每次刷新都要把画面交给 DWM，只更新变化的那块能省不少 CPU / 显卡
        r = self.dirty_rect()
        self.update(r.united(self._last_dirty))
        self._last_dirty = r

    def tick(self):
        now = time.monotonic()
        dt = min(now - self._last_tick, 0.05)
        self._last_tick = now
        k = min(1.0, dt * 14)

        # 预览弹窗开着时保持放大，图标不会从弹窗下面“缩走”
        active = (self.hovered or self.dragging or self.ext_drag or self.preview.isVisible()
                  or self.group_popup.isVisible())
        dragging_out = self.dragging and self.mouse_pos is not None and self.in_remove_zone(self.mouse_pos)
        target_h = 1.0 if active and not dragging_out else 0.0
        self.hover_amt += (target_h - self.hover_amt) * k
        if abs(self.hover_amt - target_h) < 0.004:
            self.hover_amt = target_h

        noticing = now < self.attention_until     # 来消息了，自动隐藏的 Dock 冒出来一会儿
        if self.cfg["auto_hide"] and not (active or self.menu_open or noticing) and now >= self.hide_after:
            target_r = 0.0
        else:
            target_r = 1.0
        self.reveal += (target_r - self.reveal) * min(1.0, dt * 11)
        if abs(self.reveal - target_r) < 0.004:
            self.reveal = target_r

        bouncing = any(now - it.bounce_start < BOUNCE_TIME for it in self.items + self.running_items) \
            or now - self.mascot_bounce < MASCOT_HOP
        waiting_hide = self.cfg["auto_hide"] and not active and (now < self.hide_after or noticing)
        self.update_dock()
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
            self.update_preview_target(None)
        self.kick()

    def check_hover(self):
        # 兜底：鼠标快速划出时 leaveEvent 偶尔收不到
        if self.dragging or self.menu_open:
            return
        pos = self.abs_pt(QPointF(self.mapFromGlobal(QCursor.pos())))
        _, bar, zone = self.layout()
        if zone.contains(pos) or bar.contains(pos):
            if self.mouse_pos != pos:
                self.mouse_pos = pos
                self.update()
            self.update_preview_target(self.item_at(pos))
        else:
            self.set_hovered(False)

    # ---------- 绘制

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        slots, bar, zone = self.layout()        # 抽象坐标（横着贴底边）
        W, P = self.win_rect, self.win_pt          # → 窗口坐标
        th = self.paint_theme()
        now = time.monotonic()
        B = self.B
        bar_w = W(bar)

        # 几乎透明的命中区域：分层窗口里 alpha=0 的像素会被鼠标穿透
        if self.hovered or self.dragging or self.ext_drag:
            p.fillRect(W(zone), QColor(0, 0, 0, 1))
        if self.cfg["auto_hide"] and self.reveal < 0.5:
            p.fillRect(W(QRectF(bar.left(), self.C - 3, bar.width(), 3)), QColor(0, 0, 0, 1))
        if self.reveal > 0.5:
            p.fillRect(bar_w, QColor(0, 0, 0, 1))   # 透明皮肤 / 低不透明度时也能接住鼠标

        hovered_item = drop_rect = None
        if self.reveal > 0.01:
            p.setOpacity(min(1.0, self.reveal * 1.5))
            path = bar_path(th, bar_w)
            paint_background(p, th, bar_w, self.skin_image())

            if not slots:
                p.setPen(qc(th["text"]))
                p.setFont(QFont("Microsoft YaHei UI", max(9, B // 6)))
                p.drawText(bar_w, Qt.AlignCenter, "拖程序到这里" if self.vertical else "把程序拖到这里")

            dpr = self.devicePixelRatioF()
            base_y = icon_base_y(th, bar, self.pad)
            ind_y = indicator_y(th, bar, self.pad)
            hovered_widget = group_rect = None
            for kind, item, x, w in slots:
                if kind == "sep":
                    p.setPen(QPen(qc(th["sep"]), 1))
                    cx = x + w / 2
                    p.drawLine(P(QPointF(cx, bar.top() + self.pad * 0.7)), P(QPointF(cx, bar.bottom() - self.pad * 0.7)))
                    continue
                in_slot = (self.mouse_pos is not None and self.hovered and not self.dragging
                           and x - self.gap / 2 <= self.mouse_pos.x() < x + w + self.gap / 2)
                if kind in WIDGET_KINDS:
                    wrect = W(QRectF(x, bar.top(), w, bar.height()))
                    self.paint_widget(p, kind, wrect, in_slot)
                    if in_slot:
                        hovered_widget = (kind, wrect)
                    continue
                if item is None:
                    continue
                t = now - item.bounce_start
                bounce = abs(math.sin(math.pi * t / (BOUNCE_TIME / 2))) * B * 0.38 * (1 - t / BOUNCE_TIME) \
                    if 0 <= t < BOUNCE_TIME else 0.0
                rect = W(QRectF(x, base_y - w - bounce, w, w))
                hovered = in_slot
                paint_item(p, th, item, rect, base_y, False, hovered and self.hover_amt > 0.5,
                           path, ind_y, B, dpr)
                if item.running:
                    ip = P(QPointF(x + w / 2, ind_y))
                    paint_indicator(p, th, ip.x(), ip.y(), B)
                if item.attention:
                    self.paint_badge(p, rect)
                if hovered:
                    hovered_item = (item, rect)
                if item is self.drop_on:
                    drop_rect = rect
                if item is self.group_drop:
                    group_rect = rect

            mrect = self.mascot_rect(bar)
            if mrect is not None:
                pm = self.mascot_pixmap()
                p.drawPixmap(mrect, pm, QRectF(pm.rect()))
                if now < self.mascot_say_until:
                    # 她个子高、头顶没地方，气泡放在旁边：横着放时在外侧，竖着放时朝屏幕里面
                    if self.vertical:
                        side = "right" if self.orient == "left" else "left"
                    else:
                        side = "left" if self.cfg.get("mascot_side") == "left" else "right"
                    self.draw_label(p, self.mascot_line, mrect, beside=side)

            if drop_rect is not None:
                self.draw_label(p, "移到回收站", drop_rect, red=True)
            elif group_rect is not None:
                self.draw_label(p, "放进分组" if self.group_drop.children is not None else "建立分组", group_rect)
            elif hovered_item and self.hover_amt > 0.5 and not self.preview.isVisible() \
                    and not self.group_popup.isVisible():
                used = self.usage_today(hovered_item[0])
                text = hovered_item[0].name + ("  ·  今天 " + fmt_duration(used) if used >= 60 else "")
                self.draw_label(p, text, hovered_item[1])
            elif hovered_widget and self.hover_amt > 0.5:
                self.draw_label(p, self.widget_tip(hovered_widget[0]), hovered_widget[1])

        # 正在拖动的图标
        if self.dragging and self.mouse_pos is not None:
            item = self.items[self.drag_idx]
            size = B * 1.1
            c = self.mouse_pos
            removing = self.in_remove_zone(c)
            wc = P(QPointF(c.x(), c.y()))
            wc = QPointF(min(max(wc.x(), size / 2), self.width() - size / 2),
                         min(max(wc.y(), size / 2), self.height() - size / 2))
            r = QRectF(wc.x() - size / 2, wc.y() - size / 2, size, size)
            p.setOpacity(0.55 if removing else 0.9)
            pm = item.pixmap_for(size * self.devicePixelRatioF())
            p.drawPixmap(r, pm, QRectF(pm.rect()))
            p.setOpacity(1.0)
            if removing:
                self.draw_label(p, "松开移除", r, red=True)
        p.end()

    def draw_label(self, p, text, icon_rect, red=False, beside=None):
        """名称气泡：默认放在图标朝屏幕里面的那一侧；beside 指定放在左 / 右边"""
        font = QFont("Microsoft YaHei UI")
        font.setPixelSize(max(12, round(self.B * 0.25)))
        p.setFont(font)
        fm = QFontMetricsF(font)
        direction = beside or {"up": "up", "down": "down", "right": "right", "left": "left"}[self.inward()]
        need = fm.horizontalAdvance(text) + 20
        room_right = self.width() - icon_rect.right() - 10
        room_left = icon_rect.left() - 10
        if beside and direction in ("left", "right"):
            # 看板娘的台词：这一侧放不下就换到另一侧（比如她站在最右端、右边没地方）
            here, there = (room_right, room_left) if direction == "right" else (room_left, room_right)
            if need > here and there > here:
                direction = "left" if direction == "right" else "right"
        # 实在放不下（竖着放时窗口窄）才截短
        room = self.width() - 4
        if direction == "right":
            room = room_right
        elif direction == "left":
            room = room_left
        if need > room:
            text = fm.elidedText(text, Qt.ElideRight, max(40, room - 20))
        tw, th = fm.horizontalAdvance(text) + 20, fm.height() + 8
        if direction in ("left", "right"):
            x = icon_rect.right() + 6 if direction == "right" else icon_rect.left() - tw - 6
            y = icon_rect.top() + icon_rect.height() * 0.12 if beside else icon_rect.center().y() - th / 2
        else:
            x = icon_rect.center().x() - tw / 2
            y = icon_rect.top() - th - 8 if direction == "up" else icon_rect.bottom() + 8
            if y < 2:                     # 上方放不下就放到图标下面
                y = icon_rect.bottom() + 6
        x = min(max(x, 2), self.width() - tw - 2)
        y = min(max(y, 2), self.height() - th - 2)
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
        pos = self.abs_pt(e.position())
        _, bar, zone = self.layout()
        if bar.contains(pos) or (self.cfg["auto_hide"] and pos.y() >= self.C - 4):
            self.mouse_pos = pos
            self.set_hovered(True)

    def leaveEvent(self, _):
        if not self.dragging and not self.menu_open:
            self.set_hovered(False)

    def mouseMoveEvent(self, e):
        self.mouse_win = e.position()
        pos = self.abs_pt(e.position())
        self.mouse_pos = pos
        if self.mascot_press is not None and (e.buttons() & Qt.LeftButton):
            if not self.mascot_dragging and (e.position() - self.mascot_press).manhattanLength() > 6:
                self.mascot_dragging = True
            if self.mascot_dragging:
                self.update()             # 她跟着鼠标到处跑，整窗重画
                return
        if (self.press_item is not None and self.press_item.pinned and not self.dragging
                and self.press_item in self.items
                and (e.buttons() & Qt.LeftButton) and (pos - self.press_pos).manhattanLength() > 6):
            self.dragging = True
            self.drag_idx = self.items.index(self.press_item)
            self.preview.hide()
            self.group_popup.hide()
        if self.dragging:
            if self.in_remove_zone(pos):
                self.drop_idx = self.group_drop = None
            else:
                self.group_drop = self.group_target_at(pos.x())
                self.drop_idx = None if self.group_drop is not None else self.calc_drop_idx(pos.x())
        else:
            _, bar, zone = self.layout()
            self.set_hovered(bar.contains(pos) or (self.hovered and zone.contains(pos)))
            if self.hovered:
                self.update_preview_target(self.item_at(pos))
        self.kick()

    def mousePressEvent(self, e):
        self.mouse_win = e.position()
        pos = self.abs_pt(e.position())
        self.preview_timer.stop()
        self.preview.hide()
        mrect = self.mascot_rect(self.layout()[1])
        if e.button() == Qt.LeftButton and mrect is not None and mrect.contains(e.position()):
            self.mascot_press = e.position()   # 松手时再决定是“点她”还是“拖她换边”
            return
        item = self.item_at(pos)
        if item is None or item is not self.group_popup.group:
            self.group_popup.hide()
        self.press_widget = self.widget_at(pos) if item is None else None
        if e.button() == Qt.LeftButton:
            self.press_item, self.press_pos = item, pos
        elif e.button() == Qt.MiddleButton and item is not None and item.children is None:
            self.launch(item)
        elif e.button() == Qt.RightButton:
            self.show_menu(e.globalPosition().toPoint(), item)

    def mouseReleaseEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        self.mouse_win = e.position()
        pos = self.abs_pt(e.position())
        if self.mascot_press is not None:
            if self.mascot_dragging:
                _, bar, _ = self.layout()
                side = "left" if pos.x() < bar.center().x() else "right"
                if side != self.cfg.get("mascot_side"):
                    self.set_option("mascot_side", side)
            else:
                self.poke_mascot()
            self.mascot_press, self.mascot_dragging = None, False
            self.update()
            return
        if self.dragging:
            item = self.items.pop(self.drag_idx)
            if self.group_drop is not None:
                self.merge_into_group(item, self.group_drop)
            elif not self.in_remove_zone(pos):
                self.items.insert(min(self.drop_idx or 0, len(self.items)), item)
            elif item.running and self.cfg["show_running"] and item.children is None:
                # 移除的程序还开着，就回到“正在运行”区
                item.pinned = False
                self.running_items.append(item)
            self.save_config()
            self.relayout_window()
        elif self.press_widget is not None and self.widget_at(pos) == self.press_widget:
            self.run_widget(self.press_widget, e.position())
        elif self.press_item is not None and self.item_at(pos) is self.press_item:
            self.activate(self.press_item)
        self.press_item = self.press_pos = self.press_widget = None
        self.dragging = False
        self.drag_idx = self.drop_idx = self.group_drop = None
        self.check_hover()
        self.kick()

    def wheelEvent(self, e):
        # 在音量 / 时钟 / 正在播放上滚轮 = 调音量（每格 4%）
        if self.widget_at(self.abs_pt(e.position())) in ("volume", "clock", "media"):
            steps = round(e.angleDelta().y() / 120)
            vk = VK_VOLUME_UP if steps > 0 else VK_VOLUME_DOWN
            for _ in range(abs(steps) * 2):
                press_key(vk)
            e.accept()

    # ---------- 外部拖入

    SHELL_IDLIST = 'application/x-qt-windows-mime;value="Shell IDList Array"'

    def drag_paths(self, mime):
        """拖进来的东西 → 路径列表。回收站、此电脑这类虚拟项目没有文件路径，
        Windows 只给“Shell IDList”，从里面解析出 ::{CLSID} 形式的名字"""
        if mime.hasUrls():
            return [os.path.normpath(u.toLocalFile()) if u.isLocalFile() else u.toString() for u in mime.urls()]
        paths = []
        if mime.hasFormat(self.SHELL_IDLIST):
            try:
                parent, children = shell.StringAsCIDA(bytes(mime.data(self.SHELL_IDLIST)))
                for child in children:
                    paths.append(shell.SHGetNameFromIDList(parent + child, shellcon.SIGDN_DESKTOPABSOLUTEPARSING))
            except Exception as ex:
                log_error("解析拖入的系统项目失败: %r" % ex)
        return paths

    def recycle_drop_target(self, e):
        """文件拖到回收站图标上 = 删除"""
        mime = e.mimeData()
        if not mime.hasUrls() or not all(u.isLocalFile() for u in mime.urls()):
            return None
        item = self.item_at(self.abs_pt(e.position()))
        return item if item is not None and item.is_recycle else None

    def dragEnterEvent(self, e):
        mime = e.mimeData()
        if mime.hasUrls() or mime.hasFormat(self.SHELL_IDLIST):
            e.acceptProposedAction()
            self.ext_drag = True
            self.dragMoveEvent(e)

    def dragMoveEvent(self, e):
        self.mouse_pos = self.abs_pt(e.position())
        self.drop_on = self.recycle_drop_target(e)
        if self.drop_on is not None:
            self.drop_idx = None
            e.setDropAction(Qt.MoveAction)
            e.accept()
        else:
            self.drop_idx = self.calc_drop_idx(self.mouse_pos.x())
            e.acceptProposedAction()
        self.kick()

    def dragLeaveEvent(self, _):
        self.ext_drag = False
        self.drop_idx = self.drop_on = None
        self.hide_after = time.monotonic() + 0.6
        self.kick()

    def dropEvent(self, e):
        idx = self.drop_idx if self.drop_idx is not None else len(self.items)
        paths = self.drag_paths(e.mimeData())
        to_recycle = self.drop_on is not None
        self.ext_drag = False
        self.drop_idx = self.drop_on = None
        if to_recycle:
            e.setDropAction(Qt.MoveAction)
            e.accept()
            self.send_to_recycle_bin(paths)
            self.kick()
            return
        self.add_paths(paths, idx)
        e.acceptProposedAction()

    # ---------- 回收站

    def send_to_recycle_bin(self, paths):
        """用 Windows 自己的删除（进回收站、可恢复，按系统设置弹确认框），放到线程里免得卡住 Dock"""
        def run():
            pythoncom.CoInitialize()
            try:
                shell.SHFileOperation((0, shellcon.FO_DELETE, "\0".join(paths), None, shellcon.FOF_ALLOWUNDO,
                                       None, None))
            except Exception as ex:
                log_error("移到回收站失败: %r" % ex)
            finally:
                pythoncom.CoUninitialize()
            self.monitor.poke()
        threading.Thread(target=run, daemon=True).start()

    def empty_recycle_bin(self):
        def run():
            try:
                shell.SHEmptyRecycleBin(0, None, 0)      # 0 = 照常弹出 Windows 的确认框
            except Exception:
                pass                                      # 用户在确认框点了“否”也会走到这里
            self.monitor.poke()
        threading.Thread(target=run, daemon=True).start()

    def on_recycle_state(self, full):
        changed = False
        for it in self.pinned_flat():
            if it.is_recycle and it.recycle_full != full:
                it.recycle_full = full
                img = extract_icon(IMAGERES, -54 if full else -55)
                if img is not None:
                    it.image = shrink_icon(img)
                    it._mips, it._shadow, it._reflection = {}, None, None
                    group = self.group_of(it)
                    if group is not None:
                        group.recompose()
                changed = True
        if changed:
            self.update()

    # ---------- 图标列表（分组里的图标也要算）

    def pinned_flat(self):
        for it in list(self.items):
            if it.children is not None:
                yield from it.children
            else:
                yield it

    def all_items(self):
        yield from self.pinned_flat()
        yield from list(self.running_items)

    def pinned_targets(self):
        return {it.target: it for it in self.pinned_flat() if it.target}

    def group_of(self, item):
        return next((g for g in self.items if g.children is not None and item in g.children), None)

    def add_paths(self, paths, index=None):
        existing = {os.path.normcase(it.path) for it in self.pinned_flat()}
        has_recycle = any(it.is_recycle for it in self.pinned_flat())
        index = len(self.items) if index is None else index
        for path in paths:
            if not path or os.path.normcase(path) in existing:
                continue
            if is_shell_path(path) and is_recycle_bin(path):
                if has_recycle:           # 回收站有 shell:RecycleBinFolder / ::{CLSID} 两种写法，只留一个
                    continue
                has_recycle = True
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
        pinned = self.pinned_targets()
        old_running = {it.target: it for it in self.running_items}
        new_running = []
        for it in pinned.values():
            it.hwnds = []
        owner_names = {}
        running_groups = {}               # 主程序 → [原始路径, 窗口]
        for key, path, hwnds in groups:
            owner = self.owner_of(key, hwnds[0], pinned)
            if owner is not None:
                owner.hwnds = owner.hwnds + hwnds
                owner_names[key] = owner.name
            elif self.cfg["show_running"]:
                if key not in self._roots:
                    self._roots[key] = find_app_root(key, path, hwnds[0])
                root, root_path = self._roots[key]
                running_groups.setdefault(root, [root_path, []])[1].extend(hwnds)
        for root, (root_path, hwnds) in running_groups.items():
            it = old_running.get(root)
            if it is None:
                image = extract_icon(root_path)
                if image is None:
                    image = window_icon(hwnds[0])
                it = DockItem(root_path, pinned=False, image=image)
                it.target = root
            it.hwnds = hwnds
            new_running.append(it)
            for key, (r, _) in self._roots.items():
                if r == root:
                    owner_names[key] = it.name
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
        self.clear_seen_attention()

    def windows_of(self, item):
        """这个图标名下的所有窗口（按 Z 序）"""
        if not item.target:
            return []
        pinned = self.pinned_targets()
        wins = []
        for h, key, _ in enum_app_windows():
            if key == item.target or self._roots.get(key, (None,))[0] == item.target \
                    or (item.pinned and self.owner_of(key, h, pinned) is item):
                wins.append(h)
        return wins

    # ---------- 窗口预览

    def preview_wanted(self, item):
        mode = self.cfg.get("window_preview", "multi")
        if item is None or mode == "off" or not item.running or item.children is not None:
            return False                  # 分组是点开看里面的图标，不做窗口预览
        return mode == "all" or len(item.hwnds) >= 2

    def update_preview_target(self, item):
        if item is self._preview_item:
            return
        self._preview_item = item
        if not self.preview_wanted(item):
            self.preview_timer.stop()
            if self.preview.isVisible():
                self.preview.hide_timer.start()
        elif self.preview.isVisible():
            self.show_preview()           # 已经开着预览，换个图标就立刻切过去
        else:
            self.preview_timer.start()

    def show_preview(self):
        item = self._preview_item
        if not self.preview_wanted(item) or self.dragging or self.menu_open:
            return
        # 后台每秒扫描的结果就够用（已是 Z 序），不在鼠标停下的这一刻再枚举一遍全部窗口
        hwnds = [h for h in item.hwnds if win32gui.IsWindow(h) and win32gui.IsWindowVisible(h)] \
            or self.windows_of(item)
        if not hwnds or (self.cfg.get("window_preview", "multi") == "multi" and len(hwnds) < 2):
            return
        anchor = self.popup_anchor(item)
        if anchor is not None:
            self.preview.show_for(item, hwnds, anchor, self.inward())
            self.kick()

    def activate(self, item):
        self.preview.hide()
        if item.children is not None:          # 分组：弹出里面的图标
            self.open_group(item)
            return
        item.attention = False
        wins = self.windows_of(item)
        if wins:
            fg = win32gui.GetForegroundWindow()
            if fg in wins:
                if len(wins) == 1:
                    minimize_window(fg)
                    return
                focus_window(wins[(wins.index(fg) + 1) % len(wins)])
            else:
                focus_window(wins[0])
            return
        self.launch(item)

    def launch(self, item, admin=False):
        item.bounce_start = time.monotonic()
        group = self.group_of(item)
        if group is not None:             # 从分组里启动的，Dock 上跳的是分组图标
            group.bounce_start = item.bounce_start
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
                elif not is_web_url(path) and not is_shell_path(path) and not os.path.exists(path):
                    raise FileNotFoundError("文件不存在：" + path)
                else:
                    os.startfile(launch_target(path), arguments=args, cwd=cwd)
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
        self.preview.hide()
        self.group_popup.hide()
        m = QMenu()
        m.setStyleSheet(MENU_QSS)
        if item is not None and item.children is not None:
            m.addAction("打开分组", lambda: self.open_group(item))
            m.addAction("重命名…", lambda: self.rename(item))
            m.addAction("解散分组（图标放回 Dock）", lambda: self.dissolve_group(item))
            m.addAction("从 Dock 移除整个分组", lambda: self.remove(item))
            m.addSeparator()
        elif item is not None:
            m.addAction("打开" if not item.running else "切换到窗口", lambda: self.activate(item))
            m.addAction("启动新实例", lambda: self.launch(item))
            if item.target or item.path.lower().endswith((".exe", ".lnk", ".bat", ".cmd")):
                m.addAction("以管理员身份运行", lambda: self.launch(item, admin=True))
            loc = m.addAction("打开文件所在位置", lambda: self.open_location(item))
            loc.setEnabled(not is_web_url(item.path) and not is_shell_path(item.path))
            if item.running:
                m.addAction("关闭窗口" if len(item.hwnds) == 1 else "关闭全部 %d 个窗口" % len(item.hwnds),
                            lambda: self.close_windows(item))
            if item.is_recycle:
                empty = m.addAction("清空回收站", self.empty_recycle_bin)
                empty.setEnabled(item.recycle_full)
            m.addSeparator()
            if item.pinned:
                m.addAction("重命名…", lambda: self.rename(item))
                m.addAction("从 Dock 移除", lambda: self.remove(item))
            else:
                m.addAction("固定到 Dock", lambda: self.pin(item))
            m.addSeparator()
        m.addAction("添加程序 / 文件…", self.add_file_dialog)
        m.addAction("添加文件夹…", self.add_folder_dialog)
        sysm = m.addMenu("添加系统项目")
        sysm.setStyleSheet(MENU_QSS)
        have = {os.path.normcase(it.path) for it in self.pinned_flat()}
        for title, path in SYSTEM_ITEMS:
            a = sysm.addAction(title, lambda p=path: self.add_paths([p]))
            a.setEnabled(os.path.normcase(path) not in have)
        self.build_settings_menu(m.addMenu("设置"))
        m.addAction("使用统计…", self.open_usage)
        m.addSeparator()
        m.addAction("隐藏 Dock（托盘可恢复）", lambda: self.set_user_hidden(True))
        m.addAction("退出", QApplication.quit)
        m.exec(gpos)
        m.deleteLater()
        self.menu_open = False
        self.check_hover()
        self.kick()
        QTimer.singleShot(1500, trim_memory)      # 菜单里可能开过对话框，用完的内存还回去

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

        radio("Dock 位置", "position", [("底部", "bottom"), ("顶部", "top"), ("左侧", "left"), ("右侧", "right")])
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
        radio("窗口预览", "window_preview", [("关闭", "off"), ("有多个窗口时", "multi"), ("总是", "all")])
        option("隐藏 Windows 任务栏", "hide_taskbar")
        hk = self.cfg["hotkeys"].get("desktop_icons")
        option("隐藏桌面图标" + ("\t" + hk if hk else ""), "hide_desktop_icons")
        sm.addAction("快捷键设置…", self.edit_hotkeys)
        sm.addSeparator()
        option("自动隐藏 Dock", "auto_hide")
        option("全屏程序时隐藏 Dock", "hide_on_fullscreen")
        toggle("开机自启", autostart_enabled(), set_autostart)
        sm.addSeparator()

        def sub_option(menu, title, key):
            a = menu.addAction(title)
            a.setCheckable(True)
            a.setChecked(bool(self.cfg.get(key, True)))
            a.triggered.connect(lambda c, k=key: self.set_option(k, c))

        def sub_radio(menu, title, key, options):
            sub = menu.addMenu(title)
            sub.setStyleSheet(MENU_QSS)
            group = QActionGroup(sub)
            for label, value in options:
                a = sub.addAction(label)
                a.setCheckable(True)
                a.setChecked(self.cfg.get(key) == value)
                group.addAction(a)
                a.triggered.connect(lambda _=False, v=value, k=key: self.set_option(k, v))

        tray = sm.addMenu("时钟和托盘")
        tray.setStyleSheet(MENU_QSS)
        if HAVE_MEDIA:
            sub_option(tray, "显示正在播放", "show_media")
        sub_option(tray, "显示时钟", "show_clock")
        sub_option(tray, "显示关机键", "show_power")
        sub_option(tray, "显示音量按钮", "show_volume")
        sub_option(tray, "显示托盘按钮（隐藏任务栏时）", "show_tray_button")

        sub_option(sm, "消息提醒（图标跳动 + 红点）", "notify_flash")

        mascot = sm.addMenu("看板娘")
        mascot.setStyleSheet(MENU_QSS)
        sub_option(mascot, "待机呼吸动画", "mascot_breathe")
        sub_option(mascot, "按时间段问候", "mascot_greet")
        sub_option(mascot, "来消息时说一句", "mascot_notify")
        sub_option(mascot, "打开软件时报一声", "mascot_announce")
        sub_option(mascot, "关掉软件时也说", "mascot_announce_close")
        sub_radio(mascot, "喝水提醒", "remind_water", [("关闭", 0), ("每 30 分钟", 30), ("每 1 小时", 60),
                                                    ("每 1.5 小时", 90)])
        sub_radio(mascot, "游戏 / 全屏时长提醒", "remind_game", [("关闭", 0), ("1 小时", 60), ("2 小时", 120),
                                                         ("3 小时", 180)])

    def open_skin_center(self):
        dlg = SkinCenter(self)
        dlg.show()
        dlg.activateWindow()
        dlg.exec()
        dlg.deleteLater()
        QTimer.singleShot(1500, trim_memory)

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
            self.preview.hide()
            self.group_popup.hide()
        self.update_breath_timer()

    def check_fullscreen(self):
        # 不管“全屏时隐藏”开没开都要检测：游戏时长提醒也靠它
        fs = False
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
        if fs and self.fs_since is None:
            self.fs_since, self.game_reminded = time.monotonic(), 0
        elif not fs and self.fs_since is not None:
            self.fs_since = None
            if self.pending_say:          # 全屏时攒下的提醒，回来再说
                QTimer.singleShot(800, lambda t=self.pending_say: self.queue_say(t, 4.0))
                self.pending_say = None
        hide = fs and self.cfg["hide_on_fullscreen"]
        if hide != self.fs_hidden:
            self.fs_hidden = hide
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
        self.usage.save()
        if getattr(self, "media", None) is not None:
            self.media.stop()
        for i in range(1, len(HOTKEY_ACTIONS) + 1):
            user32.UnregisterHotKey(int(self.winId()), i)
        user32.DeregisterShellHookWindow(int(self.winId()))
        self.monitor.requestInterruption()
        self.monitor.poke()
        self.monitor.wait(3000)


def place_popup(w, h, anchor, direction):
    """弹窗贴着 anchor 往 direction（up/down/left/right）方向展开，并保持在屏幕内"""
    scr = (QGuiApplication.screenAt(anchor) or QGuiApplication.primaryScreen()).geometry()
    if direction == "up":
        x, y = anchor.x() - w // 2, anchor.y() - h
    elif direction == "down":
        x, y = anchor.x() - w // 2, anchor.y()
    elif direction == "right":
        x, y = anchor.x(), anchor.y() - h // 2
    else:
        x, y = anchor.x() - w, anchor.y() - h // 2
    x = min(max(x, scr.left() + 6), scr.right() - w - 6)
    y = min(max(y, scr.top() + 6), scr.bottom() - h - 6)
    return QRect(int(x), int(y), int(w), int(h))


# ---------------------------------------------------------------- 窗口预览（DWM 实时缩略图）

class DWM_THUMBNAIL_PROPERTIES(ctypes.Structure):
    _fields_ = [("dwFlags", wintypes.DWORD), ("rcDestination", wintypes.RECT), ("rcSource", wintypes.RECT),
                ("opacity", ctypes.c_ubyte), ("fVisible", wintypes.BOOL), ("fSourceClientAreaOnly", wintypes.BOOL)]


dwmapi.DwmRegisterThumbnail.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.POINTER(ctypes.c_void_p)]
dwmapi.DwmUnregisterThumbnail.argtypes = [ctypes.c_void_p]
dwmapi.DwmUpdateThumbnailProperties.argtypes = [ctypes.c_void_p, ctypes.POINTER(DWM_THUMBNAIL_PROPERTIES)]
dwmapi.DwmQueryThumbnailSourceSize.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.SIZE)]
DWM_TNP_RECTDESTINATION, DWM_TNP_OPACITY, DWM_TNP_VISIBLE, DWM_TNP_SOURCECLIENTAREAONLY = 0x1, 0x4, 0x8, 0x10


class WindowPreview(QWidget):
    """鼠标停在有多个窗口的图标上时，弹出各窗口的实时缩略图，点哪个切到哪个"""
    CARD_W, THUMB_H, TITLE_H, PAD, GAP = 240, 150, 32, 10, 8

    def __init__(self, dock):
        # 不能用半透明窗口：DWM 缩略图画不到分层窗口上
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setMouseTracking(True)
        self.dock = dock
        self.item = None
        self.cards = []           # [{hwnd, title, rect, thumb_rect, close_rect, thumb}]
        self.hover_idx = None
        self.hover_close = False
        self.hide_timer = QTimer(self, singleShot=True, interval=300, timeout=self.maybe_hide)
        # 淡入淡出（整窗透明度，DWM 缩略图会跟着一起淡）
        self.fade = QVariantAnimation(self, duration=120)
        self.fade.valueChanged.connect(lambda v: self.setWindowOpacity(v))
        self.fade.finished.connect(self._fade_done)
        self._fading_out = False

    def warm_up(self):
        """启动后在屏幕外先显示一次：第一次弹出时就不用现场创建原生窗口、设置 DWM 属性"""
        if self.isVisible():
            return
        self.setWindowOpacity(0)
        self.setGeometry(-20000, -20000, 200, 120)
        self.show()
        self.repaint()                    # 不在屏幕上 Qt 不会自己画，强制画一次
        super().hide()
        # 字体引擎第一次画中文要初始化，也提前做掉
        pm = QPixmap(240, 32)
        pm.fill(Qt.transparent)
        q = QPainter(pm)
        font = QFont("Microsoft YaHei UI")
        font.setPixelSize(13)
        q.setFont(font)
        q.drawText(4, 22, "文件资源管理器 哔哩 - Chrome 微信 QQ…")
        q.end()

    def _fade_to(self, value, out=False):
        self._fading_out = out
        self.fade.stop()
        self.fade.setStartValue(self.windowOpacity())
        self.fade.setEndValue(value)
        self.fade.start()

    def _fade_done(self):
        if self._fading_out:
            self._fading_out = False
            super().hide()

    def fade_out(self):
        if self.isVisible() and not self._fading_out:
            self._fade_to(0.0, out=True)

    def hide(self):
        """立即关闭（点击切换窗口、拖动图标、右键菜单时）"""
        self.fade.stop()
        self._fading_out = False
        super().hide()

    def show_for(self, item, hwnds, anchor, direction="up"):
        """anchor 是全局坐标（图标朝屏幕里面那侧的中点），弹窗往 direction 方向展开"""
        if self.isVisible() and not self._fading_out and item is self.item \
                and [c["hwnd"] for c in self.cards] == list(hwnds):
            return                         # 同一个图标、同一组窗口：不用重建
        self.clear_thumbs()
        self.item = item
        self._anchor, self._direction = anchor, direction
        n = len(hwnds)
        scr = (QGuiApplication.screenAt(anchor) or QGuiApplication.primaryScreen()).geometry()
        card_w = min(self.CARD_W, (scr.width() - 2 * self.PAD - (n - 1) * self.GAP - 20) / n)
        thumb_h = card_w * self.THUMB_H / self.CARD_W
        w = int(2 * self.PAD + n * card_w + (n - 1) * self.GAP)
        h = int(2 * self.PAD + self.TITLE_H + thumb_h)
        self.setGeometry(place_popup(w, h, anchor, direction))
        self.cards = []
        for i, hwnd in enumerate(hwnds):
            rect = QRectF(self.PAD + i * (card_w + self.GAP), self.PAD, card_w, self.TITLE_H + thumb_h)
            self.cards.append(dict(
                hwnd=hwnd, title=win32gui.GetWindowText(hwnd) or item.name, rect=rect, thumb=None,
                thumb_rect=rect.adjusted(6, self.TITLE_H, -6, -6),
                close_rect=QRectF(rect.right() - 28, rect.top() + 5, 22, 22)))
        self.hover_idx, self.hover_close = None, False
        self.hide_timer.stop()
        appearing = not self.isVisible() or self._fading_out
        if not self.isVisible():
            self.setWindowOpacity(0)
        self.show()
        self.raise_()
        self.register_thumbs()
        self.update()
        if appearing:
            self._fade_to(1.0)

    def showEvent(self, e):
        super().showEvent(e)
        hwnd = int(self.winId())
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex | WS_EX_NOACTIVATE)
        corner = ctypes.c_int(2)          # Win11 圆角（DWMWA_WINDOW_CORNER_PREFERENCE = 33，ROUND = 2）
        dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd), 33, ctypes.byref(corner), ctypes.sizeof(corner))

    def hideEvent(self, e):
        self.clear_thumbs()
        super().hideEvent(e)
        self.dock.kick()

    def register_thumbs(self):
        dest, dpr = int(self.winId()), self.devicePixelRatioF()
        for c in self.cards:
            thumb = ctypes.c_void_p()
            if dwmapi.DwmRegisterThumbnail(dest, c["hwnd"], ctypes.byref(thumb)) != 0 or not thumb.value:
                continue
            c["thumb"] = thumb
            size = wintypes.SIZE()
            dwmapi.DwmQueryThumbnailSourceSize(thumb, ctypes.byref(size))
            r = c["thumb_rect"]
            if size.cx <= 0 or size.cy <= 0:
                continue
            scale = min(r.width() / size.cx, r.height() / size.cy)
            tw, th = size.cx * scale, size.cy * scale
            dst = QRectF(r.center().x() - tw / 2, r.center().y() - th / 2, tw, th)
            props = DWM_THUMBNAIL_PROPERTIES()
            props.dwFlags = DWM_TNP_RECTDESTINATION | DWM_TNP_OPACITY | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY
            props.rcDestination = wintypes.RECT(round(dst.left() * dpr), round(dst.top() * dpr),
                                                round(dst.right() * dpr), round(dst.bottom() * dpr))
            props.opacity = 255
            props.fVisible = True
            props.fSourceClientAreaOnly = False
            dwmapi.DwmUpdateThumbnailProperties(thumb, ctypes.byref(props))

    def clear_thumbs(self):
        for c in self.cards:
            if c.get("thumb"):
                dwmapi.DwmUnregisterThumbnail(c["thumb"])
                c["thumb"] = None

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        p.fillRect(self.rect(), QColor(30, 30, 35))
        p.setPen(QPen(QColor(255, 255, 255, 30), 1))
        p.drawRect(QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5))
        font = QFont("Microsoft YaHei UI")
        font.setPixelSize(13)
        p.setFont(font)
        fm = QFontMetricsF(font)
        icon = self.item.pixmap_for(18 * self.devicePixelRatioF()) if self.item else None
        for i, c in enumerate(self.cards):
            r = c["rect"]
            hovered = i == self.hover_idx
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(255, 255, 255, 26) if hovered else QColor(255, 255, 255, 8))
            p.drawRoundedRect(r, 8, 8)
            if icon is not None:
                p.drawPixmap(QRectF(r.left() + 8, r.top() + 7, 18, 18), icon, QRectF(icon.rect()))
            title_w = r.width() - 40 - (28 if hovered else 6)
            p.setPen(QColor(235, 235, 240))
            p.drawText(QRectF(r.left() + 32, r.top(), title_w, self.TITLE_H), Qt.AlignVCenter | Qt.AlignLeft,
                       fm.elidedText(c["title"], Qt.ElideRight, title_w))
            # 缩略图底板：最小化的窗口 DWM 不给画面，就显示一个大图标
            tr = c["thumb_rect"]
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(0, 0, 0, 70))
            p.drawRoundedRect(tr, 6, 6)
            if self.item is not None:
                big = self.item.pixmap_for(48 * self.devicePixelRatioF())
                p.setOpacity(0.5)
                p.drawPixmap(QRectF(tr.center().x() - 24, tr.center().y() - 24, 48, 48), big, QRectF(big.rect()))
                p.setOpacity(1.0)
            if hovered:
                cr = c["close_rect"]
                p.setBrush(QColor(196, 43, 28) if self.hover_close else QColor(255, 255, 255, 30))
                p.drawRoundedRect(cr, 5, 5)
                p.setPen(QPen(QColor(255, 255, 255), 1.6))
                m = 7
                p.drawLine(QPointF(cr.left() + m, cr.top() + m), QPointF(cr.right() - m, cr.bottom() - m))
                p.drawLine(QPointF(cr.right() - m, cr.top() + m), QPointF(cr.left() + m, cr.bottom() - m))
        p.end()

    def card_at(self, pos):
        for i, c in enumerate(self.cards):
            if c["rect"].contains(pos):
                return i, c["close_rect"].contains(pos)
        return None, False

    def mouseMoveEvent(self, e):
        idx, close = self.card_at(e.position())
        if (idx, close) != (self.hover_idx, self.hover_close):
            self.hover_idx, self.hover_close = idx, close
            self.update()

    def enterEvent(self, _):
        self.hide_timer.stop()

    def leaveEvent(self, _):
        self.hover_idx = None
        self.update()
        self.hide_timer.start()

    def mousePressEvent(self, e):
        idx, close = self.card_at(e.position())
        if idx is None:
            return
        hwnd = self.cards[idx]["hwnd"]
        if e.button() == Qt.MiddleButton or (e.button() == Qt.LeftButton and close):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            QTimer.singleShot(600, self.refresh)
        elif e.button() == Qt.LeftButton:
            self.hide()
            focus_window(hwnd)

    def refresh(self):
        """关掉一个窗口后重新排一下"""
        if not self.isVisible() or self.item is None:
            return
        hwnds = self.dock.windows_of(self.item)
        if not hwnds:
            self.hide()
            return
        self.show_for(self.item, hwnds, self._anchor, self._direction)

    def maybe_hide(self):
        if self.geometry().contains(QCursor.pos()):
            return
        if self.dock.hovered and self.dock._preview_item is self.item:
            return
        self.fade_out()


# ---------------------------------------------------------------- 使用统计窗口

class UsageList(QWidget):
    ROW_H = 40

    def __init__(self, dock):
        super().__init__()
        self.dock = dock
        self.rows = []
        self._icons = {}

    def set_rows(self, rows):
        self.rows = rows[:12]
        self.setFixedHeight(max(1, len(self.rows)) * self.ROW_H + 8)
        self.update()

    def icon_for(self, key):
        pm = self._icons.get(key)
        if pm is None:
            item = next((it for it in self.dock.all_items() if it.target == key), None)
            img = item.image if item is not None else (extract_icon(key) or load_icon_image(key))
            pm = QPixmap.fromImage(img.scaled(56, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self._icons[key] = pm
        return pm

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        f = QFont("Microsoft YaHei UI")
        f.setPixelSize(13)
        p.setFont(f)
        fm = QFontMetricsF(f)
        if not self.rows:
            p.setPen(QColor(150, 150, 160))
            p.drawText(self.rect(), Qt.AlignCenter, "还没有记录，用一会儿电脑再来看吧")
            return
        top = self.rows[0][2] or 1
        w = self.width()
        for i, (key, name, secs) in enumerate(self.rows):
            y = 4 + i * self.ROW_H
            pm = self.icon_for(key)
            p.drawPixmap(QRectF(4, y + 6, 28, 28), pm, QRectF(pm.rect()))
            p.setPen(QColor(232, 232, 238))
            p.drawText(QRectF(42, y, 150, self.ROW_H), Qt.AlignVCenter | Qt.AlignLeft,
                       fm.elidedText(name, Qt.ElideRight, 150))
            bar_x, bar_w = 200, w - 200 - 110
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(255, 255, 255, 18))
            p.drawRoundedRect(QRectF(bar_x, y + 15, bar_w, 10), 5, 5)
            p.setBrush(QColor(59, 108, 240))
            p.drawRoundedRect(QRectF(bar_x, y + 15, max(10, bar_w * secs / top), 10), 5, 5)
            p.setPen(QColor(200, 200, 210))
            p.drawText(QRectF(w - 104, y, 100, self.ROW_H), Qt.AlignVCenter | Qt.AlignRight, fmt_duration(secs))
        p.end()


class UsageWindow(QDialog):
    def __init__(self, dock):
        super().__init__()
        self.dock = dock
        self.setWindowTitle("使用统计 - " + APP_NAME)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.setStyleSheet(SKIN_CENTER_QSS + """
            QPushButton:checked { background: #3b6cf0; border-color: #3b6cf0; }""")
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(10)
        head = QHBoxLayout()
        title = QLabel("使用统计")
        title.setObjectName("title")
        head.addWidget(title)
        head.addStretch(1)
        self.buttons = []
        for label, days in (("今天", 1), ("最近 7 天", 7)):
            b = QPushButton(label)
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, d=days: self.show_days(d))
            head.addWidget(b)
            self.buttons.append((b, days))
        lay.addLayout(head)
        self.summary = QLabel()
        lay.addWidget(self.summary)
        self.list = UsageList(dock)
        lay.addWidget(self.list)
        tip = QLabel("只统计在前台、而且你正在用电脑的时间（离开超过 2 分钟不算）。数据只保存在本机。")
        tip.setStyleSheet("color: #8a8a94; font-size: 12px;")
        tip.setWordWrap(True)
        lay.addWidget(tip)
        self.show_days(1)

    def show_days(self, days):
        for b, d in self.buttons:
            b.setChecked(d == days)
        usage = self.dock.usage
        rows = [(k, usage.names.get(k) or os.path.basename(k), s) for k, s in usage.totals(days) if s >= 60]
        total = sum(s for _, _, s in rows)
        self.summary.setText(("今天" if days == 1 else "最近 7 天") + "一共用了 %s" % fmt_duration(total))
        self.list.set_rows(rows)
        self.adjustSize()


# ---------------------------------------------------------------- 分组弹窗

class GroupPopup(QWidget):
    """点分组图标时弹出：网格排着组里的图标，点一下打开，右键可以移出分组"""
    CELL_W, CELL_H, PAD, TITLE_H = 86, 90, 12, 30

    def __init__(self, dock):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setMouseTracking(True)
        self.dock = dock
        self.group = None
        self.hover = None
        self.cells = []
        self.hide_timer = QTimer(self, singleShot=True, interval=500, timeout=self.maybe_hide)

    def show_for(self, group, anchor, direction="up"):
        self.group = group
        n = len(group.children)
        cols = min(5, max(1, n))
        rows = (n + cols - 1) // cols
        w = 2 * self.PAD + cols * self.CELL_W
        h = 2 * self.PAD + self.TITLE_H + rows * self.CELL_H
        self.setGeometry(place_popup(w, h, anchor, direction))
        self.cells = [QRectF(self.PAD + (i % cols) * self.CELL_W, self.PAD + self.TITLE_H + (i // cols) * self.CELL_H,
                             self.CELL_W, self.CELL_H) for i in range(n)]
        self.hover = None
        self.hide_timer.stop()
        self.show()
        self.raise_()
        self.update()

    def showEvent(self, e):
        super().showEvent(e)
        hwnd = int(self.winId())
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex | WS_EX_NOACTIVATE)
        corner = ctypes.c_int(2)
        dwmapi.DwmSetWindowAttribute(wintypes.HWND(hwnd), 33, ctypes.byref(corner), ctypes.sizeof(corner))

    def hideEvent(self, e):
        super().hideEvent(e)
        self.dock.kick()

    def paintEvent(self, _):
        if self.group is None:
            return
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform | QPainter.TextAntialiasing)
        p.fillRect(self.rect(), QColor(30, 30, 35))
        p.setPen(QPen(QColor(255, 255, 255, 30), 1))
        p.drawRect(QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5))
        f = QFont("Microsoft YaHei UI")
        f.setPixelSize(14)
        f.setBold(True)
        p.setFont(f)
        p.setPen(QColor(240, 240, 245))
        p.drawText(QRectF(self.PAD + 6, self.PAD, self.width() - 2 * self.PAD, self.TITLE_H - 4),
                   Qt.AlignVCenter | Qt.AlignLeft, self.group.name)
        f2 = QFont("Microsoft YaHei UI")
        f2.setPixelSize(12)
        p.setFont(f2)
        fm = QFontMetricsF(f2)
        dpr = self.devicePixelRatioF()
        for i, (child, r) in enumerate(zip(self.group.children, self.cells)):
            if i == self.hover:
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(255, 255, 255, 28))
                p.drawRoundedRect(r.adjusted(3, 3, -3, -3), 10, 10)
            icon = QRectF(r.center().x() - 24, r.top() + 10, 48, 48)
            pm = child.pixmap_for(48 * dpr)
            p.drawPixmap(icon, pm, QRectF(pm.rect()))
            if child.running:
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(255, 255, 255, 220))
                p.drawEllipse(QPointF(icon.center().x(), icon.bottom() + 5), 2.5, 2.5)
            if child.attention:
                self.dock.paint_badge(p, icon)
            p.setPen(QColor(225, 225, 232))
            p.drawText(QRectF(r.left() + 4, icon.bottom() + 8, r.width() - 8, 20), Qt.AlignHCenter | Qt.AlignTop,
                       fm.elidedText(child.name, Qt.ElideRight, r.width() - 8))
        p.end()

    def cell_at(self, pos):
        return next((i for i, r in enumerate(self.cells) if r.contains(pos)), None)

    def mouseMoveEvent(self, e):
        idx = self.cell_at(e.position())
        if idx != self.hover:
            self.hover = idx
            self.update()

    def enterEvent(self, _):
        self.hide_timer.stop()

    def leaveEvent(self, _):
        self.hover = None
        self.update()
        self.hide_timer.start()

    def mousePressEvent(self, e):
        idx = self.cell_at(e.position())
        if idx is None or self.group is None:
            return
        group, child = self.group, self.group.children[idx]
        if e.button() == Qt.LeftButton:
            self.hide()
            self.dock.activate(child)
        elif e.button() == Qt.MiddleButton:
            self.hide()
            self.dock.launch(child)
        elif e.button() == Qt.RightButton:
            m = QMenu()
            m.setStyleSheet(MENU_QSS)
            m.addAction("打开", lambda: (self.hide(), self.dock.activate(child)))
            m.addAction("移出分组（放回 Dock）", lambda: (self.hide(), self.dock.take_out_of_group(group, child)))
            m.addAction("从分组删除", lambda: (self.hide(), self.dock.delete_from_group(group, child)))
            m.exec(e.globalPosition().toPoint())
            m.deleteLater()

    def maybe_hide(self):
        if self.geometry().contains(QCursor.pos()):
            return
        local = self.dock.abs_pt(QPointF(self.dock.mapFromGlobal(QCursor.pos())))
        if self.dock.isVisible() and self.dock.item_at(local) is self.group:
            self.hide_timer.start()       # 鼠标还在分组图标上，再等等
            return
        self.hide()


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
    tmenu.addAction("使用统计…", dock.open_usage)
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
