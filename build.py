# -*- coding: utf-8 -*-
"""打包成 exe：python build.py  →  release/NanmuDock-v<版本>-win64.zip

需要先 pip install -r requirements.txt pyinstaller
"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = "NanmuDock"

src = open(os.path.join(ROOT, "nanmu_dock.pyw"), encoding="utf-8").read()
version = re.search(r'APP_VERSION = "([^"]+)"', src).group(1)

subprocess.check_call([
    sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed",
    "--name", NAME, "--icon", os.path.join("assets", "icon.ico"),
    "--add-data", "%s;assets" % os.path.join("assets", "mascot_default.png"),
    "nanmu_dock.pyw",
], cwd=ROOT)

dist = os.path.join(ROOT, "dist", NAME)
with open(os.path.join(dist, "恢复任务栏和桌面图标.bat"), "w", encoding="ascii") as f:
    f.write('@echo off\r\nstart "" "%~dp0NanmuDock.exe" --restore\r\n')
shutil.copy2(os.path.join(ROOT, "LICENSE"), dist)
with open(os.path.join(dist, "使用说明.txt"), "w", encoding="utf-8-sig") as f:
    f.write("楠木 Dock v%s\r\n\r\n"
            "双击 NanmuDock.exe 启动，右键 Dock 打开设置和皮肤中心。\r\n"
            "隐藏任务栏后如果 Dock 被强制结束，双击“恢复任务栏和桌面图标.bat”即可恢复。\r\n"
            "配置保存在本文件夹的 dock_config.json，整个文件夹可以随意挪动。\r\n" % version)

os.makedirs(os.path.join(ROOT, "release"), exist_ok=True)
zip_path = os.path.join(ROOT, "release", "%s-v%s-win64.zip" % (NAME, version))
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    for folder, _, files in os.walk(dist):
        for name in files:
            full = os.path.join(folder, name)
            z.write(full, os.path.join(NAME, os.path.relpath(full, dist)))
print("\n打包完成：%s（%.1f MB）" % (zip_path, os.path.getsize(zip_path) / 1024 / 1024))
