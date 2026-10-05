#!/usr/bin/env python3
"""Build a double-clickable app with PyInstaller.

PyInstaller cannot cross-compile: run this on a Mac to get a Mac app
(dist/PrepInvoices.app) and on Windows to get a Windows app (dist/PrepInvoices.exe, a single file).

  pip install -r requirements.txt pyinstaller
  python build_app.py
"""
import sys

import PyInstaller.__main__

args = [
    "app.py",
    "--name", "PrepInvoices",
    "--windowed",   # no terminal window
    "--noconfirm",
    "--clean",
]
if sys.platform == "win32":
    args.append("--onefile")  # one self-contained PrepInvoices.exe, nothing to install
PyInstaller.__main__.run(args)
