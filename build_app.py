#!/usr/bin/env python3
"""Build a double-clickable app with PyInstaller.

PyInstaller cannot cross-compile: run this on a Mac to get a Mac app
(dist/PrepInvoices.app) and on Windows to get a Windows app (dist/PrepInvoices/PrepInvoices.exe).

  pip install -r requirements.txt pyinstaller
  python build_app.py
"""
import PyInstaller.__main__

PyInstaller.__main__.run([
    "app.py",
    "--name", "PrepInvoices",
    "--windowed",   # no terminal window
    "--noconfirm",
    "--clean",
])
