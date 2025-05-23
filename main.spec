# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_data_files

# --- !!! USER ACTION REQUIRED !!! ---
# 1.  VERIFY the 'nltk_data_path'. We've set it to the Roaming path,
#     but YOU MUST CONFIRM that 'corpora/wordnet' exists inside it.
#     If not, choose the correct path from your output.
# ------------------------------------

block_cipher = None

# --- NLTK Data Path (VERIFY THIS!) ---
# We are using the Roaming path based on your output. PLEASE CHECK IT!
nltk_data_path = 'C:\\Users\\jdufour\\AppData\\Roaming\\nltk_data'

# --- Project Paths ---
project_root = '.' # Assumes .spec is in the root

# --- Data Files ---
# Add icons, acronym list, and NLTK data.
# ('source_path', 'destination_folder_in_bundle')
datas = [
    ('ui/leaf.png', 'ui'),
    ('ui/leaf.ico', 'ui'),
    ('acronyms/acronym list.txt', 'acronyms'),
    (os.path.join(nltk_data_path, 'corpora/wordnet'), 'nltk_data/corpora/wordnet'),
    (os.path.join(nltk_data_path, 'corpora/omw-1.4'), 'nltk_data/corpora/omw-1.4'), # Wordnet 1.4, often needed
]

# --- PyQt5 Data ---
# Use PyInstaller's built-in hooks to collect PyQt5 data (plugins, etc.)
# This should pick up plugins from your Python312\Lib\site-packages\PyQt5 path.
datas += collect_data_files('PyQt5', includes=['*.dll', 'Qt5/plugins/platforms/*'])

# --- Hidden Imports ---
hiddenimports = [
    'win32com.client',
    'pywintypes',
    'PyQt5.sip',
    'nltk',
    'requests',
    'openpyxl',
    'docx',
    'PyQt5.QtWidgets',
    'PyQt5.QtGui',
    'PyQt5.QtCore',
    'wordnet',
    'numpy'
]

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Doc_Companion',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    windowed=True,
    icon='ui/leaf.ico',
)

# Builds a 'one-folder' app (recommended for debugging)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Doc_Companion_App',
)