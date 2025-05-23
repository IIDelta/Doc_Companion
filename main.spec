# main.spec
import os
import sys 

project_dir = os.getcwd()

# This is correctly set to your determined NLTK data path
nltk_data_path_base = r'C:\Users\jdufour\AppData\Roaming\nltk_data'

a = Analysis(
    ['main.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        ('ui/leaf.png', 'ui'), # For your UI icons
        # NLTK data correctly referenced
        (os.path.join(nltk_data_path_base, 'corpora/wordnet'), 'nltk_data/corpora/wordnet'),
        (os.path.join(nltk_data_path_base, 'corpora/omw-1.4'), 'nltk_data/corpora/omw-1.4'),
    ],
    hiddenimports=[
        'PyQt5.sip',
        'nltk.corpus.reader.wordnet',
        'nltk.data',
        'win32timezone',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Doc Companion',
    debug=False,
    strip=False,
    upx=False,
    console=False, 
    icon=os.path.join(project_dir, '/icons/leaf.ico'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='Doc Companion'
)