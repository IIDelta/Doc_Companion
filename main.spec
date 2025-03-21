import os

project_dir = os.getcwd()

a = Analysis(
    ['main.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    icon=os.path.join(project_dir, 'leaf.ico'),
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