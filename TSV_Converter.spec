# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Список библиотек для исключения из сборки (для уменьшения размера)
# Мы исключаем стандартные модули и библиотеки, которые точно не используются в проекте,
# но могут быть установлены в окружении.
excluded_modules = [
    'tkinter',
    'unittest',
    'email',
    'http',
    'xmlrpc',
    'numpy',
    'pandas',
    'matplotlib',
    'scipy',
    'pillow',
    'ipython',
    'notebook'
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_modules,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TSV_Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
