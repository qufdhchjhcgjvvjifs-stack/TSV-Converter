# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Расширенный список модулей для исключения (не используются в проекте)
excluded_modules = [
    # Стандартные модули
    'tkinter',
    'unittest',
    'email',
    'xmlrpc',
    'xml',
    'xml.etree',
    'xml.dom',
    'xml.sax',
    'pydoc',
    'doctest',
    
    # Научные библиотеки
    'numpy',
    'pandas',
    'matplotlib',
    'scipy',
    'pillow',
    'PIL',
    'cv2',
    'sklearn',
    
    # Jupyter/IPython
    'ipython',
    'notebook',
    'IPython',
    'jupyter',
    'nbconvert',
    'nbformat',
    
    # Тестирование
    'pytest',
    'sphinx',
    'mock',
    'nose',
    
    # Web-фреймворки
    'django',
    'flask',
    'fastapi',
    'tornado',
    'bottle',
    'cherrypy',
    
    # Базы данных
    'sqlalchemy',
    'peewee',
    
    # Неиспользуемые Qt модули (КРИТИЧНО!)
    'PySide6.QtNetwork',
    'PySide6.QtSql',
    'PySide6.QtSvg',
    'PySide6.QtXml',
    'PySide6.QtQml',
    'PySide6.QtQuick',
    'PySide6.QtQuickWidgets',
    'PySide6.Qt3D',
    'PySide6.QtBluetooth',
    'PySide6.QtCharts',
    'PySide6.QtDataVisualization',
    'PySide6.QtDBus',
    'PySide6.QtDesigner',
    'PySide6.QtHelp',
    'PySide6.QtLocation',
    'PySide6.QtMultimedia',
    'PySide6.QtMultimediaWidgets',
    'PySide6.QtNfc',
    'PySide6.QtOpenGL',
    'PySide6.QtOpenGLWidgets',
    'PySide6.QtPdf',
    'PySide6.QtPdfWidgets',
    'PySide6.QtPositioning',
    'PySide6.QtPrintSupport',
    'PySide6.QtRemoteObjects',
    'PySide6.QtScxml',
    'PySide6.QtSensors',
    'PySide6.QtSerialBus',
    'PySide6.QtSerialPort',
    'PySide6.QtSpatialAudio',
    'PySide6.QtStateMachine',
    'PySide6.QtTest',
    'PySide6.QtTextToSpeech',
    'PySide6.QtUiTools',
    'PySide6.QtWebChannel',
    'PySide6.QtWebEngine',
    'PySide6.QtWebEngineCore',
    'PySide6.QtWebEngineWidgets',
    'PySide6.QtWebSockets',
]

# Минимальный набор скрытых импортов - только используемые модули
hidden_imports = [
    'xlsxwriter',
]

# Путь к файлу иконки
import os
icon_path = 'icon.ico'
if not os.path.exists(icon_path):
    icon_path = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
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
    icon=icon_path if icon_path else None,
)
