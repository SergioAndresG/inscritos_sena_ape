# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Archivo principal
        ('automatizacion.py', '.'),
        # Icono
        ('Iconos/logoSena.ico', 'Iconos'),
        ('perfilesOcupacionales/', 'perfilesOcupacionales'),
        ('perfilesOcupacionales/perfiles_ocupacionales.json', 'perfilesOcupacionales'),
        ('funciones_formularios/', 'funciones_formularios'),
        ('funciones_loggs/', 'funciones_loggs'),
        ('URLS/', 'URLS'),
        ('.env', '.'),
    ],
    hiddenimports=[
        # Librerías estándar
        'json',
        'os',
        'pathlib',
        'queue',
        'threading',
        'tkinter',
        'logging',
        'sys',
        'time',
        
        # CustomTkinter y dependencias
        'customtkinter',
        'PIL',
        'PIL._tkinter_finder',
        
        # Selenium y dependencias
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.common.by',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        
        # Pandas y Excel
        'pandas',
        'xlrd',
        'xlwt',
        'xlutils',
        'xlutils.copy',
        
        'perfilesOcupacionales.gestorDePerfilesOcupacionales',
        'perfilesOcupacionales.dialogo_perfil',
        'perfilesOcupacionales.perfilExcepcion',
        
        # Otros
        'dotenv',
        'webdriver_manager',
        'webdriver_manager.chrome',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SENA_Automation_App',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Sin consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Iconos/logoSena.ico',
)