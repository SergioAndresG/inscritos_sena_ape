# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ===== CONFIGURACIÓN DE PATHS =====
block_cipher = None

# Recopilar datos adicionales
datas = []

# Archivos de datos necesarios
try:
    # Selenium y webdriver_manager
    datas += collect_data_files('selenium')
    datas += collect_data_files('webdriver_manager')
except:
    pass

# Agregar archivos específicos del proyecto
datas += [
    ('perfilesOcupacionales/perfiles_ocupacionales.json', 'perfilesOcupacionales'),
]

# Iconos (si existen)
import os
if os.path.exists('Iconos/logoSena.ico'):
    icon_path = 'Iconos/logoSena.ico'
else:
    icon_path = None

# ===== MÓDULOS OCULTOS =====
hiddenimports = [
    # Core Python
    'queue',
    'threading',
    'logging',
    'json',
    'pathlib',
    'shutil',
    'time',
    'datetime',
    'traceback',
    
    # Selenium y WebDriver
    'selenium',
    'selenium.webdriver',
    'selenium.webdriver.chrome',
    'selenium.webdriver.chrome.service',
    'selenium.webdriver.chrome.options',
    'selenium.webdriver.common.by',
    'selenium.webdriver.support.ui',
    'selenium.webdriver.support.expected_conditions',
    'webdriver_manager',
    'webdriver_manager.chrome',
    
    # Excel
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.styles.fills',
    'openpyxl.styles.fonts',
    'openpyxl.styles.alignment',
    'openpyxl.styles.borders',
    'openpyxl.styles.colors',
    'openpyxl.styles.patterns',
    'openpyxl.cell',
    'openpyxl.cell.cell',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.worksheet.worksheet',
    'xlrd',  # Para leer .xls antiguos
    'pandas',
    'pandas.io.excel',
    'pandas.io.excel._openpyxl',
    
    # CustomTkinter
    'customtkinter',
    'tkinter',
    'tkinter.filedialog',
    'tkinter.messagebox',
    
    # Otros
    'dotenv',
    'requests',
    'urllib3',
    'certifi',
    
    # Módulos del proyecto
    'funciones_formularios',
    'funciones_formularios.login',
    'funciones_formularios.verificacion',
    'funciones_formularios.pre_inscripcion',
    'funciones_formularios.campo_telefono_correo',
    'funciones_formularios.campo_estrato',
    'funciones_formularios.campo_sueldo',
    'funciones_formularios.experincia_laboral_campos',
    'funciones_formularios.form_campo_estado_civil',
    'funciones_formularios.form_campo_perfil_ocupacional',
    'funciones_formularios.form_campos_nacimiento',
    'funciones_formularios.form_campos_ubicacion_identificacion',
    'funciones_formularios.form_datos_residencia',
    'funciones_formularios.meses_busqueda',
    
    'funciones_excel',
    'funciones_excel.conversion_excel',
    'funciones_excel.extraccion_datos_excel',
    'funciones_excel.preparar_excel',
    
    'funciones_loggs',
    'funciones_loggs.loggs_funciones',
    
    'perfilesOcupacionales',
    'perfilesOcupacionales.perfilExcepcion',
    'perfilesOcupacionales.gestorDePerfilesOcupacionales',
    'perfilesOcupacionales.dialogo_perfil',
    
    'URLS',
    'URLS.urls',
    
    'debug_exe',
    'automatizacion',
]

# ===== ANALYSIS =====
a = Analysis(
    ['GUI.py'],  # Archivo principal
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Excluir módulos innecesarios para reducir tamaño
        'matplotlib',
        'numpy.testing',
        'scipy',
        'IPython',
        'notebook',
        'pytest',
        'unittest',
        'test',
        'tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ===== PYZ =====
pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher
)

# ===== EXE =====
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SENA_Automation_App',  # Nombre del ejecutable
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Comprimir con UPX
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # True para ver logs de debug, False para producción
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,  # Icono del ejecutable
)