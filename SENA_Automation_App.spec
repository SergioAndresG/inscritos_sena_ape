# -*- mode: python ; coding: utf-8 -*-


a = Analysis(

    ['gui.py'], 
    pathex=[],
    binaries=[],
    # Incluimos automatizacion.py en la raíz y el icono dentro de una carpeta 'Iconos'
    datas=[('automatizacion.py', '.'), ('Iconos/logoSena.ico', 'Iconos')],
    
    # Unificamos y mantenemos las importaciones útiles para Tkinter/CustomTkinter
    hiddenimports=['json', 'os', 'pathlib', 'queue', 'threading', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SENA_Automation_App',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    
    
    # 1. Añadimos el icono al ejecutable final 
    icon='Iconos/logoSena.ico', 
    
    # Cambiar a False para ocultar la ventana de consola negra
    console=False, 
    
    # 3. Quitamos hiddenimports=[] repetido
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)