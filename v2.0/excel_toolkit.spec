# excel_toolkit_v2.0.spec
# PyInstaller spec για την έκδοση 2.0 με custom icon και processors

# -- Import βασικά modules
block_cipher = None

a = Analysis(
    ['excel_toolkit_gui_v2.0.py'],  # Το κεντρικό GUI script
    pathex=[],
    binaries=[],
    datas=[
        ('processors/*.py', 'processors'),  # όλοι οι processors
        ('myicon.ico', '.'),               # το custom icon
        ('DejaVuSans.ttf', '.'),           # DejaVuSans font for PDF export
    ],
    hiddenimports=[
        'pandas', 'openpyxl', 'reportlab'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelToolkit_v2.0',      # όνομα exe
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,                 # κρύβουμε το terminal
    icon='myicon.ico'              # το δικό σου icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ExcelToolkit_v2.0'
)
