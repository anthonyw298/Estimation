# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('utils', 'utils'),
        ('systems', 'systems'),
        ('data', 'data'),
        ('config.py', '.'),
    ],
    hiddenimports=[
        'flet',
        'openpyxl',
        'reportlab',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.units',
        'reportlab.lib.colors',
        'reportlab.platypus',
        'reportlab.platypus.tables',
        'reportlab.graphics',
        'reportlab.graphics.shapes',
        'reportlab.graphics.charts',
        'reportlab.graphics.charts.piecharts',
        'sklearn',
        'sklearn.linear_model',
        'sklearn.preprocessing',
        'numpy',
        'pandas',
        'PIL',
        'supabase',
        'postgrest',
        'httpx',
        'storage3',
        'realtime',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# onedir mode: binaries/datas go into COLLECT, not EXE
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EstimatorApp',
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EstimatorApp',
)

app = BUNDLE(
    coll,
    name='EstimatorApp.app',
    bundle_identifier='com.estimation.app',
    info_plist={
        'CFBundleName': 'EstimatorApp',
        'CFBundleDisplayName': 'Estimator App',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
    },
)
