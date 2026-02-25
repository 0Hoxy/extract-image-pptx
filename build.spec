# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec 파일
# 빌드: pyinstaller build.spec

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'src', 'src.config', 'src.models', 'src.parser',
        'src.classifier', 'src.storage', 'src.extractor',
        'pptx', 'pptx.opc', 'pptx.opc.package', 'lxml', 'lxml._elementpath',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PPTX_이미지_추출기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=None,
)
