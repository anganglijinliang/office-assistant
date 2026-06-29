# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# 所有含 C 扩展/数据文件的库必须用 collect_all 收集
lxml_datas, lxml_binaries, lxml_hidden = collect_all('lxml')
docx_datas, docx_binaries, docx_hidden = collect_all('docx')
openpyxl_datas, openpyxl_binaries, openpyxl_hidden = collect_all('openpyxl')
pil_datas, pil_binaries, pil_hidden = collect_all('PIL')
reportlab_datas, reportlab_binaries, reportlab_hidden = collect_all('reportlab')
pypdfium2_datas, pypdfium2_binaries, pypdfium2_hidden = collect_all('pypdfium2')
pypdf2_datas, pypdf2_binaries, pypdf2_hidden = collect_all('PyPDF2')
pdf2docx_datas, pdf2docx_binaries, pdf2docx_hidden = collect_all('pdf2docx')
qrcode_datas, qrcode_binaries, qrcode_hidden = collect_all('qrcode')


a = Analysis(
    ['office_assistant.py'],
    pathex=['C:\\Users\\Administrator\\Desktop\\OfficeAssistant_v6.2'],
    binaries=(lxml_binaries + docx_binaries + openpyxl_binaries + pil_binaries
              + reportlab_binaries + pypdfium2_binaries + pypdf2_binaries
              + pdf2docx_binaries + qrcode_binaries),
    datas=[('utils.py', '.'), ('lib_license.py', '.'), ('modules', 'modules'), ('shared', 'shared')]
          + lxml_datas + docx_datas + openpyxl_datas + pil_datas
          + reportlab_datas + pypdfium2_datas + pypdf2_datas
          + pdf2docx_datas + qrcode_datas,
    hiddenimports=['pystray', 'win32api', 'win32com.client', 'pkg_resources']
                 + lxml_hidden + docx_hidden + openpyxl_hidden + pil_hidden
                 + reportlab_hidden + pypdfium2_hidden + pypdf2_hidden
                 + pdf2docx_hidden + qrcode_hidden,
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
    name='万能办公助手v6.2商业版',
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
