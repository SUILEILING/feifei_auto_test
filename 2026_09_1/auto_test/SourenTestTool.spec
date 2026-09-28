# -*- mode: python ; coding: utf-8 -*-
# ============================================================
#  重要：打包必须用这个 spec 文件：
#      pyinstaller --clean --noconfirm SourenTestTool.spec
#  不要用 "pyinstaller ... gui_app.py"，那会重新生成并覆盖本文件，
#  导致 pyvisa_py 后端未被打包，exe 运行时报“无法创建 ResourceManager”。
# ============================================================

import importlib.util
from PyInstaller.utils.hooks import collect_submodules, collect_data_files


def _has(mod):
    try:
        return importlib.util.find_spec(mod) is not None
    except Exception:
        return False


# ---- VISA 后端：必须显式收集 pyvisa_py，否则打包后无可用后端 ----
hiddenimports = []
hiddenimports += collect_submodules('pyvisa')
if _has('pyvisa_py'):
    hiddenimports += collect_submodules('pyvisa_py')
    hiddenimports += ['pyvisa_py']
else:
    raise SystemExit(
        "构建失败：当前 Python 未安装 pyvisa_py。请先执行: pip install pyvisa-py"
    )
for _opt in ('zeroconf', 'serial', 'usb', 'psutil'):
    if _has(_opt):
        hiddenimports.append(_opt)

# 测试报告页的交互图表依赖 pyqtgraph
if _has('pyqtgraph'):
    hiddenimports += collect_submodules('pyqtgraph')

datas = [('image.png', '.')]
datas += collect_data_files('pyvisa')
datas += collect_data_files('pyvisa_py')


a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='SourenTestTool',
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
