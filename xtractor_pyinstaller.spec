# Xtractor.spec — onedir build
from pathlib import Path
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs
import sys

# spec files don't have __file__; fall back to CWD
try:
    root = Path(__file__).parent.resolve()   # won't exist in .spec
except NameError:
    root = Path.cwd().resolve()

datas = []
binaries = []
hiddenimports = []

# --- third-party ---
datas    += collect_data_files('fitz')
binaries += collect_dynamic_libs('fitz')

datas         += collect_data_files('tkinterdnd2')
binaries      += collect_dynamic_libs('tkinterdnd2')
hiddenimports += collect_submodules('tkinterdnd2')

datas         += collect_data_files('ttkwidgets')
hiddenimports += collect_submodules('ttkwidgets')

# --- your resources ---
style_dir = root / 'app' / 'ui' / 'style'
if style_dir.is_dir():
    for p in style_dir.rglob('*'):
        if p.is_file():
            datas.append((str(p), 'style'))

tess_dir = root / 'tessdata'
if tess_dir.is_dir():
    for p in tess_dir.rglob('*'):
        if p.is_file():
            rel_parent = str((Path('tessdata') / p.relative_to(tess_dir).parent).as_posix())
            datas.append((str(p), rel_parent))

# --- icon (must be a valid .ico) ---
ICON_PATH = (root / 'app' / 'ui' / 'style' / 'Xtractor-Logo.ico').resolve()

# hard-fail if missing so you don't silently get the default icon
if not ICON_PATH.exists():
    raise SystemExit(f"Icon file not found: {ICON_PATH}")

print(f"[spec] Using icon: {ICON_PATH}")  # sanity log

a = Analysis(
    ['main.py'],
    pathex=[str(root), str(root / 'app')],
    hookspath=[str(root / 'hooks')],
    datas=datas,
    binaries=binaries,
    hiddenimports=hiddenimports,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    name='Xtractor',
    console=False,
    icon=str(ICON_PATH),     # ← absolute path, no ambiguity
    exclude_binaries=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Xtractor',
)
