# Xtractor.spec — onedir build
from pathlib import Path
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.utils.hooks import (
    collect_data_files, collect_submodules, collect_dynamic_libs
)

# Resolve project root (specs don’t always have __file__)
try:
    root = Path(__file__).parent.resolve()
except NameError:
    root = Path.cwd().resolve()

datas = []
binaries = []
hiddenimports = []

# --- Third-party deps ---
# PyMuPDF (fitz)
datas    += collect_data_files('fitz')
binaries += collect_dynamic_libs('fitz')

# tkinterdnd2 (DnD support)
datas         += collect_data_files('tkinterdnd2')
binaries      += collect_dynamic_libs('tkinterdnd2')
hiddenimports += collect_submodules('tkinterdnd2')

# ttkwidgets (CheckboxTreeview)
datas         += collect_data_files('ttkwidgets')
hiddenimports += collect_submodules('ttkwidgets')

# Pillow’s Tk bridge (sometimes missed)
hiddenimports += ['PIL.ImageTk']

# --- Your resources (map into bundle root as `style/`) ---
style_dir = root / 'app' / 'ui' / 'style'
if style_dir.is_dir():
    for p in style_dir.rglob('*'):
        if p.is_file():
            # dest folder 'style' so gui.py’s resource_path("style/...") works
            datas.append((str(p), 'style'))

# Optional: bundle tessdata/ if present (keeps folder structure)
tess_dir = root / 'tessdata'
if tess_dir.is_dir():
    for p in tess_dir.rglob('*'):
        if p.is_file():
            rel_parent = str(Path('tessdata') / p.relative_to(tess_dir).parent)
            datas.append((str(p), rel_parent))

# --- Make sure the lazily-imported standalone tools are included ---
hiddenimports += [
    'standalone.sc_pdf_dwg_list',
    'standalone.sc_dir_list',
    'standalone.sc_bulk_rename',
    'standalone.sc_bim_file_checker',
]

# --- App icon (fail early if missing) ---
ICON_PATH = (root / 'app' / 'ui' / 'style' / 'Xtractor-Logo.ico').resolve()
if not ICON_PATH.exists():
    raise SystemExit(f"Icon file not found: {ICON_PATH}")

a = Analysis(
    ['main.py'],
    pathex=[str(root), str(root / 'app')],   # 'standalone/' is under root; this covers it
    hookspath=[str(root / 'hooks')],         # keep if you have custom hooks
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
    icon=str(ICON_PATH),
    exclude_binaries=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Xtractor',
)
