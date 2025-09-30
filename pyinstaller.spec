# xtractor.spec
from PyInstaller.utils.hooks import collect_all, collect_data_files
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT
from PyInstaller.building.datastruct import Tree

# collect pymupdf
pm_datas, pm_bins, pm_hidden = collect_all('pymupdf')
# collect tkinterdnd2 (same behavior as our hook)
tkdnd_datas, tkdnd_bins, tkdnd_hidden = collect_all('tkinterdnd2')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=pm_bins + tkdnd_bins,
    datas=pm_datas + tkdnd_datas + Tree('style', prefix='style').toc,
    hiddenimports=pm_hidden + tkdnd_hidden + ['tkinterdnd2'],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, console=True, name='Xtractor')  # set console=False if you want
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, name='Xtractor')
