# build_nuitka.ps1
$ErrorActionPreference = 'Stop'

# Stop any running instance & clean old build
Stop-Process -Name main -Force -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue

# Resolve Python (prefer active venv)
$PythonCmd = $null; $PyArgs = @()
if ($env:VIRTUAL_ENV -and (Test-Path (Join-Path $env:VIRTUAL_ENV 'Scripts\python.exe'))) {
  $PythonCmd = Join-Path $env:VIRTUAL_ENV 'Scripts\python.exe'
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
  $PythonCmd = (Get-Command python).Source
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
  $PythonCmd = (Get-Command py).Source
  $PyArgs = @('-3.12')
} else { throw "No Python found. Activate your venv or install Python 3.12." }

# Ensure Nuitka is available
try { & $PythonCmd @($PyArgs + @('-c','import nuitka')) | Out-Null }
catch {
  Write-Host "Installing Nuitka in the current environment..."
  & $PythonCmd @($PyArgs + @('-m','pip','install','-U','pip','setuptools','wheel','nuitka','zstandard','ordered-set'))
}

$jobs = [int]$env:NUMBER_OF_PROCESSORS            # or leave some headroom:
# $jobs = [math]::Max(1, [int]($env:NUMBER_OF_PROCESSORS * 0.75))

# Build args (updated paths for new layout)
$NuitkaArgs = @(
  '-m','nuitka',
  '--standalone',
  '--enable-plugin=tk-inter',
  '--include-data-dir=app\ui\style=app\ui\style',
  '--include-data-dir=tessdata=tessdata',
  '--include-module=tkinterdnd2',
  '--include-module=openpyxl',
  '--include-module=pandas',
  '--include-module=numpy',
  '--include-module=pymupdf',
  '--include-module=CTkToolTip',
  '--include-module=psutil',
  '--include-module=ttkwidgets',
  '--include-package=standalone',
  '--include-package-data=customtkinter',
  '--windows-icon-from-ico=app\ui\style\Xtractor-Logo.ico',
  '--windows-console-mode=attach',
  "--jobs=$jobs",
  '--lto=no',
  '--output-dir=build',
  'main.py'
)

# Build
& $PythonCmd @($PyArgs + $NuitkaArgs)

$dist = 'build\main.dist'

# Copy ttkwidgets assets (compute paths in PS to avoid quoting issues)
try {
  $ttkPkg = & $PythonCmd @($PyArgs + @('-c','import ttkwidgets, sys; sys.stdout.write(ttkwidgets.__path__[0])'))
  if ($ttkPkg) {
    $src = Join-Path $ttkPkg 'assets'
    if (Test-Path $src) { Copy-Item -Path $src -Destination (Join-Path $dist 'ttkwidgets\assets') -Recurse -Force }
  }
} catch { Write-Warning "Could not copy ttkwidgets assets: $($_.Exception.Message)" }

# Copy CustomTkinter assets (usually already included, but safe to mirror)
try {
  $ctkPkg = & $PythonCmd @($PyArgs + @('-c','import customtkinter, sys, os; sys.stdout.write(os.path.dirname(customtkinter.__file__))'))
  if ($ctkPkg) {
    $src = Join-Path $ctkPkg 'assets'
    if (Test-Path $src) { Copy-Item -Path $src -Destination (Join-Path $dist 'customtkinter\assets') -Recurse -Force }
  }
} catch { Write-Warning "Could not copy CustomTkinter assets: $($_.Exception.Message)" }

# Run
Start-Process -FilePath (Join-Path $dist 'main.exe')
