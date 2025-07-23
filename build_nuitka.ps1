# Clean previous build
# Stop any running main.exe to release file lock
Stop-Process -Name main -Force -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue

# Compile with Nuitka
py -3.12 -m nuitka --standalone --enable-plugin=tk-inter --include-data-dir=style=style --include-data-dir=tessdata=tessdata --include-module=tkinterdnd2 --include-module=openpyxl --include-module=pandas --include-module=numpy --include-module=pymupdf --include-module=CTkToolTip --include-module=psutil --include-module=ttkwidgets --include-package-data=customtkinter --windows-icon-from-ico=style/Xtractor-Logo.ico --windows-console-mode=attach --jobs=24 --lto=no --output-dir=build main.py

# Ensure style directory is present in distribution
Copy-Item -Path style -Destination build\main.dist\style -Recurse -Force

# Ensure ttkwidgets assets directory is present
$ttkAssets = py -3.12 -c "import ttkwidgets, os; print(os.path.join(ttkwidgets.__path__[0], 'assets'))"
Copy-Item -Path $ttkAssets -Destination build\main.dist\ttkwidgets\assets -Recurse -Force

# Ensure CustomTkinter assets and themes directories are present
$ctkAssets = py -3.12 -c "import customtkinter, os; print(os.path.join(customtkinter.__path__[0], 'assets'))"
Copy-Item -Path $ctkAssets -Destination build\main.dist\customtkinter\assets -Recurse -Force
$ctkThemes = py -3.12 -c "import customtkinter, os; print(os.path.join(customtkinter.__path__[0], 'assets', 'themes'))"
Copy-Item -Path $ctkThemes -Destination build\main.dist\customtkinter\assets\themes -Recurse -Force

# Launch the executable
Start-Process -FilePath .\build\main.dist\main.exe
