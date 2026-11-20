@echo off
rem ============================================================================
rem  Fabrique l'application PCR-CI pour Windows (a lancer sur l'ordinateur du
rem  developpeur, qui a Python). Le resultat se copie sur n'importe quel PC
rem  Windows 10/11 : il n'a besoin ni de Python, ni de Git, ni d'autre chose.
rem
rem  Resultat dans le dossier dist\ :
rem    - PCR-CI-portable-AAAA.MM.JJ.zip : a dezipper, puis double-clic sur PCR-CI.exe
rem    - PCR-CI-Installation-AAAA.MM.JJ.exe : installateur (si Inno Setup 6 est installe)
rem ============================================================================
setlocal EnableExtensions
cd /d "%~dp0"
title Fabrication de l'application PCR-CI
echo.
echo  ==========================================================
echo     Fabrication de l'application PCR-CI
echo  ==========================================================
echo.

rem ---------- 1. Python (python.org, version 3.11 a 3.13) ----------
set "PY="
where py >nul 2>nul && py -3.12 -c "import sys" >nul 2>nul && set "PY=py -3.12"
if not defined PY (where py >nul 2>nul && py -3.13 -c "import sys" >nul 2>nul && set "PY=py -3.13")
if not defined PY (where py >nul 2>nul && py -3.11 -c "import sys" >nul 2>nul && set "PY=py -3.11")
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY goto :sans_python
%PY% -c "import sys; ok = sys.platform == 'win32' and (3, 11) <= sys.version_info[:2] <= (3, 13) and 'GCC' not in sys.version; sys.exit(0 if ok else 1)"
if errorlevel 1 goto :mauvais_python
echo [1/5] Python trouve :
%PY% --version

rem ---------- 2. Environnement de fabrication (separe, ne touche a rien d'autre) ----------
echo [2/5] Installation des composants (premiere fois : 3 a 5 minutes)...
if not exist ".venv-fabrication\Scripts\python.exe" (
    %PY% -m venv .venv-fabrication
    if errorlevel 1 goto :erreur
)
set "VPY=%CD%\.venv-fabrication\Scripts\python.exe"
"%VPY%" -m pip install --quiet --upgrade pip
"%VPY%" -m pip install --quiet -r requirements.txt -r requirements-app.txt pyinstaller pytest
if errorlevel 1 goto :erreur

rem ---------- 3. Tests (un code casse ne doit pas partir chez le collegue) ----------
echo [3/5] Verification du code (tests, environ 1 minute)...
"%VPY%" -m pytest -q tests
if errorlevel 1 (
    echo.
    echo  ATTENTION : des tests ont echoue ^(voir ci-dessus^).
    choice /c ON /m " Fabriquer quand meme l'application"
    if errorlevel 2 goto :fin
)

rem ---------- 4. Application (PyInstaller) ----------
echo [4/5] Fabrication de l'application (2 a 5 minutes)...
if exist "dist\PCR-CI" rmdir /s /q "dist\PCR-CI"
"%VPY%" -m PyInstaller --noconfirm --clean --windowed --name PCR-CI --icon "%CD%\packaging\icone.ico" ^
  --distpath "%CD%\dist" --workpath "%CD%\build" --specpath "%CD%\build" ^
  --add-data "%CD%\config;config" --add-data "%CD%\modeles;modeles" --add-data "%CD%\kobo;kobo" ^
  --add-data "%CD%\pcr_app\icone.png;pcr_app" ^
  --collect-all nicegui --collect-all webview --collect-all pythonnet --collect-all clr_loader ^
  --collect-all pyxform --collect-all docx --collect-data openpyxl ^
  --collect-submodules pcr --collect-submodules pcr_app --exclude-module pkg_resources ^
  --paths "%CD%" "%CD%\packaging\lancer_pcr_ci.py"
if errorlevel 1 goto :erreur
if not exist "dist\PCR-CI\PCR-CI.exe" goto :erreur

rem ---------- 5. Version a distribuer ----------
echo [5/5] Preparation des fichiers a envoyer...
for /f "delims=" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyy.MM.dd"') do set "VERSION=%%d"
if exist "dist\PCR-CI-portable-%VERSION%.zip" del /q "dist\PCR-CI-portable-%VERSION%.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\PCR-CI' -DestinationPath 'dist\PCR-CI-portable-%VERSION%.zip' -Force"
if errorlevel 1 goto :erreur

set "ISCC="
for %%i in ("%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" "%ProgramFiles%\Inno Setup 6\ISCC.exe" "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe") do if exist "%%~i" set "ISCC=%%~i"
if defined ISCC (
    "%ISCC%" /Q "/DAppVersion=%VERSION%" packaging\pcr_ci.iss
    if errorlevel 1 echo  Installateur non fabrique ^(le ZIP portable suffit^).
) else (
    echo  Inno Setup 6 absent : pas d'installateur, le ZIP portable suffit.
    echo  ^(Pour avoir aussi un installateur : https://jrsoftware.org/isdl.php puis relancer ce fichier.^)
)

echo.
echo  ==========================================================
echo     TERMINE. Fichiers a envoyer (dossier dist) :
echo  ==========================================================
dir /b dist\*.zip dist\*.exe 2>nul
echo.
echo  Sur l'ordinateur du collegue : dezipper PCR-CI-portable-%VERSION%.zip
echo  (par exemple dans Documents), puis double-clic sur PCR-CI.exe.
echo  Ou bien lancer PCR-CI-Installation-%VERSION%.exe s'il a ete fabrique.
start "" explorer "%CD%\dist"
goto :fin

:sans_python
echo.
echo  Python n'est pas installe sur CET ordinateur (celui qui fabrique l'application).
echo  1. Telecharger Python 3.12 : https://www.python.org/downloads/windows/
echo  2. A l'installation, cocher "Add python.exe to PATH".
echo  3. Relancer ce fichier.
echo  (L'ordinateur du collegue, lui, n'a besoin de rien.)
goto :fin

:mauvais_python
echo.
echo  Le Python trouve ne convient pas (il faut Python 3.11, 3.12 ou 3.13 de python.org,
echo  pas celui de MSYS2 / Git Bash). Installer Python 3.12 depuis
echo  https://www.python.org/downloads/windows/ puis relancer ce fichier.
goto :fin

:erreur
echo.
echo  ECHEC de la fabrication. Copier les dernieres lignes ci-dessus et les envoyer.
goto :fin

:fin
echo.
pause
endlocal
