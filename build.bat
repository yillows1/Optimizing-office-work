@echo off
setlocal enabledelayedexpansion

echo ============================================================
echo  Building Verizon Packing Tools (onedir)
echo ============================================================
echo.

if not exist version.py (
  echo ERROR: version.py is missing. Create it before building.
  exit /b 1
)
for /f "tokens=2 delims== " %%V in ('findstr /B "VERSION" version.py') do set VER=%%~V
echo Version: %VER%
echo.

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo [1/5] Building SplitMaterial ...
python -m PyInstaller ^
  --onedir --windowed --noupx --name SplitMaterial ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32timezone --hidden-import pythoncom ^
  --hidden-import pywintypes --hidden-import ortools ^
  --hidden-import ortools.sat.python.cp_model ^
  --collect-all ortools --collect-data win32com ^
  --add-data "Excel Bases;Excel Bases" ^
  --paths . ^
  split_material_app.py
if errorlevel 1 goto :fail

echo.
echo [2/5] Building PackingList ...
python -m PyInstaller ^
  --onedir --windowed --noupx --name PackingList ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32timezone --hidden-import pythoncom ^
  --hidden-import pywintypes --hidden-import improvedScanner ^
  --hidden-import pypdf --hidden-import docx --hidden-import openpyxl ^
  --hidden-import reportlab ^
  --collect-data win32com ^
  --add-data "Excel Bases;Excel Bases" ^
  --paths . ^
  packing_list_gui.py
if errorlevel 1 goto :fail

echo.
echo [3/5] Building FillPaymentStatement ...
python -m PyInstaller ^
  --onedir --windowed --noupx --name FillPaymentStatement ^
  --hidden-import win32com --hidden-import win32com.client ^
  --hidden-import win32timezone --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --collect-data win32com ^
  --paths . ^
  Fill_Payment_Statement.py
if errorlevel 1 goto :fail

echo.
echo [4/5] Building PDFMerge ...
python -m PyInstaller ^
  --onedir --windowed --noupx --name PDFMerge ^
  --hidden-import pypdf ^
  --paths . ^
  pdfMergeTool.py
if errorlevel 1 goto :fail

echo.
echo [5/5] Building VerizonPackingTools (launcher) ...
python -m PyInstaller ^
  --onedir --windowed --noupx --name VerizonPackingTools ^
  --paths . ^
  launcher.py
if errorlevel 1 goto :fail

echo.
echo ============================================================
echo  Assembling shippable folder ...
echo ============================================================
set OUT=dist\VerizonPackingTools

if not exist "%OUT%\logs" mkdir "%OUT%\logs"
xcopy /E /I /Y "dist\SplitMaterial\*"        "%OUT%\SplitMaterial\" >nul
xcopy /E /I /Y "dist\PackingList\*"          "%OUT%\PackingList\" >nul
xcopy /E /I /Y "dist\FillPaymentStatement\*" "%OUT%\FillPaymentStatement\" >nul
xcopy /E /I /Y "dist\PDFMerge\*"             "%OUT%\PDFMerge\" >nul

copy /Y version.py "%OUT%\version.py" >nul
echo Version %VER% > "%OUT%\VERSION.txt"

echo.
echo ============================================================
echo  DONE.  Version: %VER%
echo  Shippable folder:  %CD%\%OUT%
echo ============================================================
goto :eof

:fail
echo.
echo BUILD FAILED. Scroll up for the error.
exit /b 1