@echo off
setlocal

echo ============================================================
echo  Building Verizon Packing Tools (onedir)
echo ============================================================
echo.

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist SplitMaterial.spec del /q SplitMaterial.spec
if exist PackingList.spec del /q PackingList.spec
if exist VerizonPackingTools.spec del /q VerizonPackingTools.spec

echo [1/3] Building SplitMaterial ...
python -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name SplitMaterial ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import win32timezone ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import ortools ^
  --hidden-import ortools.sat.python.cp_model ^
  --collect-all ortools ^
  --collect-data win32com ^
  --add-data "Excel Bases;Excel Bases" ^
  split_material_app.py
if errorlevel 1 goto :fail

echo.
echo [2/3] Building PackingList ...
python -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name PackingList ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import win32timezone ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import improvedScanner ^
  --hidden-import pypdf ^
  --hidden-import docx ^
  --hidden-import openpyxl ^
  --collect-data win32com ^
  --add-data "Excel Bases;Excel Bases" ^
  --paths . ^
  packing_list_gui.py
if errorlevel 1 goto :fail

echo.
echo [3/3] Building VerizonPackingTools (launcher) ...
python -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name VerizonPackingTools ^
  launcher.py
if errorlevel 1 goto :fail

echo.
echo ============================================================
echo  Assembling shippable folder ...
echo ============================================================
set OUT=dist\VerizonPackingTools
if not exist "%OUT%" mkdir "%OUT%"

xcopy /E /I /Y "dist\VerizonPackingTools\*" "%OUT%\" >nul
xcopy /E /I /Y "dist\SplitMaterial\*" "%OUT%\SplitMaterial\" >nul
xcopy /E /I /Y "dist\PackingList\*"   "%OUT%\PackingList\"   >nul

echo.
echo ============================================================
echo  DONE.
echo.
echo  Shippable folder:  %CD%\%OUT%
echo ============================================================
goto :eof

:fail
echo.
echo BUILD FAILED. Scroll up for the error.
exit /b 1