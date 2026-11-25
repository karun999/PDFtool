@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title PDF Studio Pro - Build EXE
echo.
echo ==============================================
echo     PDF Studio Pro  -  building the .exe
echo ==============================================
echo.

if not exist "pdf_studio_pro.py" (
    echo ERROR: pdf_studio_pro.py was not found next to this file.
    pause
    exit /b 1
)

rem --- The OCR engine needs Python 3.12 or 3.11 (it crashes on 3.14) ---
set "PYCMD="
set "WITH_OCR=1"
py -3.12 --version >nul 2>&1 && set "PYCMD=py -3.12"
if not defined PYCMD (
    py -3.11 --version >nul 2>&1 && set "PYCMD=py -3.11"
)
if not defined PYCMD (
    echo Python 3.12 / 3.11 was not found.
    echo The .exe will be built WITHOUT the OCR feature, using your default Python.
    echo To include OCR: install Python 3.12 from python.org, then run this file again.
    echo.
    set "PYCMD=python"
    set "WITH_OCR=0"
)
echo Using:
!PYCMD! --version
echo.

rem --- Clean, separate environment: only the libraries the app needs ---
if exist build_env rmdir /s /q build_env
!PYCMD! -m venv build_env
if errorlevel 1 ( echo ERROR: could not create the build environment. & pause & exit /b 1 )
call build_env\Scripts\activate.bat

echo Installing libraries (first time takes a few minutes)...
python -m pip install --upgrade pip >nul
set "PKGS=PyQt5 pypdf cryptography pymupdf openpyxl pyinstaller"
rem onnxruntime 1.19.2 works with the C++ runtime that PyQt5 ships (newer ones crash)
if "!WITH_OCR!"=="1" set "PKGS=!PKGS! rapidocr_onnxruntime onnxruntime==1.19.2"
python -m pip install !PKGS!
if errorlevel 1 ( echo ERROR: installing the libraries failed. & pause & exit /b 1 )

rem --- Build ---
set "OPTS=--noconfirm --clean --windowed --name PDFStudioPro"
if "!WITH_OCR!"=="1" (
    set "OPTS=!OPTS! --collect-all rapidocr_onnxruntime --collect-binaries onnxruntime"
) else (
    set "OPTS=!OPTS! --exclude-module rapidocr_onnxruntime --exclude-module onnxruntime --exclude-module cv2"
)
echo.
echo Building the app (2-5 minutes)...
python -m PyInstaller !OPTS! pdf_studio_pro.py
if errorlevel 1 if "!WITH_OCR!"=="1" (
    echo.
    echo ----------------------------------------------------------
    echo  The build with OCR failed. Building again WITHOUT OCR so
    echo  you still get a working app (all other tools included)...
    echo ----------------------------------------------------------
    set "WITH_OCR=0"
    python -m pip uninstall -y rapidocr_onnxruntime onnxruntime opencv-python >nul 2>&1
    python -m PyInstaller --noconfirm --clean --windowed --name PDFStudioPro --exclude-module rapidocr_onnxruntime --exclude-module onnxruntime --exclude-module cv2 pdf_studio_pro.py
)
if errorlevel 1 ( echo. & echo BUILD FAILED - please send a screenshot of the messages above. & pause & exit /b 1 )
if "!WITH_OCR!"=="0" echo NOTE: this build does not include OCR.

echo.
echo ==============================================
echo  DONE!  Your app is here:
echo    %~dp0dist\PDFStudioPro\PDFStudioPro.exe
echo.
echo  To share it, copy the whole "PDFStudioPro" folder.
echo ==============================================
explorer "dist\PDFStudioPro"
pause
