@echo off
chcp 65001 >nul
title 证件核验报表工具
cd /d "%~dp0"

set "PY="
for %%C in (python.exe py.exe) do (
    if not defined PY (
        where %%C >nul 2>nul && set "PY=%%C"
    )
)
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY if exist "C:\Python313\python.exe" set "PY=C:\Python313\python.exe"
if not defined PY if exist "C:\Python314\python.exe" set "PY=C:\Python314\python.exe"

if not defined PY (
    echo.
    echo   没有找到 Python。
    echo   请先安装 Python 3.9 以上版本：https://www.python.org/downloads/
    echo   安装时记得勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

"%PY%" -c "import openpyxl, xlrd" >nul 2>nul
if errorlevel 1 (
    echo 首次运行，正在安装必要的组件，请稍候...
    "%PY%" -m pip install -q -i https://mirrors.aliyun.com/pypi/simple/ openpyxl xlrd
    if errorlevel 1 (
        echo.
        echo   组件安装失败，请检查网络后重试。
        pause
        exit /b 1
    )
)

echo 正在启动，浏览器会自动打开...
"%PY%" app.py
pause
