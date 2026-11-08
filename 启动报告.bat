@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ================================================================
echo   HIT AT 效率诊断日报 · 本地服务启动器
echo ================================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 未检测到 Python，请先安装 Python 3.9+ 并加入 PATH 后重试。
  echo.
  pause
  exit /b 1
)

echo [1/2] 检查运行环境...
python -c "import pandas, numpy" >nul 2>nul
if errorlevel 1 (
  echo     缺少依赖，正在安装 pandas / numpy ...
  python -m pip install --quiet pandas numpy
)

echo [2/2] 启动报告服务并打开浏览器...
echo     报告地址: http://127.0.0.1:8760/
echo     在网页右上角「上传」可拖入新的 ZIP 与 XLSX，自动重算更新报告。
echo     按 Ctrl+C 可停止服务。
echo.

python "scripts\server.py" --port 8760 --open
pause
