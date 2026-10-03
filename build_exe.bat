@echo off
chcp 65001 >nul
cd /d %~dp0

echo ============================================
echo  [1/2] 安装依赖 (pyinstaller / openpyxl / xlrd)
echo ============================================
python -m pip install pyinstaller openpyxl xlrd

echo ============================================
echo  [2/2] 正在打包 exe，约需 1-2 分钟...
echo ============================================
set EXCLUDES=--exclude-module torch --exclude-module transformers --exclude-module pandas --exclude-module numpy --exclude-module scipy --exclude-module sklearn --exclude-module matplotlib --exclude-module PIL --exclude-module sympy --exclude-module networkx --exclude-module pygments --exclude-module rich --exclude-module sqlalchemy --exclude-module pydantic --exclude-module huggingface_hub --exclude-module opentelemetry --exclude-module xlsxwriter --exclude-module cv2 --exclude-module onnxruntime --exclude-module safetensors --exclude-module tokenizers --exclude-module fastapi --exclude-module aiohttp --exclude-module IPython --exclude-module jupyter --exclude-module numba --exclude-module numexpr --exclude-module pyarrow --exclude-module tqdm --exclude-module pytest --exclude-module setuptools

python -m PyInstaller --noconfirm --onefile --windowed --name 邮件群发助手 源代码\mail_gui.py %EXCLUDES%

echo.
echo 打包完成！exe 位于: dist\邮件群发助手.exe
pause
