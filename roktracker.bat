".\platform-tools\adb.exe" kill-server
".\platform-tools\adb.exe" connect localhost:5555
cd .\
REM Check if pip is installed
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo pip is not installed, installing...
    python -m ensurepip --default-pip
    if %errorlevel% neq 0 (
        echo Failed to install pip, please install manually
        exit /b
    )
)
pip install --upgrade pip
pip install Pillow==8.4.0
pip install opencv-python==4.6.0.66
pip install -r requirements.txt
python roktracker.py
PAUSE
