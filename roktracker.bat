".\platform-tools\adb.exe" kill-server
".\platform-tools\adb.exe" connect localhost:5555
cd .\
pip install --upgrade pip
pip install Pillow==8.4.0
pip install opencv-python==4.6.0.66
pip install -r requirements.txt
python roktracker.py
PAUSE
