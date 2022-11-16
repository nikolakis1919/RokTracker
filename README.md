# RokTracker
Open Source Rise of Kingdoms Stats Management Tool. Track TOP (50-975) Players in Power/KillPoints. Creates a .xls file with in-game Id, Power, Kill Points, T1/T2/T3/T4/T5 Kills, Dead troops, RSS Assistance, Alliance **!NEW FEATURE!**.

# Required
1. Bluestacks 5 Installation
https://cdn3.bluestacks.com/downloads/windows/nxt/5.4.100.1026/0129e8eb74f84fc396a1500329365a09/BlueStacksMicroInstaller_5.4.100.1026_native.exe?filename=BlueStacksMicroInstaller_5.4.100.1026_native_5ffb0694218e1b99e7000bed6dcbe547_0.exe
2. **Python 3.6.8 Installation(MAKE SURE YOU CLICK ADD TO PATH OPTION IN INSTALLER)** https://www.python.org/ftp/python/3.6.8/python-3.6.8-amd64.exe
3. Tesseract-OCR Installation https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.0.0.20211201.exe
4. Adb Platform Tools Download and Extract(See Important Notes) https://dl.google.com/android/repository/platform-tools_r31.0.3-windows.zip
5. Windows 10-Tested, Windows 7 or Windows 11 could be fine.

# Bluestacks 5 Settings
![screen-20 40 24 14 12 2021](https://user-images.githubusercontent.com/96141261/146060069-d0c138e6-a083-4add-96a3-9b3d41f27420.png)
![screen-20 43 02 14 12 2021](https://user-images.githubusercontent.com/96141261/146060189-acc8cba8-5f06-4f1d-8cfe-d9aaf03344b8.png)
![screen-20 43 59 14 12 2021](https://user-images.githubusercontent.com/96141261/146060299-01dc3881-44a3-4a5f-97f8-b220bdda52d5.png)

Everything should be like this and no more emulators should be running concurrent. If you are running more emulators, make sure, that in Settings->Advanced, port is 5555 for the emulator you want to use for the scan.

# Important Notes
1. Every requirement that tracker needs in order to run, will be installed automatically, by running the .bat file.
2. You must run the .bat file without administration privileges.
3. Change the installation path of Tesseract in .py file if your path is different. Open this file with your notepad and change the path to yours at 18th line of code. My path is C:\Program Files\Tesseract-OCR\tesseract.exe
4. Platform tools folder should be extracted as a folder inside the Tracker's folder. Like the following image
![image](https://user-images.githubusercontent.com/96141261/191936152-f0fbd0c0-a026-48d0-bd1e-07947e2805bd.png)
5. In order to get only your kingdoms ranks, the character that is currently logged in game must be in HOME KINGDOM, else you will get all the players in your KvK including players from different kingdoms.
6. Game Language should be English. Anything else will cause trouble in detecting inactive governors. Change it only for scan, if yours is different and then switch back.
7. The view before running the programme should be at the top of power rankings or at the top of kill points rankings. No move should be made in this window until scanning is done.
8. Account must be lower in ranks than the amount of players you want to scan. e.g. Cannot scan top 100 when character's rank is 85. Use a farm account instead.
9. You can see tool's progress in CMD when it's running.
10. Chinese letters are not shown properly in CMD but they are visible in the final .xls file.
11. Bluestacks settings must be the same as in pictures above.
12. You can do whatever you want in your computer when tool is scanning.
13. Resume Scan option starts scanning the middle governor that is displayed in screen. The 4th in order. So before starting the tool make sure that you are in the correct view in bluestacks.
14. BE CAREFUL to always copy the .xls file from the RokTracker folder when it is created, because in the next capture, there is a chance to overwrite.
15. V2.0 New Added Feature. With \ button you can terminate the scan, saving current progress to .xls file. It will scan the current scanning governor and it won't continue to the next. Appropriate messages will appear in CMD.

# Contact and Support
Any bugs that you may find or any suggestions you may have, please feel free to contact me in my discord: nikos#4469\
If you like what i created, you can always support me.\
https://www.paypal.com/donate/?hosted_button_id=55G95MNYPVX72

# Usage
Options Screen. Here you can select the amount of players you want to scan and your kingdom. \
Added option in v2.0 to resume Scan. The results will be saved in new file. You can merge manually if you wish.\
![image](https://user-images.githubusercontent.com/96141261/146555364-d55cc627-cbf9-45de-8962-ec05e03ad009.png)


Single result in cmd.\
![image](https://user-images.githubusercontent.com/96141261/146094135-9b869feb-722b-43cb-8623-f2cbfc7d0052.png)


Results in excel file.\
![screen-01 19 42 15 12 2021](https://user-images.githubusercontent.com/96141261/146095176-96dcacb2-9c3e-48c7-8b8f-ac2e91973901.png)






