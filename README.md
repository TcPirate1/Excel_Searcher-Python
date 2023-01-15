# Excel Searcher - Python

This is a Python code that I have written to search through my excel spreadsheet (that I use to keep track of my card collection).

The script runs on both Windows and Linux OS but relies on the file being in Documents/Spreadsheets/ directory. It also needs the file to be called FFTCG_automated.xlsx but is possible to change the name if needed.

## How to use?
- Make sure the excel file is in the Documents/Spreadsheets/ directory and named FFTCG_automated.xlsx
- Download the code by cloning the repository
- Run the code in an IDE or go to the folder its in and open the command-prompt/terminal inside the folder
- When the command-prompt/terminal opens it should already be in the folder's directory
- Create a Python virtual environment with `py -m venv <file name>` or `python3 -m venv <file name>`

Open the .bat (on Windows) or .desktop file (on Linux) to run the [script](#.bat,-.desktop-and-.sh-file-contents).

## Video (for Windows)

https://user-images.githubusercontent.com/88120195/211103204-5766d9cc-09d3-4c73-96e5-893844f4c894.mp4

## Video (for Linux)

https://user-images.githubusercontent.com/88120195/211176332-3a26de9b-ffa8-4e1e-8f30-5b005fe26070.mp4



## .bat, .desktop and .sh file contents
.bat:

`CALL C:\Users\teren\Desktop\Python_stuff\Programs\Excel_Searcher\env\Scripts\activate.bat
@py.exe C:\Users\teren\Desktop\Python_stuff\Programs\Excel_Searcher\Search_fftcg.py %*
@pause`

.desktop:

`[Desktop Entry]
Name=card_script
Exec=/home/tc/Desktop/Python_Scripts/Excel_Searcher-Python/card_script.sh
Type=Application
Categories=GTK;GNOME;Utility;
Icon=
Comment=
Path=
Terminal=true
StartupNotify=false`

.sh:

`#!/usr/bin/env bash
activate() {
source /home/tc/Desktop/Python_Scripts/Excel_Searcher-Python/env/bin/activate
}
activate
python3 /home/tc/Desktop/Python_Scripts/Excel_Searcher-Python/Search_fftcg.py
deactivate
bash`

## Goals
- Turn the script into an executable with a GUI.
- Use Polymorphism to shorten code length
