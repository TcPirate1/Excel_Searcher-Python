# Excel Searcher - Python

This is a Python code that I have written to search through my excel spreadsheet (that I use to keep track of my card collection).

The script runs on both Windows and Linux OS but relies on the file being in Documents/Spreadsheets/ directory. It also needs the file to be called FFTCG.xlsx but is possible to change the name if needed.

## How to use?
- Make sure the excel file is in the Documents/Spreadsheets/ directory and named FFTCG_automated.xlsx
- Download the code by cloning the repository
- Run the code in an IDE or go to the folder its in and open the command-prompt/terminal inside the folder
- When the command-prompt/terminal opens it should already be in the folder's directory
- Create a Python virtual environment with `py -m venv <file name>` or `python3 -m venv <file name>`

Open the .bat or .sh file to run the [script](#.bat-and-.sh-file-contents).

## Video examples (for Windows)



https://user-images.githubusercontent.com/88120195/211103204-5766d9cc-09d3-4c73-96e5-893844f4c894.mp4


## .bat and .sh file contents
.bat:

`CALL C:\Users\teren\Desktop\Python_stuff\Programs\Excel_Searcher\env\Scripts\activate.bat
@py.exe C:\Users\teren\Desktop\Python_stuff\Programs\Excel_Searcher\Search_fftcg.py %*
@pause`

.sh:

## Goals
- Record videos for it running on Linux and upload it to this readme.
