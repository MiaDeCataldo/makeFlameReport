# makeFlameReport
This program is designed to make a report from the Marsland Lab eLabNext download format, and put into the requested report format for the FLAME study. 

## Set-up
1. Running this program requires Python. Check that Python is installed on your computer.
   - From Terminal (Mac) or Command Prompt (Windows) enter ```python --version```. If python is not installed download from https://www.python.org/downloads/
3. Download the source code. Save to a convenient location for you.

If program was sent to you as an executable, you will not need to download the **.py** file. Follow the "Using the program" steps below. Make sure the ***makeFlameReport.exe*** and the ***_internal*** file stay saved together.

## Using the Program: 
1. Start the program by double clicking the ***makeFlameReport.py*** file. Alternatively you can start the program via CommandPrompt using ```<filePath>/makeFlameReport.py```
2. When prompted, Enter the file path for the .xlsx file with the data to clean and hit Enter.
3. If there are multiple sheets in your file, enter the Sheet Name and hit Enter (if there is only 1 sheet in your file, you can just hit enter to continue) 
   a. If there is an issue with your file path or sheet name you will get a message and the option to either restart or exit the program. 
   b. If the file is open the program will close or throw an error. Close the file and restart the program.
4. The program will then clean and reformat the data. It expects the columns "Name", "Sample type", "Location path", "Location", and "Date". If any of these columns are not in your data, it will prompt you to either input the correct name for that column, or to restart the program. 
5. The program will save the data as <filename>_clean_<date>.xlsx in the same folder as the original data. You will be prompted to either exit or restart the program. 
   a. if there is another file with the name <filename>_clean_<date>.xlsx already saved in that folder, the program will save this version as <filename>_clean_<date>_<time>.xlsx
6. If at any point the program closes unexpectedly, there was likely an error while running the file. If this happens follow the steps below and contact Mia DeCataldo (MID80@pitt.edu) with the error message.
   - Copy the dialouge from this command prompt session into a document and share when contacting Mia for help.   

## If the Program is not working as expected: 
For any issues not resolved with the steps below, contact Mia DeCataldo (MID80@pitt.edu) for support. 
1. (For executable version only) If the program is not opening 
   - check that the makeFlameReport.exe and the _internal folder are saved in the same place location. Check that the _internal folder is not empty. If the _internal folder is empty, contact Mia or re-download the folder.
2. If the program is quitting unexpectedly
   - to get information on why the program is quitting unexpectedly you will have tostart the program via CommandPrompt using ```<filePath>/makeFlameReport.py```. If you are using the .exe version sent to you, enter ```<filePath>/makeFlameReport.exe```: 
	1. Follow the prompts as normal until an error message shows
	2. Copy all of the dialouge of this command prompt session into a document and share when contacting Mia for help with the error. 
3. If the program cannot find your files
   - check that your file path is correct. You can get your file path by clicking the Address bar in the folder the files are saved to. Make sure to append your excel file name to this path. 
   - This version of the program only takes files in .xlsx format. Convert other file types to .xlsx for this program to run.
