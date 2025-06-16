# makeFlameReport
This program is designed to make a report from the eLabNext download format, and put into the requested  format for the FLAME study. 

Save the program makeFlameReport.exe and the folder labeled "_internal" together. 

Using the Program:
1. Click on the makeFlameReport.exe to start the program. At any point you can close the program or stop running 
the program by pressing Ctrl+c.
2. A command prompt window will open up. Enter the file path for the .xlsx file 
with the data to clean and hit Enter.
3. If there are multiple sheets in your file, enter the Sheet Name and hit Enter
   a. If there is an issue with your file path or sheet name you will get a message and the option 
to either restart or exit the program. 
   b. If the file is open the program will close. Close the file and restart the program
4. The program will then clean and reformat the data. It expects the columns "Name", "Sample type", 
"Location path", "Location", and "Date". If any of these columns are not in your data, it will prompt you to either
input the correct name for that column, or to restart the program. 
5. The program will save the data as <filename>_clean_<date>.xlsx in the same folder that the original data was 
saved. You will be prompted to either exit or restart the program (in case you wish to run another file). 
   a. if there is another file with the name <filename>_clean_<date>.xlsx already saved in that folder, the program
will save this version as <filename>_clean_<date>_<time>.xlsx
6. If at any point the program closes, there was an unexpected error while running the file. If this happens 
consistently follow the steps below and contact Mia DeCataldo (MID80@pitt.edu) for support.  


If the Program is not working as expected: 

1. If the program is not opening
	- check that the makeFlameReport.exe and the _internal folder are saved in the same place location. Check that 
the _internal folder is not empty. 
2. If the program is quitting unexpectedly
	- to get information on why the program is quitting unexpectedly you will have to run it from the command prompt maunally
using these steps: 
	1. open the folder makeFlameReport.exe is saved to. Open a command prompt window in this location (Type "cmd" then Enter in 
the Address bar). 
	2. Enter dir to get display the directory of files in this folder. 
	3. Enter .\makeFlameReport.exe to begin the program 
	4. Follow the prompts as normal until an error message shows
	5. Copy all of the dialouge of this command prompt session into a document and share when contacting Mia for help 
with the error. 
3. If the program cannot find your files
	- check that your file path is correct. You can get your file path by clicking the Address bar in the folder the files are
saved to. Make sure to append your excel file name to this path. 
	- This version of the program only takes files in .xlsx format. Convert other document types to .xlsx for this program to run 
