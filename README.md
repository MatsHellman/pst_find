PST_Find Solution
-------

Solution to find and move PST files from the user's directories in Windows 7.
If a PST file is found help the user to move the file to a location that wont be
synced to OneDrive For Business.
Solution can be used as something independedt but is written to be triggered
with Configuration Manager CI's.
Functionality and logic is still very basic.

Basic Funtcionality of the scripts:

* FindPSTFiles.ps1
    * [x] Designed to run with user credentials
    * [x] Recurse through the user folders to find any PST files
        * [x] Set the variable for user home folder on Windows 
    * [x] Store the file paths to send to be read by MovePSTFile.PS1
        * [x] Set final variable for storing the CSV file containing the users
                PST file locations

* MovePSTFiles.ps1
    * [x] Read output from FindPSTFiles
    * [x] Show user information that Outlook needs to be closed, as we wont
          check if the PST file is connected in Outlook and it might be in use.
        * [x] Close Outlook.exe process
    * [x] Display move PST file GUI with recommendations on where to store file
    * [ ] Separate the information and text to manage localization of messages

* Solution tests
    * [x] Verify functionality on Windows 7
    * [x] Test Configuration Item in ConfigMGR
    * [ ] Outlook start after PST file has been moved

Future development
-------
* MovePSTFiles.ps1
    * [ ] Read different Outlook versions REG Keys and change the mapping to the new file location
