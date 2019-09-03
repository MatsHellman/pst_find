Solution to find and move PST files from user directories.
Solution can be used as something independedt but is written to be triggered
with Configuration Manager CI's.

Funtcionality:

FindPSTFiles.ps1
    [] Designed to run with user credentials
    [] Recurse through the user folders to find any PST files.
    [] Store the file paths to send to the second PS1 file 

MovePSTFiles.ps1
    [] Read output from FindPSTFiles
    [] Show user information that Outlook needs to be closed
    [] Display move PST file GUI with recommendations
    