#Main Function
function main {
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    
    #File and Folder information
    $InFile             = $env:USERPROFILE + "\FindPST.csv"
    $OutFile            = $env:USERPROFILE + "\Desktop\PstFilesMovedTo.txt"

    #Import the list of PST files to move
    $Files              = Get-Content $InFile
    $FileCount          = $Files.Count
    #Default PowerShell version in Windows 7 does not handle .Count like the newer ones and will show an empty filecount
    #if there is only 1 file to move.
    #Since this script is triggered by the non-compliance we know we have a minimum of one in the array, so to cut corners
    #if FileCount is empty at this point, let's set it to one.
    if (!$FileCount) {
        $FileCount = 1
    }
    $MoveNowMessage     =   "We are preparing your workstation for Windows 10 upgrade. There are " + $FileCount +
                            " e-mail archiving PST-file(s) on your device that we need to move to a different location." +
                            " Moving these files will take only few seconds and has no impact on your work. Would you " +
                            "like to continue moving the file(s) now?"
    $MoveNowWindowname  =   "PST files found in your personal folders"
    $MoveNowButton      =   "YesNo"
    $MoveNowType        =   "Error"

    $msgMoveNow = [System.Windows.Messagebox]::Show( $MoveNowMessage, $MoveNowWindowname,$MoveNowButton, $MoveNowType)
        switch ($msgMoveNow) {
            Yes { 
                $Answer = $True
                #Check if Outlook is running and if it's ok to close it.
                $Answer = CheckOutlookRunning
            }
            Default {
                $Answer = $False
            }
        }

    if($Answer){

        # Start moving the found files
        $SelectedTarget = SelectTarget
        #Start moving the PST Files
        foreach ($File in $Files){
            
            #Move File to selected target
            MovePST $File $SelectedTarget
            #Store the places in a file so the user and support can find them in the future
            Add-Content -Path  $OutFile -Value $SelectedTarget
        }
        $DoneMovingMessage     =    "You have now successfully moved the file(s) and the OneDrive for Business sync can" +
                                    " begin. You can find a text file on your computer’s desktop which has the selected" +
                                    " location for the e-mail archiving files for future reference."
        $DoneMovingWindowname  =    "Where are my PST files? "
        $DoneMovingButton      =    "Ok"
        $DoneMovingType        =    "Information"
        [System.Windows.Messagebox]::Show( $DoneMovingMessage, $DoneMovingWindowname, $DoneMovingButton, $DoneMovingType  )

    }
    #Perform Cleanup before exiting
    Remove-Variable InFile -ErrorAction SilentlyContinue
    Remove-Variable OutFile -ErrorAction SilentlyContinue
    Remove-Variable Files -ErrorAction SilentlyContinue
    Remove-Variable FileCount -ErrorAction SilentlyContinue
    Remove-Variable MoveNowMessage -ErrorAction SilentlyContinue
    Remove-Variable MoveNowWindowname -ErrorAction SilentlyContinue
    Remove-Variable MoveNowButton -ErrorAction SilentlyContinue
    Remove-Variable MoveNowType -ErrorAction SilentlyContinue
    Remove-Variable DoneMovingMessage -ErrorAction SilentlyContinue
    Remove-Variable DoneMovingWindowname -ErrorAction SilentlyContinue
    Remove-Variable DoneMovingButton -ErrorAction SilentlyContinue
    Remove-Variable DoneMovingType -ErrorAction SilentlyContinue

    
    #Done
    Return 0
}

function MovePST () {
    param(
        $PSTFile,
        $Target
    )

    Move-Item -Path $PSTFile -Destination $Target -ErrorAction SilentlyContinue
}

function SelectTarget {
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Please select a folder. Ex. C:\Temp, all files will be moved here."
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
    
}

function CheckOutlookRunning {
    $OutlookRunningMessage  =   "Your Outlook is open and it needs to be closed before we continue. Please save" +
                                " all your on-going work. By selecting YES, your Outlook will be automatically closed." +
                                " Select NO if you would like to continue later."
    # get Outlook process
    $outlook = Get-Process Outlook -ErrorAction SilentlyContinue

    if ($outlook) {
        $msgCloseOutlook = [System.Windows.Messagebox]::Show(
            $OutlookRunningMessage,
            'Outlook needs to be closed',
            'YesNo',
            'Error'
        )
        switch ($msgCloseOutlook) {
            Yes { 
                $Answer = $True 

                # Try gracefully first
                $outlook.CloseMainWindow()
                # Kill after fifteen seconds
                Start-Sleep 15
                if (!$outlook.HasExited) {
                $outlook | Stop-Process -Force
                }
            }
            Default {
                $Answer = $False
            }
        }
    }
    Return $Answer
}

main