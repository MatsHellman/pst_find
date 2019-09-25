#Main Function
function main {
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    
    #File and Folder information
    $InFile             = $env:USERPROFILE + "\FindPST.csv"
    $OutFile            = $env:USERPROFILE + "\Desktop\PstFilesMovedTo.txt"

    #Import the list of PST files to move
    $Files              = Get-Content $InFile
    $FileCount          = $Files.Length
    $MoveNowMessage     =   "Email archive file(s) have been found in your user folder, theese files" +
                            " need to be moved for OneDrive for Business to be able to sync files " +
                            "successfully. You currently have, " + $FileCount + " PST file(s) in your personal folders." +
                            "Do you want to move the file(s) now?"
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

        #Start moving the PST Files
        foreach ($File in $Files){
            # Start moving the found files
            $SelectedTarget = SelectTarget
            
            #Move File to selected target
            MovePST $File $SelectedTarget
            #Store the places in a file so the user and support can find them in the future
            Add-Content -Path  $OutFile -Value $SelectedTarget
        }
        $DoneMovingMessage     =   "A text file has been placed on your desktop with your selected PST file location for future reference."
        $DoneMovingWindowname  =   "Where are my PST files? "
        $DoneMovingButton      =   "Ok"
        $DoneMovingType        =   "Information"
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
    $foldername.Description = "Please select a folder. Ex. C:\Temp"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
    
}

function CheckOutlookRunning {
    $OutlookRunningMessage  =   "Outlook is running and needs to be closed to continue moving the file(s)." +
                                "Select YES to continue and the script will close Outlook, press NO to exit and" +
                                " continue later."
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