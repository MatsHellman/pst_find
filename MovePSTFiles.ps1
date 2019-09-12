#Main Function
function main {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    #File and Folder information
    $InFile         = $env:USERPROFILE + "/FindPST.csv"
    $OutDir         = $env:USERPROFILE

    #Import the list of PST files to move
    $Files = Get-Content $InFile
    $FileCount = $Files.Count
    $MoveNowMessage =   "Email archive file(s) have been found in your user folder, theese files" +
                        "need to be moved for OneDrive for Business to be able to sync files " +
                        "successfully. You currently have, " + $FileCount + " PST file(s) in your folders." +
                        "Do you want to move theese files now?"

    $msgMoveNow = [System.Windows.Messagebox]::Show(
        $MoveNowMessage,
        'PST files found in your personal folders',
        'YesNo',
        'Error'
        )
    switch ($msgMoveNow) {
        Yes { 
            $Answer = $True 
        }
        Default {
            $Answer = $False
        }
    }
    $Answer = CheckOutlookRunning
    
    if($Answer){

        #If Outlook is running we need to close it so it isn't locking the PST file
        CheckOutlookRunning


        #Start moving the PST Files
        foreach ($File in $Files){
            # Start moving the found files
            #$SelectedTarget = SelectTarget
            Write-Host $SelectedTarget
            Write-Host $File
            #MovePST $File $SelectedTarget
        }
        
    }
    #Perform Cleanup before exiting
    Remove-Variable InFile
    Remove-Variable OutDir 
    Remove-Variable Files 
    
    #Done
    Return 0
}

function MovePST () {
    param(
        $PSTFile,
        $Target
    )
    Write-Host "PSTFile to be moved is:" + $PSTFile
    Write-Host "The target is:" + $Target
    Move-Item -Path $PSTFile -Destination $Target
    Write-Host "In MovePST"
    Remove-Variable PSTFile
    Remove-Variable Target
}

function SelectTarget {
    Write-Host "In SelectTarget"
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
    $outlook = Get-Process notepad -ErrorAction SilentlyContinue

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

                # try gracefully first
                $outlook.CloseMainWindow()
                # kill after five seconds
                Sleep 15
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