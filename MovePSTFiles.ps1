
#File and Folder information
$InFile = "/Users/mahellma/FindPST.csv"
$OutDir = "/Users/mahellma/temp"
$Files = Get-Content $InFile

#Messages for the user
$msgBoxOutlookRunning = [System.Windows.Messagebox]::Show(
    'Outlook.exe is running and needs to be closed to move your email Archives.
    Please close Outlook and click Ok',
    'Outlook.exe is running',
    'YesNo',
    'Error'
)

function MovePST {
    param (
        $PSTFile,
        $Target
    )
    Move-Item -Path $PSTFile -Destination $Target
    
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

function ThrowMessage {
    param (
        $FileCount = $Files.Count
    )
    $msgBoxInput1 = [System.Windows.Messagebox]::Show(
        'Email archive file(s) have been found in your user folder, theese files
        need to be moved for OneDrive for Business to be able to sync files 
        successfully. You currently have, $FileCount PST file(s) in your folders. 
        Do you want to move theese files now?',
        'PST files found in your personal folders',
        'YesNo',
        'Error'
        )
}

function OutlookRunning {
    While (Get-Process "Outlook.exe"){
        $msgBoxOutlookRunning
    }
    
}

while ($Files){
    $ms

    
}