$InFile = "/Users/mahellma/FindPST.csv"
Get-Content $InFile | ForEach-Object { Write-Host $_ }

$msgBoxInput1 = [System.Windows.Messagebox]::Show(
    'Email archive file(s) have been found in your user folder, theese files
    need to be moved for OneDrive for Business to be able to sync files 
    successfully. Do you want to move theese files now?',
    'PST files found in your personal folders',
    'YesNo',
    'Error'
    )

switch ($msgBoxInput1) {
    yes {  

    }
    no {

    }
}