$Search = $env:USERPROFILE
$OutFile = $env:USERPROFILE + "\FindPST.csv"

if (Get-ChildItem $OutFile -ErrorAction SilentlyContinue){
    Remove-Item $OutFile
}


$Result = Get-ChildItem -Recurse -Path $Search -ErrorAction SilentlyContinue | Where-Object { $_.name -like '*.pst'}


if ( $Result ){
    $Result | ForEach-Object { Add-Content -Path  $OutFile -Value $_ }
    Return "Non-Compliant"
}
else{
    Return "Compliant"
}
