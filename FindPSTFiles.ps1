$Search = $env:USERPROFILE + "/*.pst"
$OutFile = $env:USERPROFILE + "/FindPST.csv"

if (Get-ChildItem $OutFile){
    Remove-Item $OutFile
}


$Result = Get-ChildItem -Recurse -Path $Search
Write-Host $Result

if ( $Result ){
    $Result | ForEach-Object { Add-Content -Path  $OutFile -Value $_ }
    Return "Non-Compliant"
}
else{
    Return "Compliant"
}
