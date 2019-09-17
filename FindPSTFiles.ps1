$Search = $env:USERPROFILE + "/*.pst"
$OutFile = $env:USERPROFILE + "/FindPST.csv"

if (Get-ChildItem $OutFile -ErrorAction SilentlyContinue){
    Remove-Item $OutFile
}


$Result = Get-ChildItem -Recurse -Path $Search -ErrorAction SilentlyContinue

if ( $Result ){
    $Result | ForEach-Object { Add-Content -Path  $OutFile -Value $_ }
    Return "Non-Compliant"
}
else{
    Return "Compliant"
}
