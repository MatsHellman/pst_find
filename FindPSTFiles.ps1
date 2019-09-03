$Search = "/Users/mahellma/Downloads/*.pst"
$OutFile = "/Users/mahellma/FindPST.csv"

$Result = Get-ChildItem -Recurse -Path $Search
Write-Host $Result

if ( $Result ){

    $Result | ForEach-Object { Add-Content -Path  $OutFile -Value $_ }
    Return "Non-Compliant"
}
else{
    Return "Compliant"
}
