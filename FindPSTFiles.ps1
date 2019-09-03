$Search = "/Users/mahellma/Documents/*.pst"
$OutFile = "/Users/mahellma/FindPST.csv"

$Result = Get-ChildItem -Recurse -Path $Search

if ( $Result -ne '' ){

    $Result | ForEach-Object { Add-Content -Path  $OutFile -Value $_ }
    Return "Non-Compliant"
}
else{
    Return "Compliant"
}
