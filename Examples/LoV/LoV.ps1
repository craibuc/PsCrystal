Import-Module PsCrystal -Force

$lovPath = Convert-Path ".\LoV.txt"
$rptPath = Convert-Path ".\Lov.rpt"

Open-Report $rptPath | Import-Lov -path $lovPath -paramName "stringParam" -verbose | Save-Report -path ($rptPath -replace ".rpt", ("." + (Get-Date -f "yyyyMMdd-HHmmss") + ".rpt")) -verbose | Close-Report -verbose