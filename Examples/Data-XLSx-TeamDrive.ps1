$excelPath     = "$env:TEMP\Test.xlsx"
Remove-Item $excelPath -ea SilentlyContinue

Get-Process |  Select-Object -first 50 -Property Name, cpu, pm, handles, company |
        Export-Excel  $excelPath -WorkSheetname Processes `
                -IncludePivotTable -PivotRows Company -PivotData PM -NoTotalsInPivot -PivotDataToColumn -Activate `
                -IncludePivotChart -ChartType PieExploded3D -ShowCategory -ShowPercent  -NoLegend

$myTeam        = Get-GraphUser -Teams | Select-Object -First 1
$teamDrive     = Get-GraphTeam -Team $myTeam -Drive
$generalfolder = '/drives/' + $teamDrive.id + '/root:/general/'

$URL = Copy-ToGraphFolder -Path $excelPath -Destination $generalfolder -Verbose
Start-Process $url