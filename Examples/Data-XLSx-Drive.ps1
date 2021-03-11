$excelPath     = "$env:TEMP\Test.xlsx"
Remove-Item $excelPath -ErrorAction SilentlyContinue

Get-Process |  Select-Object -first 50 -Property Name, CPU, PM, Handles, Company |
        Export-Excel  $excelPath -WorkSheetname Processes `
                -IncludePivotTable -PivotRows Company -PivotData PM -NoTotalsInPivot -PivotDataToColumn -Activate `
                -IncludePivotChart -ChartType PieExploded3D -ShowCategory -ShowPercent  -NoLegend

$myTeam        = Get-GraphUser -Teams | Select-Object -First 1
$teamDrive     = Get-GraphTeam -Team $myTeam -Drive
$generalfolder = '/drives/' + $teamDrive.id + '/root:/general/'

$file = Copy-ToGraphFolder -Path $excelPath -Destination $generalfolder -Verbose
Start-Process $file.webUrl