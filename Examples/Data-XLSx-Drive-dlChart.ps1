<#
    .Description
        Creates a very simple Excel file with a chart, and uses the Graph API to upload it
        Then uses the Excel part of the Graph API to extract the chart as a picture, back to the local machine.
        A roundabout way of getting a chart as a picture from data in Powershell :-)
#>

Param (
    $ExcelFile   = 'C:\temp\tempchart.xlsx',
    $Destination = 'root:/Documents'

)
#Remove old files
Remove-Item   C:\temp\graph.png,  $ExcelFile -ea SilentlyContinue

#Get some trivial data - the names and lengths of PowerShell files in this module.
#Export it to Excel and chart it
$excel = Get-ChildItem -path (get-module msftgraph).ModuleBase  -Recurse -Include *.ps1 | Select-Object name,length |
    Export-Excel -PassThru -ChartType BarClustered -AutoNameRange -Path $ExcelFile
Add-ExcelChart -Worksheet $excel.Sheet1 -ChartType ColumnClustered -XRange "Name" -YRange "Length" -column 4 -SeriesHeader "file size"
Close-ExcelPackage $excel

#Upload the file to current users' one drive. Open the file in Excel web app to see it
$graphfile  = Copy-ToGraphFolder -Path $ExcelFile -Destination $Destination
Start-Process $graphfile.webUrl

#We can add /workbook to the end of the URI which refers to a one drive object, so use this to find the first chart ...
$chartsURI  = "https://graph.microsoft.com/v1.0/me/drive/items/$($graphfile.id)/workbook/worksheets/sheet1/charts"
$chartName  = (Invoke-RestMethod -Uri $chartsURI  -method Get -Headers @{Authorization = "Bearer $AccessToken"} ).value |
    Select-Object -First 1 -ExpandProperty Name

#We can download the chart image - it will come down as base-64 encoded PNG, so convert it back and save it , then open the PNG
$imagechars = (Invoke-RestMethod -Uri "$chartsURI/$chartName/image" -Method Get -Headers @{Authorization = "Bearer $AccessToken"} ).value -as [char[]]
$imagebytes = [convert]::FromBase64CharArray($imagechars , 0 , $imagechars.Count)
[System.IO.File]::WriteAllBytes("C:\temp\graph.png",$imagebytes)
Start-Process C:\temp\graph.png

