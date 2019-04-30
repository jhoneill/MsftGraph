
Function GetWorkbook {
    (Invoke-RestMethod -Uri  "$fileuri/workbook?`$expand=worksheets" -method get -Headers @{Authorization = "Bearer $AccessToken"})
}

Function GetWorksheet {
    $worksheet = (Invoke-RestMethod -Uri  "$fileuri/workbook/worksheets/sheet1?$`expand=tables,charts,names" -method get -Headers @{Authorization = "Bearer $AccessToken"})

    $worksheet.names | Format-Table Scope, Type, Name,value,visible,Comment
    $worksheet.charts | Format-Table name, left,top,width,height

}

Function GetChart {
    (Invoke-RestMethod -Uri  "$fileuri/workbook/worksheets/sheet1/charts/chart83b3?`$expand=series,legend,title,format,datalabels,axes" -method get -Headers @{Authorization = "Bearer $AccessToken"})

}

#Function Set-Chart {}
