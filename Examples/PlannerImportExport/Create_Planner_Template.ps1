#requires -modules Microsoft.Graph.PlusPlus, importExcel
<#
.synopsis
    Use an existing team's planner to create a .xlsx file as a template for tasks to import
#>
param (
        #The team which owns the planner. The signed in user must be a member of the team. Being an owner but not a member will fail
        $TeamName  = 'Consultants'  ,
        #The name of the plan to base the template on
        $PlanName  = 'Team Planner' ,
        #Path for the Excel file to create as the template.
        $ExcelPath = '.\Planner-Template.xlsx'
)

$myteam       = Get-GraphUser -Teams         | Where-Object -Property title -eq  $teamName # assumes user is 'me'
$teamplanner  = Get-GraphTeam $myteam -Plans | Where-Object -Property title -eq  $PlanName # my team's planner named above

#region export Plan buckets and team members to a "Values" sheet in the workbook, and insert categories as column headings in a "Plan" sheet" but well to right; then put the other column headings in to the left of them
$excelPackage = Get-GraphPlan -Plan $teamplanner -Buckets |
                    Select-Object @{n="BucketName"; e={$_.name}},PlanTitle,ID |
                         Export-Excel -Path $excelPath -WorksheetName Values -ClearSheet -BoldTopRow -AutoSize -PassThru

$excelPackage = Get-GraphTeam $myteam -Members |
                    Select-Object @{n='User';e={$_.displayName}},Jobtitle,mail,ID |
                        Export-Excel -ExcelPackage $excelPackage  -worksheetname Values -StartColumn 12 -BoldTopRow -AutoSize -PassThru

#Hide IDs: we can spot new team members if they don't have an ID. and if a bucket is renamed in the spreadsheet, we can update it if we have the ID
Set-Excelrange -Range $excelPackage.Workbook.Worksheets['Values'].Column(15) -Hidden
Set-Excelrange -Range $excelPackage.Workbook.Worksheets['Values'].Column(3) -Hidden

#Now export the catgegories - create a new worksheet named 'plan' and put them on the right in the top row.
$excelPackage = Get-GraphPlan  $teamplanner -Details |
    Select-Object -ExpandProperty categorydescriptions |
        Export-Excel -ExcelPackage $excelPackage -WorksheetName  Plan -ClearSheet -StartColumn 10 -NoHeader -BoldTopRow -FreezeTopRowFirstColumn -Activate -PassThru

#put the fixed column names in on the left of the top row in the 'plan' sheet
$planSheet = $excelPackage.Workbook.Worksheets['Plan']
$col = 1 ;
'Task Title' , 'Bucket' , 'Start Date', 'Due Date', '% Complete',  'Assign To', 'Check list', 'Description' ,'Links' | ForEach-Object {
    $Address = [OfficeOpenXml.ExcelAddress]::new(1,$col,1,$col).address
    $PlanSheet.Cells[$address].Value = $_
    $col ++
}
#endregion

#region set column widths,  number formats and data validation rules on the "plan" sheet
Set-ExcelRange -WorkSheet $PlanSheet -Range '1:1' -Bold
$PlanSheet.Cells.AutoFitColumns()
Set-ExcelRange -Range $planSheet.Cells['A:A'] -Width 35                                    #Title
Set-ExcelRange -Range $planSheet.Cells['B:B'] -Width $excelPackage.Values.Column(1).width  #Make Bucket column as wide as the bucket-name column on the values sheet
Set-ExcelRange -Range $planSheet.Cells['C:D'] -Width 11 -NumberFormat 'Short Date'         #Format Start-date and Due-date columns as dates
Set-ExcelRange -Range $planSheet.Cells['F:F'] -Width $excelPackage.Values.Column(13).width #Make Assign-To column as wide as the email-address column on the values sheet
Set-ExcelRange -Range $planSheet.Cells['G:H'] -Width 20 -WrapText                          #Check-list and Description columns
Set-ExcelRange -Range $planSheet.Cells['I:I'] -Width 35 -WrapText                          #Links - tried setting a smaller font but excel applies its own hyperlink style when you add one.
Set-ExcelRange -Range $planSheet.cells -VerticalAlignment Center
$params = @{'ShowErrorMessage'=$true; 'ErrorStyle'='stop'; 'ErrorTitle'='Invalid Data'; 'worksheet'=$planSheet }
Add-ExcelDataValidationRule @params -Range 'B2:B1001' -ValidationType List    -Formula 'values!$a$2:$a$1000'         -ErrorBody "You must select an item from the list.`r`nYou can add to the list on the values page" #Bucket
Add-ExcelDataValidationRule @params -Range 'F2:F1001' -ValidationType List    -Formula 'values!$M$2:$M$1000'         -ErrorBody 'You must select an item from the list'               # Assign to
Add-ExcelDataValidationRule @params -Range 'J2:O1001' -ValidationType List    -ValueSet @('yes','YES','Yes')         -ErrorBody "Enter Yes or leave blank for no"                     # Categories
Add-ExcelDataValidationRule @params -Range 'E2:E1001' -ValidationType Integer -Operator between -Value 0 -Value2 100 -ErrorBody 'Percentage must be a whole number between 0 and 100' # Pecent complete
#endregion
Close-ExcelPackage $excelPackage