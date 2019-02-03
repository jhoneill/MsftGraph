Param (
    $excelPath = '.\Planner-Export.xlsx',
    $TeamName  = 'Consultants',
    $PlanName  = 'PlanName'
)

ipmo C:\Users\mcp\Documents\WindowsPowerShell\MSGraphAPI\msgraph.psd1
ipmo C:\Users\mcp\documents\github\importexcel\ImportExcel.psd1       -Force

$myteam       = Get-GraphUser -Teams -Name $TeamName                         # assumes user is 'me'
$teamplanner  = Get-GraphTeam $myteam -Plans | where title -eq 'Team Planner'  # my team's planner named 'team planner'
$teamMembers  = Get-GraphTeam $myteam -Members
$teamMembers  | foreach-object -begin {$MemberHash = @{}} -Process {$memberhash[$_.id] =$_.mail}

#region export Plan buckets and team members to "Values" workseet in the workbook.
$excelPackage = Get-GraphPlan -Plan $teamplanner -buckets |
                    Select-Object @{n="BucketName"; e={$_.name}},PlanTitle,ID |
                         Export-excel -path $excelPath -worksheetname Values -ClearSheet -BoldTopRow -AutoSize -PassThru

$excelPackage = $teamMembers | Select @{n='User';e={$_.displayName}},Jobtitle,mail,ID |
        Export-Excel -ExcelPackage $excelPackage  -worksheetname Values -StartColumn 12 -BoldTopRow -AutoSize -PassThru
#Hide IDs: we can spot new team members if they don't have an ID. and  if a bucket is renamed in the spreadsheet, we can update it if we have the ID
Set-Excelrange -Range $excelPackage.Workbook.Worksheets['Values'].Column(15) -Hidden
Set-Excelrange -Range $excelPackage.Workbook.Worksheets['Values'].Column(3) -Hidden
#endregion

#region export the plan to a "Plan" worksheet in the workbook
#Export the tasks, for intitially name the category columns 'category1' to 'category6'
$excelPackage = Get-GraphPlan -Plan $teamplanner -FullTasks | Sort-Object orderHint|
    Select-Object -Property @{n='Task Title' ; e={   $_.title          }},
                            @{n='Bucket'     ; e={   $_.BucketName     }},
                            @{n='Start Date' ; e={   [datetime]$_.StartDateTime  }},
                            @{n='Due Date'   ; e={   [datetime]$_.dueDatetime    }},
                            @{n='%Complate'  ; e={   $_.percentComplete }},
                            @{n='AssignTo'   ; e={  ($_.assignments.psobject.properties.name | foreach {$MemberHash[$_]} )-join "; "      }},
                            @{n="Checklist"  ; e={  ($_.checklist.psobject.Properties.value  | sort-object orderHint | select -expand title) -join "; "} },
                            @{n='Description'; e={   $_.description     }},
                            @{n="Links"      ; e={  ($_.references.psobject.Properties.name -replace "%2E","." -replace "%3A",":" -replace "%25","%") -join "; "}},
                            @{n="Category1"  ; e={if($_.appliedCategories.Category1) {'Yes'} else {$null}  } },
                            @{n="Category2"  ; e={if($_.appliedCategories.Category2) {'Yes'} else {$null}  } },
                            @{n="Category3"  ; e={if($_.appliedCategories.Category3) {'Yes'} else {$null}  } },
                            @{n="Category4"  ; e={if($_.appliedCategories.Category4) {'Yes'} else {$null}  } },
                            @{n="Category5"  ; e={if($_.appliedCategories.Category5) {'Yes'} else {$null}  } } ,
                            @{n="Category6"  ; e={if($_.appliedCategories.Category6) {'Yes'} else {$null}  } } ,  ID |
        Export-Excel -ExcelPackage $excelPackage -AutoFilter -AutoSize -FreezeTopRowFirstColumn -BoldTopRow -WorksheetName Plan -ClearSheet -Activate -PassThru

#Now give the category columns the right names by exporting catgegories to the right place in the plan sheet
$excelPackage = Get-GraphPlan  $teamplanner -Details |
    Select-Object -ExpandProperty categorydescriptions |
        Export-Excel -ExcelPackage $excelPackage -WorksheetName  Plan -noheader -StartColumn 10 -BoldTopRow -AutoSize -PassThru #Categories in columns j to o

$planSheet = $excelPackage.Workbook.Worksheets['Plan']
#Hide the IDs column: a new task won't have an ID and we can find tasks to update using the ID; format dates as short date
Set-ExcelRange -Range $plansheet.Column(16)   -Hidden
Set-ExcelRange -Range $planSheet.cells['C:D'] -Width 11 -NumberFormat 'Short Date'
Set-ExcelRange -Range $planSheet.cells -VerticalAlignment Center
#if name, description and/or checklist are too wide, make them narrower
if ($planSheet.Column(1).Width -gt 35) { 
    Set-ExcelRange -Range $planSheet.cells['A:A'] -Width 35 -WrapText #Title
}
if ($planSheet.Column(7).Width -gt 20) {
    Set-ExcelRange -Range $planSheet.cells['G:G'] -Width 20 -WrapText #Check list
}
if ($planSheet.Column(8).Width -gt 20) {
    Set-ExcelRange -Range $planSheet.cells['H:H'] -Width 20 -WrapText #Description
}
if ($planSheet.Column(9).Width -gt 35) {
    $linksRange = "i2:i" + $PlanSheet.Dimension.end.row
    Set-ExcelRange -Range $planSheet.Cells[$linksRange] -FontSize 8 
    Set-ExcelRange -Range $planSheet.cells['i:i'] -Width 35 -WrapText #Links
}
#Put a data bar on the percent complete; make sure it goes 0-100 not min value to max value
$PercentRange            = "E2:E" +  $planSheet.Dimension.End.Row
$databar                 = Add-ConditionalFormatting -WorkSheet $planSheet -Address $PercentRange -DataBarColor LightBlue -PassThru
$databar.LowValue.type   = [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType]::Num
$databar.HighValue.type  = [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType]::Num
$databar.LowValue.Value  = 0
$databar.HighValue.Value = 100
#endregion 

#Create Validation rules. Bucket Name and user must come from the values page, 6 Categories must be Yes or blank, Percentage is an integer from 0 to 100
$VParams = @{WorkSheet = $PlanSheet; ShowErrorMessage=$true; ErrorStyle='stop'; ErrorTitle='Invalid Data' }
Add-ExcelDataValidationRule @VParams -Range 'B2:B1001' -ValidationType List    -Formula 'values!$a$2:$a$1000'         -ErrorBody "You must select an item from the list.`r`nYou can add to the list on the values page" #Bucket
Add-ExcelDataValidationRule @VParams -Range 'F2:F1001' -ValidationType List    -Formula 'values!$M$2:$M$1000'         -ErrorBody 'You must select an item from the list'                # Assign to
Add-ExcelDataValidationRule @VParams -Range 'J2:o1001' -ValidationType List    -ValueSet @('yes','YES','Yes')         -ErrorBody "Select Yes or leave blank for no"                     # Categories
Add-ExcelDataValidationRule @VParams -Range 'E2:E1001' -ValidationType Integer -Operator between -Value 0 -Value2 100 -ErrorBody 'Percentage must be a whole number between 0 and 100'  # Percent complete

Close-ExcelPackage -ExcelPackage $excelPackage -Show
