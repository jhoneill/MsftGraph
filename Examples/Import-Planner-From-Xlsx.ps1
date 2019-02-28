#requires -modules msftGraph, importExcel

param (
    #File to import from
    $excelPath = '.\planner-Import.xlsx',
    #The team which owns the planner. The signed in user must be a member of the team. Being an owner but not a member will fail
    $TeamName  =  'Consultants'  ,
    #The name of the plan to import to
    $PlanName  =  'Team Planner'
)

Write-Progress -Activity 'Importing plan' -Status 'Getting information about the plan, its team, and the team members'
$teamplanner         = Get-GraphTeam -ByName $TeamName -Plans | Where-Object title -eq $PlanName

#region ensure team members in the sheet are really in the team
#Get the members of the team and create two hash tables, one to get Mail from ID and one to get ID from mail
$existingteamMembers = Get-GraphTeam $myteam -Members | Where-Object {$_.mail}
$existingteamMembers | ForEach-Object -Begin {$memberMailHash = @{}; $memberIDHash = @{} } -Process {
                                              $memberMailHash[$_.mail] = $_.id
                                              if ($_.id) {$memberIDHash[$_.id]  =  $_.mail  }
                       }

$importedTeamMembers = Import-Excel -Path $excelPath -WorksheetName values -StartColumn 12
#If any team members have no ID, and mail is not in the hash of existing users ...
# ...Look them up (assume for this demo mail = upn) and add them to the team.
$importedTeamMembers.Where({$mail -and -not $_.id -and -not $memberMailHash[$_.Mail]})  | ForEach-Object {
    Write-Progress -Activity 'Importing plan' -Status "Processing new team member '$($_.mail)'"
    $user = $null
    $user = Get-GraphUser -UserID $_.mail -ErrorAction SilentlyContinue
    if ($user) {
        $_.id = $user.id
        Add-GraphGroupMember -Group $myteam -Member $user
    }
    else {Write-Warning "($_.mail) Doesn't Seem to be a valid user"}
}
#endregion

#region ensure the plan's 6 category labels match the ones in the sheet
#6 category lables are at I1:N1 in the Plan sheet; Import with no header so they will be P1..P6 as properties on an object. Make that into 6 objects with a name and value
Write-Progress -Activity 'Importing plan' -Status 'Checking categories'
$importedCategories  = (Import-Excel -path $excelPath -WorksheetName 'Plan' -NoHeader -StartColumn 10 -EndRow 1 -EndColumn 14).psobject.Properties | Sort-Object name
#Transform categories returned by the server into hash table of p1..P6 --> name;  to compare with the imported ones
$existingCategories  = Get-GraphPlan $teamplanner -Details | Select-Object -ExpandProperty categorydescriptions
$existingCategories.psobject.Properties | ForEach-Object -Begin {$catHash = @{} } -Process {$catHash[($_.name -replace 'category','p')] = $_.value}

#for our 6 imported categories check them against the corresponding entry in catHash; if different pop in a hash table that can be splatted into Set-GraphplanDetails
$importedCategories.where({$catHash[($_.name)] -ne $_.value}) |
    ForEach-Object -Begin {$newCategories= @{} } -Process  { $newCategories[($_.name -replace 'p','Category')] = $_.value}
if ($newCategories.Count -gt 0)  {
    Write-Progress -Activity 'Importing plan' -Status 'Updating categories'
    Set-GraphPlanDetails $teamplanner @newCategories
}
#endregion

#region ensure buckets in the the sheet are in the plan
#Get buckets from the server and make a hash of name--> ID , and import the bucket list from the Values sheet
Write-Progress -Activity 'Importing plan' -Status 'Checking Buckets'
$existingBuckets     = Get-GraphPlan $teamplanner -buckets
$existingBuckets     | foreach-object -Begin {$bucketHash = @{}} -Process {$bucketHash[$_.Name] = $_.id}
$importedBuckets     = Import-Excel -path $excelPath   -WorksheetName values -StartColumn 1 -EndColumn 3

#NewBuckets here is new or changed buckets. We don't cope with bucket A being renamed, and then bucket B being changed to "A"
$newbuckets          = $importedBuckets.where({-not $bucketHash[$_.bucketName]})

#Buckets with an ID but no match in the hash table must have been reanmed, and those without an ID are new...
foreach ($bucket in $newbuckets.where({$_.id}) ) {
    Write-Progress -Activity 'Importing plan' -Status 'Renaming bucket to ' -CurrentOperation $bucket.bucketName
    Rename-GraphPlanBucket -Bucket $bucket.id -NewName $bucket.bucketName
    $bucketHash[$bucket.bucketName] = $bucket.id
}
foreach ($bucket in $newbuckets.where({-not $_.id}) ) {
    Write-Progress -Activity 'Importing plan' -Status 'Adding new bucket ' -CurrentOperation $bucket.bucketName
    $newbucket = New-GraphPlanBucket -Plan $teamPlanner -Name $bucket.bucketName
    $bucketHash[$newbucket.Name] = $newbucket.id
}
#endregion

Write-Progress -Activity 'Importing plan' -Status 'Checking tasks'
#region get existing tasks - fiddle the results so they look like what we export and we are going to import next.
$existingTasks       = Get-GraphPlan $teamplanner -FullTasks |
    Sort-Object -Property ID|
        Select-Object -Property @{n='Title'          ; e={   $_.title          }},
                                @{n='Bucket'         ; e={   $_.BucketName     }},
                                @{n='StartDate'      ; e={   [datetime]$_.StartDateTime  }},
                                @{n='DueDate'        ; e={   [datetime]$_.dueDatetime    }},
                                @{n='PercentComplete'; e={   $_.percentComplete }},
                                @{n='AssignTo'       ; e={  ($_.assignments.psobject.properties.name) -join "; "      }},
                                @{n="Checklist"      ; e={  ($_.checklist.psobject.Properties.value  | sort-object orderHint | Select-Object -expand title) -join "; "} },
                                @{n='Description'    ; e={   $_.description     }},
                                @{n="Links"          ; e={  ($_.references.psobject.Properties.name -replace "%2E","." -replace "%3A",":" -replace "%25","%") -join "; "}},
                                @{n="Category1"      ; e={if($_.appliedCategories.Category1) {'Yes'} else {$null}  } },
                                @{n="Category2"      ; e={if($_.appliedCategories.Category2) {'Yes'} else {$null}  } },
                                @{n="Category3"      ; e={if($_.appliedCategories.Category3) {'Yes'} else {$null}  } },
                                @{n="Category4"      ; e={if($_.appliedCategories.Category4) {'Yes'} else {$null}  } },
                                @{n="Category5"      ; e={if($_.appliedCategories.Category5) {'Yes'} else {$null}  } } ,
                                @{n="Category6"      ; e={if($_.appliedCategories.Category6) {'Yes'} else {$null}  } } ,
                                @{n="Task"           ; e={$_.id } }

Write-Progress -Activity 'Importing plan' -Completed # it isn't from the progress message point of view it is.
$existingTasks | foreach-object -Begin {$taskHash = @{}} -Process {$taskHash[$_.task] = $true }
#endregion
#
Import the tasks from the sheet. The Category names will be customized so use our own header names, so make them usable,  And pull other columns to match add-GraphPlanTask , Set-GraphPlanTask commands
$importedTasks = Import-Excel -Path $excelPath -WorksheetName 'Plan' -HeaderName Title, Bucket, StartDate, DueDate, PercentComplete, AssigneeMail, Checklist, Description, Links, Category1, Category2, Category3, Category4, Category5, Category6,Task  |
                    ForEach-Object {
                        if ($_.AssigneeMail) {Add-Member -InputObject $_ -MemberType NoteProperty -Name AssignTo -Value $memberMailHash[$_.AssigneeMail]}
                        else                 {Add-Member -InputObject $_ -MemberType NoteProperty -Name AssignTo -Value ""}
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name CategoryNumbers -value $(foreach ($n in (1..6)) {if ($_."Category$n" -eq 'Yes') {$n} }) -PassThru
                    } | Sort-Object -Property Task

#If a task has no ID it is new, so add it. Check it has a title so if blank rows were imported, we don't try to process them
$importedTasks.Where({$_.title -and -not $_.task}) | Add-GraphPlanTask -Plan $teamplanner

#$existingTasks and Imported tasks can be compared - so compare them, show the results in Gridview to allow the user to see any which should not be modified
$propsToCompare= @('Task', 'Title', 'Bucket', 'StartDate', 'DueDate', 'PercentComplete', 'assignto', 'Checklist', 'Description', 'Links', 'Category1', 'Category2', 'Category3', 'Category4', 'Category5', 'Category6')
$comparison = Compare-Object  $existingTasks $importedTasks.Where({$_.Task -and $taskhash[$_.Task]}) -Property $propsToCompare | Sort-Object Task,sideindicator
$comparison | Select-Object  @{n='Side';e={if ($_.sideindicator -eq "<=") {'Existing'} else {'Import'}}}, 'Title', 'Bucket', 'StartDate', 'DueDate',
          'PercentComplete', @{n='Asignee';e={$memberIDHash[$_.assignto]}}, 'Checklist', 'Description', 'Links',
          'Category1', 'Category2', 'Category3', 'Category4', 'Category5', 'Category6' | Out-GridView -Title 'Review the changes - import may overwrite newer data on the server'

$Comparison.Where({$_.sideindicator -eq '=>'})  |  Set-GraphPlanTask -Confirm

