using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models

function Get-GraphPlan           {
    <#
      .Synopsis
        Gets information about plans used in the Planner app.
      .Example
        >Get-GraphTeam -Plans | where title -eq "team planner" | get-graphplan -FullTasks
        Gets the Plan(s) for the current user's team(s), and isolates those with the name "Team Planner" ;
        for each of these plans gets the tasks, expanding the name, bucket name, and assignee names
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param   (
        #The ID of the plan or a plan object with an ID property. if omitted the current users planner will be assumed.
        [Parameter( ValueFromPipeline=$true,Position=0)]
        $Plan,
        #If Specified returns only the details of the plan
        [Parameter(Mandatory=$true, ParameterSetName="Details")]
        [switch]$Details,
        #If specified returns a list of plan tasks.
        [Parameter(Mandatory=$true, ParameterSetName="Tasks")]
        [switch]$Tasks,
        #If specified gets a list of plan buckets which tasks can be assigned to
        [Parameter(Mandatory=$true, ParameterSetName="Buckets")]
        [switch]$Buckets,
        #If specified fills in the plan name, Assignee Name(s) and bucket name for each task.
        [Parameter(Mandatory=$true, ParameterSetName="FullTask")]
        [switch]$FullTasks
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if (-not $Plan)         {$Plan = Invoke-GraphRequest -Uri "$GraphUri/me/planner/plans" -ValueOnly -AsType ([MicrosoftGraphPlannerPlan]) -ExcludeProperty '@odata.etag' | Select-Object -First 1 }
        if ($Plan.title)        {$planTitle = $Plan.title}
        if ($Plan.id)           {$Plan      = $Plan.id}
        if ($Plan -is [string]) {$Uri       = "$GraphUri/planner/plans/$Plan" }
        else                    {
            Write-Warning "Could not get a plan ID from the information provided"
        }
        if     ($Tasks -or
                $FullTasks)     {
            #we need @odata.etag for changing items, but it isn't in the object definition ... grrr.
            $response = Invoke-GraphRequest  -Uri "$uri/Tasks" -ValueOnly | Sort-Object -Property orderHint
            $result   = foreach ($r in $response) {
                $etag = $r.'@odata.etag'
                [void]$r.remove( '@odata.etag') ;
                $taskobj = New-Object -Property $r -TypeName MicrosoftGraphPlannerTask
                if ($planTitle) { Add-Member -InputObject $taskobj -NotePropertyName  PlanTitle -NotePropertyValue $planTitle}

                Add-Member -InputObject $taskobj -NotePropertyName  etag -NotePropertyValue $etag -PassThru
            }
            if ($FullTasks) {$result | Expand-GraphTask }
            else            {$result}
        }
        elseif ($Details )      {
            Invoke-GraphRequest  -Uri "$uri/Details" -AsType ([MicrosoftGraphPlannerPlanDetails]) -ExcludeProperty '@odata.etag','@odata.context'
        }
        elseif ($Buckets)       {
            #we need @odata.etag for changing items, but it isn't in the object definition ... grrr.
            Invoke-GraphRequest   -Uri "$uri/Buckets" -ValueOnly | Sort-Object -Property orderHint | ForEach-Object {
                $etag = $_.'@odata.etag'
                [void]$_.remove('@odata.etag')
                $bucketobj = New-object -Property $_ -TypeName MicrosoftGraphPlannerBucket |
                    Add-Member -PassThru -NotePropertyName  etag -NotePropertyValue $etag
                if ($planTitle) {Add-Member -PassThru -InputObject $bucketobj -NotePropertyName PlanTitle -NotePropertyValue $planTitle     }
                else            {$bucketobj}
            }
        }
        else                    {
            #we need @odata.etag for changing items, but it isn't in the object definition ... grrr.
            $result    =  Invoke-GraphRequest  -Uri "$uri`?`$expand=details"
            $etag      =  $result.'@odata.etag'
            $odatakeys =  $result.Keys.Where({$_ -match "@odata\."})
            foreach ($k in $odatakeys) {[void]$result.Remove($k)}
            $planObj = New-Object  -Property $result -TypeName MicrosoftGraphPlannerPlan |
                Add-Member -PassThru -NotePropertyName  etag -NotePropertyValue $etag

            if ($planObj.owner) {
                $owner = (Invoke-GraphRequest  -Uri "$GraphUri/directoryobjects/$($planObj.owner)").displayname
                Add-Member -InputObject $planObj -NotePropertyName OwnerName -NotePropertyValue $owner
            }
            if ($planObj.createdBy.user.id -and $planObj.createdBy.user.id  -eq $planObj.owner) {
                Add-Member -InputObject $planObj -MemberType NoteProperty -Name CreatorName -Value $owner
            }
            elseif ($planObj.createdBy.user.id) {
                $creator = (Invoke-GraphRequest  -Uri "$GraphUri/directoryobjects/$($planObj.createdBy.user.id)").displayname
                Add-Member -InputObject $planObj -MemberType NoteProperty -Name CreatorName -Value $creator
            }
            return $planObj
        }
    }
}

function Set-GraphPlanDetails    {
    <#
    .Synopsis
        Sets the category labels on a Plan
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="Detail would be incorrect")]
    Param(
        #The ID of the Plan or a Plan object with an ID property.
        [Parameter(Mandatory=$true, Position=0,ValueFromPipeline=$true)]
        $Plan,
        #Label for category 1
        [AllowNull()]
        [string]
        $Category1 ,
        #Label for category 2
        [AllowNull()]
        [string]
        $Category2 ,
        #Label for category 3
        [AllowNull()]
        [string]
        $Category3 ,
        #Label for category 4
        [AllowNull()]
        [string]
        $Category4 ,
        #Label for category 5
        [AllowNull()]
        [string]
        $Category5 ,
        #Label for category 6
        [AllowNull()]
        [string]
        $Category6,
        #If specified the plan will updated without confirmation
        [switch]$Force
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if ($Plan.id) {$detailsURI = "$GraphUri/planner/plans/$($plan.id)/details" ; $planTitle = $Plan.Title}
        else          {$detailsURI = "$GraphUri/planner/plans/$plan/details"       ; $planTitle = "."   }
        try {
            $tag = (Invoke-GraphRequest   -Uri $detailsURI -ErrorAction Stop ).'@odata.etag'
        }
        catch          {throw "Failed to get tag from $detailsURI" ; return }
        if (-not $tag) {throw "Failed to get tag from $detailsURI" ; return }
        Write-Verbose -Message "SET-GRAPHPLANDETAILS Details uri is $detailsURI  will match etag of $tag"

        $CategorySettings = @{}
        foreach ($x in (1..6)) {
            if ($PSBoundParameters.ContainsKey("Category$x")) {
                $CategorySettings["category$x"] = $PSBoundParameters["category$x"]
            }
        }
        if ($CategorySettings.Count -eq 0) {throw "You need to specify a setting to change "}
        else {$Settings = @{"categoryDescriptions" = $CategorySettings} }
        $webParams = @{ Method      = "Patch"
                        URI         = $detailsURI
                        Headers     = @{"If-Match" = $tag}
                        Contenttype = "application/json"
                        body        =  ((ConvertTo-Json $settings) -replace '""','null')

        }
        Write-Debug   $webParams.body
        if ($Force -or $PSCmdlet.ShouldProcess($PlanTitle,"Update Plan Details")) {Invoke-GraphRequest @webParams }
    }
}

function Remove-GraphPlan  {
    <#
      .synopsis
        Removes a plan from a plan the
    #>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]
    param (
        #The ID of the plan or a plan object with an ID property.
        [Parameter( ValueFromPipeline=$true,Position=0)]
        $Plan,
        #If specified the plan will be removed without prompting for confirmation; by default confirmation IS requested.
        [switch]$Force
    )
    begin {
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if (-not $Plan)         {$Plan = Invoke-GraphRequest -Uri "$GraphUri/me/planner/plans" -ValueOnly -AsType ([MicrosoftGraphPlannerPlan]) -ExcludeProperty '@odata.etag' | Select-Object -First 1 }

        if ($Plan.Title )   {$target = $Plan.Title}
        if ($Plan.etag)     {$tag    = $Plan.etag}
        if ($Plan.id )      {$Plan   = $Plan.ID}
        $uri =  "$GraphUri/planner/Plans/$Plan"
        if (-not $tag)  {
            $plandetails   = Invoke-GraphRequest  -Uri $uri
            $tag           = $plandetails.'@odata.etag'
            $target        = $plandetails.title
        }
        if (-not $target)  {$target=$plan}
        if($Force -or $PSCmdlet.ShouldProcess($target,'Delete Plan')) {
            Invoke-GraphRequest -Method Delete -Uri $uri -Headers @{'If-Match' = $tag}
        }
    }
}


function Add-GraphPlanBucket     {
    <#
      .Synopsis
        Creates a task-bucket in an exsiting plan
      .Example
        > New-GraphPlanBucket -Plan $NewTeamplan -Name 'Backlog', 'To-Do','Not Doing'
        Creates 3 buckets in the same plan.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The ID of the Plan or a Plan object with an ID property.
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        $Plan,
        #The Name of the new bucket.
        [Parameter(Mandatory=$true,Position=1)]
        $Name,
        #If Specified the bucket will be added without confirmation
        [switch]$Force
    )
    begin {
        $webParams = @{ 'Method'      = "Post"
                        'URI'         = "$GraphUri/planner/buckets"
                        'Contenttype' = "application/json"

        }
        $orderHint = " !"
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if     ($Plan.id)           {$Planid = $plan.id}
        elseif ($Plan -is [String]) {$planid = $Plan}
        else   {Write-Warning 'Could not get the plan ID' ; return }
        foreach ($bucketName in $name) {
            $json      = (ConvertTo-Json ([ordered]@{"planId"=$Planid; "name"=$bucketName; "orderHint"= $orderHint}))
            Write-Debug $json
            if ($force -or $PSCmdlet.ShouldProcess($Name,"Add Bucket to plan $($Plan.title)")){
            $result    = Invoke-GraphRequest @webParams -Body $json
            $etag = $result.'@odata.etag'
            [void]$result.remove('@odata.etag')
            [void]$result.remove('@odata.context')
            $bucketobj = New-object -Property $result -TypeName MicrosoftGraphPlannerBucket |
                Add-Member -PassThru -NotePropertyName  etag -NotePropertyValue $etag
            if ($plan.Title) {Add-Member -PassThru -InputObject $bucketobj -NotePropertyName PlanTitle -NotePropertyValue $plan.Title     }
            else             {$bucketobj}
            }
        }
    }
}

function Rename-GraphPlanBucket  {
    [CmdletBinding(SupportsShouldProcess)]
    <#
      .Synopsis
        Renames a bucket in a plan
      .Example
        Get-GraphPlan $teamplanner -Buckets | where name -eq "wish list" | Rename-GraphPlanBucket -NewName "Wish-List"
        Gets a list of a buckets and finds the one named "Wish list" and reanmes is.
    #>
    Param(
        #Bucket to update either as an ID or a Bucket object with an ID
        [Parameter(ValueFromPipeline=$true,Mandatory=$true, Position=0)]
        $Bucket,
        #The new name for the Bucket.
        [Parameter(Mandatory=$true, Position=1)]
        $NewName,
        #If specified the bucket will be renamed without prompting for confirmation; this is the default unless $ConfirmPreference is set
        [Switch]$Force
    )

    if ($Bucket.id)   {$uri = "$GraphUri/planner/buckets/$($Bucket.id)"}
    else              {$uri = "$GraphUri/planner/buckets/$Bucket"  }
    if ($Bucket.etag) {$tag =  $Bucket.etag}
    else              {$tag = (Invoke-GraphRequest  -URI $uri ).'@odata.etag' }

    $body    = "{  ""name"": ""$NewName"" }"
    if ($Force -or $PSCmdlet.ShouldProcess($NewName,'Apply new name to bucket')) {
        Invoke-GraphRequest -Method Patch -URI $uri  -Headers @{'If-Match'=$tag} -Body $body -ContentType 'application/json'
    }
}

function Remove-GraphPlanBucket  {
    <#
      .synopsis
        Removes a bucket from a plan in planner
    #>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]
    param (
        #The bucket to remove
        [parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
        $Bucket,
        #If specified the bucket will be removed without prompting for confirmation; by default confirmation IS requested.
        [switch]$Force
    )
    begin {
    }
    process {
        if ($Bucket.name )  {$target = $Bucket.name}
        if ($Bucket.etag)   {$tag    = $Bucket.etag}
        if ($Bucket.id )    {$Bucket = $Bucket.ID}
        $uri =  "$GraphUri/planner/buckets/$Bucket"
        if (-not $tag)  {
            $bucketdetails = Invoke-GraphRequest  -Uri $uri
            $tag           = $bucketdetails.'@odata.etag'
            $target        = $bucketdetails.name
        }
        if (-not $target)  {$target=$Bucket}
        if($Force -or $PSCmdlet.ShouldProcess($target,'Delete Plan Bucket')) {
            Invoke-GraphRequest -Method Delete -Uri $uri -Headers @{'If-Match' = $tag}
        }

    }
}

function Get-GraphBucketTaskList {
    [CmdletBinding()]
    Param(
        #Bucket to query either as an ID or a Bucket object with an ID
        [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
        [Alias('ID')]
        $Bucket,
        #If specified IDs will be updated to their names, and extended properties (e.g. Checklist) will be added
        [Alias('FullTasks')]
        [Switch]$Expand
    )
    process {
        if ($Bucket.id) {$Bucket = $Bucket.ID}
        #we need etag for chaning items, but it isn't in the object definition ... grrr.
        $response      = Invoke-GraphRequest  -URI "$GraphUri/planner/buckets/$Bucket/tasks"
        $result        = $response.value
        while ($response.'@odata.nextLink') {
            $response  = Invoke-GraphRequest  -URI $response.'@odata.nextLink'
            $result   += $response.value
        }
        $taskObjs = foreach ($r in $result) {
            $etag      =  $r.'@odata.etag'
            [void]$r.remove( "@odata.etag") ;
            New-Object -Property $r -TypeName MicrosoftGraphPlannerTask |
                Add-Member -PassThru -NotePropertyName  etag -NotePropertyValue $etag
        }
        if ($Expand) { $taskObjs | Expand-GraphTask }
        else         { $taskobjs }
    }
}

function Add-GraphPlanTask       {
    <#
      .Synopsis
        Adds a task to an exsiting plan
      .Description
        Multiple items may be piped in, to be added to the same plan.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The ID of the Plan or a Plan object with an ID property.
        [Parameter(Mandatory=$true, Position=0)]
        $Plan,
        #The title of the new task.
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        $Title,
        #Longer description of the task
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Description,
        #User(s) to assign the task to either as a UPN name (bob@contoso.com) or ID
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $AssignTo,
        #Bucket to place the task in - it must exist already
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $Bucket,
        #Start date for the task
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Nullable[datetime]] $StartDate,
        #Date by when task should be completed
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Nullable[datetime]]$DueDate,
        #Percentage complete (note the planner app doesn't show percentages, only "Not started", "In Progress", and "Complete")
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,100)]
        [int]$PercentComplete,
        #Category tabs by number (1=Magenta, 2=Red, 3=Orange, 4=Green, 5=Teal, 6=Cyan)
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        # [ValidateRange(1,6)] #doesn't work if piped and values are null.
        [AllowNull()]
        [int[]]$CategoryNumbers,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        #A single item, or an array of items to display as a list with check boxes on the task
        [string[]]$Checklist,
        #HyperLinks (a.k.a. references): a single item, a string with items seperated with ';' an array of strings or as a hash table of URI=Label.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $Links,
        #if specified the task will be added without confirmation. (This is the default unless $confirmPreference has been changed)
        [switch]$Force,
        #By default, the task is added without returning a result. -Passthru specifies the new task should be returned.
        [Alias('PT')]
        [switch]$Passthru
    )
    begin   {
        if ($Plan.owner)  {$owner = $plan.owner}
        if ($Plan.id)     {$Plan = $Plan.id}

        try {
            Write-Progress -Activity 'Adding Task' -Status 'Getting buckets and team memmbers for this plan'
            if (-not $owner) {$owner = (Get-GraphPlan -Plan $plan).owner }
            $PlanUserHash = @{}
            Get-GraphTeam -Team $owner -Members | ForEach-Object {$PlanUserHash[$_.Mail]=$_.ID}

            $planBucketshash = @{}
            Get-GraphPlan -Buckets -Plan $Plan  | ForEach-Object {$planBucketshash[$_.Name]=$_.ID}
        }
        catch { throw "An error occured while get information about the plan" ; return }

        $webParams = @{ Method      = "Post"
                        URI         = "$GraphUri/planner/tasks"
                        Contenttype = "application/json"
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        $settings =  [ordered]@{"planId"=$Plan; "title"=$title}

        if ($Bucket) {
            if     ($Bucket.id)                              {$settings["bucketId"]=$Bucket.Id}
            elseif ($planBucketshash.ContainsValue($Bucket)) {$settings["bucketId"]=$Bucket}
            elseif ($planBucketshash[$Bucket])               {$settings["bucketId"]=$planBucketshash[$Bucket]}
            else   {throw "$Bucket is not a valid bucket name or ID"}
        }

        if ($DueDate )               {$settings["dueDateTime"]   =   $DueDate.ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ssZ")  } # 'o' for ISO date format may work here
        if ($StartDate)              {$settings["startDateTime"] = $StartDate.ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ssZ")  }

        If ($PercentComplete -ge 0) { #need to use this to catch Percent complete being 0
                                $settings["percentComplete"] = $PercentComplete
        }
        if ($AssignTo) {
            $settings["assignments"] = @{}
            ForEach ($a in $AssignTo) {
                try {
                    if ($a -match "\w+@\w+") {
                    Write-Progress -Activity 'Adding Task' -Status 'Getting system ID for user' -CurrentOperation $a
                    $a = (Invoke-GraphRequest   -Uri "$GraphUri/users/$a" -ErrorAction stop).id}
                }
                catch {throw "Couldn't resolve user $a"; return}
                $settings.assignments[$a] = @{'@odata.type'= "#microsoft.graph.plannerAssignment"; 'orderHint'= " !" }}
        }
        if ($CategoryNumbers) {
            $Settings["appliedCategories"] = @{}
            foreach ($n in $CategoryNumbers) {
               if ($n -lt 1-or $n -gt 6) {throw "$n is not a valid category - valid numbers are 1..6"; return}
               else {$settings.appliedCategories["category$n"] = $true}
            }
        }
        $json =  (ConvertTo-Json $settings)
        Write-Debug $json
        if ($Force -or $PSCmdlet.ShouldProcess($Title,"Add Task") ) {
            Write-Progress -Activity 'Adding Task' -Status 'Saving new task'
            $result  = Invoke-GraphRequest @webParams -body $Json
            if     ($Description -and $Checklist) {Set-GraphTaskDetails -PSC $PSCmdlet -Task $result -Description $Description -CheckList $Checklist }
            elseif ($Description )                {Set-GraphTaskDetails -PSC $PSCmdlet -Task $result -Description $Description  }
            elseif ($Checklist   )                {Set-GraphTaskDetails -PSC $PSCmdlet -Task $result -CheckList $Checklist }
            if     ($Links)                       {Set-GraphTaskDetails -PSC $PSCmdlet -Task $result -Links $Links }
            Write-Progress -Activity 'Adding Task' -Completed
            if ($Passthru) {
                $etag      =  $result.'@odata.etag'
                $odatakeys =  $result.Keys.Where({$_ -match "@odata\."})
                foreach ($k in $odatakeys) {[void]$result.Remove($k)}
                New-Object -Property $result -TypeName  MicrosoftGraphPlannerTask |
                        Add-Member -NotePropertyName  etag -NotePropertyValue $etag -PassThru
            }
        }
    }
}

function Get-GraphPlanTask       {
    <#
      .Synopsis
        Gets a task from a plan in planner, and optionally expands IDs to names and fetches extended properties
    #>
    [cmdletbinding()]
    param (
        #The Task to get, either an ID or a Task object with an ID property.
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)]
        $Task,
        #If specified IDs will be updated to their names, and extended properties (e.g. Checklist) will be added
        [Alias('FullTasks')]
        [Switch]$Expand
    )
    process {
        if ($Task.ID)   {$Task = $Task.ID}
        #we need odata.etag for changing items, but it isn't in the object definition ... grrr.
        $result    = Invoke-GraphRequest  -URI "$GraphUri/planner/tasks/$Task"
        $etag      =  $result.'@odata.etag'
        $odatakeys =  $result.Keys.Where({$_ -match "@odata\."})
        foreach ($k in $odatakeys) {[void]$result.Remove($k)}
        $taskobj  = New-Object -Property $result -TypeName  MicrosoftGraphPlannerTask |
                        Add-Member -NotePropertyName  etag -NotePropertyValue $etag -PassThru
        if ($Expand) { $taskobj | Expand-GraphTask}
        else         {$taskobj}
    }
}

function Set-GraphPlanTask       {
    <#
      .Synopsis
        Update an a existing task in a planner plan
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The Task to update, either an ID or a Task object with an ID property.
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true, Position=0)]
        [alias('ID')]
        $Task,
        #The new title of for task.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $Title,
        #Longer description of the task
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [string]$Description,
        #User(s) to assign the task to either as a UPN name (bob@contoso.com) or ID. They must already be part of the team.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $AssignTo,
        #Bucket to place the task in - it must exist already
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $Bucket,
        #Start date for the task
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Nullable[datetime]] $StartDate,
        #Date by when task should be completed
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Nullable[datetime]]$DueDate,
        #Percentage complete (note the planner app doesn't show percentages, only "Not started", "In Progress", and "Complete")
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateRange(0,100)]
        [int]$PercentComplete,
        #Category tabs by number (1=Magenta, 2=Red, 3=Orange, 4=Green, 5=Teal, 6=Cyan)
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        # [ValidateRange(1,6)] #doesn't work if piped and values are null.
        [AllowNull()]
        [int[]]$CategoryNumbers,
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        #If specified, any existing check-list will be removed
        [switch]$ClearList,
        #A single item, A string with items seperated with ";" or an array of items to display as a list with check boxes on the task.
        $Checklist,
        #If specified, any existing links will be removed
        [switch]$ClearLinks,
        #HyperLinks (a.k.a. references): a single item, a string with items seperated with ';' an array of strings or as a hash table of URI=Label.
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $Links,
        #Specified no confirmation will occur
        [switch]$Force,
        #If Specified returns the modified task.
        [Alias('PT')]
        [switch]$Passthru
    )
    begin   {
        $planHash = @{}
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #Did we get a task object with an ID , a title, a Plan ID and an etag ? Or and ID with the need to look up the others up
        $tag = $plan = $promptTitle = $null
        if ($Task.planID)        {$plan        = $Task.planID}
        if ($task.etag)          {$tag         = $Task.etag}
        if ($Task.title)         {$promptTitle = $Task.title}
        if ($Task.ID)            {$Task        = $Task.ID}
        if (-not ($tag -and $plan -and $promptTitle) ) {
            Write-Progress -Activity "Updating task" -Status 'Getting task information'
            try {$taskobj =   Get-GraphPlanTask -Task $Task }
            catch { throw "Could not get the task: Server response code was $($_.exception.response.statuscode.value__)" ; return }
            $plan        = $taskobj.planId
            $tag         = $taskobj.etag
            $promptTitle = $taskobj.title
        }
        #If we have not seen this Plan before get its users and buckets
        if (-not $planHash[$plan] ) {
            try {
                Write-Progress -Activity "Updating task" -Status 'Getting team members'
                $owner = (Get-GraphPlan -Plan $plan).owner
                $PlanUserHash = @{}
                Get-GraphTeam -Team $owner -Members | ForEach-Object {$PlanUserHash[$_.Mail]=$_.ID}

                Write-Progress -Activity "Updating task" -Status 'Getting plan buckets'
                $planBucketshash = @{}
                Get-GraphPlan -Buckets -Plan $Plan  | ForEach-Object {$planBucketshash[$_.Name]=$_.ID}

                $planHash[$Plan] = $true
            }
            catch { throw "An error occured while get information about the plan" ; return }
        }

        #Build up a hash table of the settings, and then convert it to JSON. Some people would rather wrangle JSON text ...
        $settings =  [ordered]@{}
        #start by adding bucket and assigned to - if they are not in the plan already, bail out.
        if ($Bucket)   {
            if     ($planBucketshash.Containsvalue($Bucket)) {$settings["bucketId"]=$Bucket}
            elseif ($planBucketshash[$Bucket])               {$settings["bucketId"]=$planBucketshash[$Bucket]}
            else   {throw ("$Bucket is not a valid bucket name or ID; Names are: '" + ($planBucketshash.Keys -join "', '") + "'" )}
        }

        if ($AssignTo) {
            $settings["assignments"] = @{}
            ForEach ($a in $AssignTo) {
                if     ($a -match "\w+@\w+")             {$assigneeID = $PlanUserHash[$a]}
                elseif ($PlanUserHash.ContainsValue($a)) {$assigneeID = $a }
                else   {throw "User $a is not a user of this plan "; return}
                $settings.assignments[$assigneeID] = @{'@odata.type'= "#microsoft.graph.plannerAssignment"; 'orderHint'= " !" }}
        }
        #Add category numbers next. If outside the range 1..6, bail out.
        if ($CategoryNumbers) {
            $Settings["appliedCategories"] = @{}
            foreach ($n in $CategoryNumbers) {
               if   ($n -lt 1-or $n -gt 6) {throw "$n is not a valid category - valid numbers are 1..6"; return}
               else {$settings.appliedCategories["category$n"] = $true}
            }
        }
        #Now everything else, dates become strings in a specific format. All the names are case sensitive BTW.
        if ($Title)                  {$settings["title"]           = $title}
        if ($DueDate )               {$settings["dueDateTime"]     = $DueDate.ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ssZ")  }
        if ($StartDate)              {$settings["startDateTime"]   = $StartDate.ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ssZ")  }
        If ($PSBoundParameters.ContainsKey('PercentComplete')) {
                                      $settings["percentComplete"] = $PercentComplete
        }

        $json =  (ConvertTo-Json $settings)
        Write-Debug $json
        $webParams = @{ URI     = "$GraphUri/planner/tasks/$Task"
                    Headers     = @{'If-Match' = $tag ; 'Prefer' = 'return=representation'  }
                    Contenttype = 'application/json'
                    body        = $json
        }
        if (($settings.count -gt 0) -and ($Force -or $PSCmdlet.ShouldProcess($promptTitle,"Update Task")) ) {
            Write-Progress -Activity "Updating task" -Status 'Updating Task'
            #by specifying a 'return' preference in the headers we get the task back, and we can use that when calling set-graphtaskDetails, and return it if asked to.
            $UpdatedTask = Invoke-GraphRequest -Method Patch @webParams
        }
        #The only warnings we get from Set-GraphTaskDetails are 'This check list item/ This link' is already there' - supress those because if we have a changed task, that's expected.
        if     ($Description -and $Checklist) {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -CheckList   $Checklist   -WarningAction SilentlyContinue -ClearList:$ClearList  -Description $Description }
        elseif ($Checklist   )                {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -CheckList   $Checklist   -WarningAction SilentlyContinue -ClearList:$ClearList  }
        elseif ($Description )                {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -Description $Description}
        if     ($Links)                       {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -Links       $Links       -WarningAction SilentlyContinue -ClearLinks:$ClearLinks}
        Write-Progress -Activity "Updating task" -Completed
        if ($Passthru) {
            $etag      =  $UpdatedTask.'@odata.etag'
            $odatakeys =  $UpdatedTask.Keys.Where({$_ -match "@odata\."})
            foreach ($k in $odatakeys) {[void]$UpdatedTask.Remove($k)}
            New-Object -Property $UpdatedTask -TypeName  MicrosoftGraphPlannerTask |
                            Add-Member -NotePropertyName  etag -NotePropertyValue $etag -PassThru
        }

    }
}

function Remove-GraphPlanTask    {
    <#
      .synopsis
        Removes a task from a plan in planner
    #>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]
    param   (
        #The task to remove, either as an ID, or as a Task object containing an ID.
        [parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
        $Task,
        #If specified the Task will be removed without prompting for confirmation; by default confirmation IS requested.
        [switch]$Force
    )
    begin   {
    }
    process {
        if ($Task.title )        {$target = $Task.title}
        if ($Task.etag)          {$tag    = $Task.etag}
        if ($Task.id )           {$Task   = $Task.ID}
        $uri =  "$GraphUri/planner/Tasks/$Task"
        if (-not $tag)  {
            $Taskdetails = Invoke-GraphRequest   -Uri $uri
            $tag           = $Taskdetails.'@odata.etag'
            $target        = $Taskdetails.title
        }
        if (-not $target)  {$target=$Task}
        if($Force -or $PSCmdlet.ShouldProcess($target,'Delete Plan Task')) {
            Invoke-GraphRequest -Method Delete -Uri $uri -Headers  @{'If-Match' = $tag}
        }
    }
}

function Expand-GraphTask        {
    <#
      .Synopsis
        Adds Assignees, buckname, plan name. Checklist, links, Preview and description fields in an existing task
      .Description
        This is not exported - it is called in Get-GraphPlan -FullTasks and Get-GraphPlanTask -Expand
    #>
    param   (
        #ID of a task or a task object contining an ID
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Task
    )
    begin   {
        $allTasks   = @()
        $planhash   = @{}
        $bucketHash = @{}
        $userHash   = @{}
    }
    process {
        $allTasks += $Task
    }
    end     {
        Write-Progress -Activity "Getting task details" -Status "Getting plan and bucket names"
        $planids      = $allTasks.planid | Sort-Object -Unique
        foreach ($p  in $planids) {
            $planhash[$p] = (Invoke-GraphRequest  -Uri "$GraphUri/planner/plans/$P" ).title
            Invoke-GraphRequest   -Uri "$GraphUri/planner/plans/$p/buckets"  -ValueOnly |
                ForEach-Object  {$bucketHash[$_.id] = $_.name}
        }
        Write-Progress -Activity "Getting task details" -Status "Getting name(s) for assignee ID(s)"
        $userIDs = $allTasks.Assignments.Keys | Sort-object -unique
        foreach ($u in $userIDs)  {
            $uData = Invoke-GraphRequest  -Uri  "$GraphUri/users/$u"
            if ($uData) {$userHash[$uData.id]=$uData.displayname}
        }
        $i = 0 #Counter for progress bar.
        Write-Progress -Activity "Getting task details" -Status "Extending Tasks" -PercentComplete 0
        foreach ($t in $allTasks) {
            $assignees = $t.assignments.keys |  foreach-object {$userhash[$_]}
            $details   = Invoke-GraphRequest  -Uri "$GraphUri/planner/tasks/$($t.id)/details"
            $expandedTask = $t | Select-Object -Property * -ExcludeProperty keys,values,additionalproperties,count   |
                Add-Member -Force -PassThru -NotePropertyName Assignees   -NotePropertyValue ($assignees -join ", ") |
            Add-Member -Force -PassThru -NotePropertyName Bucketname  -NotePropertyValue  $buckethash[$t.bucketId]   |
            Add-Member -Force -PassThru -NotePropertyName PlanTitle   -NotePropertyValue  $planhash[$t.planID]       |
            Add-Member -Force -PassThru -NotePropertyName DetailTag   -NotePropertyValue  $details.'@odata.etag'     |
            Add-Member -Force -PassThru -NotePropertyName References  -NotePropertyValue  $details.references        |
            Add-Member -Force -PassThru -NotePropertyName Checklist   -NotePropertyValue  $details.checklist         |
            Add-Member -Force -PassThru -NotePropertyName Description -NotePropertyValue  $details.description       |
            Add-Member -Force -PassThru -NotePropertyName PreviewType -NotePropertyValue  $details.previewType
            $expandedTask.pstypeNames.Add("GraphExtendedTask")
            $i += 100 #To give percentage
            Write-Progress -Activity "Getting task details" -Status "Extending Tasks" -PercentComplete ($i/$allTasks.count)
            $expandedTask
        }
        Write-Progress -Activity "Getting task details" -Completed

    }
}

function Set-GraphTaskDetails    {
    <#
      .Synopsis
        Adds Checklist, links, Preview and/or description to an existing task
      .Description
        This is not exported - it is called in Add-GraphPlanlTasks and Set-GraphPlanTask

    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="Detail would be incorrect")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification="False positives when initializing variable in begin block")]

    param (
        #ID of a task or a task object contining an ID
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Task ,
        #Task description field
        [string]$Description,
        #Preview style for the task
        [ValidateSet("automatic", "noPreview", "checklist", "description", "reference")]
        $PreviewType,
        #If specified, any existing check-list will be removed
        [switch]$ClearList,
        #A single item, A string with items seperated with ";" or an array of items to display as a list with check boxes on the task.
        $CheckList,
        #If specified, any existing links will be removed
        [switch]$ClearLinks,
        #HyperLinks (a.k.a. references): a single item, a string with items seperated with ';' an array of strings or as a hash table of URI=Label.
        $Links,
        #If specified the tasks will be updated without prompting
        [Switch]$Force,
        #used to pass state should process state from another command.
        $PSC
    )
    #See https://docs.microsoft.com/en-us/graph/api/plannertaskdetails-update?view=graph-rest-1.0

    $referencesHash = $checklistHash = $null
    if (-not $psc) {$psc = $PSCmdlet}
    if ($task.id ) {$detailsURI = "$GraphUri/planner/tasks/$($task.id)/details" ; $taskTitle =$Task.title}
    else           {$detailsURI = "$GraphUri/planner/tasks/$task/details"       ; $taskTitle = "."       }
    try   {
        if ($task.DetailTag -and -not $ClearChecks -and -not $ClearReferences) {
            $tag            = $task.DetailTag
            $existingChecks = $task.checklist.psobject.Properties.value.title
            $existingRefs   = $task.references.psobject.Properties.name
        }
        else {
            Write-Progress -Activity "Updating task" -Status 'Updating Task' -CurrentOperation 'Fetching suplementary details'
            $taskdetails    = Invoke-GraphRequest   -Uri $detailsURI
            $tag            = $taskdetails.'@odata.etag'
            if ($ClearChecks) {
                               $taskdetails.checklist.psobject.Properties.name |
                                 ForEach-Object -begin {$checklistHash=[ordered]@{} } -Process {$checklistHash[$_] = $null}
                               $existingChecks = @()
            }
            else             { $existingChecks = $taskdetails.checklist.psobject.Properties.value.title}
            if ($ClearLinks) {
                               $taskdetails.checklist.references.Properties.name |
                                 ForEach-Object -begin {$referencesHash=[ordered]@{} } -Process {$referencesHash[$_] = $null}
                               $existingRefs = @()
            }
            else             { $existingRefs = $taskdetails.references.psobject.Properties.name}
        }
    }
    catch {
        if ($_.exception.response.statuscode.value__ -eq 404) {
            Write-Warning "Retrying connection to get taskdetails"
            Start-Sleep -Seconds 5
            $taskdetails    = Invoke-GraphRequest   -Uri $detailsURI
            $tag            = $taskdetails.'@odata.etag'
            $existingChecks = $taskdetails.checklist.psobject.Properties.value.title
        }
        else {  throw "Failed to get tag from $detailsURI" ;  return}
    }
    if (-not $tag) {throw "Failed to get detail tag " ; return }
    Write-Verbose -Message "SET-GRAPHPLANDETAILS Details uri is $detailsURI  will match etag of $tag"

    #build up settings which will be converted into JSON later
    $Settings = @{}

    if ($CheckList) {
        if (-not $checklistHash) {$checklistHash=[ordered]@{} }
        #if Checklist is a single string with items split with ; split at the ; and include spaces either side of it.
        if     ($Checklist -is [string] )     {$Checklist = $Checklist -split '\s*;\s*'}
        foreach ($c in $CheckList) {
            if ($c -notin $existingChecks) {
                $guid = (New-Guid) -as [string]
                $checklistHash[$guid] = @{'@odata.type' = 'microsoft.graph.plannerChecklistItem' ;  'title'= $c;  }
            }
        }
        if (-not $PreviewType) { $settings["previewType"] = "checklist" }
    }
    if ($checklistHash.count -gt 0) {$settings["checklist"] = $checklistHash}

    #see https://docs.microsoft.com/en-us/graph/api/resources/plannerexternalreferences?view=graph-rest-1.0
    if     ($Links -is [hashtable] -or $links -is  [System.Collections.Specialized.OrderedDictionary]) {
        if (-not $referencesHash) {$referencesHash=[ordred]@{} }
        $orderhint = " !"
        foreach ($key in $links.keys) {
            $l = $links[$Key]  -replace "%","%25" -replace ":","%3A" -replace "\.","%2E"
            if ($l -notin $existingRefs ){
                $referencesHash[$l] = @{
                    '@odata.type'        = 'microsoft.graph.plannerExternalReference'
                    "previewPriority"    =  $orderhint
                    "alias"              =  $key
                }
                $orderhint = " $orderhint!"
            }
            else {Write-Warning -Message "$($Links[$key]) is already part of the task"}
        }
    }
    elseif ($Links)       {
        if ($Links -is [string]) {$Links = $Links -split "\s*;\s*"}  #Support semi-colon seperated list; remove any spaces adjacent to the semi-colon
        $referencesHash=[Ordered]@{}
        $orderhint = " !"
        foreach ($link in $Links) {
            #property names in Open Types cannot contain the following characters: ., :, % so they need to be encoded.
            $l = $link  -replace "%","%25" -replace ":","%3A" -replace "\.","%2E"
            if ($l -notin $existingRefs ){
                $referencesHash[$l] = @{
                    '@odata.type' = 'microsoft.graph.plannerExternalReference'
                    "previewPriority" =  $orderhint
                }
                $orderhint = " $orderhint!"
            }
            else {Write-Warning -Message "$link is already part of the task"}
        }
    }
    if ($referencesHash.Count -gt 0) {$settings["references"] = $referencesHash}
    if ($Description) {
        $settings["description"] = $Description
        if (-not $PreviewType) { $settings["previewType"] = "description"}
    }
    if ($PreviewType) { $settings["previewType"] = $PreviewType}

    #Now send a PATCH to the details URI with the if-match header and the settings in JSON Form
    $webParams = @{ Method      = "Patch"
                    URI         = $detailsURI
                    Headers     = @{"If-Match" = $tag}
                    Contenttype = "application/json"
                    body        = (ConvertTo-Json $settings)}
    Write-Debug $webParams.body
    if (($Settings.Count -gt 0 ) -and  ($Force -or $PSC.ShouldProcess($taskTitle,"Set details on task"))) {
        Write-Progress -Activity "Updating task" -Status 'Updating Task' -CurrentOperation 'Updating suplementary details'
        Invoke-GraphRequest @webParams | Out-Null
    }
    Write-Progress -Activity "Updating task" -Completed
}
