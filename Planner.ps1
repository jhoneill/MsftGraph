Function Get-GraphPlan           {
    <#
      .Synopsis
        Gets information about plans used in the Planner app.
      .Example
        >Get-GraphTeam -Plans | where title -eq "team planner" | get-graphplan -FullTasks
        Gets the Plan(s) for the current user's team(s), and isolates those with the name "Team Planner" ;
        for each of these plans gets the tasks, expanding the name, bucket name, and assignee names
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param   (
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
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Method = "Get"
                       Headers = $Script:DefaultHeader
        }
        if ($Plan.title)    {$planTitle = $Plan.title}
        if ($Plan.id)       {$Plan = $Plan.id}
        if ($Plan)          {$Uri  = "https://graph.microsoft.com/v1.0/planner/plans/$Plan" }
        else                {$Uri  = "https://graph.microsoft.com/v1.0/me/planner/"}
        Write-Verbose -Message "Geting information from $uri"
        if     ($Tasks)     {
            $results = Invoke-RestMethod @webParams -Uri "$uri/Tasks"
            $t       = $results.value | Sort-Object -Property orderHint
            foreach ($task in $t) {$task.pstypenames.add("GraphTask")}
              if ($planTitle) {
                $t = $t  | Add-Member -PassThru   -MemberType NoteProperty -Name PlanTitle -value $planTitle
            }
            return $t
        }
        elseif ($FullTasks) {
            $results = (Invoke-RestMethod @webParams -Uri "$uri/tasks"   ).value | Sort-Object -Property orderHint
            if ($planTitle) {
                $results  = $results  | Add-Member -PassThru  -MemberType NoteProperty -Name PlanTitle -value $planTitle
            }
            $results |  Expand-GraphTask
        }
        elseif ($Details -and
                $Plan)      {  Invoke-RestMethod @webParams -Uri "$uri/Details" }
        elseif ($Buckets -and
                $Plan)      {
            $results =  Invoke-RestMethod @webParams -Uri "$uri/Buckets"
            $b       = $results.value | Sort-Object -Property orderHint
            foreach ($bucket in $b) {
                $bucket.pstypenames.add("GraphBucket")
                if ($planTitle) {Add-Member -InputObject $bucket -MemberType NoteProperty -Name PlanTitle -value $planTitle     }
            }
            return $b
        }
        elseif ($Buckets  -or
                $Details)   {  Write-Warning -Message "You need to specify a Plan when using -Buckets or -Details"}
        elseif ($plan)      {
            $result =  Invoke-RestMethod @webParams -Uri "$uri`?`$expand=details"
            $result.pstypenames.add("GraphPlan")
            if ($result.owner) {
                $owner = (Invoke-RestMethod  @webparams -Uri "https://graph.microsoft.com/v1.0/directoryobjects/$($result.owner)").displayname
                Add-Member -InputObject $result -MemberType NoteProperty -Name OwnerName -Value $owner
            }
            if ($result.createdBy.user.id -and $result.createdBy.user.id  -eq $result.owner) {
                Add-Member -InputObject $result -MemberType NoteProperty -Name CreatorName -Value $owner
            }
            elseif ($result.createdBy.user.id) {
                $creator = (Invoke-RestMethod  @webparams -Uri "https://graph.microsoft.com/v1.0/directoryobjects/$($result.createdBy.user.id)").displayname
                Add-Member -InputObject $result -MemberType NoteProperty -Name CreatorName -Value $creator
            }
            return $result
        }
        else                {
            $result =  Invoke-RestMethod @webParams -Uri  $uri
            $result.pstypenames.add("GraphPlan")
            return $result
        }
    }
}

Function New-GraphTeamPlan       {
    <#
      .Synopsis
        Creates new a plan for a team.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param   (
        #The ID of the team
        [parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0)]
        $Team,
        #Name(s) of the plan(s) to add to this team.
        [parameter(Mandatory=$true, Position=1)]
        $PlanName,
        #If Specified the plan will be added without confirmation
        [Switch]$Force
    )
    begin   {
        Connect-MSGraph
    }
    process {
        if ($Team.id) {$Team = $Team.id}
        $settings =  @{owner = $team}

        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        foreach ($p in $PlanName) {
            $settings["title"] = $p
            $webParams = @{Method      = "Post"
                           Headers     = $Script:DefaultHeader
                           URI         = "https://graph.microsoft.com/beta/planner/plans"
                           Contenttype = "application/json"
                           Body        = (ConvertTo-Json $settings)
            }
            if ($Force -or  $PSCmdlet.ShouldProcess($P,"Add Team Planner")) {
                $result = Invoke-RestMethod @webParams
                $result.pstypenames.add("GraphPlan")
                Add-Member -InputObject $result -MemberType NoteProperty -Name Team -Value $Team
                return $result
            }
        }
    }
}

Function Set-GraphPlanDetails    {
    <#
    .Synopsis
        Sets the category labels on a Plan
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param(
        #The ID of the Plan or a Plan object with an ID property.
        [Parameter(Mandatory=$true, Position=0)]
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
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    if ($Plan.id) {$detailsURI = "https://graph.microsoft.com/v1.0/planner/plans/$($plan.id)/details" ; $planTitle = $Plan.Title}
    else          {$detailsURI = "https://graph.microsoft.com/v1.0/planner/plans/$plan/details"       ; $planTitle = "."   }
    try {
        $tag = (Invoke-RestMethod -Method Get -Headers $Script:DefaultHeader -Uri $detailsURI -ErrorAction Stop ).'@odata.etag'
    }
    catch          {throw "Failed to get tag from $detailsURI" ; return }
    if (-not $tag) {throw "Failed to get tag from $detailsURI" ; return }
    Write-Verbose -Message "Details uri is $detailsURI  will match etag of $tag"

    $CategorySettings = @{}
    foreach ($x in (1..6)) {
        if ($PSBoundParameters.ContainsKey("Category$x")) {
            $CategorySettings["category$x"] = $PSBoundParameters["category$x"]
        }
    }
    if ($CategorySettings.Count -eq 0) {throw "You need to specifiy "}
    else {$Settings = @{"categoryDescriptions" = $CategorySettings} }
    $webParams = @{ Method      = "Patch"
                    URI         = $detailsURI
                    Headers     = @{Authorization = $Script:AuthHeader; "If-Match" = $tag}
                    Contenttype = "application/json"
                    body        =  ((ConvertTo-Json $settings) -replace '""','null')

    }
    write-Verbose -Message  $webParams.body
    if ($Force -or $PSCmdlet.ShouldProcess($PlanTitle,"Update Plan Details")) {Invoke-RestMethod @webParams }
}

Function Add-GraphPlanBucket     {
    <#
      .Synopsis
        Adds a task-bucket to an exsiting plan
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param   (
        #The ID of the Plan or a Plan object with an ID property.
        [Parameter(Mandatory=$true,Position=0)]
        $Plan,
        #The Name of the new bucket.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Name,
        #If Specified the bucket will be added without confirmation
        [switch]$Force
    )
    process {
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if ($Plan.id) {$Plan = $plan.id}


        $webParams = @{ 'Method'      = "Post"
                        'URI'         = "https://graph.microsoft.com/v1.0/planner/buckets"
                        'Headers'     = $Script:DefaultHeader
                        'Contenttype' = "application/json"
                        'body'        = (convertto-json ([ordered]@{"planId"=$Plan; "name"=$Name ; "orderHint"= " !"}))
        }
        Write-Debug $webParams.body
        if ($force -or $PSCmdlet.ShouldProcess($Name,"Add Bucket")){
            $null = Invoke-RestMethod @webParams
        }
    }
}

Function Rename-GraphPlanBucket  {
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

    if ($Bucket.id) { $uri = "https://graph.microsoft.com/v1.0/planner/buckets/$($Bucket.id)"}
    else             {$uri = "https://graph.microsoft.com/v1.0/planner/buckets/$Bucket"       }
    if ($Bucket.'@odata.etag') {$tag = $Bucket.'@odata.etag'}
    else                       {$tag = (Invoke-RestMethod -Method Get -URI $uri -Headers $Script:DefaultHeader).'@odata.etag' }

    $headers = @{'If-Match'=$tag} + $Script:DefaultHeader
    $body    = "{  ""name"": ""$NewName"" }"
    if ($Force -or $PSCmdlet.ShouldProcess($NewName,'Apply new name to bucket')) {
        Invoke-RestMethod -Method Patch -URI $uri  -Headers $headers -Body $body -ContentType 'application/json'
    }
}

Function Remove-GraphPlanBucket  {
    <#
      .synopsis
        Removes a bucket from a plan in planner
    #>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]
    Param (
        #The bucket to remove
        [parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
        $Bucket,
        #If specified the bucket will be removed without prompting for confirmation; by default confirmation IS requested.
        [switch]$Force
    )
    begin {
        Connect-MSGraph
    }
    process {
        if ($Bucket.name )         {$target = $Bucket.name}
        if ($Bucket.'@odata.etag') {$tag    = $Bucket.'@odata.etag'}
        if ($Bucket.id )           {$Bucket = $Bucket.ID}
        $uri =  "https://graph.microsoft.com/v1.0/planner/buckets/$Bucket"
        if (-not $tag)  {
            $bucketdetails = Invoke-RestMethod -Method Get -Headers $Script:DefaultHeader -Uri $uri
            $tag           = $bucketdetails.'@odata.etag'
            $target        = $bucketdetails.name
        }
        if (-not $target)  {$target=$Bucket}
        $headers = @{'If-Match' = $tag} + $Script:DefaultHeader
        if($Force -or $PSCmdlet.ShouldProcess($target,'Delete Plan Bucket')) {
            Invoke-RestMethod -Method Delete -Uri $uri -Headers $headers
        }

    }
}

Function Get-GraphBucketTaskList {
    [CmdletBinding()]
    Param(
        #Bucket to query either as an ID or a Bucket object with an ID
        [Parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
        $Bucket,
         #If specified IDs will be updated to their names, and extended properties (e.g. Checklist) will be added
        [Switch]$Expand
    )
    if ($Bucket.id) {$Bucket = $BucketID}
    $response = Invoke-RestMethod -Method Get -URI "https://graph.microsoft.com/v1.0/planner/buckets/$Bucket/tasks" -Headers $Script:DefaultHeader
    $value    = $response.value
    while ($response.'@odata.nextLink') {
        $response = Invoke-RestMethod -Method Get -URI $response.'@odata.nextLink' -Headers $Script:DefaultHeader
        $value += $response.value
    }
    if ($Expand) {$value | Expand-GraphTask}
    else {
        foreach ($v in $value) { $v.pstypenames.add("GraphTask")}
        return $value
    }
}

Function New-GraphPlanTask       {
    <#
      .Synopsis
        Adds a task to an exsiting plan
      .Description
        Multiple items may be piped in, to be added to the same plan.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param   (
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
    Begin   {
        if ($Plan.owner)  {$owner = $plan.owner}
        if ($Plan.id)     {$Plan = $Plan.id}

        Connect-MSGraph
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
                        URI         = "https://graph.microsoft.com/v1.0/planner/tasks"
                        Headers     =  $Script:DefaultHeader
                        Contenttype = "application/json"
        }
    }
    Process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $settings =  [ordered]@{"planId"=$Plan; "title"=$title}

        if ($Bucket) {
            if     ($planBucketshash.ContainsValue($Bucket)) {$settings["bucketId"]=$Bucket}
            elseif ($planBucketshash[$Bucket])               {$settings["bucketId"]=$planBucketshash[$Bucket]}
            else   {throw "$Bucket is not a valid bucket name or ID"}
        }

        if ($AssignTo) {
            $settings["assignments"] = @{}
            ForEach ($a in $AssignTo) {
                if     ($a -match "\w+@\w+")           {$assigneeID = $PlanUserHash[$a]}
                elseif ($PlanUserHash.ContainsKey($a)) {$assigneeID = $a }
                else   {throw "User $a is not a user of this plan "; return}
                $settings.assignments[$assigneeID] = @{'@odata.type'= "#microsoft.graph.plannerAssignment"; 'orderHint'= " !" }}
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
                    Write-Progress -Activity 'Adding Task' -Status 'Getting system ID for user' -Id $a
                    $a = (Invoke-RestMethod -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/users/$a" -ErrorAction stop).id}
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
        Write-Verbose -Message $json
        if ($Force -or $PSCmdlet.ShouldProcess($Title,"Add Task") ) {
            Write-Progress -Activity 'Adding Task' -Status 'Saving new task'
            $task  = Invoke-RestMethod @webParams -body $Json
            if     ($Description -and $Checklist) {Set-GraphTaskDetails -PSC $PSCmdlet -Task $task -Description $Description -CheckList $Checklist }
            elseif ($Description )                {Set-GraphTaskDetails -PSC $PSCmdlet -Task $task -Description $Description  }
            elseif ($Checklist   )                {Set-GraphTaskDetails -PSC $PSCmdlet -Task $task -CheckList $Checklist }
            if     ($Links)                       {Set-GraphTaskDetails -PSC $PSCmdlet -Task $task -Links $Links }
            Write-Progress -Activity 'Adding Task' -Completed
            if ($Passthru) {
                $task.pstypenames.add("GraphTask")
                return $task
            }
        }
    }
}

Function Get-GraphPlanTask       {
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
        [Switch]$Expand

    )

    if ($Task.ID)   {$Task = $Task.ID}
    $response = Invoke-RestMethod -Method Get -URI "https://graph.microsoft.com/v1.0/planner/tasks/$Task" -Headers $Script:DefaultHeader
    if ($Expand) {$response | Expand-GraphTask }
    else         {
        $response.pstypenames.add("GraphTask")
        return $response
    }
}

Function Set-GraphPlanTask       {
    <#
      .Synopsis
        Update an a existing task in a planner plan
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param   (
        #The Task to update, either an ID or a Task object with an ID property.
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true, Position=0)]
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
        Connect-MSGraph
        $planHash = @{}
    }
    Process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        #Did we get a task object with an ID , a title, a Plan ID and an etag ? Or and ID with the need to look up the others up
        $tag = $plan = $promptTitle = $null
        if ($Task.planID)        {$plan        = $Task.planID}
        if ($task.'@odata.etag') {$tag         = $Task.'@odata.etag'}
        if ($Task.title)         {$promptTitle = $Task.title}
        if ($Task.ID)            {$Task        = $Task.ID}
        if (-not ($tag -and $plan -and $promptTitle) ) {
            Write-Progress -Activity "Updating task" -Status 'Getting task information'
            try {$taskobj =   Get-GraphPlanTask -Task $Task }
            catch { throw "Could not get the task: Server response code was $($_.exception.response.statuscode.value__)" ; return }
            $plan        = $taskobj.planId
            $tag         = $taskobj.'@odata.etag'
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
        Write-Verbose -Message $json
        $webParams = @{ URI     = "https://graph.microsoft.com/v1.0/planner/tasks/$Task"
                    Headers     = @{'If-Match' = $tag ; 'Prefer' = 'return=representation'  } + $Script:DefaultHeader
                    Contenttype = 'application/json'
                    body        = $json
        }
        if (($settings.count -gt 0) -and ($Force -or $PSCmdlet.ShouldProcess($promptTitle,"Update Task")) ) {
            Write-Progress -Activity "Updating task" -Status 'Updating Task'
            #by specifying a 'return' preference in the headers we get the task back, and we can use that when calling set-graphtaskDetails, and return it if asked to.
            $UpdatedTask = Invoke-RestMethod -Method Patch @webParams
        }
        #The only warnings we get from Set-GraphTaskDetails are 'This check list item/ This link' is already there' - supress those because if we have a changed task, that's expected.
        if     ($Description -and $Checklist) {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -CheckList   $Checklist   -WarningAction SilentlyContinue -ClearList:$ClearList  -Description $Description }
        elseif ($Checklist   )                {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -CheckList   $Checklist   -WarningAction SilentlyContinue -ClearList:$ClearList  }
        elseif ($Description )                {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -Description $Description}
        if     ($Links)                       {Set-GraphTaskDetails -Task $UpdatedTask -PSC $PSCmdlet -Links       $Links       -WarningAction SilentlyContinue -ClearLinks:$ClearLinks}
        Write-Progress -Activity "Updating task" -Completed
        if ($Passthru) {
            $updatedtask.pstypenames.add('GraphTask')
            return $UpdatedTask
        }

    }
}

Function Remove-GraphPlanTask    {
    <#
      .synopsis
        Removes a task from a plan in planner
    #>
    [CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]
    Param   (
        #The task to remove, either as an ID, or as a Task object containing an ID.
        [parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
        $Task,
        #If specified the Task will be removed without prompting for confirmation; by default confirmation IS requested.
        [switch]$Force
    )
    begin   {
        Connect-MSGraph
    }
    process {
        if ($Task.title )        {$target = $Task.title}
        if ($Task.'@odata.etag') {$tag    = $Task.'@odata.etag'}
        if ($Task.id )           {$Task   = $Task.ID}
        $uri =  "https://graph.microsoft.com/v1.0/planner/Tasks/$Task"
        if (-not $tag)  {
            $Taskdetails = Invoke-RestMethod -Method Get -Headers $Script:DefaultHeader -Uri $uri
            $tag           = $Taskdetails.'@odata.etag'
            $target        = $Taskdetails.title
        }
        if (-not $target)  {$target=$Task}
        $headers = @{'If-Match' = $tag} + $Script:DefaultHeader
        if($Force -or $PSCmdlet.ShouldProcess($target,'Delete Plan Task')) {
            Invoke-RestMethod -Method Delete -Uri $uri -Headers $headers
        }

    }
}

Function Expand-GraphTask        {
    <#
      .Synopsis
        Adds Assignees, buckname, plan name. Checklist, links, Preview and description fields in an existing task
      .Description
        This is not exported - it is called in Get-GraphPlan -FullTasks and Get-GraphPlanTask -Expand
    #>
    Param   (
        #ID of a task or a task object contining an ID
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Task
    )
    begin   {
        Connect-MSGraph
        $webParams  = @{Method = "Get"
                        Headers = $Script:DefaultHeader
        }
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
            $planhash[$p] = (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/planner/plans/$P" ).title
            (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/planner/plans/$p/buckets" ).value |
                ForEach-Object  {$bucketHash[$_.id] = $_.name}
        }
        Write-Progress -Activity "Getting task details" -Status "Getting name(s) for assignee ID(s)"
        $userIDs = $allTasks | ForEach-Object {$_.assignments.psobject.Properties} | Select-Object -ExpandProperty Name | Sort-object -unique
        foreach ($u in $userIDs)  {
            $uData = Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/users/$u"
            if ($uData) {$userHash[$uData.id]=$uData.displayname}
        }
        $i = 0 #Counter for progress bar.
        Write-Progress -Activity "Getting task details" -Status "Extending Tasks" -PercentComplete 0
        foreach ($t in $allTasks) {
            $details   = Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($t.id)/details"
            $assignees = $t.assignments.psobject.Properties | Select-Object -ExpandProperty Name | foreach-object {$userhash[$_]}
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name Assignees    -Value ($assignees -join ", ")
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name Bucketname   -Value $buckethash[$t.bucketId]
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name PlanTitle    -Value $planhash[$t.planID]
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name DetailTag    -Value $details.'@odata.etag'
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name references   -Value $details.references
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name checklist    -Value $details.checklist
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name description  -Value $details.description
            Add-Member -Force -InputObject $t -MemberType NoteProperty -Name previewType  -Value $details.previewType
            $t.pstypeNames.Add("GraphExtendedTask")
            $i += 100 #To give percentage
            Write-Progress -Activity "Getting task details" -Status "Extending Tasks" -PercentComplete ($i/$allTasks.count)
        }
        Write-Progress -Activity "Getting task details" -Completed
        return $allTasks
    }
}

Function Set-GraphTaskDetails    {
    <#
      .Synopsis
        Adds Checklist, links, Preview and/or description to an existing task
      .Description
        This is not exported - it is called in Add-GraphPlanlTasks and Set-GraphPlanTask

    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
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
    if ($task.id ) {$detailsURI = "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details" ; $taskTitle =$Task.title}
    else           {$detailsURI = "https://graph.microsoft.com/v1.0/planner/tasks/$task/details"       ; $taskTitle = "."       }
    try   {
        if ($task.DetailTag -and -not $ClearChecks -and -not $ClearReferences) {
            $tag            = $task.DetailTag
            $existingChecks = $task.checklist.psobject.Properties.value.title
            $existingRefs   = $task.references.psobject.Properties.name
        }
        else {
            Write-Progress -Activity "Updating task" -Status 'Updating Task' -CurrentOperation 'Fetching suplementary details'
            $taskdetails    = Invoke-RestMethod -Method Get -Headers $Script:DefaultHeader -Uri $detailsURI
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
            $taskdetails    = Invoke-RestMethod -Method Get -Headers $Script:DefaultHeader -Uri $detailsURI
            $tag            = $taskdetails.'@odata.etag'
            $existingChecks = $taskdetails.checklist.psobject.Properties.value.title
        }
        else {  throw "Failed to get tag from $detailsURI" ;  return}
    }
    if (-not $tag) {throw "Failed to get detail tag " ; return }
    Write-Verbose -Message "Details uri is $detailsURI  will match etag of $tag"

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
                    Headers     = @{Authorization = $Script:AuthHeader; "If-Match" = $tag}
                    Contenttype = "application/json"
                    body        = (ConvertTo-Json $settings)}
    Write-Verbose -Message $webParams.body
    if (($Settings.Count -gt 0 ) -and  ($Force -or $PSC.ShouldProcess($taskTitle,"Set details on task"))) {
        Write-Progress -Activity "Updating task" -Status 'Updating Task' -CurrentOperation 'Updating suplementary details'
        Invoke-RestMethod @webParams | Out-Null
    }
    Write-Progress -Activity "Updating task" -Completed
}

Function Add-GraphPlannerTab     {
    <#
      .Synopsis
        Adds a planner tab to a team-channel for a pre-existing plan
      .Description
        This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
        which requires consent to use the Group.ReadWrite.All scope.
      .Example
        >
        >$channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
        >$plan   = Get-GraphTeam -ByName accounts  -Plans | where title -Like "year end*"
        >Add-GraphPlannerTab -Plan $plan -Channel $channel -TabLabel "Planner"
        The first line gets the 'year-end' channel for the accounts team
        The second gets a plan with tile which matches 'year end'
        and the third creates a tab labelled 'Planner' in the channel for that plan.
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        #An ID or Plan object for a plan within the team
        [Parameter(Mandatory=$true,Position=0)]
        $Plan,
        #An ID or Channel object for a channel (which may contain the team ID)
        [Parameter(Mandatory=$true,Position=1)]
        $Channel,
        #A team ID, or a team object, if not specified as part of the channel
        $Team,
        #The label for the tab.
        $TabLabel,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        $PassThru,
        #If Specified the tab will be added without confirming
        $Force
    )

    #We got a team ID use it. If the the channel had a team, use that. If we didn't get a team, throw an error.
    if       ($Team.id)      {$Team = $Team.id}
    elseif   ($Channel.Team) {$Team = $Channel.Team}
    if ( -not $Team)         {throw 'Can not determine the team from the channel; please specify it explicitly' }
    if ((-not $TabLabel) -and $Plan.Title) {
        Write-Verbose -Message "No Tab label was specified, using the Plan title '$($Plan.Title)'"
        $TabLabel = $Plan.Title
    }
    #If Plan and/or channel were objects with IDs use the ID
    if       ($Channel.id) {$Channel = $Channel.id}
    if       ($Plan.id)    {$Plan    = $Plan.id}
    $tabURI = "https://tasks.office.com/{0}/Home/PlannerFrame?page=7&planId={1}" -f $Script:TenantId , $Plan

    $webparams = @{'Method'      = 'Post';
                   'Uri'         = "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs" ;
                   'Headers'     =  $Script:DefaultHeader;
                   'ContentType' = 'application/json'
    }

    $json = ConvertTo-Json ([ordered]@{
                'name'          = $TabLabel
                'TeamsAppId'    = 'com.microsoft.teamspace.tab.planner'
                'configuration' = [ordered]@{
                   'entityId'   = $plan
                   'contentUrl' = $tabURI
                   'websiteUrl' = $tabURI
                   'removeUrl'  = $tabURI
                }
            })
    Write-Debug $json
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
        $result = Invoke-RestMethod @webParams -body $json
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'Planner'
            return $result
        }
    }
}

