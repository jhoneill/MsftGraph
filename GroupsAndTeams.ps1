#see also Get-MsolGroup ; Get-MsolGroupMember ; add-MsolGroupMember; Remove-MsolGroupMember
Function Get-GraphGroupList {
    <#
      .Synopsis
        Gets a list of groups
      .Example
        >Get-GraphGroupList | format-table -autosize  Displayname, SecurityEnabled, Mailenabled, Mail, ID
        Displays a table of groups in the current tennant
      .Example
        >(Get-GraphGroupList -Name consult | Get-GraphTeam -Site).weburl
        Gets any group whose name begins "Consult" , finds its sharepoint site, and returns the site's URL
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param(
        #if specified limits the groups returned to those with names begining...
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
        [string]$Name,
        #Field(s) to select
        [ValidateSet("classification", "createdDateTime", "description", "displayName", "groupTypes",
                    "id", "mail", "mailEnabled", "mailNickname",      "onPremisesLastSyncDateTime",
                    "onPremisesProvisioningErrors", "onPremisesSecurityIdentifier", "onPremisesSyncEnabled",
                    "preferredDataLocation",  "proxyAddresses", "renewedDateTime", "securityEnabled", "visibility")]
        [string[]]$Select,
        #An oData order by string
        [string]$OrderBy,
        #An oData filter string; there is a graph limitation  that you can't filter by description or Visibility.
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )
    Connect-MSGraph
    #https://docs.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0
    $webParams   = @{Method = "Get"
                     Headers = $Script:DefaultHeader
    }
    $uri         = 'https://graph.microsoft.com/v1.0/Groups/'
    $JoinChar    = "?"
    if ($Select) {
        $uri     = $uri + '?$select=' + ($Select -join ',')
        $JoinChar= "&"
    }
    if ($Name)   {
        if ($Name -match '\*') {$Name = $Name -replace "\*",""}
        #The StartsWith oData function is case insensitive so we don't need to fix case.
        $uri     = $uri + $JoinChar + ("`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $Name )
        $JoinChar= "&"
    }
    if ($OrderBy){
        $uri     = $uri + $JoinChar + '$OrderBy=' + $OrderBy
        $JoinChar= "&"
    }
    if ($Filter) {
      $uri       = $uri + $JoinChar + '$Filter='  +$Filter
      $JoinChar  = "&"
    }
    Write-progress -Activity "Finding Groups"
    $groups = (Invoke-RestMethod @webParams -Uri $uri ).value
    foreach ($g in $groups) {$g.pstypenames.Add("GraphGroup") }
    Write-progress -Activity "Finding Groups" -Completed
    #return groups as they are if they have been sorted server-side, otherwise sort them here.
    if ($OrderBy) {return $groups}
    else          {       $groups | Sort-Object -Property DisplayName}
#(Invoke-RestMethod -Method get -Headers @{Authorization = "Bearer $script:AccessToken"} -Uri 'https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+''Unified'')').VALUE
}

Function New-GraphGroup {
    <#
      .Synopsis
        Adds a new group/team
      .Description
        Every team is also a group, but not every group is team enabled.
        This Command has an alias of New-GraphTeam so you call it as team or group
        By default it creates the group as a team UNLESS you specify -NoTeam.
        A non-Teams enabled group can be teams enabled with Set-GraphGroup -EnableTeam
        Creating and modifying groups requires consent to use the Group.ReadWrite.All scope
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias("New-GraphTeam")]
    Param(
        #The Name of the group / team
        [parameter(Mandatory=$true, Position=0)]
        [string]$Name,
        #The group/team's mail nickname
        [string]$MailNickName,
        #A description for the group
        [string]$Description,
        #The visibility of the group, Public by default, it can be 'private' or 'hidden membership'
        [ValidateSet('private', 'public', 'hiddenmembership')]
        [string]$Visibility = 'public',
        #Ordinary Members of the group - assumed to be users, given by their User Principal Name or ID or as objects
        $Members,
        #Owners of the group - assumed to be users, given by their User Principal Name or ID or as objects
        $Owners,
        #By default the group is configured as a team unless -NoTeam is specified
        [Switch]$NoTeam,
        #if specified group will be added without prompting
        [Switch]$Force
    )
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

    $webParams = @{ Headers     = $Script:DefaultHeader  }
    if ( (Invoke-RestMethod -Method Get @webParams -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$Name'" ).value) {
        throw "There is already a group with the display name '$Name'." ; return
    }
    #Server-side is case-sensitive for [most] JSON so make sure hashtable names and constants have the right case!
    if (-not $MailNickName) {$MailNickName = $Name -replace "\W",'' }
    $settings = @{  'displayName'          = $Name ;
                    'mailNickname'         = $MailNickName;
                    'mailEnabled'          = $true;
                    'securityEnabled'      = $false;
                    'visibility'           = $Visibility.ToLower() ;
                    'groupTypes'           = @("Unified") ;
    }
    if ($Description)        {$settings['description'] = $Description}
    #if we got owners or users with no ID, fix them at the end, if they have an ID add them now
    if ($Members) {
        $settings['members@odata.bind']= @();
        foreach ($m in $Members) {
            if  ($m.id) {$settings['members@odata.bind'] += "https://graph.microsoft.com/v1.0/users/$($m.id)"}
            else        {$settings['members@odata.bind'] += "https://graph.microsoft.com/v1.0/users/$m"}
         #   else        {$noIDMembers += $m}
        }
    }
    #If we make someone else the owner of the group, we can't make it a team, so only set owners here if we are not making a team.
    if ($noTeam -and $Owners) {
        $settings['owners@odata.bind']= @()
        foreach    ($o in $Owners)  {
            if     ($o.id) {$settings['owners@odata.bind']  += "https://graph.microsoft.com/v1.0/users/$($o.id)"}
            else{           $settings['owners@odata.bind']  += "https://graph.microsoft.com/v1.0/users/$o"}
        }
    }
    $webParams["contentType"] = 'application/json'
    #Don't add URI or body to web params as we are going to make two calls ...
    $uri       = "https://graph.microsoft.com/v1.0/groups"
    $json = ConvertTo-Json $settings
    Write-Debug $json

    if ($Force -or $PSCmdlet.shouldProcess($Name,"Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-RestMethod @webParams -Method Post -uri $uri -body $json
        foreach ($m in $group.members) {if ($m.'@odata.type' -match "user") {$m.pstypenames.add("GraphUser")}}
        if ($NoTeam) {
            $group.pstypenames.Add("GraphGroup")
            Write-Progress -Activity 'Creating Group/Team' -Completed
            return $group
        }
        else {
            $uri = $uri + "/" + $group.id + "/team"
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name"
            $team   = Invoke-RestMethod @webParams -Method Put -uri $uri -Body "{ }"
            $team.pstypenames.Add("GraphTeam")
            Add-Member -InputObject $team -MemberType NoteProperty -Name DisplayName -Value $group.displayName
            Add-Member -InputObject $team -MemberType NoteProperty -Name Description -Value $group.description
            Add-Member -InputObject $team -MemberType NoteProperty -Name Members     -Value $group.members
            Add-Member -InputObject $team -MemberType NoteProperty -Name Mail        -Value $group.Mail
            Add-Member -InputObject $team -MemberType NoteProperty -Name visibility  -Value $group.visibility
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Setting Group ownership on $Name"
            if ($Owners) {
                $Owners | Add-GraphGroupMember -Group $group -AsOwner -Force
            }
            Write-Progress -Activity 'Creating Group/Team' -Completed
            return $team
        }
    }
}
<#
See also https://docs.microsoft.com/e#n-gb/graph/api/team-post?view=graph-rest-beta

POST https://graph.microsoft.com/beta/teams
Content-Type: application/json
{
 "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates/standard",
 "displayName": "My Sample Team",
 "description": "My Sample Team’s Description",
}
#>

Function Set-GraphGroup {
    <#
      .synopsis
        Sets options on a group
      .Description
        Allows or blocks external senders, changes visibility or description and enables the group as a team.
        Other options for a team are set via Set-GraphTeam.
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    Param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        $Group ,
        #If specified, the group can receive external email; the option can be disabled with -AllowExternalSenders:$false.
        [switch]$AllowExternalSenders,
        #A new description for the group
        [string]$Description,
        #The visibility of the group; groups are created as public by default, it can be 'private' or 'hidden membership'
        [ValidateSet('private', 'public', 'hiddenmembership')]
        [string]$Visibility,
        #Enables team functionality on a group which does not yet have it enabled
        [switch]$EnableTeam,
        #If specified the group will be updated without prompting for confirmation.
        [switch]$Force
    )
    if ($Group.Id)  {$uri = "https://graph.microsoft.com/v1.0/groups/$($Group.ID)"}
    else            {$uri = "https://graph.microsoft.com/v1.0/groups/$Group"}
    $settings = @{}
    if ($Visibility)                         {$settings['visibility']            = $Visibility.ToLower()}
    if ($Description)                        {$settings['description']           = $Description}
    if ($PSBoundParameters.ContainsKey(
                    'AllowExternalSenders')) {$settings['allowExternalSenders']  = [bool]$AllowExternalSenders}
    if ($settings.Count -and ($Force -or $PSCmdlet.ShouldProcess($group.displayname,'Update Group'))) {
        Connect-MSGraph
        Invoke-RestMethod -Uri $uri -Headers $Script:DefaultHeader -Method Patch -ContentType 'application/json' -Body (ConvertTo-Json $settings)

        $g = Invoke-RestMethod -Uri $uri -Headers $Script:DefaultHeader -Method Get
        if ($EnableTeam -and $g.resourceProvisioningOptions -notcontains 'Team') {
             $uri = $uri +  "/team"
             Invoke-RestMethod -Uri $uri -Headers $Script:DefaultHeader -Method Put -ContentType 'application/json' -Body "{ }"
        }
        elseif ($EnableTeam) {Write-Warning  "$($g.displayName) is already team enabled" }
    }
}

Function Remove-GraphGroup {
    <#
      .Synopsis
        Removes a group/team
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeam")]
    Param(
        #The ID of the Group / team
        [parameter(Mandatory=$true, Position=0,ValueFromPipeline=$true )]
        [Alias("Team")]
        $Group,
        #If specified the group will be removed without prompting
        $Force
    )
    Begin   {
        Connect-MSGraph
    }
    Process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if ($Group.displayName) {$displayName = $Group.DisplayName}
        if  ($Group.id)         {$Group   = $Group.id}

        $webParams = @{Headers = $Script:DefaultHeader ; uri =  "https://graph.microsoft.com/v1.0/groups/$Group/"}
        if (-not $displayName){
            try   {  $g  = Invoke-RestMethod -Method Get @webParams }
            catch        {throw "Could not get the thread to delete"; return}
            if (-not $g) {throw "Could not get the thread to delete"; return}
            else         {$displayName = $g.displayname}
        }
         if ($PSCmdlet.ShouldProcess($DisplayName,"Delete Group")) { Invoke-RestMethod -Method Delete  @webParams }
    }
}

# Groups in the recycle bin (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group").value
# DELETE /directory/deletedItems/{id}                permanent delete
# POST /directory/deletedItems/{id}/restore          restore item

Function Add-GraphGroupMember {
    <#
      .Synopsis
        Adds a user (or group) to a group/team as either a member or owner.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias("Add-GraphTeamMember")]
    Param(
        #The group / team either as an ID or a group/team object with an IDn
        [parameter(Mandatory=$true, Position=0)]
        [Alias("Team")]
        $Group,
        #The user or nested-group to add, either as a UPN or ID or as a object with an ID
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        $Member,
        #If specified the user will be added as an owner, otherwise they will be a standard member
        [switch]$AsOwner,
        #If specified group member will be added without prompting
        [Switch]$Force
    )
    Begin   {
        Connect-MSGraph
        if ($Group.id) {$Group       = $Group.id}
        $webParams = @{'Method'      = 'Post'
                       'Headers'     = $Script:DefaultHeader
                       'ContentType' = 'application/json'
        }
        if ($AsOwner) {
              $webParams['URI']      = "https://graph.microsoft.com/v1.0/groups/$Group/owners/`$ref"
        }
        else {$webParams['URI']      = "https://graph.microsoft.com/v1.0/groups/$Group/members/`$ref"}
    }
    Process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if   ($Member.displayName) {$membername = $Member.displayName}
        if   ($Member.id)          {$Member     = $Member.id}
        else {
            try {$u = Get-GraphUser -User $Member
                $Member     = $u.id
                $membername = $u.displayName
            }
            catch {throw "Could not get a user matching $Member"; return }
            if (-not $Member) {throw "Could not get a member ID" ; return}
        }

        $settings  = @{'@odata.id'   = "https://graph.microsoft.com/v1.0/directoryObjects/$Member"   }
        $json      = convertto-Json $settings
        Write-Debug $json
        if ($Force -or $PSCmdlet.shouldProcess($membername,"Add to Group")) {
            Invoke-RestMethod @webParams -Body $json  }
    }
}

Function Remove-GraphGroupMember {
    <#
      .Synopsis
        Removes a user (or group) from a group/team
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeamMember")]
    Param(
        #The ID of the Group / team
        [parameter(Mandatory=$true, Position=0)]
        [Alias("Team")]
        $Group,
        #A group object with an ID field, or a user object, user ID or UPN
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        $Member,
        #If specified the user will be removed without prompting for confirmation
        $Force
    )
    Process {
        if   ($Group.id)   {$Group   = $Group.id}
        if   ($Member.id) {$Member = $Member.id}
        else {
            try {
                $u =  Get-GraphUser -User $Member
                $Member =$u.id
                $userName = $u.displayName
            }
            catch {throw "Could not get a user matching $Member"; return }
            if (-not $Member) {throw "Could not get a member ID" ; return}
        }
        #https://docs.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Method      = "Delete"
                       URI         = "https://graph.microsoft.com/v1.0/groups/$Group/members/$Member/`$ref"
                       Headers     = $Script:DefaultHeader
                       contentType = 'application/json'
                    }
        if ($Force -or $PSCmdlet.ShouldProcess($userName,"Remove from Group")) {Invoke-RestMethod @webParams }
    }
}

Function Get-GraphTeam {
    <#
      .Synopsis
        Gets information about an office 365 team
      .Description
        Takes a a team ID or team object as a parameter and gets information about the team
        The teams Apps, Calendar, Channels, Drive, Members or Planners can be requested.
      .Example
        >get-graphuser -teams | get-graphteam -Plans | select -last 1 | get-graphplan -FullTasks  | ft PlanTitle,Bucketname,Title,DueDateTime,PercentComplete,Assignees
         Gets the current user's Teams, and gets the plans for each; selects just the last one, and gets its task details, showing the result as a table.
      .Example
        >(Get-GraphTeam -Site).lists | where name -match document
        Gets team(s) for the current user and returns the associated site(s).
        Site objects include a lists property, which holds a collection of lists
        this command will fiter the lists down to those where name matches "document"
      .Example
        >(Get-GraphTeam -Drive).root.children.where({$_.folder}) | Select  name, weburl, id,@{n="drive";e={$_.parentReference.driveId}}
        Gets team(s) for the current user and returns the associated drive(s)
        Drive objects include a root property, which holds an object describing the root folder;
        this in turn has a children property which contains files and folder objects in the root folder.
        This command filters the children collection to folders and returns their name,
        WebURl and the item ID and Drive ID needed to access them from one
      .Example
        >Get-GraphTeam -Notebooks | select -ExpandProperty sections | where "Displayname" -eq "General_Notes"
        Gets team(s) for the current user and returns the associated notebooks(s)
        Notebook objects include a Sections property, which holds a collection of OneNote sections in the notebook;
        This command gets returns any section in a team notebook which has the name "General_Notes"
      .Example
        > Get-GraphTeam -threads | where {[datetime]$_.lastDeliveredDateTime -gt [datetime]::Now.AddDays(-7) }
        Gets the teams conversation threads which have been updated in the last 7 days.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Alias("Get-GraphGroup")]
    Param(
        #The name of a team.
        #One more Team IDs or team objects containing and ID. If omitted the current user's teams will be used.
        [parameter(ValueFromPipeline=$true, Position=0)]
        [Alias("ID","Group")]
        $Team ,
        #If specified the Team parameter is treated as a name not an ID
        [Switch]$ByName,
        #If specified returns the teams Apps
        [Parameter(Mandatory=$true, ParameterSetName='Apps')]
        [switch]$Apps,
        #If specified gets the team's Calendar (a team only has one)
        [Parameter(Mandatory=$true, ParameterSetName='Calendar')]
        [switch]$Calendar,
        #If specified gets the team's channels
        [Parameter(Mandatory=$true, ParameterSetName='Channels')]
        [switch]$Channels,
        #If Specified, retrun team's conversations (usually better to use threads)
        [Parameter(Mandatory=$true, ParameterSetName='Conversations' )]
        [switch]$Conversations,
        #If specified gets the Team's one drive
        [Parameter(Mandatory=$true, ParameterSetName='Drive')]
        [switch]$Drive,
        #If specified returns the members of the team
        [Parameter(Mandatory=$true, ParameterSetName='Members')]
        [switch]$Members,
        #If specified returns the Owners of the team
        [Parameter(Mandatory=$true, ParameterSetName='Owners')]
        [switch]$Owners,
        #If specified returns the team's notebook(s)
        [Parameter(Mandatory=$true, ParameterSetName='Notebooks')]
        [switch]$Notebooks,
        #if Specified, returns the teams Planners.
        [Parameter(Mandatory=$true, ParameterSetName='Planners')]
        [switch]$Plans,
        #If Specified, retrun team's threads
        [Parameter(Mandatory=$true, ParameterSetName='Threads' )]
        [switch]$Threads,
        #if Specified, returns the teams site.
        [Parameter(Mandatory=$true, ParameterSetName='Site')]
        [switch]$Site,
        #limits searches for appsby name.
        [Parameter(ParameterSetName='Apps')]
        [String]$AppName,
        #limits searches for channels by name. Other's cant be filtered by name ...  perhaps notebooks can but a group only has one.
        [Parameter(ParameterSetName='Channels')]
        [String]$ChannelName
    )
    begin {
        Connect-MSGraph
        $webParams = @{Method = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($ByName -and $Team -isnot [string]) {Write-Warning 'The team parameter does not look like a name'; return}
        elseif ($ByName)    {$Team = Get-GraphGroupList -Name $Team}
        elseif (-not $Team) {$Team  = Get-GraphUser      -Teams }
        if     (-not $Team) {Write-Warning 'Could not Get a team from the parameters provided' ; return}
        foreach ($t in   $Team) {
            if  ($t.id) {$teamid = $t.id}
            else        {$teamid = $t }
            $groupURI = "https://graph.microsoft.com/v1.0/groups/$teamid"
            $teamURI  = "https://graph.microsoft.com/v1.0/teams/$teamid"
            try {
                #For each of the switches get the data from /groups{id}/whatever or /teams/{id}.whatever
                #Add a type to PS Type names so we can format it, and add any properties we expect to want later.
                if     ($Site)          {
                    Write-Progress -Activity 'Getting Group Site Information'
                    $result  =  Invoke-RestMethod  @webparams -Uri ("$groupURI/sites/root"  + '?expand=drives,sites,lists(expand=columns,contenttypes,drive)')
                    foreach ($s in $result) {
                        $s.pstypenames.Add("GraphSite")
                        foreach ($l in $s.lists) {
                            $l.pstypenames.add('GraphList')
                            Add-Member -InputObject $l -MemberType NoteProperty   -Name SiteID    -Value  $s.id
                            Add-Member -InputObject $l -MemberType NoteProperty   -Name ParentUrl -Value  $s.weburl
                            Add-Member -InputObject $l -MemberType ScriptProperty -Name Template  -Value {$this.list.template}
                            $l.columns | ForEach-Object {$_.pstypenames.add('GraphColumn')}
                        }
                        $s.drives      | ForEach-Object {$_.pstypenames.add('GraphDrive') }
                    }
                    Write-Progress -Activity 'Getting Group Site Information' -Completed
                    return $result
                }
                elseif ($Calendar)      {
                    Write-Progress -Activity 'Getting Group Calendar'
                    $result = Invoke-RestMethod  @webparams -Uri  "$groupURI/calendar"
                    $result.pstypenames.Add("GraphCalendar")
                    Add-Member -InputObject $result -MemberType NoteProperty -Name GroupID -Value $teamid
                    Write-Progress -Activity 'Getting Group Calendar' -Completed
                    return $result
                }
                elseif ($Drive)         {
                    Write-Progress -Activity 'Getting Group Drive information'
                    $result = Invoke-RestMethod  @webparams -Uri ("$groupURI/drive" + '?$expand=root($expand=children)' )
                    $result.pstypenames.Add("GraphDrive")
                    Write-Progress -Activity 'Getting Group Drive information' -Completed
                    return $result
                }
                elseif ($Members)       {
                    Write-Progress -Activity 'Getting Group Members'
                    $result = (Invoke-RestMethod  @webparams -Uri  "$groupURI/members")
                    $users  = $result.value  #do we need a  while ($result.'@odata.nextLink') { irm nextlink, add value to users} ??
                    foreach ($u in $users) {
                         if ($u.'@odata.type'  -match 'user') {$u.psTypenames.Add("GraphUser")}
                    }
                    Write-Progress -Activity 'Getting Group Members' -Completed
                    return $users
                }  #can do group ?$expand=Memebers, the others don't expand
                elseif ($Owners)       {
                    Write-Progress -Activity 'Getting Group Owners'
                    $result = (Invoke-RestMethod  @webparams -Uri  "$groupURI/Owners")
                    $users  = $result.value  #do we need a  while ($result.'@odata.nextLink') { irm nextlink, add value to users} ??
                    foreach ($u in $users) {
                         if ($u.'@odata.type'  -match 'user') {$u.psTypenames.Add("GraphUser")}
                    }
                    Write-Progress -Activity 'Getting Group Owners' -Completed
                    return $users
                }  #can do group ?$expand=Memebers, the others don't expand
                elseif ($Notebooks)     {
                    Write-Progress -Activity 'Getting Group OneNote Notebooks' -Completed
                    #if groups can have more than then add if name ... uri = blah + "?`$expand=sections&`$filter=startswith(tolower(displayname),'$name')"
                    $results = (Invoke-RestMethod  @webparams -Uri ("$groupURI/onenote/notebooks" + '?$expand=sections'  ) )
                    $books   = $results.value
                    foreach ($b in $books) {
                        $b.pstypenames.add("GraphOneNoteBook")
                        foreach ($s in $b.sections) {
                            Add-Member -InputObject $s -MemberType NoteProperty -Name ParentNotebookID -Value $b.id
                            $s.pstypeNames.add("GraphOneNoteSection")
                        }
                    }
                    Write-Progress -Activity 'Getting Group OneNote Notebooks' -Completed
                    return $books
                }
                elseIf ($Plans)         {
                    Write-Progress -Activity 'Getting Group Planner Plans'
                    $result = (Invoke-RestMethod  @webparams -Uri  "$groupURI/planner/plans") #would like to have expand details here but it only works with a single plan.
                    $planList  = $result.value
                    while ($result.'@odata.nextLink') {
                        $result = Invoke-RestMethod  @webparams -Uri $result.'@odata.nextLink'
                        $planList += $result.value
                    }
                    $dirObjectsHash = @{}
                    @() + $planList.owner + $planList.createdby.user.id  |
                         Sort-Object -Unique | ForEach-Object  {
                            $dirObjectsHash[$_] = (Invoke-RestMethod  @webparams -Uri "https://graph.microsoft.com/v1.0/directoryobjects/$_").displayname
                    }
                    foreach ($p in $planList) {
                        $p.pstypenames.add("GraphPlan")
                        Add-Member -InputObject $P -MemberType NoteProperty -Name OwnerName   -Value $dirObjectsHash[$p.owner]
                        Add-Member -InputObject $P -MemberType NoteProperty -Name CreatorName -Value $dirObjectsHash[$p.createdBy.user.id]
                    }
                    Write-Progress -Activity 'Getting Group Planner Plans' -Completed
                    return $planList
                }
                elseif ($Threads)       {
                    Write-Progress -Activity 'Getting Group Conversation threads'
                    $results = (Invoke-RestMethod  @webparams -Uri  "$groupURI/threads")
                    $threadList = $results.value #do we need a  while ($result.'@odata.nextLink') { irm nextlink, add value to threads} ??
                    foreach ($t in $threadList) {
                        $t.pstypenames.add("GraphThread")
                        Add-Member -InputObject $t -MemberType NoteProperty -Name Team -Value $teamid
                    }
                    Write-Progress -Activity 'Getting Group Conversation threads' -Completed
                    return $threadList
                }
                elseif ($Conversations) {
                    Write-Progress -Activity 'Getting Group Conversations'
                    $results  = (Invoke-RestMethod  @webparams -Uri ("$groupURI/conversations" +'?$expand=Threads')  )
                    $convList = $results.value #do we need a  while ($result.'@odata.nextLink') { irm nextlink, add value to convList} ??
                    foreach ($c in $convList) {
                        $c.pstypenames.add("GraphConversation")
                        Add-Member -InputObject $c -MemberType NoteProperty -Name Team -Value $teamid
                        foreach ($t in $c.threads) {
                            $t.pstypenames.add("GraphThread")
                            Add-Member -InputObject $t -MemberType NoteProperty -Name Team -Value $teamid
                        }
                    }
                    Write-Progress -Activity 'Getting Group Conversations' -Completed
                    return $convList
                }
                elseif ($Channels)      {
                    if   ($ChannelName) {
                        $uri =  "$teamURI/channels?`$filter=startswith(tolower(displayname), '$ChannelName')"
                    }
                    else {
                        $uri =  "$teamURI/channels"
                    }
                    Write-Progress -Activity 'Getting Team Channels'
                    $results  = (Invoke-RestMethod  @webparams -Uri $uri)
                    $chanList = $results.value
                    foreach ($c in $chanList) {
                         $c.pstypenames.add("GraphChannel")
                         Add-Member -InputObject $c -MemberType NoteProperty -Name Team -Value $teamid
                    }
                    Write-Progress -Activity 'Getting Team Channels' -Completed
                    return $chanList
                }
                elseif ($Apps)          {
                    Write-Progress -Activity 'Getting Team Apps'
                    if ($AppName) {
                        $uri = "$teamURI/installedApps" +
                                    '?$expand=teamsAppDefinition&$filter=' +
                                    "startswith(tolower(teamsappdefinition/displayname),'$($AppName.ToLower())')"
                    }
                    else {
                        $uri = ("$teamURI/installedApps" + '?$expand=teamsAppDefinition')
                    }
                    $results  = (Invoke-RestMethod  @webparams -Uri $uri)
                    $appsList = $results.value
                    foreach ($a in $appsList) {$a.pstypenames.add("GraphApp")}
                    Write-Progress -Activity 'Getting Team Apps' -Completed
                    return $appsList
                }
                else                    {
                    Write-Progress -Activity 'Getting Group/Team information'
                    $g =  Invoke-RestMethod  @webparams -Uri "$groupuri`?`$expand=members"
                    if ($g.resourceProvisioningOptions -contains 'Team') {
                        $t = Invoke-RestMethod  @webparams -Uri  "$teamURI"
                        $t.pstypenames.Add("GraphTeam")
                        $memberList = $g.members | Select-Object id,department,displayname,mail,UserPrincipalName,usertype,businessPhones,MobilePhone,OfficeLocation
                        foreach ($m in $memberList) {if ($m.'@odata.type' -match "user") {$m.pstypenames.add("GraphUser")}}
                        Add-Member -InputObject $t -MemberType NoteProperty -Name DisplayName -Value $g.displayName
                        Add-Member -InputObject $t -MemberType NoteProperty -Name Description -Value $g.description
                        Add-Member -InputObject $t -MemberType NoteProperty -Name Members     -Value $memberList
                        Add-Member -InputObject $t -MemberType NoteProperty -Name Mail        -Value $g.Mail
                        Add-Member -InputObject $t -MemberType NoteProperty -Name visibility  -Value $g.visibility
                        Write-Progress -Activity 'Getting Group/Team information' -Completed
                        return $t
                    }
                    else {
                        $g.pstypenames.Add("GraphGroup")
                        Write-Progress -Activity 'Getting Group/Team information' -Completed
                        return $g
                    }
                }
            }
            catch {
                if ($_.exception -match"403\) Forbidden") {
                    Write-warning -Message "Server returned a 403 (Forbidden) error ; you must be a memeber of the team to view some things [admin does not give access]. "
                }
                else {throw $_  }
            }
        }
    }
}
#>(irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/groupsettings/$($team.id)") ##may be empty

Function Set-GraphTeam {
    <#
      .Synopsis
        Updates the settings for a team
    #>
    [cmdletbinding()]
    Param (
        #The team to update either as an ID or a team object with and ID.
        $Team ,
        #Allow members to add or remove apps
        [nullable[bool]]$AllowMemberAddRemoveApps,
        #Allow members to create update or remove connectors
        [nullable[bool]]$AllowMemberCreateUpdateRemoveConnectors,
        #Allow members to create update or remove Tabs
        [nullable[bool]]$AllowMemberCreateUpdateRemoveTabs,
        #Allow members to create or update Channels
        [nullable[bool]]$AllowMemberCreateUpdateChannels,
        #Allow members to delete Channels
        [nullable[bool]]$AllowMemberDeleteChannels,
        #Allow guests to create or update Channels
        [nullable[bool]]$AllowGuestCreateUpdateChannels,
        #Allow guests to delete Channels
        [nullable[bool]]$AllowGuestDeleteChannels,
        #Allow members to edit their own messages
        [nullable[bool]]$AllowUserEditMessages,
        #Allow members to delete their own messages
        [nullable[bool]]$AllowUserDeleteMessages,
        #Allow owners to delete mssages
        [nullable[bool]]$AllowOwnerDeleteMessages,
        #Allow mentions of teams in messages
        [nullable[bool]]$AllowTeamMentions,
        #Allow mentions of channels in messages
        [nullable[bool]]$AllowChannelMentions,
        #Allow giphy graphics
        [nullable[bool]]$AllowGiphy,
        #Rating for giphy graphics; either moderate or strict
        [ValidateSet('moderate', 'strict')]
        [string]$GiphyContentRating,
        #Allow stickers and memes
        [nullable[bool]]$AllowStickersAndMemes,
        #Allow Custom memes
        [nullable[bool]]$AllowCustomMemes
    )
    Connect-MSGraph
    $webParams = @{Method      =  'PATCH'
                  ContentType  =  'application/json'
                  Headers      =  $Script:DefaultHeader }

    if ($Team.id) {$webParams['Uri'] = "https://graph.microsoft.com/v1.0/teams/$($Team.id)"}
    else          {$webParams['Uri'] = "https://graph.microsoft.com/v1.0/teams/$Team"}

    $settings = @{}
    if ($PSBoundParameters.ContainsKey('AllowMemberAddRemoveApps') -or
        $PSBoundParameters.ContainsKey('AllowMemberCreateUpdateChannels') -or
        $PSBoundParameters.ContainsKey('AllowMemberCreateUpdateRemoveConnectors') -or
        $PSBoundParameters.ContainsKey('AllowMemberCreateUpdateRemoveTabs') -or
        $PSBoundParameters.ContainsKey('AllowMemberDeleteChannels')
       ) {
       $settings['memberSettings'] = @{}
       if ($PSBoundParameters.ContainsKey('AllowMemberAddRemoveApps'))                {$settings.memberSettings.allowAddRemoveApps                = $AllowMemberAddRemoveApps}
       if ($PSBoundParameters.ContainsKey('AllowMemberCreateUpdateChannels'))         {$settings.memberSettings.allowCreateUpdateChannels         = $AllowMemberCreateUpdateChannels}
       if ($PSBoundParameters.ContainsKey('AllowMemberCreateUpdateRemoveConnectors')) {$settings.memberSettings.allowCreateUpdateRemoveConnectors = $AllowMemberCreateUpdateRemoveConnectors}
       if ($PSBoundParameters.ContainsKey('AllowMemberCreateUpdateRemoveTabs'))       {$settings.memberSettings.allowCreateUpdateRemoveTabs       = $AllowMemberCreateUpdateRemoveTabs}
       if ($PSBoundParameters.ContainsKey('AllowMemberDeleteChannels'))               {$settings.memberSettings.allowDeleteChannels               = $AllowMemberDeleteChannels}
    }

    if ($PSBoundParameters.ContainsKey('AllowGuestCreateUpdateChannels') -or
        $PSBoundParameters.ContainsKey('AllowGuestDeleteChannels')
       ) {
        $settings['guestSettings'] = @{}
        if ($PSBoundParameters.ContainsKey('AllowGuestCreateUpdateChannels'))         {$settings.guestSettings.allowCreateUpdateChannels          = $AllowGuestCreateUpdateChannels}
        if ($PSBoundParameters.ContainsKey('AllowGuestDeleteChannels'))               {$settings.guestSettings.allowDeleteChannels                = $AllowGuestDeleteChannels}
    }

    if ($PSBoundParameters.ContainsKey('AllowUserEditMessages') -or
        $PSBoundParameters.ContainsKey('AllowUserDeleteMessages') -or
        $PSBoundParameters.ContainsKey('AllowOwnerDeleteMessages') -or
        $PSBoundParameters.ContainsKey('AllowTeamMentions') -or
        $PSBoundParameters.ContainsKey('AllowChannelMentions')
       ) {
        $settings['messagingSettings'] = @{}
        if ($PSBoundParameters.ContainsKey('AllowUserEditMessages'))                  {$settings.messagingSettings.allowUserEditMessages          = $AllowUserEditMessages}
        if ($PSBoundParameters.ContainsKey('AllowUserDeleteMessages'))                {$settings.messagingSettings.allowUserDeleteMessages        = $AllowUserDeleteMessages}
        if ($PSBoundParameters.ContainsKey('AllowOwnerDeleteMessages'))               {$settings.messagingSettings.allowOwnerDeleteMessages       = $AllowOwnerDeleteMessages}
        if ($PSBoundParameters.ContainsKey('AllowTeamMentions'))                      {$settings.messagingSettings.allowTeamMentions              = $AllowTeamMentions}
        if ($PSBoundParameters.ContainsKey('AllowChannelMentions'))                   {$settings.messagingSettings.allowChannelMentions           = $AllowChannelMentions}
    }

    if ($PSBoundParameters.ContainsKey('AllowGiphy') -or
        $PSBoundParameters.ContainsKey('GiphyContentRating') -or
        $PSBoundParameters.ContainsKey('AllowStickersAndMemes') -or
        $PSBoundParameters.ContainsKey('AllowCustomMemes')
       ){
        $settings['funSettings'] = @{}
        if ($PSBoundParameters.ContainsKey('AllowGiphy'))                             {$settings.funSettings.allowGiphy                           = $AllowGiphy}
        if ($PSBoundParameters.ContainsKey('GiphyContentRating'))                     {$settings.funSettings.giphyContentRating                   = $GiphyContentRating}
        if ($PSBoundParameters.ContainsKey('AllowStickersAndMemes'))                  {$settings.funSettings.allowStickersAndMemes                = $AllowStickersAndMemes}
        if ($PSBoundParameters.ContainsKey('AllowCustomMemes'))                       {$settings.funSettings.allowCustomMemes                     = $AllowCustomMemes}
    }

    $json = ConvertTo-Json $settings -Depth 10
    Write-Debug $json

   Invoke-RestMethod @webParams -Body $json
}

Function Get-GraphGroupConversation {
    <#
      .Synopsis
        Gets details of group converstation from outlook, or its threads,
      .Example
        Get-GraphGroupList -Name consult | Get-GraphGroup -Conversations | Get-GraphGroupConversation -Threads
        Gets group(s) matching the name "consult*" , finds their conversations and for each one gets the threads in the conversation
        Note, unless you are dealing with conversations which have multiple threads, it is easier to do Get-GraphGroup -Threads
    #>
    [cmdletbinding()]
    [Alias("Get-GraphTeamConversation","Get-GraphConversation")]  #Strictly Conversations belong to a group in Outlook, not a Team in Microsoft teams, but let either name be used.
    Param(
        #The Conversation, either as an ID or an object.
        [parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0)]
        $Conversation,
        #The group where the conversation is found, either as an ID or as an object, if it can't be found from the conversation
        [Alias("Team")]
        $Group,
        #If specified selects the conversations threads, otherwise returns an object representing the conversation itself
        [Switch]$Threads
    )
    process {
        if     ($Group.id)          {$Group = $Group.ID}
        elseif ($Conversation.Team) {$Group = $Conversation.Team}
        if     ($Conversation.id)   {$Conversation = $Conversation.id}
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Method = "Get"
                    Headers = $Script:DefaultHeader
        }
        if ($Threads) { (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/groups/$Group/conversations/$conversation/Threads").value |
                            ForEach-Object {$_.pstypenames.add("GraphThread"); $_ }  |
                            Add-Member -PassThru -MemberType NoteProperty -Name Team         -Value $Group |
                            Add-Member -PassThru -MemberType NoteProperty -Name Conversation -Value $Conversation
        }
        else     {
            $c = (Invoke-RestMethod @webParams -Uri ("https://graph.microsoft.com/v1.0/groups/$Group/conversations/$conversation"  +'?$expand=Threads'))
            $c.pstypenames.add("GraphConversation")
            return $c
        }
    }
}

Function Get-GraphGroupThread {
    <#
      .Synopsis
        Gets a thread in a Group conversation in outlook, or its posts
      .Example
        >Get-GraphUser -Teams  | Get-GraphGroup -Threads | Get-GraphGroupThread -Posts |
             ft -a -Wrap  @{n="from";e={$_.from.emailaddress.name}},CreatedDateTime,Topic,@{n="Body";e={$_.body.content}}
        Gets a users teams, for each one gets their threads, and for each thread gets the outlook posts
        Displays the result as a table showing from, message date, thread topic and message body
        Note this uses Get-GraphGroup as an alias for Get-GraphTeams
    #>
    [cmdletbinding()]
    [Alias("Get-GraphTeamThread")]
    Param(
        #The group thread, either as an ID or as a thread object (which may have the team/group as property)
        [parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0)]
        $Thread,
        #The group holding the thread, if it can't be found drm -thread
        [Alias("Team")]
        $Group,
        #If specified, returns the posts in the thread
        [Switch]$Posts
    )

    Process {
        if     ($Thread.topic) {$Topic   = $Thread.topic} else {$topic = "-"}
        if     ($Group.id)     {$Group   = $Group.ID}
        elseif ($Thread.Team)  {$Group   = $Thread.Team}
        if     ($Thread.id)    {$Thread  = $Thread.id}
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Method  = "Get"
                       Headers = @{Authorization = $Script:AuthHeader ; "Prefer" ='outlook.body-content-type="text"' }
        }
        if ($Posts) { (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/groups/$Group/Threads/$thread/posts").value |
                            ForEach-Object {$_.pstypenames.add("GraphPost") ; $_ } |
                            Add-Member -PassThru -MemberType NoteProperty -Name Team   -Value $Group  |
                            Add-Member -PassThru -MemberType NoteProperty -Name Thread -Value $Thread |
                            Add-Member -PassThru -MemberType NoteProperty -Name Topic  -Value $Topic

        }
        else          {
                 $t = (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/groups/$Group/Threads/$thread")
                 $t.pstypenames.Add("GraphThread")
                 return $t
        }
    }
}

Function Add-GraphGroupThread {
    <#
      .Synopsis
        Starts a new thread in a group in outlook.
    #>
    [cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    Param (
        #The group where the thread will be added
        [Parameter(Mandatory=$true,Position=0)]
        [Alias("Team")]
        $Group,
        #The subkect line for the thread
        [Parameter(Mandatory=$true, Position=1)]
        [Alias("Subject")]
        $ThreadTopic,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true, Position=2)]
        [String]$Content,
        #The content type, (Text by default) or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )

    if    ($Group.ID) {$Group  = $Group.ID}

    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    $webParams = @{ Headers     = $Script:DefaultHeader ;
                    uri         =  "https://graph.microsoft.com/v1.0/groups/$Group/threads/";
                    method      = "Post";
                    contentType = "application/json"
    }
    $Settings = @{  topic       = $ThreadTopic;
                    posts       = @( @{body= @{content=$Content; contentType=$ContentType}})}
    if ($force -or $PSCmdlet.ShouldProcess($ThreadTopic,"Create New thread")) {
        Invoke-RestMethod  @webparams -Body (ConvertTo-Json $settings -Depth 5 )
    }
}

Function Remove-GraphGroupThread {
    <#
      .Synopsis
        Removes a thread from a group in outlook
    #>
    [cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    Param (
        #The thread to remove, either as an ID or a thread object containing an ID, and possibly a group ID
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Thread,
        #The group from which the thread is to be removed, either as an ID or a group object containing an ID
        [Alias("Team")]
        $Group,
        #if Specified the thread will be deleted without prompting.
        [switch]$Force
    )
    process {
        if     ($Thread.Team) {$Group  = $Thread.Team}
        elseif    ($Group.ID) {$Group  = $Group.ID}

        if       ($Thread.ID) {$Thread = $Thread.id  }

        if (-not ($Thread -and $Group)) {throw "Could not find Group and Thread IDs from supplied parameters."; Return}

        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Headers = $Script:DefaultHeader }
        $uri =  "https://graph.microsoft.com/v1.0/groups/$Group/threads/$thread"
        try   {$thread  = Invoke-RestMethod -Method Get -uri $uri @webParams }
        catch             {throw "Could not get the thread to delete"; return}
        if (-not $thread) {throw "Could not get the thread to delete"; return}
        if ($PSCmdlet.ShouldProcess($thread.topic,"Delete thread")) {
            Invoke-RestMethod -Method Delete -uri $uri @webParams
        }
    }
}

Function Send-GraphGroupReply {
    <#
      .Synopsis
        Replies to a group's post in outlook.
    #>
    [cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    Param (
        #The Post being replied to, either as an ID or a post object containing an ID which may identify the thread and group
        [Parameter(Mandatory=$true,Position=0)]
        $Post,
        #The tread containing the post (if not embedded in the post itself), as an ID or object, which may identify the group
        $Thread,
        #The group containing the thread (if not embedded in the Post or thread) as an ID or object
        [Alias("Team")]
        $Group,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The type of content, text by default or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )

    if       ($Post.Team) {$Group  = $Post.Team}
    elseif ($Thread.Team) {$Group  = $Thread.Team}
    elseif    ($Group.ID) {$Group  = $Group.ID}

    if     ($Post.Thread) {$Thread = $Post.Thread}
    elseif   ($Thread.ID) {$Thread = $Thread.id  }

    if         ($Post.ID) {$Post   = $Post.ID}

    if (-not ($Post -and $Thread -and $Group)) {throw "Could not find Group, Thread and Post IDs from supplied parameters."; Return}

    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    $webParams = @{Headers = $Script:DefaultHeader }
    $uri =  "https://graph.microsoft.com/v1.0/groups/$Group/threads/$thread"
    try   {$thread  = Invoke-RestMethod -Method Get -uri $uri @webParams }
    catch             {throw "Could not get the thread to post to"; return}
    if (-not $thread) {throw "Could not get the thread to post to"; return}
    if ($Force -or $PSCmdlet.ShouldProcess($thread.topic,"Reply to thread")) {
        $uri = $uri + "/posts/$post/Reply"
        $Settings = @{ Post = @{body= @{content=$Content; contentType=$ContentType}}}
        Invoke-RestMethod -Method Post -Uri $URI  @webparams -Body (ConvertTo-Json $settings) -ContentType "application/json"
    }
}

Function Get-GraphChannel {
    <#
      .Synopsis
        Gets details of a channel, or its Tabs or messages shown in Teams
      .Example
         >get-graphuser -teams | where displayname -eq "Consultants" | get-graphteam -Channels | where displayname -eq "general" |
                 Get-GraphChannel -Tabs |  ft @{n="Name";e={[System.Web.HttpUtility]::UrlDecode($_.displayname)}},@{n="AppName";e={$_.teamsApp.displayName}}
        Gets teams for the current user, and selects the one named "Consultants", gets this team's channels, and selects the one
        named "general"; gets this channels tabs and formats them as a table showing the tab name and application behind it.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Alias("Get-GraphTeamChannel")]
    Param(
        #The channel either as an ID or as a channel object (which may contain the team as a property)
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        $Channel,
        #The ID of the team if it is not in the channel object.
        $Team,
        #If specified gets the channel's Tabs
        [parameter(ParameterSetName="Tabs", Mandatory=$true)]
        [switch]$Tabs,
        #if Specified uses the beta api to get the channel's messages.
        [parameter(ParameterSetName="Msgs", Mandatory=$true)]
        [Alias("Msgs")]
        [switch]$Messages
    )
    process {
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        $webParams = @{Method = "Get"
                       Headers = $Script:DefaultHeader
        }
        foreach ($ch in $channel) {
            if ($ch.Team)      { $Team    = $ch.team }
            elseif  ($Team.ID) { $Team    = $Team.ID      }
            if ($ch.id  )      { $ch = $ch.ID   }
            if (-not ($team -and $ch)) {Write-warning -Message "You need to provide a team ID and a Channel ID"; return}
            elseif ($Messages) {
                Write-progress -Activity 'Getting messages' -CurrentOperation "Reading $($ch.displayname) Messages"
                $result = Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/beta/teams/$Team/channels/$ch/messages"
                $msgList  = $result.value
                while ($result.'@odata.nextLink') {
                    $result = Invoke-RestMethod  @webparams -Uri $result.'@odata.nextLink'
                    $msgList += $result.value
                }
                $userHash = @{}
                Write-progress -Activity 'Getting messages' -CurrentOperation "Expanding User information"
                $msglist.from.user.id | Sort-Object -Unique | foreach-object {
                    $userHash[$_] = ( Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/directoryObjects/$_").displayName
                }
                Write-progress -Activity 'Getting messages' -Completed
                foreach ($msg in $msgList) {
                    $msg.pstypeNames.add("GraphTeamMsg")
                    if ($msg.from.user.id) {
                        Add-Member -InputObject $msg -MemberType NoteProperty -Name FromUserName -Value $userHash[$msg.from.user.id]
                    }
                }
                return $msgList
            }
            elseif ($Tabs)     {
                $results = Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/teams/$Team/channels/$ch/tabs?`$expand=teamsApp"
                $t       = $results.value
                foreach ($tab in $t) {
                    $tab.pstypeNames.add('GraphTab')
                    #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -InputObject $tab -MemberType ScriptProperty -Name teamsAppId   -Value {$this.teamsApp.ID}
                    Add-Member -InputObject $tab -MemberType ScriptProperty -Name teamsAppName -Value {$this.teamsApp.displayName}
                }
                return $t
            }
            else               {
                $result = Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/teams/$Team/channels/$ch"
                $result.pstypeNames.add("GraphChannel")
                Add-Member -InputObject $result -MemberType NoteProperty -Name Team -Value $Team
                return $result
            }
        }
    }
}

Function New-GraphChannel {
    <#
      .Synopsis
        Adds a channel to a team
      .Description
        This requires the Group.ReadWrite.All scope.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias("Add-GraphTeamChannel")]
    Param(
        #The team where the channel will be added, either as an ID or a team object
        [parameter( Mandatory=$true, Position=0)]
        $Team,
        #Display name for the new channel
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Alias("DisplayName")]
        [String]$Name,
        #Description for the new channel
        [String]$Description
    )
    begin  {
        if ($Team.id) {$Team = $Team.id}
        Connect-MSGraph
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if (Get-GraphTeam -Channels $team | Where-Object -property displayname -eq $Name) {
            Write-warning -Message "Channel $Name already exists in this team"
        }
        $webParams = @{Method = "POST"
                    Headers = $Script:DefaultHeader
                    URI    = "https://graph.microsoft.com/v1.0/teams/$Team/channels"
                    ContentType = "application/json"
        }
        $Settings  = @{"displayName" = $Name}
        if ($Description) {$settings["description"] = $Description}
        if ($PSCmdlet.ShouldProcess($Name,"Create channel")) {
            $channel =  Invoke-RestMethod @webParams -body (ConvertTo-Json $settings)
            $channel.psTypenames.add('GraphChannel')
            Add-Member -InputObject $channel -MemberType NoteProperty -Name Team -Value $Team
            return $channel
        }
    }
}

Function Remove-GraphChannel {
    <#
      .Synopsis
        Removes a channel from a team
    #>[cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
    Param(
        #The channel to delete; either as an ID, or a channel object
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        $Team,
        #if Specified the channel will be deleted without prompting
        [switch]$Force
    )
    process {

        if ($Channel.Team) { $Team    = $Channel.team }
        elseif  ($Team.id) { $Team    = $Team.ID}
        if ($Channel.id  ) { $Channel = $Channel.ID   }

        try {
            $c = Get-GraphChannel -Channel $Channel -Team $Team
        }
        Catch {
            throw "Could not get the channel" ; return
        }
        if (-not $c)  {throw "Could not get the channel" ; return }
        if ($force -or $PSCmdlet.ShouldProcess($c.displayname, "Delete Channel")) {
            Connect-MSGraph
            Invoke-RestMethod -Method "Delete" -Headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/teams/$Team/channels/$Channel"
            }
        }
}

Function Add-GraphChannelThread {
    <#
      .Synopsis
        Adds a new thread in a channel in Teams.
    #>

    [cmdletbinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
    Param(
        #The channel to post to either as an ID or a channel object.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #Normally the message is added 'silently'. If passthru is specified, the new message will be returned.
        [Alias('PT')]
        [switch]$Passthru,
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    process {
        if ($Channel.Team) { $Team    = $Channel.team }
        elseif  ($Team.id) { $Team    = $Team.ID}
        if   ($Channel.id) { $Channel = $Channel.ID   }

        try {$c = Get-GraphChannel -Channel $Channel -Team $Team }
        Catch         {throw "Could not get the channel" ; return}
        if (-not $c)  {throw "Could not get the channel" ; return }
        $webParams = @{ 'Method'      = 'POST'
                        'Headers'     = $Script:DefaultHeader
                        'URI'         = "https://graph.microsoft.com/beta/teams/$Team/channels/$channel/chatThreads"
                        'ContentType' = 'application/json'
        }
        $Settings = @{ rootMessage = @{body= @{content=$Content;}}}
        if ($ContentType -eq 'HTML') {$settings.rootMessage.body['contentType'] = 1}
        else                         {$settings.rootMessage.body['contentType'] = 2}
        $json =  (ConvertTo-Json $settings)
        Write-Debug $json
        if ($force -or $PSCmdlet.ShouldProcess("Create Message")) {
            $result = Invoke-RestMethod @webparams  -Body $json
            If ($Passthru) {
                $URI    = "https://graph.microsoft.com/beta/teams/$Team/channels/$channel/Messages/$($Result.id)"
                $msg    = Invoke-RestMethod -Uri $uri -Method Get -Header $Script:DefaultHeader
                $msg.pstypenames.add('GraphTeammsg')
                return $msg
            }

        }
    }
}
# can get replies in a thread , but can't post to a reply. #  https://graph.microsoft.com/beta/teams/{group-id-for-teams}/channels/{channel-id}/messages/{message-id}/replies/{reply-id}
# Doesn't seem to be a delete or a patch ?


Function Add-GraphWikiTab {
    <#
      .Synopsis
        Adds a wiki tab to a channel in teams
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        #An ID or Channel object which may contain the team ID
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab
        $TabLabel = "Wiki",
        #If specified the tab will be added without prompting for confirmation
        [switch]$Force,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        [Alias('PT')]
        [switch]$PassThru
    )
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    if  ($Channel.Team) {$Team    = $Channel.Team }
    elseif   ($Team.id) {$Team    = $Team.id      }
    elseif (-not $team) {throw "Could not determine the team from the channel. Please Specify the team explicitly."}
    if    ($Channel.id) {$Channel = $Channel.id   }

    $webParams = @{'Method'      = 'Post'
                   'Uri'         = "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs"
                   'Headers'     =  $Script:DefaultHeader
                   'ContentType' = 'application/json'
    }
    $json = ConvertTo-Json ([ordered]@{
                    'name'       = $TabLabel
                    'TeamsAppId' = 'com.microsoft.teamspace.tab.wiki'
            })
    Write-Debug $json
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Create wiki tab")) {
        $result = Invoke-RestMethod @webParams -body $json
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'Wiki'
            return $result
        }
    }
}
# Adding tab https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
# https://products.office.com/en-us/microsoft-teams/appDefinitions.xml

