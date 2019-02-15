#see also Get-MsolGroup ; Get-MsolGroupMember ; add-MsolGroupMember; Remove-MsolGroupMember
function Get-GraphGroupList {
    <#
      .Synopsis
        Gets a list of groups
      .Description
        This list of groups returned can be filtered by name (the beginning of the displayname and mail
        address are checked) or with a custom filter string. Or it can be sorted, Or specific fields can be selected
        However there is a limitation in the graph API which prevent these being combined.
      .Example
        >Get-GraphGroupList | format-table -autosize  Displayname, SecurityEnabled, Mailenabled, Mail, ID
        Displays a table of groups in the current tennant
      .Example
        >(Get-GraphGroupList -Name consult | Get-GraphTeam -Site).weburl
        Gets any group whose name begins "Consult" , finds its sharepoint site, and returns the site's URL
    #>
    [cmdletbinding(DefaultparameterSetName="None")]
    param (
        #if specified limits the groups returned to those with names begining...
        [parameter(Mandatory=$true, parameterSetName='FilterByName')]
        [string]$Name,
        #Field(s) to select: ID and displayname are always included;
        #The following are only available when getting a single group:
        #'allowExternalSenders','autoSubscribeNewMembers','isSubscribedByMail' 'unseenCount',
        [ValidateSet( 'assignedLicenses', 'classification', 'createdDateTime', 'description', 'groupTypes',
                    'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'onPremisesLastSyncDateTime',
                    'onPremisesProvisioningErrors', 'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled',
                    'preferredDataLocation', 'proxyAddresses', 'renewedDateTime', 'securityEnabled', 'visibility')]
        [parameter(Mandatory=$true, parameterSetName='SelectFields')]
        [string[]]$Select,
        #An oData order by string
        [parameter(Mandatory=$true, parameterSetName='Sort')]
        [string]$OrderBy,
        #An oData filter string; there is a graph limitation  that you can't filter by description or Visibility.
        [parameter(Mandatory=$true, parameterSetName='FilterByString')]
        [string]$Filter
    )
    #investigate '?$filter=groupTypes/any(c:c+eq+''Unified'')'

    Connect-MSGraph

    $webparams    = @{'Method'  = "Get"
                      'Headers' = $Script:DefaultHeader
    }
    $uri          = 'https://graph.microsoft.com/v1.0/Groups/'

    if     ($Select)  {$uri += '?$select='  + ((@('id','displayName') + $Select) -join ',')}
    elseif ($OrderBy) {$uri += '?$OrderBy=' + $OrderBy}
    elseif ($Filter)  {$uri += '?$Filter='  + $Filter }
    elseif ($Name)    {
        #for once we don't need to fix case.If * is specified , remove it.
        if ($Name -match '\*') {$Name = $Name -replace "\*",""}
        $uri      += ("?`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $Name )
    }

    Write-progress -Activity "Finding Groups"
    $groups = (Invoke-RestMethod @webparams -Uri $uri ).value
    #If Selecting fields, don't set a type to display fields we probably do not have.
    if (-not $Select) {
         foreach ($g in $groups) {$g.pstypenames.Add("GraphGroup") }
    }
    Write-progress -Activity "Finding Groups" -Completed

    $groups
}

function New-GraphGroup {
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
    [cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("New-GraphTeam")]
    param(
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

    $webparams = @{ Headers     = $Script:DefaultHeader  }
    if ( (Invoke-RestMethod -Method Get @webparams -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayname eq '$Name'" ).value) {
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
    if ($Description) {
                  $settings['description'] = $Description
    }
    #if we got owners or users with no ID, fix them at the end, if they have an ID add them now
    if ($Members) {
        $settings['members@odata.bind']= @();
        foreach ($m in $Members) {
            if  ($m.id) {$settings['members@odata.bind'] += "https://graph.microsoft.com/v1.0/users/$($m.id)"}
            else        {$settings['members@odata.bind'] += "https://graph.microsoft.com/v1.0/users/$m"}
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
    $webparams["contentType"] = 'application/json'
    #Don't add URI or body to web params as we are going to make two calls ...
    $uri       = "https://graph.microsoft.com/v1.0/groups"
    $json = ConvertTo-Json $settings
    Write-Debug $json

    if ($Force -or $PSCmdlet.shouldprocess($Name,"Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-RestMethod @webparams -Method Post -uri $uri -body $json
        foreach ($m in $group.members) {if ($m.'@odata.type' -match "user") {$m.pstypenames.add("GraphUser")}}
        if ($NoTeam) {
            $group.pstypenames.Add("GraphGroup")
            Write-Progress -Activity 'Creating Group/Team' -Completed
            return $group
        }
        else {
            $uri = $uri + "/" + $group.id + "/team"
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name"
            $team   = Invoke-RestMethod @webparams -Method Put -uri $uri -Body "{ }"
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

            $team
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

function Set-GraphGroup {
    <#
      .synopsis
        Sets options on a group
      .Description
        Allows or blocks external senders, changes visibility or description and enables the group as a team.
        Other options for a team are set via Set-GraphTeam.
    #>
    [cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
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
    Connect-MSGraph

    if     ($Group.Id)            {$uri = "https://graph.microsoft.com/v1.0/groups/$($Group.ID)"}
    elseif ($Group -is [string])  {$uri = "https://graph.microsoft.com/v1.0/groups/$Group"}
    else   {Write-Warning -Message 'Could not process group paramaeter' ; return}
    $webparams = @{'uri'         = $uri
                   'Headers'     = $Script:DefaultHeader
                   'ContentType' = 'application/json'
    }

    $settings = @{}
    if ($Visibility)        {$settings['visibility']            = $Visibility.ToLower()}
    if ($Description)       {$settings['description']           = $Description}
    if ($PSBoundparameters.ContainsKey('AllowExternalSenders')) {
                             $settings['allowExternalSenders']  = [bool]$AllowExternalSenders
    }
    $json = ConvertTo-Json $settings
    Write-Debug $json
    if (($settings.Count -or $EnableTeam) -and
        ($Force -or $PSCmdlet.Shouldprocess($group.displayname,'Update Group'))) {
        if ($settings.Count) {
                  Invoke-RestMethod @webparams -Method Patch -Body $json | Out-Null
        }
        if ($EnableTeam)     {
            $g  = Invoke-RestMethod @webparams -Method Get
            if ($g.resourceProvisioningOptions -notcontains 'Team') {
                  $webparams['uri'] +=  "/team"
                  Invoke-RestMethod @webparams -Method Put -Body "{ }"   | Out-Null
            }
            elseif ($EnableTeam) {Write-Warning  "Group $($g.displayName) is already team-enabled." }
        }
    }
}

function Remove-GraphGroup {
    <#
      .Synopsis
        Removes a group/team
    #>
    [cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeam")]
    param(
        #The ID of the Group / team
        [parameter(Mandatory=$true, Position=0,ValueFromPipeline=$true )]
        [Alias("Team")]
        $Group,
        #If specified the group will be removed without prompting
        $Force
    )
    begin   {
        Connect-MSGraph
    }
    process {
        if     (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($Group.displayName)  {$displayName = $Group.DisplayName}
        if     ($Group.ID)           {$groupID     = $Group.ID}
        elseif ($Group -is [String]) {$groupID     = $Group   }
        else   {Write-Warning -Message 'Could not process Group parameter.'; return }

        $webparams = @{'Headers' = $Script:DefaultHeader
                       'uri'     = "https://graph.microsoft.com/v1.0/groups/$groupID/"
        }
        if (-not $displayName){
            try   {  $g  = Invoke-RestMethod -Method Get @webparams }
            catch        {throw "Could not get the thread to delete"; return}
            if (-not $g) {throw "Could not get the thread to delete"; return}
            else         {$displayName = $g.displayname}
        }
        if ($PSCmdlet.Shouldprocess($DisplayName,"Delete Group")) {
            Invoke-RestMethod -Method Delete  @webparams
        }
    }
}

# Groups in the recycle bin (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group").value
# DELETE /directory/deletedItems/{id}                permanent delete
# POST /directory/deletedItems/{id}/restore          restore item

function Add-GraphGroupMember {
    <#
      .Synopsis
        Adds a user (or group) to a group/team as either a member or owner.
    #>
    [cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("Add-GraphTeamMember")]
    param   (
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
    begin   {
        Connect-MSGraph
        if     ($Group.ID)           {$groupID  = $Group.ID}
        elseif ($Group -is [String]) {$groupID  = $Group   }
        else   {Write-Warning -Message 'Could not process Group parameter.'; return }

        if ($AsOwner) {$uri   = "https://graph.microsoft.com/v1.0/groups/$groupID/owners/`$ref" }
        else          {$Uri   = "https://graph.microsoft.com/v1.0/groups/$groupID/members/`$ref"}

        $webparams = @{'Method'      = 'Post'
                       'uri'         = $uri
                       'Headers'     = $Script:DefaultHeader
                       'ContentType' = 'application/json'
        }
    }
    process {
        if   (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if   ($Member.id)          {$memberID     = $Member.id}
        else {
            try   {
                $Member     = Get-GraphUser -User $Member
                $memberid   = $Member.id
            }
            catch {throw "Could not get a user matching $Member"; return }
            if (-not $Member) {throw "Could not get a member ID"; return }
        }

        $settings  = @{'@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$memberID"   }
        $json      = ConvertTo-Json $settings
        Write-Debug $json
        if ($Force -or $PSCmdlet.shouldprocess($Member.displayname,"Add to Group")) {
            Invoke-RestMethod @webparams -Body $json
        }
    }
}

function Remove-GraphGroupMember {
    <#
      .Synopsis
        Removes a user (or group) from a group/team
    #>
    [cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeamMember")]
    param   (
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
    process {
        if     (-not $Script:WorkOrSchool) {
                 Write-Warning -Message "This command only works when you are logged in with a work or school account."
                 return
        }
        if     ($Group.id)           {$groupid = $Group.id}
        elseif ($Group -is [string]) {$groupId = $Group }
        else   {Write-Warning -Message "Could not resolve the group parameter."; Return}
        if     ($Member.id) {
                $memberid = $Member.id
                Connect-MSGraph
        }
        else {
            try {
                $Member   = Get-GraphUser -User $Member
                $memberid = $Member.id
            }
            catch {throw "Could not get a user matching $Member";  return}
            if (-not $Memberid) {throw "Could not get a member ID" ; return}
        }

        #https://docs.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0
        $webparams = @{Method      = 'Delete'
                       URI         = "https://graph.microsoft.com/v1.0/groups/$groupid/members/$memberid/`$ref"
                       Headers     =  $Script:DefaultHeader
                       contentType = 'application/json'
                    }
        if ($Force -or $PSCmdlet.Shouldprocess($member.displayName,"Remove from Group")) {
            Invoke-RestMethod @webparams
        }
    }
}

function Get-GraphTeam {
    <#
      .Synopsis
        Gets information about a group and associated office 365 team
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
    [cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphGroup")]
    param   (
        #The name of a team.
        #One more Team IDs or team objects containing and ID. If omitted the current user's teams will be used.
        [parameter(ValueFromPipeline=$true, Position=0)]
        [Alias("ID","Group")]
        $Team ,
        #If specified the Team parameter is treated as a name not an ID
        [Switch]$ByName,
        #If specified returns the teams Apps
        [parameter(parameterSetName='Apps')]
        [switch]$Apps,
        #If specified gets the team's Calendar (a team only has one)
        [parameter(Mandatory=$true, parameterSetName='Calendar')]
        [switch]$Calendar,
        #If specified gets the team's channels
        [parameter(parameterSetName='Channels')]
        [switch]$Channels,
        #If Specified, retrun team's conversations (usually better to use threads)
        [parameter(Mandatory=$true, parameterSetName='Conversations' )]
        [switch]$Conversations,
        #If specified gets the Team's one drive
        [parameter(Mandatory=$true, parameterSetName='Drive')]
        [switch]$Drive,
        #If specified returns the members of the team
        [parameter(Mandatory=$true, parameterSetName='Members')]
        [switch]$Members,
        #If specified returns the Owners of the team
        [parameter(Mandatory=$true, parameterSetName='Owners')]
        [switch]$Owners,
        #If specified returns the team's notebook(s)
        [parameter(Mandatory=$true, parameterSetName='Notebooks')]
        [switch]$Notebooks,
        #if Specified, returns the teams Planners.
        [parameter(Mandatory=$true, parameterSetName='Planners')]
        [switch]$Plans,
        #If Specified, retrun team's threads
        [parameter(Mandatory=$true, parameterSetName='Threads' )]
        [switch]$Threads,
        #if Specified, returns the teams site.
        [parameter(Mandatory=$true, parameterSetName='Site')]
        [switch]$Site,
        #limits searches for appsby name.
        [parameter(parameterSetName='Apps')]
        [String]$AppName,
        #limits searches for channels by name. Other's cant be filtered by name ...  perhaps notebooks can but a group only has one.
        [parameter(parameterSetName='Channels')]
        [String]$ChannelName,
         #Field(s) to select: ID and displayname are always included
        #The following are only available when getting a single group:
        [ValidateSet('allowExternalSenders','autoSubscribeNewMembers','isSubscribedByMail', 'unseenCount',
                     'assignedLicenses', 'classification', 'createdDateTime', 'description', 'groupTypes',
                     'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'onPremisesLastSyncDateTime',
                     'onPremisesProvisioningErrors', 'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled',
                     'preferredDataLocation', 'proxyAddresses', 'renewedDateTime', 'securityEnabled', 'visibility')]
        [parameter(Mandatory=$true, parameterSetName='SelectFields')]
        [string[]]$Select

    )
    begin   {
        Connect-MSGraph
        $webparams = @{Method = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    process {
        if     (-not $Script:WorkOrSchool)          {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($ByName -and $Team -isnot [string]) {Write-Warning 'The team parameter does not look like a name'; return}
        elseif ($ByName)    {$Team = Get-GraphGroupList -Name $Team}
        elseif (-not $Team) {$Team = Get-GraphUser      -Teams }
        if     (-not $Team) {Write-Warning 'Could not Get a team from the parameters provided' ; return}
        foreach ($t in   $Team) {
            if  ($t.id) {$teamid = $t.id}
            else        {$teamid = $t }
            $groupURI = "https://graph.microsoft.com/v1.0/groups/$teamid"
            $teamURI  = "https://graph.microsoft.com/v1.0/teams/$teamid"
            try   {
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
                    Add-Member -InputObject $result -MemberType NoteProperty -Name groupID -Value $teamid
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
                elseif ($Owners)        {
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
                    Write-Progress -Activity 'Getting Group OneNote Notebooks'
                    #if groups can have more than onebook , then add if name ... uri = blah + "?`$expand=sections&`$filter=startswith(tolower(displayname),'$name')"
                    $results = (Invoke-RestMethod  @webparams -Uri ("$groupURI/onenote/notebooks" + '?$expand=sections'  ) )
                    $books   = $results.value
                    foreach ($b in $books) {
                        $b.pstypenames.add("GraphOneNoteBook")
                        #Section fetched this way won't have parentNotebook, so make sure it is available when needed
                        $bookobj =new-object -TypeName psobject -Property @{'id'=$b.id; 'displayname'=$b.displayName; 'Self'=$b.self}
                        foreach ($s in $b.sections) {
                            Add-Member -InputObject $s -MemberType NoteProperty -Name ParentNotebook   -Value $bookobj
                            $s.pstypeNames.add("GraphOneNoteSection")
                        }
                    }
                    Write-Progress -Activity 'Getting Group OneNote Notebooks' -Completed
                    return $books
                }
                elseIf ($Plans)         {
                    Write-Progress -Activity 'Getting Group Planner Plans'
                    $result   = Invoke-RestMethod  @webparams -Uri  "$groupURI/planner/plans" #would like to have expand details here but it only works with a single plan.
                    $planList = $result.value
                    while ($result.'@odata.nextLink') {
                        $result = Invoke-RestMethod  @webparams -Uri $result.'@odata.nextLink'
                        $planList += $result.value
                    }
                    if (-not $planList) {
                        Write-Host "The team $($Team.DisplayName) has not created any plans"
                        return
                    }
                    $dirObjectsHash = @{}
                    if ($t.displayName) {$dirObjectsHash[$teamId] = $t.displayName}
                    @() + $planList.owner + $planList.createdby.user.id  |ForEach-Object  {
                        if (-not $dirObjectsHash[$_]) {
                            $dirObjectsHash[$_] = (Invoke-RestMethod  @webparams -Uri "https://graph.microsoft.com/v1.0/directoryobjects/$_").displayname
                        }
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
                        Add-Member -InputObject $t -MemberType NoteProperty -Name Group -Value $teamid
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
                        Add-Member -InputObject $c -MemberType NoteProperty -Name Group -Value $teamid
                        foreach ($t in $c.threads) {
                            $t.pstypenames.add("GraphThread")
                            Add-Member -InputObject $t -MemberType NoteProperty -Name Group -Value $teamid
                        }
                    }
                    Write-Progress -Activity 'Getting Group Conversations' -Completed
                    return $convList
                }
                elseif ($Channels -or
                        $ChannelName)   {
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
                elseif ($Apps -or
                        $AppName)       {
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
                elseif ($Select)        {
                    $SelectList = (@('id','displayName') + $Select ) -join','
                    Invoke-RestMethod  @webparams -Uri ($groupuri + '?$Select=' + $SelectList)
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
                        $t #<= No return here, because we want to keep loopin
                    }
                    else {
                        $g.pstypenames.Add("GraphGroup")
                        Write-Progress -Activity 'Getting Group/Team information' -Completed
                        $g  #<= No return here, because we want to keep looping
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

function Set-GraphTeam {
    <#
      .Synopsis
        Updates the settings for a team
      .Example
        >Get-GraphTeam -byname accounts | Set-GraphTeam -AllowGiphy:$false
        Gets a the team(s) with a name that begins with accounts, and turns off Giphy content
        Note the use of -SwitchName:$false.


    #>
    [cmdletbinding()]
    param (
        #The team to update either as an ID or a team object with and ID.
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Team ,
        #Allow members to add or remove apps
        [switch]$AllowMemberAddRemoveApps,
        #Allow members to create update or remove connectors
        [switch]$AllowMemberCreateUpdateRemoveConnectors,
        #Allow members to create update or remove Tabs
        [switch]$AllowMemberCreateUpdateRemoveTabs,
        #Allow members to create or update Channels
        [switch]$AllowMemberCreateUpdateChannels,
        #Allow members to delete Channels
        [switch]$AllowMemberDeleteChannels,
        #Allow guests to create or update Channels
        [switch]$AllowGuestCreateUpdateChannels,
        #Allow guests to delete Channels
        [switch]$AllowGuestDeleteChannels,
        #Allow members to edit their own messages
        [switch]$AllowUserEditMessages,
        #Allow members to delete their own messages
        [switch]$AllowUserDeleteMessages,
        #Allow owners to delete mssages
        [switch]$AllowOwnerDeleteMessages,
        #Allow mentions of teams in messages
        [switch]$AllowTeamMentions,
        #Allow mentions of channels in messages
        [switch]$AllowChannelMentions,
        #Allow giphy graphics
        [switch]$AllowGiphy,
        #Rating for giphy graphics; either moderate or strict
        [ValidateSet('moderate', 'strict')]
        [string]$GiphyContentRating,
        #Allow stickers and memes
        [switch]$AllowStickersAndMemes,
        #Allow Custom memes
        [switch]$AllowCustomMemes
    )
    Connect-MSGraph
    $webparams = @{Method      =  'PATCH'
                  ContentType  =  'application/json'
                  Headers      =  $Script:DefaultHeader }

    if     ($Team.id)          {$webparams['Uri'] = "https://graph.microsoft.com/v1.0/teams/$($Team.id)"}
    elseif ($Team -is [string]) {$webparams['Uri'] = "https://graph.microsoft.com/v1.0/teams/$Team"}
    else   {Write-Warning -Message 'Could not resolve the team'; return}

    $settings          = @{}
    $memberSettings    = @{}
    $guestSettings     = @{}
    $messagingSettings = @{}
    $funSettings       = @{}

    if ($PSBoundparameters.ContainsKey('AllowMemberAddRemoveApps'))                {$memberSettings.allowAddRemoveApps                = [Bool]$AllowMemberAddRemoveApps}
    if ($PSBoundparameters.ContainsKey('AllowMemberCreateUpdateChannels'))         {$memberSettings.allowCreateUpdateChannels         = [Bool]$AllowMemberCreateUpdateChannels}
    if ($PSBoundparameters.ContainsKey('AllowMemberCreateUpdateRemoveConnectors')) {$memberSettings.allowCreateUpdateRemoveConnectors = [Bool]$AllowMemberCreateUpdateRemoveConnectors}
    if ($PSBoundparameters.ContainsKey('AllowMemberCreateUpdateRemoveTabs'))       {$memberSettings.allowCreateUpdateRemoveTabs       = [Bool]$AllowMemberCreateUpdateRemoveTabs}
    if ($PSBoundparameters.ContainsKey('AllowMemberDeleteChannels'))               {$memberSettings.allowDeleteChannels               = [Bool]$AllowMemberDeleteChannels}
    if ($PSBoundparameters.ContainsKey('AllowGuestCreateUpdateChannels'))          {$guestSettings.allowCreateUpdateChannels          = [Bool]$AllowGuestCreateUpdateChannels}
    if ($PSBoundparameters.ContainsKey('AllowGuestDeleteChannels'))                {$guestSettings.allowDeleteChannels                = [Bool]$AllowGuestDeleteChannels}
    if ($PSBoundparameters.ContainsKey('AllowUserEditMessages'))                   {$messagingSettings.allowUserEditMessages          = [Bool]$AllowUserEditMessages}
    if ($PSBoundparameters.ContainsKey('AllowUserDeleteMessages'))                 {$messagingSettings.allowUserDeleteMessages        = [Bool]$AllowUserDeleteMessages}
    if ($PSBoundparameters.ContainsKey('AllowOwnerDeleteMessages'))                {$messagingSettings.allowOwnerDeleteMessages       = [Bool]$AllowOwnerDeleteMessages}
    if ($PSBoundparameters.ContainsKey('AllowTeamMentions'))                       {$messagingSettings.allowTeamMentions              = [Bool]$AllowTeamMentions}
    if ($PSBoundparameters.ContainsKey('AllowChannelMentions'))                    {$messagingSettings.allowChannelMentions           = [Bool]$AllowChannelMentions}
    if ($PSBoundparameters.ContainsKey('AllowGiphy'))                              {$funSettings.allowGiphy                           = [Bool]$AllowGiphy}
`   if ($PSBoundparameters.ContainsKey('GiphyContentRating'))                      {$funSettings.giphyContentRating                   = [Bool]$GiphyContentRating}
    if ($PSBoundparameters.ContainsKey('AllowStickersAndMemes'))                   {$funSettings.allowStickersAndMemes                = [Bool]$AllowStickersAndMemes}
    if ($PSBoundparameters.ContainsKey('AllowCustomMemes'))                        {$funSettings.allowCustomMemes                     = [Bool]$AllowCustomMemes}

    if ($memberSettings.Count)    {$settings['memberSettings']    = $memberSettings}
    if ($guestSettings.Count )    {$settings['guestSettings']     = $guestSettings}
    if ($messagingSettings.Count) {$settings['messagingSettings'] = $messagingSettings}
    if ($funSettings.Count)       {$settings['funSettings']       = $funSettings}

    if ($settings.Count) {
        $json = ConvertTo-Json $settings -Depth 10
        Write-Debug $json

        Invoke-RestMethod @webparams -Body $json
    }
    else {Write-Warning -Message "Nothing to set"}
}

function Get-GraphGroupConversation {
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
    param(
        #The Conversation, either as an ID or an object.
        [parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0, ParameterSetName='OneConversation')]
        $Conversation,
        #The group where the conversation is found, either as an ID or as an object, if it can't be found from the conversation
        [parameter(ParameterSetName='AllInTeam', Mandatory=$true )]
        [parameter(ParameterSetName='OneConversation', Position=1)]
        [Alias("Team")]
        $Group,
        #If specified selects the conversation's threads, otherwise an object representing the conversation itself is returned.
        [parameter(ParameterSetName='OneConversation', Position=1)]
        [Switch]$Threads
    )
    begin   {
        Connect-MSGraph
        $webparams = @{Method  = "Get"
                       Headers = $Script:DefaultHeader
        }

    }
    process {
        if ($Group  -and -not $Conversation) {
            Get-GraphGroup -Group $Group -Conversations
            return
        }
        if     (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($Conversation.Group)       {$groupID = $Conversation.Group}
        elseif ($Group.id)                 {$groupId = $Group.ID}
        elseif ($Group -is [String])       {$groupid = $Group}
        else   {Write-Warning 'Could not resolve the group - please specify the group explicitly' ; return}
        if     ($Conversation.id)          {$Conversation = $Conversation.id}

        if ($Threads) {
            $uri    = "https://graph.microsoft.com/v1.0/groups/$groupID/conversations/$conversation/Threads"
            $result = Invoke-RestMethod @webparams -Uri $uri
            foreach ($thread in $result.value) {
                $thread.pstypenames.add("GraphThread")
                Add-Member -InputObject $thread -MemberType NoteProperty -Name Group        -Value $GroupID
                Add-Member -InputObject $thread -MemberType NoteProperty -Name Conversation -Value $Conversation
            }
            return $result.value
        }
        else     {
            $c = (Invoke-RestMethod @webparams -Uri ("https://graph.microsoft.com/v1.0/groups/$groupID/conversations/$conversation"  +'?$expand=Threads'))
            $c.pstypenames.add("GraphConversation")
            Add-Member -PassThru -InputObject $c -MemberType NoteProperty -Name Group -Value $GroupID
        }
    }
}

function Get-GraphGroupThread {
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
    param   (
        #The group thread, either as an ID or as a thread object (which may have the team/group as property)
        [parameter(ParameterSetName='SingleThread', Position=0, ValueFromPipeline=$true, Mandatory=$true)]
        $Thread,
        #The group holding the thread, if it can't be found drm -thread
        [Alias("Team")]
        [parameter(ParameterSetName='AllGroupThreads', Mandatory=$true)]
        [parameter(ParameterSetName='SingleThread', Position=1)]
        $Group,
        #If specified, returns the posts in the thread
        [parameter(ParameterSetName='SingleThread')]
        [Switch]$Posts
    )
    begin   {
        Connect-MSGraph
        $webparams = @{Method  = "Get"
                       Headers = @{Authorization = $Script:AuthHeader ; "Prefer" ='outlook.body-content-type="text"' }
        }
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        If     ($Group -and -not $Thread) {
            Get-GraphGroup -Group $Group -Threads
            return
        }
        if     ($Thread.Group)         {$groupid  = $Thread.group}
        elseif ($Group.id)             {$groupid  = $Group.ID}
        elseif ($Group -is [string])   {$groupid  = $Group}
        else   {Write-Warning -Message 'Could not resolve group ID'; return}

        if     ($Thread.topic)         {$Topic    = $Thread.topic} else {$topic = "-"}
        if     ($Thread.id)            {$threadID = $Thread.id}
        elseif ($Thread -is [string])  {$threadID = $Thread}
        else   {Write-Warning -Message 'Could not resolve thread ID'; return}

        if ($Posts) {
            $results = (Invoke-RestMethod @webparams -Uri "https://graph.microsoft.com/v1.0/groups/$Groupid/Threads/$threadID/posts").value
            foreach  ($post in $results) {
                $Post.pstypenames.add("GraphPost")
                Add-Member -InputObject $post -MemberType NoteProperty -Name Group  -Value $Groupid
                Add-Member -InputObject $post -MemberType NoteProperty -Name Thread -Value $threadID
                Add-Member -InputObject $post -MemberType NoteProperty -Name Topic  -Value $Topic
            }
            return $results
        }
        else        {
            $t = (Invoke-RestMethod @webparams -Uri "https://graph.microsoft.com/v1.0/groups/$Groupid/Threads/$threadid")
            $t.pstypenames.Add("GraphThread")
            Add-Member -PassThru -InputObject $t -MemberType NoteProperty -Name Group -Value $Groupid
        }
    }
}

function Add-GraphGroupThread {
    <#
      .Synopsis
        Starts a new thread in a group in outlook.
    #>
    [cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The group where the thread will be added
        [parameter(Mandatory=$true,Position=0)]
        [Alias("Team")]
        $Group,
        #The subkect line for the thread
        [parameter(Mandatory=$true, Position=1)]
        [Alias("Subject")]
        $ThreadTopic,
        #The Message body - text by default, specify -contentType if using HTML
        [parameter(Mandatory=$true, Position=2)]
        [String]$Content,
        #The content type, (Text by default) or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting; this is the default, unless $confirm preference has been changed
        [switch]$Force,
        #if Specified an object containing the Thread ID will be returned
        [switch]$PassThru
    )

    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

    if     ($Group.ID)           {$groupID  = $Group.ID}
    elseif ($Group -is [String]) {$groupID  = $Group   }
    else   {Write-Warning -Message 'Could not process Group parameter.'; return }

    $webparams = @{ 'uri'         = "https://graph.microsoft.com/v1.0/groups/$groupID/threads/"
                    'method'      = 'Post'
                    'contentType' = 'application/json'
                    'Headers'     = $Script:DefaultHeader
    }
    $Settings  = @{ 'topic'       = $ThreadTopic
                    'posts'       = @( @{body= @{'content'     = $Content
                                                 'contentType' = $ContentType}})
    }
    $json      = ConvertTo-Json $settings -Depth 5

    if ($force -or $PSCmdlet.Shouldprocess($ThreadTopic,"Create New thread")) {
        $t = Invoke-RestMethod  @webparams -Body $json
        if ($PassThru) {return $t}
    }
}

function Remove-GraphGroupThread {
    <#
      .Synopsis
        Removes a thread from a group in outlook
    #>
    [cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param (
        #The thread to remove, either as an ID or a thread object containing an ID, and possibly a group ID
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Thread,
        #The group from which the thread is to be removed, either as an ID or a group object containing an ID
        [Alias("Team")]
        $Group,
        #if Specified the thread will be deleted without prompting.
        [switch]$Force
    )
    process {
        Connect-MSGraph
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

        if     ($Thread.group)        {$groupid  = $Thread.group}
        elseif ($Group.ID)            {$groupid  = $Group.ID}
        elseif ($Group -is [string])  {$groupid  = $Group}
        else   {Write-Warning 'Could not resolve the group ID' ; return}

        if     ($Thread.ID)           {$threadid = $Thread.id  }
        elseif ($Thread -is [string]) {$threadid = $Thread.id  }
        else   {Write-Warning 'Could not resolve the group ID' ; return}


        $webparams = @{'Headers' = $Script:DefaultHeader
                       'uri'    =  "https://graph.microsoft.com/v1.0/groups/$GroupID/threads/$threadID"
        }
        Write-Progress -Activity "Deleting thread" -Status "Checking existing thread"
        try   {$thread  = Invoke-RestMethod -Method Get @webparams }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-warning 'Thread not found, it may have been deleted already'
                return
            }
            else {
                throw $_ ;
                return
            }
        }
        if (-not $thread) {throw "Could not get the thread to delete"; return}
        Write-Progress -Activity "Deleting thread" -Completed
        if ($Force -or $PSCmdlet.Shouldprocess($thread.topic,"Delete thread")) {
            Write-Progress -Activity "Deleting thread" -Status "Sending delete instruction"
            Invoke-RestMethod -Method Delete  @webparams
            Write-Progress -Activity "Deleting thread" -Completed
        }
    }
}

function Send-GraphGroupReply {
    <#
      .Synopsis
        Replies to a group's post in outlook.
    #>
    [cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The Post being replied to, either as an ID or a post object containing an ID which may identify the thread and group
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Post,
        #The tread containing the post (if not embedded in the post itself), as an ID or object, which may identify the group
        $Thread,
        #The group containing the thread (if not embedded in the Post or thread) as an ID or object
        [Alias("Team")]
        $Group,
        #The Message body - text by default, specify -contentType if using HTML
        [parameter(Mandatory=$true)]
        [String]$Content,
        #The type of content, text by default or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }


    if     ($Post.Group)          {$groupID  = $Post.group}
    elseif ($Thread.Group)        {$groupID  = $Thread.group}
    elseif ($Group.ID)            {$groupID  = $Group.ID}
    elseif ($Group -is [string])  {$groupID  = $Group}
    else   {Write-warning -Message 'Could not resolve the group ID.' ; return}

    if     ($Post.Thread)         {$threadID = $Post.Thread}
    elseif ($Thread.ID)           {$threadID = $Thread.id  }
    elseif ($Thread -is [String]) {$threadID = $Thread.id  }
    else   {Write-warning -Message 'Could not resolve the Thread ID.' ; return}

    if     ($Post.ID)             {$PostID   = $Post.ID}
    elseif ($Post -is [String])   {$PostID = $Post  }
    else   {Write-warning -Message 'Could not resolve the Post ID.' ; return}

    if (-not ($PostID -and $threadID -and $groupID)) {throw "Could not find Group, Thread and Post IDs from supplied parameters."; Return}

    $webparams = @{Headers = $Script:DefaultHeader }
    $uri       =  "https://graph.microsoft.com/v1.0/groups/$groupID/threads/$threadID/posts/$postid"
    Write-Progress -Activity 'Posting reply to thread' -Status 'Checking parent message'
    try   {$thread  = Invoke-RestMethod -Method Get -uri $uri @webparams }
    catch             {throw "Could not get the post to reply to"; return}
    if (-not $thread) {throw "Could not get the Post to reply to"; return}
    Write-Progress -Activity 'Posting reply to thread' -Completed

    $Settings  = @{ Post = @{body= @{content=$Content; contentType=$ContentType}}}
    $Json      = ConvertTo-Json $settings
    Write-Debug $Json

    if ($Force -or $PSCmdlet.Shouldprocess($thread.topic,"Reply to thread")) {
        $uri     += "/Reply"
        Write-Progress -Activity 'Posting reply to thread' -Status 'sending reply'
        Invoke-RestMethod -Method Post -Uri $URI  @webparams -Body $Json -ContentType "application/json"
        Write-Progress -Activity 'Posting reply to thread' -Completed
    }
}

function Get-GraphChannel {
    <#
      .Synopsis
        Gets details of a channel, or its Tabs or messages shown in Teams
      .Example
        >Get-GraphGroup -ByName consultants -ChannelName general | Get-GraphChannel -Tabs
        Gets channels for the team(s) with a name beginning 'Consultants' and selects channel(s)
        with a name beginning "general"; then gets the tabs shown in Teams for this channel
      .Example
        >Get-GraphGroup -ByName consultants -ChannelName general | Get-GraphChannel -Messages
        This followes the same method for getting the Teams but this time returns messaes in the channel
       .Example
        >
        >$chan = Get-GraphGroup -ByName consultants -ChannelName general
        >Get-GraphChannel -Messages
        >
        This followes the same method for getting the Teams but this time returns messaes in the channel
    #>
    [cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeamChannel")]
    param(
        #The channel either as an ID or as a channel object (which may contain the team as a property)
        [parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        $Channel,
        #The ID of the team if it is not in the channel object.
        $Team,
        #If specified gets the channel's Tabs
        [parameter(parameterSetName="Tabs", Mandatory=$true)]
        [switch]$Tabs,
        #if Specified uses the beta api to get the channel's messages.
        [parameter(parameterSetName="Msgs")]
        [Alias("Msgs")]
        [switch]$Messages,
        #If specified, returns the top n messages, otherwise the command will attempt to get all messages. The server may return more than the specified number.
        [parameter(parameterSetName="Msgs")]
        $Top
    )
    begin   {
        Connect-MSGraph
        $webparams = @{Method = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        foreach ($ch in $channel) {
            if     ($ch.Team)           {$teamID    = $ch.team }
            elseif ($Team.ID)           {$teamID    = $Team.ID }
            elseif ($Team -is [string]) {$teamID    = $Team    }
            else   {Write-Warning -Message 'Could not resolve the team for this channel'; return}
            if     ($ch.id  )           {$channelID = $ch.ID   }
            elseif ($ch -is [string])   {$channelID = $ch      }
            else   {Write-Warning -Message 'Could not resolve the channel ID'; return}
            if (-not ($teamid -and $channelID)) {Write-warning -Message "You need to provide a team ID and a Channel ID"; return}
            elseif ($Messages -or $Top) {
                Write-progress -Activity 'Getting messages' -Status "Reading $($ch.displayname) Messages"
                $uri      =  "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/messages"
                if ($Top) {$uri += '?$top=' + $Top }
                $result   = Invoke-RestMethod @webparams -Uri $uri
                $msgList  = @() + $result.value
                while ($result.'@odata.nextLink' -and $result.'@odata.count' -gt 0 ) {
                    Write-Verbose  $result.'@odata.count'
                    Write-progress -Activity 'Getting messages' -Status "Reading $($ch.displayname) Messages" -CurrentOperation "$($msglist.count) so far"
                    $result   = Invoke-RestMethod  @webparams -Uri $result.'@odata.nextLink'
                    $msgList += $result.value
                }
                $userHash = @{}
                Write-progress -Activity 'Getting messages' -Status "Expanding User information"
                $msglist.from.user.id | Sort-Object -Unique | foreach-object {
                    $userHash[$_] = ( Invoke-RestMethod @webparams -Uri  "https://graph.microsoft.com/v1.0/directoryObjects/$_").displayName
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
                $results = Invoke-RestMethod @webparams -Uri  "https://graph.microsoft.com/v1.0/teams/$teamID/channels/$channelID/tabs?`$expand=teamsApp"
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
                $result = Invoke-RestMethod @webparams -Uri  "https://graph.microsoft.com/v1.0/teams/$teamID/channels/$channelId"
                $result.pstypeNames.add("GraphChannel")
                Add-Member -InputObject $result -MemberType NoteProperty -Name Team -Value $teamID

                $result
            }
        }
    }
}

function New-GraphChannel {
    <#
      .Synopsis
        Adds a channel to a team
      .Description
        This requires the Group.ReadWrite.All scope.
    #>
    [cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("Add-GraphTeamChannel")]
    param(
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
        $webparams = @{Method = "POST"
                    Headers = $Script:DefaultHeader
                    URI    = "https://graph.microsoft.com/v1.0/teams/$Team/channels"
                    ContentType = "application/json"
        }
        $Settings  = @{"displayName" = $Name}
        if ($Description) {$settings["description"] = $Description}
        if ($PSCmdlet.Shouldprocess($Name,"Create channel")) {
            $channel =  Invoke-RestMethod @webparams -body (ConvertTo-Json $settings)
            $channel.psTypenames.add('GraphChannel')
            Add-Member -InputObject $channel -MemberType NoteProperty -Name Team -Value $Team

            $channel
        }
    }
}

function Remove-GraphChannel {
    <#
      .Synopsis
        Removes a channel from a team
    #>[cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param(
        #The channel to delete; either as an ID, or a channel object
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
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
        if ($force -or $PSCmdlet.Shouldprocess($c.displayname, "Delete Channel")) {
            Connect-MSGraph
            Invoke-RestMethod -Method "Delete" -Headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/teams/$Team/channels/$Channel"
            }
        }
}

function Add-GraphChannelThread {
    <#
      .Synopsis
        Adds a new thread in a channel in Teams.
    #>
    [cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param(
        #The channel to post to either as an ID or a channel object.
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [parameter(Mandatory=$true)]
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
        if     ($Channel.Team)      {$teamID    = $Channel.team }
        elseif ($Team.id)           {$teamID    = $Team.ID}
        elseif ($Team -is [string]) {$teamID    = $Team}
        else   {Write-Warning -Message 'Could not determine the team ID'; return}

        if     ($Channel.id)           {$channelID = $Channel.ID   }
        elseif ($Channel -is [string]) {$channelID = $channel  }
        else   {Write-Warning -Message 'Could not determine the channel ID'; return}

        try {$c = Get-GraphChannel -Channel $channelID -Team $teamID }
        Catch         {throw "Could not get the channel" ; return}
        if (-not $c)  {throw "Could not get the channel" ; return }
        $webparams = @{ 'Method'      = 'POST'
                        'Headers'     = $Script:DefaultHeader
                        'URI'         = "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/chatThreads"
                        'ContentType' = 'application/json'
        }
        $Settings = @{ rootMessage = @{body= @{content=$Content;}}}
        if ($ContentType -eq 'HTML') {$settings.rootMessage.body['contentType'] = 1}
        else                         {$settings.rootMessage.body['contentType'] = 2}
        $json =  (ConvertTo-Json $settings)
        Write-Debug $json
        if ($force -or $PSCmdlet.Shouldprocess("Create Message")) {
            $result = Invoke-RestMethod @webparams  -Body $json
            If ($Passthru) {
                $URI    = "https://graph.microsoft.com/beta/teams/$teamid/channels/$channelid/Messages/$($result.id)"
                $msg    = Invoke-RestMethod -Uri $uri -Method Get -Header $Script:DefaultHeader
                $msg.pstypenames.add('GraphTeammsg')

                $msg
            }
        }
    }
}
# can get replies in a thread , but can't post to a reply. #  https://graph.microsoft.com/beta/teams/{group-id-for-teams}/channels/{channel-id}/messages/{message-id}/replies/{reply-id}
# Doesn't seem to be a delete or a patch ?

function Add-GraphWikiTab {
    <#
      .Synopsis
        Adds a wiki tab to a channel in teams
    #>
    [CmdletBinding(SupportsShouldprocess=$true)]
    param(
        #An ID or Channel object which may contain the team ID
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
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
    if     ($Channel.Team)         {$teamID  = $Channel.Team }
    elseif ($Team.id)              {$teamID  = $Team.id      }
    elseif ($Team -is [String])    {$teamID  = $Team}
    else   {Write-Warning -Message 'Could not determine the team from the channel. Please Specify the team explicitly.'; return}
    if     ($Channel.id)           {$channelID = $Channel.id }
    elseif ($Channel -is [string]) {$channelID = $Channel    }
    else   {Write-Warning -Message 'Could not determine the channel ID.'; return}
    $webparams = @{'Method'      = 'Post'
                   'Uri'         = "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/tabs"
                   'Headers'     =  $Script:DefaultHeader
                   'ContentType' = 'application/json'
    }
    $json = ConvertTo-Json ([ordered]@{
                    'name'       = $TabLabel
                    'TeamsAppId' = 'com.microsoft.teamspace.tab.wiki'
            })
    Write-Debug $json
    if ($Force -or $PSCmdlet.Shouldprocess($TabLabel,"Create wiki tab")) {
        $result = Invoke-RestMethod @webparams -body $json
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'Wiki'

            $result
        }
    }
}
# Adding tab https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
# https://products.office.com/en-us/microsoft-teams/appDefinitions.xml
