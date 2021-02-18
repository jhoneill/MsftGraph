using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Globalization

$Script:GraphUri  = "https://graph.microsoft.com/v1.0"
$Script:GUIDRegex = "^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$"


function Get-GraphGroupList         {
    <#
      .Synopsis
        Gets a list of Groups in Microsoft Graph.
      .Description
        The list of groups returned can be filtered by name (the beginning of the displayname and mail
        address are checked) or with a custom filter string, or it can be sorted, or specific fields can
        be selected. However there is a limitation in the graph API which prevent these being combined.
        Requires consent to use the Group.Read.All scope.
      .Example
        >Get-GraphGroupList | Format-Table -Autosize  Displayname, SecurityEnabled, Mailenabled, Mail, ID
        Displays a table of groups in the current tennant
      .Example
        >(Get-GraphGroupList -Name consult | Get-GraphTeam -Site).weburl
        Gets any group whose name begins "Consult" , finds its sharepoint site, and returns the site's URL
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup])]
    param (
        #if specified limits the groups returned to those with names begining...
        [Parameter(Mandatory=$true, parameterSetName='FilterByName', Position=1)]
        [string]$Name,
        #Field(s) to select: ID and displayname are always included;
        #The following are only available when getting a single group:
        #'allowExternalSenders','autoSubscribeNewMembers','isSubscribedByMail' 'unseenCount',
        [ValidateSet( 'acceptedSenders', 'allowExternalSenders', 'appRoleAssignments', 'assignedLabels', 'assignedLicenses',
                'autoSubscribeNewMembers', 'calendar', 'calendarView', 'classification', 'conversations', 'createdDateTime',
                'createdOnBehalfOf', 'deletedDateTime', 'description', 'displayName', 'drive', 'drives', 'events',
                'expirationDateTime', 'extensions', 'groupLifecyclePolicies', 'groupTypes', 'hasMembersWithLicenseErrors',
                'hideFromAddressLists', 'hideFromOutlookClients', 'id', 'isArchived', 'isSubscribedByMail',
                'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'memberOf', 'members', 'membershipRule',
                'membershipRuleProcessingState', 'membersWithLicenseErrors', 'onenote', 'onPremisesDomainName',
                'onPremisesLastSyncDateTime', 'onPremisesNetBiosName', 'onPremisesProvisioningErrors',
                'onPremisesSamAccountName', 'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'owners',
                'permissionGrants', 'photo', 'photos', 'planner', 'preferredDataLocation', 'preferredLanguage',
                'proxyAddresses', 'rejectedSenders', 'renewedDateTime', 'securityEnabled', 'securityIdentifier', 'settings',
                'sites', 'team', 'theme', 'threads', 'transitiveMemberOf', 'transitiveMembers', 'unseenCount', 'visibility')]
        [Parameter(Mandatory=$true, parameterSetName='SelectFields')]
        [string[]]$Select,
        #An oData order by string
        [Parameter(Mandatory=$true, parameterSetName='OrderBy')]
        [ValidateSet('allowExternalSenders', 'assignedLabels', 'assignedLicenses', 'autoSubscribeNewMembers', 'classification',
                'createdDateTime', 'deletedDateTime', 'description', 'displayName', 'expirationDateTime', 'groupTypes',
                'hasMembersWithLicenseErrors', 'hideFromAddressLists', 'hideFromOutlookClients', 'id', 'isArchived',
                'isSubscribedByMail', 'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'membershipRule',
                'membershipRuleProcessingState', 'onPremisesDomainName', 'onPremisesLastSyncDateTime', 'onPremisesNetBiosName',
                'onPremisesProvisioningErrors', 'onPremisesSamAccountName', 'onPremisesSecurityIdentifier',
                'onPremisesSyncEnabled', 'preferredDataLocation', 'preferredLanguage', 'proxyAddresses', 'renewedDateTime',
                'securityEnabled', 'securityIdentifier', 'theme', 'unseenCount', 'visibility')]
        [string]$OrderBy,

        [Parameter(parameterSetName='Sort')]
        [Switch]$Descending,
        #An oData filter string; there is a graph limitation  that you can't filter by description or Visibility.
        [Parameter(Mandatory=$true, parameterSetName='FilterByString')]
        [string]$Filter
    )
    process {
        #xxxx to do: investigate "groupTypes/any(c:  c eq 'Unified')"  -filter "groupTypes/any(x: x eq 'DynamicMembership')"
        # check access to scopes  Group.Read.All

        if     ($Select)  {
            if ("id" -notin $select)          {$select += 'id'}
            if ("displayName" -notin $select) {$select += 'displayName'}
            $uri  =  $GraphUri +  '/Groups/?$select='  + ($Select -join ',')
        }
        elseif ($Filter)  {$uri =  $GraphUri +  '/Groups/?$Filter='  + $Filter }
        elseif ($Name)    {
            #for once we don't need to fix case.If * is specified , remove it.
            if ($Name -match '\*') {$Name = $Name -replace "\*",""}
                $uri = ( $GraphUri +  "/Groups/?&`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $Name)
        }
        else   {$uri =  $GraphUri +  '/Groups/?$OrderBy=displayname' }
        Write-Progress -Activity "Finding Groups"
        Invoke-GraphRequest -Uri $uri -AllValues -ExcludeProperty 'creationOptions' -AsType ([MicrosoftGraphGroup])
        Write-Progress -Activity "Finding Groups" -Completed
    }
}

function Get-GraphGroup             {
    <#
      .Synopsis
        Gets information about a Group and any associated Office 365 Team
      .Description
        Takes a Group/Team ID or object as a parameter and gets information about it.
        Apps, Calendar, Channels, Drive, Members or Planners can be requested.
        Depending on which aspect of the group are queried, may need access to the following
        Scopes Group.Read.All, Files.Read, Sites.Read.All, Notes.Create, Notes.Read,
      .Example
        >Get-GraphUser -teams | Get-GraphTeam -Plans | select -last 1 | Get-GraphPlan -FullTasks  | ft PlanTitle,Bucketname,Title,DueDateTime,PercentComplete,Assignees
         Gets the current user's Teams, and gets the plans for each;
         Note that because we are refering to "Teams" the command is calling using its alias
         of Get-GraphTeam. The last plan is selected and details of the plan are fetched,
         showing the result as a table.
      .Example
        >(Get-GraphGroup -Site).lists | where name -match document
        If no Group/Team is provided the command gets those associated with the
        current user; it this case it returns their associated site(s).
        Site objects include a lists property, which holds a collection of lists
        this command will fiter the lists down to those where name matches "document"
      .Example
        >(Get-GraphGroup -Drive).root.children.where({$_.folder}) | Select  name, weburl, id,@{n="drive";e={$_.parentReference.driveId}}
        As with the previous example gets this command gets Groups/Teams for current user,
        in this case the command returns their associated drive(s)
        Drive objects include a root property, which holds an object describing the root folder;
        this in turn has a children property which contains files and folder objects in the root folder.
        This example filters the children collection to folders and returns their name,
        WebURl and the item ID and Drive ID needed to access them
      .Example
        >Get-GraphGroup -Notebooks | select -ExpandProperty sections | where "Displayname" -eq "General_Notes"
        Again gets Groups/Teams for the current user and returns their associated notebooks(s)
        Notebook objects include a Sections property, which holds a collection of OneNote sections in the notebook;
        This command gets returns any section in a team notebook which has the name "General_Notes"
      .Example
        > Get-GraphTeam -threads | where LastDeliveredDateTime -gt [datetime]::Now.AddDays(-7)
        Gets the teams conversation threads which have been updated in the last 7 days.
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeam")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',  Justification='Write-warning could be used, but the is informational non-output.')]
    param   (
        #The name of a team.
        #One more Team IDs or team objects containing and ID. If omitted the current user's teams will be used.
        [Parameter(ValueFromPipeline=$true, Position=1)]
        [Alias("Team","Group")]
        [ArgumentCompleter([GroupCompleter])]
        $ID ,
        #If specified returns the teams Apps
        [Parameter(parameterSetName='Apps')]
        [switch]$Apps,
        #If specified gets the team's Calendar (a team only has one)
        [Parameter(Mandatory=$true, parameterSetName='Calendar')]
        [switch]$Calendar,
        #If specified gets the team's channels
        [Parameter(parameterSetName='Channels')]
        [switch]$Channels,
        #If Specified, retrun team's conversations (usually better to use threads)
        [Parameter(Mandatory=$true, parameterSetName='Conversations' )]
        [switch]$Conversations,
        #If specified gets the Team's one drive
        [Parameter(Mandatory=$true, parameterSetName='Drive')]
        [switch]$Drive,
        #If specified returns the members of the team
        [Parameter(Mandatory=$true, parameterSetName='Members')]
        [switch]$Members,
        #If specified returns the Owners of the team
        [Parameter(Mandatory=$true, parameterSetName='Owners')]
        [switch]$Owners,
        #If specified returns the team's notebook(s)
        [Parameter(Mandatory=$true, parameterSetName='Notebooks')]
        [switch]$Notebooks,
        #if Specified, returns the teams Planners.
        [Parameter(Mandatory=$true, parameterSetName='Planners')]
        [switch]$Plans,
        #If Specified, retrun team's threads
        [Parameter(Mandatory=$true, parameterSetName='Threads' )]
        [switch]$Threads,
        #if Specified, returns the teams site.
        [Parameter(Mandatory=$true, parameterSetName='Site')]
        [switch]$Site,
        #limits searches for appsby name.
        [Parameter(parameterSetName='Apps')]
        [String]$AppName,
        #limits searches for channels by name. Other's cant be filtered by name ...  perhaps notebooks can but a group only has one.
        [Parameter(parameterSetName='Channels')]
        [String]$ChannelName,
         #Field(s) to select: ID and displayname are always included
        #The following are available when getting a single group:
        [ValidateSet('acceptedSenders', 'allowExternalSenders', 'appRoleAssignments', 'assignedLabels', 'assignedLicenses',
                'autoSubscribeNewMembers', 'calendar', 'calendarView', 'classification', 'conversations', 'createdDateTime',
                'createdOnBehalfOf', 'deletedDateTime', 'description', 'displayName', 'drive', 'drives', 'events',
                'expirationDateTime', 'extensions', 'groupLifecyclePolicies', 'groupTypes', 'hasMembersWithLicenseErrors',
                'hideFromAddressLists', 'hideFromOutlookClients', 'id', 'isArchived', 'isSubscribedByMail',
                'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'memberOf', 'members', 'membershipRule',
                'membershipRuleProcessingState', 'membersWithLicenseErrors', 'onenote', 'onPremisesDomainName',
                'onPremisesLastSyncDateTime', 'onPremisesNetBiosName', 'onPremisesProvisioningErrors',
                'onPremisesSamAccountName', 'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'owners',
                'permissionGrants', 'photo', 'photos', 'planner', 'preferredDataLocation', 'preferredLanguage', 'proxyAddresses',
                'rejectedSenders', 'renewedDateTime', 'securityEnabled', 'securityIdentifier', 'settings', 'sites', 'team',
                'theme', 'threads', 'transitiveMemberOf', 'transitiveMembers', 'unseenCount', 'visibility')]
        [Parameter(Mandatory=$true, parameterSetName='SelectFields')]
        [string[]]$Select,
        [Parameter(Mandatory=$true, parameterSetName='BareGroups')]
        [switch]$NoTeamInfo
    )
    begin   {
        $usersAndGroups = @()
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx toDo check scopes - Scopes Group.Read.All, Files.Read, Sites.Read.All, Notes.Create, Notes.Read, depending on params passed.
        # if we didn't get passed a group but something about a group or groups was wanted, get the current user's groups,
        # if we got a single string that looks like a name (not a GUID) resolve it.
        # If we got nothing return the list, We'll loop through an array and (or single object) with either GUIDs or objects.
        if      ($PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters}) -and -not $ID) {
                       $ID = Get-GraphUser -Current -MemberOf
        }
        elseif  ($ID -is [string] -and  $ID -notmatch $guidregex)   {
                       $ID = Get-GraphGroupList -Name $id
        }
        elseif  (-not  $ID) {Get-GraphGroupList ; return }

        foreach ($i in $ID) {
            <# not all teams have team set in resource procisioning options
                if  ($i.ResourceProvisioningOptions -is [array] -and
                 $i.ResourceProvisioningOptions -notcontains "Team" -and
                ($Channels -or $ChannelName -or $Apps)) {
                Write-Verbose "$($i.DisplayName) is a group but not a team"
                continue
            }#>
            if  ($i -is [string] -and  $ID -notmatch $guidregex)   {$i = Get-GraphGroupList -Name $i}
            if  ($i.DisplayName)  {$displayname       = $i.DisplayName}
            else                  {$displayname       = $i            }
            if  ($i.id)           {$groupid = $teamid = $i.id         }
            else                  {$groupid = $teamid = $i            }
            $groupURI = "$GraphURI/groups/$groupid"
            $teamURI  = "$GraphURI/teams/$teamid"
            try   {
                #For each of the switches get the data from /groups{id}/whatever or /teams/{id}.whatever
                #Add a type to PS Type names so we can format it, and add any properties we expect to want later.
                Write-Progress -Activity 'Getting Group Information'
                if     ($Site)          {
                    $uri = ("$groupURI/sites/root?expand=drives,sites,lists(expand=columns,contenttypes,drive)")
                    $result  =  Invoke-GraphRequest -Uri $uri -ExcludeProperty 'sites@odata.context', '@odata.context', 'drives@odata.context', 'lists@odata.context' -AsType ([MicrosoftGraphSite]) |
                            Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    foreach ($siteObj in $result) {
                        foreach ($l in $siteObj.lists) {
                            Add-Member -InputObject $l -NotePropertyName SiteID    -NotePropertyValue  $r.id
                            Add-Member -InputObject $l -NotePropertyName ParentUrl -NotePropertyValue  $r.weburl
                            Add-Member -InputObject $l -MemberType ScriptProperty -Name Template -Value {$this.list.template}
                        }
                        $siteobj
                    }
                    continue
                }
                elseif ($Calendar)      {
                    Invoke-GraphRequest -Uri  "$groupURI/calendar" -ExcludeProperty "@odata.context" -AsType ([MicrosoftGraphCalendar]) |
                        Add-Member -PassThru -NotePropertyName GroupID      -NotePropertyValue $groupid   |
                        Add-Member -PassThru -NotePropertyName CalendarPath -NotePropertyValue "groups/$groupid/Calendar" |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    continue
                }
                elseif ($Drive)         {
                    $uri = ("$groupURI/drive" + '?$expand=root($expand=children)' )
                    Invoke-GraphRequest  -Uri $uri -ExcludeProperty "@odata.context", "root@odata.context" -AsType ([MicrosoftGraphDrive]) |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    continue
                }
                elseif ($Members)       { #can do group ?$expand=Memebers, the others don't expand
                    $usersAndGroups += Invoke-GraphRequest  -Uri  "$groupURI/members"  -AllValues |
                        ForEach-Object {$_['GroupName'] =  $displayname ; $_ }
                }
                elseif ($Owners)        {
                    $usersqAndGroups +=  Invoke-GraphRequest  -Uri  "$groupURI/Owners" -AllValues|
                        ForEach-Object {$_['GroupName'] =  $displayname ; $_ }
                }
                elseif ($Notebooks)     {
                    #if groups can have more than one book , then add if name ... uri = blah + "?`$expand=sections&`$filter=startswith(tolower(displayname),'$name')"
                    $uri = $groupURI + '/onenote/notebooks?$expand=sections'
                    $result = Invoke-GraphRequest  -Uri $uri -ValueOnly -ExcludeProperty 'sections@odata.context'  -AsType ([MicrosoftGraphNotebook]) |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    foreach ($bookobj in $result) {
                        #Section fetched this way won't have parentNotebook, so make sure it is available when needed
                        foreach ($s in $bookobj.sections) {
                             $s.ParentNotebook.id          = $r.id
                             $s.ParentNotebook.displayname = $r.displayName;
                             $s.ParentNotebook.Self        = $r.self
                        }
                        $bookobj
                    }
                    continue
                }
                elseIf ($Plans)         {
                    #would like to have expand details here but it only works with a single plan.
                    $result  = Invoke-GraphRequest  -Uri  "$groupURI/planner/plans"  -AllValues -ExcludeProperty  "@odata.etag" -AsType ([MicrosoftGraphPlannerPlan]) |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    if (-not $result) { Write-Host "The team $($ID.DisplayName) has not created any plans" ;   continue}
                    $dirObjectsHash = @{}
                    if ($i.displayName) {$dirObjectsHash[$teamId] = $i.displayName}
                    @() + $result.owner + $result.createdby.user.id  |ForEach-Object  {
                        if (-not $dirObjectsHash[$_]) {
                            $dirObjectsHash[$_] = (Invoke-GraphRequest  -Uri "$GraphURI/directoryobjects/$_").displayname
                        }
                    }
                    foreach ($r in $result) {
                        Add-Member -PassThru  -InputObject $r  -NotePropertyName OwnerName   -NotePropertyValue $dirObjectsHash[$r.owner] |
                        Add-Member -PassThru                   -NotePropertyName CreatorName -NotePropertyValue $dirObjectsHash[$r.createdBy.user.id]
                    }
                }
                elseif ($Threads)       {
                    Invoke-GraphRequest  -Uri  "$groupURI/threads"  -AllValues -AsType ([MicrosoftGraphConversationThread]) |
                        Add-Member -PassThru -NotePropertyName Group       -NotePropertyValue $groupid  |
                        Add-Member -PassThru -NotePropertyName GroupName   -NotePropertyValue $displayname
                }
                elseif ($Conversations) {
                     $result = Invoke-GraphRequest  -Uri ($groupURI + '/conversations?$expand=Threads')  -AllValues -As ([MicrosoftGraphConversation]) |
                        Add-Member -PassThru -NotePropertyName Group        -NotePropertyValue $groupid  |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    foreach ($convobj in $result) {
                        foreach ($t in $convObj.threads) {
                            Add-Member -InputObject $t -NotePropertyName Group      -NotePropertyValue $groupid  |
                            Add-Member -InputObject $t -NotePropertyName GroupName  -NotePropertyValue $displayname
                        }
                        $convObj
                    }
                    continue
                }
                elseif ($Channels -or
                        $ChannelName)   {
                    if ($ChannelName)   { $uri =  "$teamURI/channels?`$filter=startswith(tolower(displayname), '$($ChannelName.ToLower())')"}
                    else                { $uri =  "$teamURI/channels"}
                    Invoke-GraphRequest  -Uri $uri -ValueOnly -As ([MicrosoftGraphChannel]) |
                        Add-Member -PassThru -NotePropertyName Team      -NotePropertyValue $teamid |
                        Add-Member -PassThru -NotePropertyName TeamName  -NotePropertyValue $displayname
                }
                elseif ($Apps -or
                        $AppName)       {
                    $uri = $teamURI + '/installedApps?$expand=teamsAppDefinition'
                    if ($AppName) { $uri = $URI + '&$filter=' +
                                    "startswith(tolower(teamsappdefinition/displayname),'$($AppName.ToLower())')"
                    }
                    Invoke-GraphRequest  -Uri $uri -ValueOnly  -As ([MicrosoftGraphTeamsAppDefinition]) |
                        Add-Member -PassThru -NotePropertyName Team      -NotePropertyValue $teamid |
                        Add-Member -PassThru -NotePropertyName TeamName  -NotePropertyValue $displayname
                }
                elseif ($Select)        {
                    $SelectList = (@('id','displayName') + $Select ) -join','
                    Invoke-GraphRequest  -Uri ($groupuri + '?$Select=' + $SelectList ) -ExcludeProperty '@odata.context' -As ([MicrosoftGraphGroup])
                }
                else                    {
                    $g =  Invoke-GraphRequest  -Uri "$groupuri`?`$expand=members"  -ExcludeProperty '@odata.context','creationOptions' -As ([MicrosoftGraphGroup])
                    # consider adding $MyInvocation.InvocationName -ne 'Get-GraphTeam'
                    if ($g.resourceProvisioningOptions -notcontains 'Team' -or $NoTeamInfo) { $g }
                    else {
                        $t = Invoke-GraphRequest  -Uri  "$teamURI"                  -ExcludeProperty '@odata.context' -As ([MicrosoftGraphTeam])
                        $t.members = $g.Members
                        $t
                    }
                }
            Write-Progress -Activity 'Getting Group/Team information' -Completed
            }
            catch {
                if ($_.exception -match"Forbidden") {
                    Write-warning -Message "Server returned a 403 (Forbidden) error; you must be a memeber of the team $($t.displayname) to view some things [admin does not give access]. "
                }
                if ($_.exception.response.statuscode.value__ -eq 404) {
                    Write-Verbose -Message "GET-GROUP: Nothing found the group $displayname"
                }
                else {throw $_  }
            }
        }
    }
    end     {
        foreach( $g in $UsersAndGroups.where({$_.'@odata.type' -match 'group$'})) {
            $displayname = $g.GroupName
            [void]$g.Remove('GroupName')
            [void]$g.remove('@odata.type')
            [void]$g.remove('@odata.context')
            [void]$g.remove('creationOptions')
            New-Object -Property  $g -TypeName MicrosoftGraphGroup |
                Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
        }
        foreach( $u in $UsersAndGroups.where({$_.'@odata.type' -match 'user$'})) {
            $displayname = $u.GroupName
            [void]$u.Remove('GroupName')
            [void]$u.Remove('@odata.type')
            [void]$u.Remove('@odata.context')
            New-Object -Property $u -TypeName MicrosoftGraphUser |
                Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
        }
    }
}

function New-GraphGroup             {
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
    [Cmdletbinding(SupportsShouldprocess=$true,DefaultParameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup])]
    [Alias("New-GraphTeam")]
    param(
        #The Name of the group / team
        [Parameter(Mandatory=$true, Position=1)]
        [string]$Name,

        #Unless specified groups will be mail enabled "unfied" / Microsoft365 groups
        #The Graph API doesn't allow mail-enabled & security-enabled,  or mail-disabled & unified
        #Only unified groups can be made into teams. Unified groups can only contain users,
        #Security groups can contain other security principals
        [parameter(ParameterSetName='Security',Mandatory=$true)]
        [Switch]$AsSecurity,

        #By default the group is configured as a team unless -NoTeam is specified
        [parameter(ParameterSetName='Team',Mandatory=$true)]
        [Switch]$AsTeam,

        #A description for the group
        [string]$Description,

        #The group/team's mail nickname
        [string]$MailNickName,

        #The visibility of the group, Public by default, it can be 'private' or 'hidden membership'
        [ValidateSet('private', 'public', 'hiddenmembership')]
        [string]$Visibility = 'public',

        #Ordinary Members of the group - assumed to be users, given by their User Principal Name or ID or as objects
        $Members,

        #Owners of the group - assumed to be users, given by their User Principal Name or ID or as objects
        [parameter(ParameterSetName='Owners')]
        [parameter(ParameterSetName='Security')]
        $Owners,
        #if specified group will be added without prompting
        [Switch]$Force
    )
    ContextHas -WorkOrSchoolAccount -BreakIfNot

    if (Invoke-GraphRequest -Uri "$GraphURI/groups?`$filter=displayname eq '$Name'" -ValueOnly) {
        throw "There is already a group with the display name '$Name'." ; return
    }
    #Server-side is case-sensitive for [most] JSON so make sure hashtable names and constants have the right case!
    if (-not $MailNickName) {$MailNickName = $Name -replace "\W",'' }
    $settings = @{  'displayName'          = $Name
                    'mailNickname'         = $MailNickName
                    'mailEnabled'          = -not $AsSecurity
                    'securityEnabled'      = $AsSecurity -as [bool]
                    'visibility'           = $Visibility.ToLower()
                    'groupTypes'           = @()
    }
    if   (-not $AsSecurity ) {
          $settings.groupTypes            += "Unified"
          if ($MyInvocation.InvocationName -eq 'New-GraphTeam' -and -not $PSBoundParameters.ContainsKey('AsTeam')) {
              $AsTeam = $true
          }
    }

    if ($Description) {
          $settings['description']         = $Description
    }
    #if we got owners or users with no ID, fix them at the end, if they have an ID add them now
    if ($Members) {
        $settings['members@odata.bind']    = @();
        foreach ($m in $Members) {
            if  ($m.id) {$settings['members@odata.bind'] += "$GraphURI/users/$($m.id)"}
            else        {$settings['members@odata.bind'] += "$GraphURI/users/$m"}
        }
    }
    #If we make someone else the owner of the group, we can't make it a team,
    #so parameter sets should ensure we can't get owners here if we are making a team.
    if ($Owners) {
        $settings['owners@odata.bind']     = @()
        foreach    ($o in $Owners)  {
            if     ($o.id) {$settings['owners@odata.bind']  += "$GraphURI/users/$($o.id)"}
            else{           $settings['owners@odata.bind']  += "$GraphURI/users/$o"}
        }
    }
    $webparams = @{
        Method     = 'Post'
        Uri        = "$GraphURI/groups"
        Body       = (ConvertTo-Json $settings)
        ContentType = 'application/json'
    }
    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.shouldprocess($Name,"Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-GraphRequest @webparams -As ([MicrosoftGraphGroup]) -Exclude "@odata.context","creationOptions"

        if (-not $AsTeam) {
            Write-Progress -Activity 'Creating Group/Team' -Completed
            return $group
        }
        elseif ($Group.GroupTypes) {
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name"
            $webparams.Uri   +=  "/$($group.id)/team"
            $webparams.Method = 'Put'
            $webparams.Body   = '{ }'

            $team             = Invoke-GraphRequest @webparams -Exclude '@odata.context' -As ([MicrosoftGraphTeam]) |
                                    Add-Member -PassThru -NotePropertyName Mail -NotePropertyValue $group.Mail
            $team.Description = $group.description
            $team.Members     = $group.members
            $team.visibility  = $group.visibility

            if ($Owners) { $
                Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Setting Group ownership on $Name"
                Owners | Add-GraphGroupMember -Group $group -AsOwner -Force
            }
            Write-Progress -Activity 'Creating Group/Team' -Completed
            $team
        }
    }
}

function Set-GraphGroup             {
    <#
      .Synopsis
        Sets options on a group
      .Description
        Allows or blocks external senders, changes visibility or description and enables the group as a team.
        Other options for a team are set via Set-GraphTeam.
        Requires consent to use the Group.ReadWrite.All scope.
      .Example
       Get-GraphGroupList -Name consult | Set-GraphGroup -Description "People who do consulting work" -Force
       Finds the group(s) with a name which matches Consult* and sets the description without a confirmation prompt.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    param   (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([GroupCompleter])]
        $Group ,
        #If specified, updates the group's displayName
        $DisplayName,
        #If specified, the group can receive external email; the option can be disabled with -AllowExternalSenders:$false.
        [switch]$AllowExternalSenders,
        #A new description for the group
        [string]$Description,
        #Enables team functionality on a group which does not yet have it enabled
        [switch]$EnableTeam,
        #If specified the group will be updated without prompting for confirmation.
        [switch]$Force
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #ensure we have an ID for the group(s) we were passed. If we got a GUID in a string, we'll confirm it's a group and get the display name.
        $Group = foreach ($g in $Group) {
                  if     ($g.ID) {$g}
                  elseif ($g -is [String]) {Get-GraphGroup $g -ErrorAction Stop  }
                  else   {throw -Message 'Could not process Group parameter.'; return }
        }
        foreach ($g in $Group)  {
            $settings = @{}  #theme, preferredLanguage, preferredDataLocation, hideFromOutlookClients , hideFromAddressLists , classification, displayname
            if ($Description)       {$settings['description']           = $Description}
            if ($DisplayName)       {$settings['displayName']           = $DisplayName}
            if ($PSBoundparameters.ContainsKey('AllowExternalSenders')) {
                                     $settings['allowExternalSenders']  = [bool]$AllowExternalSenders
            }
            $webparams = @{
                 uri         = "$GraphUri/groups/$($g.ID)"
                 Method      = 'Patch '
                 Body        = (ConvertTo-Json $settings)
                 ContentType = 'application/json'
            }
            Write-Debug $webparams.Body
            if (($settings.Count -or $EnableTeam) -and
                ($Force -or $PSCmdlet.Shouldprocess($g.displayname,'Update Group'))) {
                if ($settings.Count) {
                    Invoke-GraphRequest @webparams | Out-Null
                }
                if ($EnableTeam ) {
                    $response  = Invoke-GraphRequest -Uri $webparams.uri
                    if ($response.resourceProvisioningOptions -contains 'Team')     {
                        Write-Warning  "Group $($g.displayName) is already team-enabled."
                    }
                    else {
                        $webparams.uri    += '/team'
                        $webparams.Method  = 'Put'
                        $webparams.Body    = "{ }"
                        Invoke-GraphRequest  @webparams | Out-Null
                    }
                }
            }
        }
    }
}

function Set-GraphTeam              {
    <#
      .Synopsis
        Updates the settings for a team
      .Description
        Requires consent to use the  Group.ReadWrite.All scope
      .Example
        >Get-GraphTeam -byname accounts | Set-GraphTeam -AllowGiphy:$false
        Gets a the team(s) with a name that begins with accounts, and turns off Giphy content
        Note the use of -SwitchName:$false.
    #>
    [Cmdletbinding(SupportsShouldProcess=$true)]
    param (
        #The team to update either as an ID or a team object with and ID.
        [ArgumentCompleter([GroupCompleter])]
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

    $webparams = @{Method      =  'PATCH'
                  ContentType  =  'application/json'}

    Write-Progress -Activity "Updating Team" -Status "Checking team is valid"
    if     ($Team.id)           {
        $group =  Invoke-GraphRequest -method get "$Graphuri/groups/$($Team.id)" -Headers $DefaultHeader
    }
    elseif ($Team -is [string] -and $team -match $GuidRegex ) {
        $group =  Invoke-GraphRequest -method get "$Graphuri/groups/$Team" -Headers $DefaultHeader
    }
    elseif ($Team -is [string]  ) {
        $group =  Get-GraphGroupList -Name $Team
    }
    if ($group.id -and $group.displayName -and $group.resourceProvisioningOptions -contains 'Team') {
        $webparams['Uri'] = "$Graphuri/teams/$($group.id)"
    }
    else   {
        Write-Progress -Activity "Updating Team" -Completed
        Write-Warning -Message 'Could not resolve the team';
        return
    }
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
    if ($PSBoundparameters.ContainsKey('AllowStickersAndMemes'))                   {$funSettings.allowStickersAndMemes                = [Bool]$AllowStickersAndMemes}
    if ($PSBoundparameters.ContainsKey('AllowCustomMemes'))                        {$funSettings.allowCustomMemes                     = [Bool]$AllowCustomMemes}
    if ($PSBoundparameters.ContainsKey('GiphyContentRating'))                      {$funSettings.giphyContentRating                   = $GiphyContentRating}  #the only string

    if ($memberSettings.Count)    {$settings['memberSettings']    = $memberSettings}
    if ($guestSettings.Count )    {$settings['guestSettings']     = $guestSettings}
    if ($messagingSettings.Count) {$settings['messagingSettings'] = $messagingSettings}
    if ($funSettings.Count)       {$settings['funSettings']       = $funSettings}

    if ($settings.Count) {
        $json = ConvertTo-Json $settings -Depth 10
        Write-Debug $json
        if ($PSCmdlet.ShouldProcess($group.displayName,'Update Team settings')) {
            Write-Progress -Activity "Updating Team" -CurrentOperation $group.displayName -Status "Committing changes"
            Invoke-GraphRequest @webparams -Body $json
            Write-Progress -Activity "Updating Team" -Completed
        }
    }
    else {Write-Warning -Message "Nothing to set"}
}

function Remove-GraphGroup          {
    <#
      .Synopsis
        Removes a group/team
      .Description
        Requires consent to use the Group.ReadWrite.All scope.
        The group may remain visible for a short time.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeam")]
    param(
        #The ID of the Group / team
        [Parameter(Mandatory=$true, Position=0,ValueFromPipeline=$true )]
        [ArgumentCompleter([GroupCompleter])]
        [Alias("Team")]
        $Group,
        #If specified the group will be removed without prompting
        [switch]$Force
    )
    process {
        $Group = foreach ($g in $Group) {
                if     ($g.ID) {$g}
                elseif ($g -is [String]) {Get-GraphGroup $g -ErrorAction Stop  }
                else   {throw 'Could not process Group parameter.'; return }
        }
        foreach ($g in $Group){
            if ($Force -or $PSCmdlet.Shouldprocess("'$($g.displayname)'","Delete Group")) {
                Invoke-GraphRequest -Method Delete  -Uri "$GraphUri/groups/$($g.id)/"
                Write-Verbose "REMOVED GROUP $($g.displayname)"
            }
        }
    }
<#
    Groups in the recycle bin (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group").value
   DELETE /directory/deletedItems/{id}                permanent delete
   POST /directory/deletedItems/{id}/restore          restore item
#>
}

function Add-GraphGroupMember       {
    <#
      .Synopsis
        Adds a user (or group) to a group/team as either a member or owner.
      .Description
        Because the group may be a team the this command has alias of Add-GraphTeamMember
        requires consent to use the Group.ReadWrite.All, Directory.ReadWrite.All, or
        Directory.AccessAsUser.All scope.
      .Example
        >
        >$newGroup = New-GraphGroup -Name Test101
        >Get-GraphUserList -Filter "Department eq 'Accounts'" | Add-GraphGroupMember -Group $newGroup
        Creates a new group; then gets a list of users and adds them to the group.
      .Example
        >Add-GraphTeamMember -Team $Newteam -Member alex@contoso.com -AsOwner
        Adds an owner to a team, using aliases for both the command and the group parameter
    #>
    [Cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("Add-GraphTeamMember")]
    param   (
        #The group / team either as an ID or a group/team object with an IDn
        [Parameter(Mandatory=$true, Position=0)]
        [ArgumentCompleter([GroupCompleter])]
        [Alias("Team")]
        $Group,
        #The user or nested-group to add, either as a UPN or ID or as a object with an ID
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        $Member,
        #If specified the user will be added as an owner, otherwise they will be a standard member
        [switch]$AsOwner,
        #If specified group member will be added without prompting
        [Switch]$Force
    )
    begin   {
        #ensure we have an ID for the group(s) we were passed. If we got a GUID in a string, we'll confirm it's a group and get the display name.
        if (ContextHas -WorkOrSchoolAccount){
            $Group = foreach ($g in $Group) {
                if     ($g.ID) {$g}
                elseif ($g -is [String]) {Get-GraphGroup $g -select id,displayName -ErrorAction Stop  }
                else   {throw -Message 'Could not process Group parameter.'; return }
            }
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot

        foreach ($g in $Group) {
            #group(s) resolved in begin block so should have an ID and display name.
            if ($AsOwner) {$uri   = "$GraphUri/groups/$($g.ID)/owners/`$ref" }
            else          {$uri   = "$GraphUri/groups/$($g.ID)/members/`$ref"}
            #I'm not really expecting an array of users so I have left this is one call for each user.
            #To optimize it piped users could be collected in the process block and the work done in the end block
            foreach ($m in $member) {
                #if we weren't passed as a user as a an object, resolve what we did get ...
                if   (-not $m.id)  {
                    try   {$m     = Get-GraphUser -User $m -Select id,displayname}
                    catch {throw "Could not get a user matching $m"; return }
                    if (-not $m) {throw "Could not get a member ID"; return }
                }
                $body = ConvertTo-Json @{'@odata.id' = "$GraphUri/directoryObjects/$($m.id)"   }
                Write-Debug $body
                if ($Force -or $PSCmdlet.shouldprocess($m.displayname,"Add to Group '$($g.displayname)'")) {
                    Invoke-GraphRequest -Method post -Uri $uri -Body $body -ContentType 'application/json'
                    Write-Verbose "ADDED $($m.displayname) to group $($g.displayname)"
                }
            }
        }
    }
}

function Remove-GraphGroupMember    {
    <#
      .Synopsis
        Removes a user (or group) from a group/team
      .Description
        Because the group may be a team the command has an alias of Remove-GraphTeamMember.
        It requires consent to use the Group.ReadWrite.All, Directory.ReadWrite.All, or
        Directory.AccessAsUser.All scope.
      .Example
        Remove-GraphGroupMember -Group $g -FromOwners -Member alex@contoso.com -Force
        Remmvoes a user from the owners of a group without prompting for confirmation.
      .Example
        Get-GraphUserList -Filter "Department eq 'Accounts'" | Remove-GraphGroupMember -Group $g
        Gets a list of users and removes them from from a group.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeamMember")]
    param   (
        #The ID of the Group / team
        [Parameter(Mandatory=$true, Position=0)]
        [Alias("Team")]
        $Group,
        #A group object with an ID field, or a user object, user ID or UPN
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        $Member,
        #If specified the member will be removed from the owners rather than members
        [switch]$FromOwners,
        #If specified the member will be removed without prompting for confirmation
        [switch]$Force
    )
        begin   {
        #ensure we have an ID for the group(s) we were passed. If we got a GUID in a string, we'll confirm it's a group and get the display name.
        $Group = foreach ($g in $Group) {
            if     ($g.ID) {$g}
            elseif ($g -is [String]) {Get-GraphGroup $g -select id,displayName -ErrorAction Stop  }
            else   {throw  'Could not process Group parameter.'; return }
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        foreach ($g in $Group) {
            #I'm not really expecting an array of users so I have left this is one call fo.r each user.
            #To optimize it piped users could be collected in the process block and the work done in the end block
            foreach ($m in $member) {
                if   (-not $m.id)  {
                    try   {$m = Get-GraphUser -User $m -Select ID,displayName}
                    catch {throw "Could not get a user matching $m"; return }
                    if (-not $m) {throw "Could not get a member ID"; return }
                }
                #group(s) resolved in begin block so should have an ID and display name.

                if ($FromOwners) { $uri = "$GraphUri/groups/$($g.ID)/owners/$($m.id)/`$ref" }
                else {             $Uri = "$GraphUri/groups/$($g.ID)/members/$($m.id)/`$ref"}
                if ($Force -or $PSCmdlet.Shouldprocess($m.displayName,"Remove from Group $($g.displayname)")) {
                    try {
                        Invoke-GraphRequest -Method Delete -Uri $uri
                        Write-Verbose "REMOVED $($m.displayname) from group $($g.displayname)"
                    }
                    catch {
                        If (($_.exception.response.statuscode.value__ -eq 404)) {
                            Write-Warning "Member '$($m.displayName)' was not found in the group $($g.displayname)"
                        }
                        else {$_}
                    }
                }
            }
        }
    }
}

function Export-GraphGroupMember    {
<#
    .synopsis
        Exports a list of group memberships to a CSV file
    .description
        Takes a list of groups (as a parameter or from the pipeline)  and creates four columns
        * Action is either Add or Remove - on export it will always be add
        * MemberOf the name of ONE group the user should be added to or removed from
        * UserPrincipalName the name which will be used for add/remove operations.
        * Displayname just to make things easier to read, especially if UPNs are opaque
        If a file is specified it will be treated as CSV file for export,
        otherwise the objects are output
#>
    param  (
        [Parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        #One or more group(s) to export
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #Destination for CSV output
        $Path,
        #If specified , output will be in Group name order (default is User name.)
        [switch]$OrderByGroup
    )
    begin  {
    $list = @()
    }
    process{
        foreach ($g in $group) {
            if ($g.DisplayName) {$groupName = $g.DisplayName}
            else                {$groupname = $g}
            $list += Get-GraphGroup $g -Members |
                    Select-Object -Property @{n='Action';  e={'Add'}} ,
                                            @{n='MemberOf';e={$groupName}},
                                            UserPrincipalName,
                                            Displayname
        }
    }
    end    {
        if ($OrderByGroup) {$list = $list | Sort-Object -Property Memberof, UserPrincipalName }
        else               {$list = $list | Sort-Object -Property UserPrincipalName, Memberof }
        if (-not $path) {return $list}
        else   {$list | Export-Csv -Path $Path -NoTypeInformation }
    }
}

Function Import-GraphGroupMember    {
<#
    .synopsis
       Imports a list of group memberships from a CSV file
    .description
        Takes a list of CSV files and looks for three columns
        * Action is either Add or Remove - other values will cause the row to be ignored
        * MemberOf the name of ONE group the user should be added to or removed from
        * UserPrincipalName the name which will be used for add/remove operations.
        for each named group the command fetches the membership, users in the group,
        who are marked "remove" in the file  will be removed, and users
        marked "add" in the file who are not in the group will be added.
#>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='high')]
    param   (
        #One or more files to read for input.
        [Parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Usually the command will prompt for confirmation -Force disables this primpt
        [switch]$Force,
        #Supresses output of Added, Removed, or No action messages for each row in the file.
        [switch]$Quiet
    )
    begin   {
        $list = @()
    }
    process {
        foreach ($p in $path) {
            if (Test-Path $p) {$list += Import-Csv -Path $p}
            else { Write-Warning -Message "Cannot find $p" }
        }
    }
    end     {
        if (-not $Quiet) { $InformationPreference = 'continue'  }
        $groups = ($List | Group-Object -NoElement -Property memberof).Name
        foreach ($g in $groups) {
            $w = $Null
            $Members = (Get-GraphGroup $g -Members -WarningAction SilentlyContinue -WarningVariable W).UserPrincipalName
            if ($w) {Write-Warning "Skipping Group $g it did not match a group." ; continue}
            foreach    ($member in $list.where({$_.memberof -eq $g}) ) {
                $upn =  $member.UserPrincipalName
                if    (($member.Action -eq 'Add' -and $upn -notin $Members) -and
                       ($force -or $PSCmdlet.ShouldProcess($upn,"Add user to group'$g'"))) {
                        Add-GraphGroupMember -Force -Group $g -Member $upn
                        Write-Information "Added $UPN user to group'$g'"
                }
                elseif (($member.Action -eq 'Remove' -and $upn -in $Members) -and
                        ($force -or $PSCmdlet.ShouldProcess($upn,"Remove member from group'$g'"))){
                        Remove-GraphGroupMember -Force -Group $g -Member $upn
                        Write-Information "Removed $UPN user from group'$g'"
                }
                else   {Write-Information -Message "No action needed for $g / $upn"}
            }
        }
    }
}

Function Import-GraphGroup          {
<#
    .synopsis
       Imports a list of groups from a CSV file
    .description
        Takes a list of CSV files and looks for four columns
        * Action is either Add or Remove - other values will cause the row to be ignored
        * DisplayName the name which will be used for add/remove operations.
        * Description - the longer text describing the group
        * Type is either Security to configure a non-mail-enabled Security group,
          or Team, to teams enable a group. Blank or other values will create a non-security
          email enabled group which can be teams-enabled later.
        * Visibility - one of 'private', 'public', 'hiddenmembership'
        The command fetches the list of existing groups, any  marked "remove" in the file
        will be removed, and marked "add" who are not in the group will be added using
        the type, visibility, and description settings.
        IF the group exists no check is done to see that it matches the file settings.
#>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #One or more files to read for input.
        [Parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Disables any prompt for confirmation
        [switch]$Force,
        #Supresses output of Added, Removed, or No action messages for each row in the file.
        [switch]$Quiet
    )
    begin   {
        $list = @()
    }
    process {
        foreach ($p in $path) {
            if (Test-Path $p) {$list += Import-Csv -Path $p}
            else { Write-Warning -Message "Cannot find $p" }
        }
    }
    end     {
        if (-not $Quiet) { $InformationPreference = 'continue'  }
        $existingGroups = Get-GraphGroupList
        $existingNames  = $existingGroups.DisplayName
        foreach ($group in $list) {
            $displayName = $group.DisplayName
            if (($Group.Action -eq 'Remove' -and $displayname -in $existingNames) -and
                ($force -or $PSCmdlet.ShouldProcess($displayname,"Remove group "))){
                        Remove-GraphGroup -Force -Group $displayname
                        Write-Information "Removed group'$displayname'"
            }
            elseif (($Group.Action -eq 'Add' -and $displayname -notin $existingNames) -and
                ($force -or $PSCmdlet.ShouldProcess($displayname,"Add new group"))){
                    $params = @{Force=$true; Name=$displayName}
                    if ($group.Type -match 'Security') {$params['AsSecurity'] = $true}
                    if ($group.Type -match 'Team')     {$params['AsTeam'] = $true}
                    if ($group.Visibility)             {$params['Visibility'] = $group.Visibility}
                    if ($group.Description)            {$params['Description'] = $group.Description}
                    New-GraphGroup @params
                    Write-Information "Added group'$displayName'"
            }
            else {  Write-Information "No action taken for group '$displayName'"}
        }
    }
}

function New-GraphTeamPlan          {
    <#
      .Synopsis
        Creates new a plan for a team.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [outputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphPlannerPlan])]
    param   (
        #The ID of the team
        [parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0)]
        [Alias("Group")]
        [ArgumentCompleter([GroupCompleter])]
        $Team,
        #Name(s) of the plan(s) to add to this team.
        [parameter(Mandatory=$true, Position=1)]
        $PlanName,
        #If Specified the plan will be added without confirmation
        [Switch]$Force
    )
    begin   {
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if     ($Team.id)                   {$settings =  @{owner = $team.id} }
        elseif ($Team -is [string] -and
                $Team -match $GUIDRegex )   {$settings =  @{owner = $team} }
        elseif ($Team -is [string]) {
                $Team = Get-GraphTeam $Team; $settings =  @{owner = $team.id}
        }

        foreach ($p in $PlanName) {
            $settings["title"] = $p
            $webParams = @{Method      = 'Post'
                           URI         = "$GraphUri/planner/plans"
                           Contenttype = 'application/json'
                           Body        = (ConvertTo-Json $settings)
            }
            Write-Debug $webParams.Body
            if ($Force -or  $PSCmdlet.ShouldProcess($P,"Add Team Planner")) {
                $result    = Invoke-GraphRequest @webParams -ErrorAction Stop
                $etag      =  $result.'@odata.etag'
                $odatakeys =  $result.Keys.Where({$_ -match "@odata\."})
                foreach ($k in $odatakeys) {$result.Remove($k)}
                $planobj = New-Object  -Property $result -TypeName MicrosoftGraphPlannerPlan |
                    Add-Member -PassThru -NotePropertyName  etag -NotePropertyValue $etag |
                    Add-Member -PassThru -NotePropertyName  Team -NotePropertyValue $Team
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
                $planObj
            }
        }
    }
}

function Get-GraphGroupConversation {
    <#
      .Synopsis
        Gets details of group converstation from outlook, or its threads.
      .Description
        Requires consent to use the Group.Read.All scope
      .Example
        Get-GraphGroupList -Name consult | Get-GraphGroup -Conversations | Get-GraphGroupConversation -Threads
        Gets group(s) matching the name "consult*" , finds their conversations and for each one gets the threads in the conversation
        Note, unless you are dealing with conversations which have multiple threads, it is easier to do Get-GraphGroup -Threads
    #>
    [Cmdletbinding()]
    [Alias("Get-GraphTeamConversation","Get-GraphConversation")]  #Strictly Conversations belong to a group in Outlook, not a Team in Microsoft teams, but let either name be used.
    param   (
        #The Conversation, either as an ID or an object.
        [Parameter(ValueFromPipeline=$true, Mandatory=$true, Position=1, ParameterSetName='OneConversation')]
        $Conversation,
        #The group where the conversation is found,it is not part of can't be found from the conversation object
        [Parameter(ParameterSetName='InTeam')]
        [Parameter(ParameterSetName='OneConversation')]
        [ArgumentCompleter([GroupCompleter])]
        [Alias("Team")]
        $Group,
        #When selecting the Conversations for a group narrows the list by the name of the topic
        [Parameter(ParameterSetName='InTeam', Position=3)]
        $Topic = "*",
        #If specified selects the conversation's threads, otherwise an object representing the conversation itself is returned.
        [Switch]$Threads
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if ( -not $Conversation) {
            $conversations = Get-GraphGroup -Group $Group -Conversations | Where-Object -Property Topic -like $Topic
            if ($Threads) {$conversations | Get-GraphGroupConversation -Threads}
            else          {$conversations}
            return
        }
        if     ($Conversation.Group)       {$groupID = $Conversation.Group}
        elseif ($Group.ID)                 {$groupID = $Group.ID}
        elseif ($Group -is [String]  -and
                $Group -match $GUIDRegex)  {$groupID = $Group}
        elseif ($Group -is [String])       {$groupID = (Get-GraphGroup -Group $group -NoTeamInfo ).id}
        if     ($groupID -notmatch $GUIDRegex) {
                Write-Warning -Message 'Could not resolve group ID'; return
        }
        if     ($Conversation.id)          {$Conversation = $Conversation.id}
        if ($Threads) {
            $uri    = "$GraphUri/groups/$groupID/conversations/$conversation/Threads"
            Invoke-GraphRequest  -Uri $uri -ValueOnly -AsType ([MicrosoftGraphConversationThread]) |
                Add-Member -PassThru -NotePropertyName Group        -NotePropertyValue $GroupID   |
                Add-Member -PassThru -NotePropertyName Conversation -NotePropertyValue $Conversation
        }
        else     {
            Invoke-GraphRequest  -Uri ("$GraphUri/groups/$groupID/conversations/$conversation"  +'?$expand=Threads') -AsType ([MicrosoftGraphConversation]) -ExcludeProperty '@odata.context' |
                 Add-Member -PassThru -NotePropertyName Group -NotePropertyValue $GroupID
        }
    }
}

function Get-GraphGroupThread       {
    <#
      .Synopsis
        Gets a thread in a Group conversation in outlook, or its posts
      .Description
        Requires consent to use the Group.Read.All scope
      .Example
        >Get-GraphUser -Teams  | Get-GraphGroup -Threads | Get-GraphGroupThread -Posts |
             ft -a -Wrap  @{n="from";e={$_.from.emailaddress.name}},CreatedDateTime,Topic,@{n="Body";e={$_.body.content}}
        Gets a users teams, for each one gets their threads, and for each thread gets the outlook posts
        Displays the result as a table showing from, message date, thread topic and message body
        Note this uses Get-GraphGroup as an alias for Get-GraphTeams
    #>
    [Cmdletbinding()]
    [Alias("Get-GraphTeamThread")]
    param   (
        #The group thread, either as an ID or as a thread object (which may have the team/group as property)
        [Parameter(ParameterSetName='SingleThread', Position=0, ValueFromPipeline=$true, Mandatory=$true)]
        $Thread,
        #The group holding the thread (s), if thread is either not passed or is just the ID of a thread.
        [Alias("Team")]
        [Parameter(ParameterSetName='GroupThreads')]
        [Parameter(ParameterSetName='SingleThread', Position=1)]
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #When selecting the threads for a group narrows the list by the name of the topic
        [Parameter(ParameterSetName='GroupThreads')]
        $Topic = '*',
        #If specified, returns the posts in the thread
        [Switch]$Posts
    )
    begin   {

        $webparams = @{Headers         = @{"Prefer" ='outlook.body-content-type="text"' }
                       AsType          =  ([MicrosoftGraphConversationThread])
                       ExcludeProperty = '@odata.context' }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if  (-not $Thread) {
            $threads = Get-GraphGroup -Group $Group -Threads | Where-Object -Property Topic -like $topic
            if ($posts) {$threads | Get-GraphGroupThread -Posts}
            else        {$threads}
            return
        }
        if     ($Thread.Group)                 {$groupid  = $Thread.group}
        elseif ($Group.ID)                     {$groupID  = $Group.ID}
        elseif ($Group -is [String] -and
                $Group -match $GUIDRegex)      {$groupID  = $Group}
        elseif ($Group -is [String])           {$groupID  = (Get-GraphGroup -Group $group -NoTeamInfo ).id}

        if     ($groupID -notmatch $GUIDRegex) {Write-Warning -Message 'Could not resolve group ID'; return }

        if     ($Thread.id)                    {$threadID = $Thread.id}
        elseif ($Thread -is [string])          {$threadID = $Thread}
        else   {Write-Warning -Message 'Could not resolve thread ID'; return}

        $t = Invoke-GraphRequest @webparams -Uri "$GraphUri/groups/$Groupid/Threads/$threadID`?`$expand=Posts"  |
                Add-Member -PassThru -NotePropertyName Group -NotePropertyValue $Groupid
        foreach  ($post in $t.posts) {
            Add-Member -InputObject $post -MemberType NoteProperty -Name Group  -Value $groupid
            Add-Member -InputObject $post -MemberType NoteProperty -Name Thread -Value $t.ID
            Add-Member -InputObject $post -MemberType NoteProperty -Name Topic  -Value $t.Topic
        }
        if ($Posts) {$t.posts}
        else        {$t}
    }
}

function Add-GraphGroupThread       {
    <#
      .Synopsis
        Starts a new thread in a group in outlook.
      .Description
        Requires consent to use the Group.ReadWrite.All scope
      .Example
        >
        >$G = Get-GraphGroup  consultants
        >Add-GraphGroupThread -Group $G -Subject "Running tests.." -Content "We will be running a full test pass this afternoon"
        Gets a group by name and creates a new thread with a message using a plain text body.
      .Example
        >$thread = Add-GraphGroupThread -passthru -Group $G -Subject "Ruuning tests.." -ContentType HTML -Content "<b><i>Drum-Roll...</i>A full test pass is running... Watch this space</i>"
        Uses the group from the previous example, and creates a thread with an HTML body, and keeps a reference to it.
      .link
        Send-GraphGroupReply
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The group where the thread will be added
        [Parameter(Mandatory=$true,Position=0)]
        [Alias("Team")]
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #The subject line for the thread
        [Parameter(Mandatory=$true, Position=1)]
        [Alias("Subject")]
        $ThreadTopic,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true, Position=2)]
        [String]$Content,
        #The content type, (Text by default) or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting; this is the default, unless $confirm preference has been changed
        [switch]$Force,
        #if Specified an object containing the Thread ID will be returned
        [switch]$PassThru
    )
    ContextHas -WorkOrSchoolAccount -BreakIfNot

    if     ($Group.ID)                {$groupID  = $Group.ID}
    elseif ($Group -is [String] -and
            $Group -match $GUIDRegex) {$groupID = $Group}
    elseif ($Group -is [String])      {$groupID =  (Get-GraphGroup -Group $group -NoTeamInfo ).id}

    if ($groupID -notmatch $GUIDRegex)  {Write-Warning -Message 'Could not process Group parameter.'; return }

    $Settings  = @{ 'topic'       = $ThreadTopic
                    'posts'       = @( @{body= @{'content'     = $Content
                                                 'contentType' = $ContentType}})
    }
    $webparams = @{ 'URI'         = "$GraphUri/groups/$groupID/threads/"
                    'Method'      = 'Post'
                    'ContentType' = 'application/json'
                    'Body'        = (ConvertTo-Json $settings -Depth 5)
    }
    Write-Debug $webparams.Body

    if ($force -or $PSCmdlet.Shouldprocess($ThreadTopic,"Create New thread")) {
        $t = Invoke-GraphRequest  @webparams
        if ($PassThru) {
            Start-Sleep -Seconds 2
            Get-GraphGroupThread -Group $Groupid -Thread $t.id
        }
    }
}

function Remove-GraphGroupThread    {
    <#
      .Synopsis
        Removes a thread from a group in outlook
      .Example
        Get-GraphGroup -ByName consultants -Threads | where topic -eq "Today's tests..."  | Remove-GraphGroupThread
        Finds the threads for a named group; isolates one by topic name, and removes it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param   (
        #The thread to remove, either as an ID or a thread object containing an ID, and possibly a conversation ID and group ID
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Thread,
        #The conversation the thread is part of.
        $Conversation,
        #The group from which the thread is to be removed, either as an ID or a group object containing an ID
        [Alias("Team")]
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #if Specified the thread will be deleted without prompting.
        [switch]$Force
    )
    process {
        contexthas -WorkOrSchoolAccount -BreakIfNot
        if     ($Thread.group)             {$groupid  = $Thread.group}
        elseif ($Group.ID)                 {$groupID = $Group.ID}
        elseif ($Group -is [String]  -and
                $Group -match $GUIDRegex)  {$groupID = $Group}
        elseif ($Group -is [String])       {$groupID = (Get-GraphGroup -Group $group -NoTeamInfo ).id}
        if     ($groupID -notmatch $GUIDRegex) {
                Write-Warning -Message 'Could not resolve group ID'; return
        }
        if     ($Thread.Conversation)         {$conversationID = $thread.Conversation}
        elseif ($Conversation.id)             {$conversationID = $Conversation.id}
        elseif ($conversationID -is [string]) {$conversationID = $Conversation}
        if (-not $conversationID) {
                Write-Warning -Message 'Could not resolve Conversation ID'; return
        }

        if     ($Thread.ID)           {$threadid = $Thread.id  }
        elseif ($Thread -is [string]) {$threadid = $Thread.id  }
        else   {Write-Warning 'Could not resolve the Thread ID' ; return}

        $webparams = @{ 'uri'    =  "$GraphUri/groups/$GroupID/conversations/$conversationID/threads/$threadID" }
        Write-Progress -Activity "Deleting thread" -Status "Checking existing thread"
        try   {$thread  = Invoke-GraphRequest -Method Get @webparams }
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
            Invoke-GraphRequest -Method Delete  @webparams
            Write-Progress -Activity "Deleting thread" -Completed
        }
    }
}

function Send-GraphGroupReply       {
    <#
      .Synopsis
        Replies to a group's post in outlook.
      .Example
        >$thread.posts[0] | Send-GraphGroupReply -content '<b><font color="green">Success!!</font> Go team!</b>' -ContentType HTML
        One of the examples for Add-GraphGroupThread left the result of a creating a new thread in $thread
        This takes the only post in the new thread and creates a reply to it with the content in HTML format.
      .Example
        >
        >$post = Get-GraphGroup -ByName consultants -Threads | where topic -eq "Today's tests..." | Get-GraphGroupThread -Posts | select -last 1
        >Send-GraphGroupReply $post -Content "Please join a celebration of the successful test at 4PM"
        This example finds threads for the consultants group, Isolates the one with the topic of
        "Today's Tests..." and finds the last post in the thread. It then posts are reply with the content as plain text.
      .link
        Add-GraphGroupThread
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The Post being replied to, either as an ID or a post object containing an ID which may identify the thread and group
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Post,
        #The thread containing the post (if not embedded in the post itself), as an ID or object, which may identify the group
        $Thread,
        #The group containing the thread (if not embedded in the Post or thread) as an ID or object
        [Alias("Team")]
        [ArgumentCompleter([GroupCompleter])]
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
    ContextHas -WorkOrSchoolAccount -BreakIfNot

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


    $uri       =  "$GraphUri/groups/$groupID/threads/$threadID/posts/$postid"
    Write-Progress -Activity 'Posting reply to thread' -Status 'Checking parent message'
    try   {   $p  = Invoke-GraphRequest -Method Get -uri $uri }
    catch {       throw "Could not get the post to reply to"; return}
    if (-not $p) {throw "Could not get the post to reply to"; return}

    $Settings  = @{ 'Post' = @{'body'= @{'content'=$Content; 'contentType'=$ContentType}}}
    $Json      = ConvertTo-Json $settings
    Write-Debug $Json

    if ($Force -or $PSCmdlet.Shouldprocess($thread.topic,"Reply to thread")) {
        $uri     += "/Reply"
        Write-Progress -Activity 'Posting reply to thread' -Status 'sending reply'
        Invoke-GraphRequest -Method Post -Uri $URI  -Body $Json -ContentType "application/json"
        Write-Progress -Activity 'Posting reply to thread' -Completed
    }
}

Function Get-ChannelMessagesByURI   {
    <#
      .synopsis
        Helper function to add get and expand messages or replies to messages
    #>
    param (
        [parameter(Position=0,ValueFromPipeline=$true)]
        $URI,
        $Top = 20
    )

    process {
        $msglist = @()
        Write-progress -Activity 'Getting messages' -Status "Reading Messages"
        $result   = (Invoke-GraphRequest -Uri $uri)
        $msgList  += $result.value
        while ($result.'@odata.nextLink' -and $result.'@odata.count' -gt 0 -and $msgList.Count -lt $top ) {
            Write-progress -Activity 'Getting messages' -Status "Reading $($ch.displayname) Messages" -CurrentOperation "$($msglist.count) so far"
            $result   = Invoke-GraphRequest -Uri $result.'@odata.nextLink'
            $msgList += $result.value
        }
        $msgList |
            ForEach-Object {New-Object -TypeName MicrosoftGraphChatMessage -Property $_ } |
                Sort-Object -Property createdDateTime -Descending| Select-Object -First $top |
                Add-Member -PassThru -MemberType ScriptProperty -name Team     -Value {$this.ChannelIdentity.TeamID} |
                Add-Member -PassThru -MemberType ScriptProperty -name Channel  -Value {$this.ChannelIdentity.ChannelId}
       <# $userHash = @{}
        Write-Progress -Activity 'Getting messages' -Status "Expanding User information"
        $msglist.from.user.id | Sort-Object -Unique | foreach-object {
            $userHash[$_] = ( Invoke-GraphRequest -Uri  "$GraphUri/directoryObjects/$_").displayName
        }#>
        Write-progress -Activity 'Getting messages' -Completed
        <#foreach ($msg in $msgList) {
            if ($msg.from.user.id) {
                Add-Member -InputObject $msg -NotePropertyName FromUserName -NotePropertyValue $userHash[$msg.from.user.id]
            }
            Add-Member -PassThru -MemberType ScriptProperty -name Team     -Value {$this.ChannelIdentity.TeamID} -InputObject $msg |
            Add-Member -PassThru -MemberType ScriptProperty -name Channel  -Value {$this.ChannelIdentity.ChannelId}
        }#>
    }
}

function Get-GraphChannel           {
    <#
      .Synopsis
        Gets details of a channel, or its Tabs or messages shown in Teams
      .Example
        >Get-GraphTeam -ByName consultants -ChannelName general | Get-GraphChannel -Tabs
        Gets channels for the team(s) with a name beginning 'Consultants' and selects channel(s)
        with a name beginning "general"; then gets the tabs shown in Teams for this channel
      .Example
        >Get-GraphTeam -ByName consultants -ChannelName general | Get-GraphChannel -Messages
        This follows the same method for getting the Teams but this time returns messaes in the channel
      .Example
        >Get-GraphChannel -Team $c -ByName general -Messages
        This is a variation on the previous example - here $c holds an object describing
        the consultants Team and the channel and its messages are retieved using a single command.
      .Example
        >Get-GraphChannel -Team $c -ByName -channel ""
        This previous example didn't explictly specify the channel parameter when using the
        ByName switch; this version does and specifies and empty string so it will return all
        channels (channel is a required parameter, but it can be an empty string)
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeamChannel")]
    param(
        #The channel either as a name, an ID or as a channel object (which may contain the team as a property)
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=1, parameterSetName="CHMsgs")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=1, parameterSetName="CHTabs")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=1, parameterSetName="CHFolder")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=1, parameterSetName="CHFiles")]
        $Channel,
        #The ID of the team if it is not in the channel object. If not specified the current users teams are tried
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #If specified gets the channel's Tabs
        [Parameter(parameterSetName="NoCHTabs", Mandatory=$true)]
        [Parameter(parameterSetName="CHTabs", Mandatory=$true)]
        [switch]$Tabs,
        #If specified gets the channel's Tabs
        [Parameter(parameterSetName="NoCHFolder", Mandatory=$true)]
        [Parameter(parameterSetName="CHFolder", Mandatory=$true)]
        [switch]$Folder,
        [Parameter(parameterSetName="NoCHFiles", Mandatory=$true)]
        [Parameter(parameterSetName="CHFiles", Mandatory=$true)]
        [switch]$Files,
        #if Specified uses the beta api to get the channel's messages.
        [Parameter(parameterSetName="NoCHMsgs")]
        [Parameter(parameterSetName="CHMsgs")]
        [Alias("Msgs")]
        [switch]$Messages,
        #If specified, returns the top n messages, otherwise the command will attempt to get all messages. The server may return more than the specified number.
        [Parameter(parameterSetName="NoCHMsgs")]
        [Parameter(parameterSetName="CHMsgs")]
        $Top
    )

    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if ($Channel -is [string] -and $Channel -notmatch '@thread') {
            $Channel = Get-GraphTeam -Team $Team -Channels -ChannelName $channel
        }
        elseif (-not $Channel) {
            $Channel = Get-GraphTeam -Team $Team -Channels
        }
        #ByName might return multiple channels. Support -channel being given an array of channels.
        foreach ($ch in $channel) {
            if     ($ch.Team)           {$teamID    = $ch.team }
            elseif ($Team.ID)           {$teamID    = $Team.ID }
            elseif ($Team -is [string]) {$teamID    = $Team    }
            else   {Write-Warning -Message 'Could not resolve the team for this channel'; return}
            if     ($ch.id  )           {$channelID = $ch.ID   }
            elseif ($ch -is [string])   {$channelID = $ch      }
            else   {Write-Warning -Message 'Could not resolve the channel ID'; return}
            if (-not ($teamid -and $channelID)) {Write-warning -Message "You need to provide a team ID and a Channel ID"; return}
            elseif   ($Messages -or $Top)              {
                $uri      =  "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/messages"
                if ($top) {Get-ChannelMessagesByURI -URI $uri -Top $Top}
                else      {Get-ChannelMessagesByURI -URI $uri}
                return
            }
            elseif   ($Tabs)                           {
                if ($ch.DisplayName) {Write-Progress -Activity 'Getting Tab information' -CurrentOperation $ch.DisplayName}
                else                 {Write-Progress -Activity 'Getting Tab information' }
                $uri = "$GraphUri/teams/$teamID/channels/$channelID/tabs?`$expand=teamsApp"
                Invoke-GraphRequest -Uri $uri  -ValueOnly -astype ([MicrosoftGraphTeamsTab]) | #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -PassThru -MemberType ScriptProperty -Name TeamsAppName -Value {$this.teamsApp.displayName}
                    #Add-Member -PassThru -MemberType ScriptProperty -Name teamsAppId   -Value {$this.teamsApp.ID}
                Write-Progress -Activity 'Getting Tab information' -Completed
            }
            elseif   ($Folder -or $Files)              {
                if ($ch.DisplayName) {Write-Progress -Activity 'Getting Folder information' -CurrentOperation $ch.DisplayName}
                else                 {Write-Progress -Activity 'Getting Folder information' }
                $uri = "$GraphUri/teams/$teamID/channels/$channelID//filesFolder"
                $f = Invoke-GraphRequest -Uri $uri -AsType  ([MicrosoftGraphDriveItem]) -ExcludeProperty '@odata.context'
                if ($folder) {$f}
                else         {
                    $uri = "$GraphUri/drives/$($f.ParentReference.DriveId)/items/$($f.id)/children"
                    Invoke-GraphRequest -Uri $uri -ValueOnly -ExcludePropert '@odata.etag', '@microsoft.graph.downloadUrl' -AsType ([MicrosoftGraphDriveItem])
                }
                Write-Progress -Activity 'Getting Folder information' -Completed
            }
            elseif   ($ch -is [MicrosoftGraphChannel]) {
                #Have already fetched the channel once so don't fetch it again
                $ch
            }
            else     {
                if ($ch.DisplayName) {Write-Progress -Activity 'Getting Channel information' -CurrentOperation $ch.DisplayName}
                else                 {Write-Progress -Activity 'Getting Channel information' }
                Invoke-GraphRequest -Uri  "$GraphUri/teams/$teamID/channels/$channelId" -ExcludeProperty '@odata.context' -AsType ([MicrosoftGraphChannel]) |
                    Add-Member -PassThru -NotePropertyName Team -NotePropertyValue $teamID
                Write-Progress -Activity 'Getting Channel information' -Completed
            }
        }
    }
}

function New-GraphChannel           {
    <#
      .Synopsis
        Adds a channel to a team
      .Description
        This requires the Group.ReadWrite.All scope.
      .Example
       >$newChannel  = New-GraphChannel -Team $newTeam -Name $newProjectName -Description "For anything about project $newProjectName"
       $newTeam holds the result of creating a team with New-GraphTeam...
       $newProjectName holds the name of a project the team will be working on.
       This command creates a new channel in Teams, and stores the result in a variable
       which can then be used to post messages to the channel, or add tabs to it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("Add-GraphTeamChannel")]
    param(
        #The team where the channel will be added, either as an ID or a team object
        [Parameter( Mandatory=$true, Position=0)]
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #Display name for the new channel
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Alias("DisplayName")]
        [String[]]$Name,
        #Description for the new channel
        [String]$Description
    )
    begin  {
        $webparams = @{Method = "POST"
                       ContentType = "application/json"
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if     ($Team.id) {$team = $team.id }
        elseif ($Team  -is [string] -and $Team -notmatch $GUIDRegex) {
                $Team = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
        }       #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams
        #intentionally fail if the previous step returns an array or nothing - and it won't return groups which aren't team enabled. We should now have a GUID
        if     ($Team  -is [string] -and $Team -match $GUIDRegex) {
                $webparams['uri'] = "$GraphUri/teams/$Team/channels"
        }
        else  { Write-Warning "Could not resolve $($PSBoundParameters['team']) to team-enabled group." ; return        }

        foreach ($n in $Name) {
            if (Get-GraphTeam  $team -ChannelName $n  ) {
                Write-Warning -Message "Channel '$n' already exists in team '$($PSBoundParameters['team'])'."
                continue
            }

            $Settings  = @{"displayName" = $n}
            if ($Description) {$settings["description"] = $Description}
            $webparams['body'] = ConvertTo-Json $settings
            Write-Debug $webparams['body']

            if ($PSCmdlet.Shouldprocess($n,"Create channel")) {
                Invoke-GraphRequest @webparams -ExcludeProperty '@odata.context' -AsType ([MicrosoftGraphChannel]) |
                        Add-Member -PassThru -NotePropertyName Team -NotePropertyValue $team
            }
        }
    }
}

function Remove-GraphChannel        {
    <#
      .Synopsis
        Removes a channel from a team
      .Description
        This requires the Group.ReadWrite.All scope.
      .Example
        >Get-GraphTeam -ByName Developers -ChannelName "Project Firebird" | Remove-GraphChannel
        Finds a channel by name from a named team , and removes it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param(
        #The channel to delete; either as an ID, or a channel object
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        [ArgumentCompleter([TeamCompleter])]
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

            Invoke-GraphRequest -Method "Delete" -Uri "$GraphUri/teams/$Team/channels/$Channel"
            }
        }
}

function New-GraphChannelMessage    {
    <#
      .Synopsis
        Adds a new thread in a channel in Teams.
      .Description
      .Example
        >
        >$General = Get-GraphTeam $newTeam -ChannelName "General"
        >Add-GraphChannelMessage -Channel $General -Content "Project Firebird now has its own channel."
        This adds a message
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param(
        #The channel to post to either as an ID or a channel object.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if     ($Channel.Team)             {$teamID    = $Channel.team }
        elseif ($Team.id)                  {$teamID    = $Team.ID}
        elseif ($Team -is [string] -and
                $Team -match $GUIDRegex)   {$teamID    = $Team}
        elseif ($Team -is [string])        {
                $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
        }       #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams

        if     ($Channel.id)               {$channelID = $Channel.ID   }
        elseif ($Channel -is [string] -and
                $Channel -match '@thread') {$channelID = $channel  }
        elseif ($Channel -is [string]) {
                $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
        }

        if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                  $channelID -is [string] -and $channelID -match '@thread'))  {
            #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
            Write-Warning -Message 'Could not determine the team and channel IDs'; return
        }
        if ($Channel-is [MicrosoftGraphChannel] ) {$c= $Channel}
        else {
            try {$c = Get-GraphChannel -Channel $channelID -Team $teamID }
            catch         {throw "Could not get the channel" ; return}
        }
        if (-not $c)  {throw "Could not get the channel" ; return }
        $webparams = @{ 'Method'      = 'POST'
                        'URI'         = "$GraphUri/teams/$teamID/channels/$channelID/messages" # "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/chatThreads"
                        'ContentType' = 'application/json'
                        'AsType'      = ([MicrosoftGraphChatMessage])
                        'ExcludeProperty' = '@odata.context'

        }
        #$Settings = @{ rootMessage = @{body= @{content=$Content;}}}
        #if ($ContentType -eq 'HTML') {$settings.rootMessage.body['contentType'] = 1}
        #else                         {$settings.rootMessage.body['contentType'] = 2}
        $webparams['body'] = ConvertTo-Json (@{body = @{content=$Content ; contentType = $ContentType}})
        Write-Debug $webparams.body

        if ($force -or $PSCmdlet.Shouldprocess("Create Message")) {
            $result = Invoke-GraphRequest @webparams
            $result.channelIdentity.TeamID = $teamID
            $result.channelIdentity.ChannelId = $channelID
            Add-Member -PassThru -MemberType ScriptProperty -name Team  -Value {$this.ChannelIdentity.TeamID} -InputObject $result |
            Add-Member -PassThru -MemberType ScriptProperty -name Channel  -Value {$this.ChannelIdentity.ChannelId}
        }
    }
}

function New-GraphChannelReply      {
    <#
      .Synopsis
        Posts a reply to a message in a Teams channel
    #>
    [Cmdletbinding(SupportsShouldProcess=$true)]
    param (
        #The Message to reply to as an ID or a message object
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Message,
        #If Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        $Channel,
        #If the message or channel parameters don't included the team ID, the team either as an ID or an objec containing the ID
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #Normally the reply is added 'silently'. If passthru is specified, the new message will be returned.
        [Alias('PT')]
        [switch]$Passthru,
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    $webparams = @{ 'Method'      = 'POST'
                    'ContentType' = 'application/json'
                    'AsType'      = ([MicrosoftGraphChatMessage])
                    'ExcludeProperty' = '@odata.context'
    }
    #region convert the information from the message (and optionally channel and team) into a URI to post to
    if     ($message.ChannelIdentity.TeamId)    {$teamId    = $message.ChannelIdentity.TeamId }
    elseif ($Message.team)                      {$teamid    = $Message.team}
    elseif ($Channel.team)                      {$teamid    = $Channel.team}
    elseif ($Team.id)                           {$teamid    = $team.id}
    elseif ($Team -is [string] -and
            $Team -match $GUIDRegex)            {$teamID    = $Team}
    elseif ($Team -is [string])                 {
            $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
    }       #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams

    if     ($Message.ChannelIdentity.ChannelId) {$channelid = $Message.ChannelIdentity.ChannelId}
    elseif ($Message.channel)                   {$channelid = $Message.channel}
    elseif ($Channel.id)                        {$channelid = $channel.id}
    elseif ($Channel -is [string] -and
            $Channel -match '@thread')          {$channelID = $channel}
    elseif ($Channel -is [string]) {
                $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
    }

    if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                  $channelID -is [string] -and $channelID -match '@thread'))  {
            #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
            Write-Warning -Message 'Could not determine the team and channel IDs'; return
    }

    if     ($Message.ID)           {$msgID     = $Message.ID}
    elseif ($Message -is [string]) {$msgID     = $Message }
    else   {Write-Warning 'Could not determine the ID for the message.'; return}

    $webparams['uri'] =  "$GraphUri/teams/$teamid/channels/$channelid/Messages/$msgID/replies"
    #endregion

    $webparams['body'] = ConvertTo-Json  @{body= @{content=$Content; 'contentType'=$ContentType}}
    Write-Debug $webparams.body

    if ($force -or $PSCmdlet.Shouldprocess("Post Reply")) {Invoke-GraphRequest @webparams }
}

function Get-GraphChannelReply      {
    <#
      .Synopsis
        Returns replies to messages in Teams channels
      .Description
        Access to channel messages is currently in the BETA API
        It is possible to start a new thread, but not to reply to the thread.
      .Example
        >Get-GraphChannel $General -Messages | Get-GraphChannelReply -PassThru
        The GraphAPI does not return replies when requesting messages
        from a channel in Teams. By piping the messages to Get-GraphChannelReply
        it is possible to get the replies; and if -Passthru is specified
        the messages will returned, followed by their replies.
        So if $General is a channel object, the first message and the its first
        reply might be output like this.

        From          Created          Isreply Deleted Importance Content
        ----          -------          ------- ------- ---------- -------
        James O'Neill 17/02/2019 11:42 False   False   normal     Project Firebird now has its own channel.
        James O'Neill 17/02/2019 13:06 True    False   normal     And the channel has its own planner

    #>
    [Cmdletbinding()]
    param (
        #The Message to reply to as an ID or a message object containing an ID (and possibly the team and channel ID)
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Message,
        #If the Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        $Channel,
        #If the message or channel parameters don't included the team ID, the team either as an ID or an objec containing the ID
        $Team,
        #If specified returns the message, followed by its replies. (Otherwise , only the replies are returned)
        [switch]$PassThru
    )
    process {
        ContextHas -scopes 'ChannelMessage.Read.All' -BreakIfNot
        #region convert the information from the message (and optionally channel and team) into a URI to post to
        if     ($message.ChannelIdentity.TeamId)    {$teamId    = $message.ChannelIdentity.TeamId }
        elseif ($Message.team)                      {$teamid    = $Message.team}
        elseif ($Channel.team)                      {$teamid    = $Channel.team}
        elseif ($Team.id)                           {$teamid    = $team.id}
        elseif ($Team -is [string] -and
                $Team -match $GUIDRegex)            {$teamID    = $Team}
        elseif ($Team -is [string])                 {
                $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
        }        #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams

        if     ($Message.ChannelIdentity.ChannelId) {$channelid = $Message.ChannelIdentity.ChannelId}
        elseif ($Message.channel)                   {$channelid = $Message.channel}
        elseif ($Channel.id)                        {$channelid = $channel.id}
        elseif ($Channel -is [string] -and
                $Channel -match '@thread')          {$channelID = $channel}
        elseif ($Channel -is [string]) {
                    $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
        }

        if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                        $channelID -is [string] -and $channelID -match '@thread'))  {
                #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
                Write-Warning -Message 'Could not determine the team and channel IDs'; return
        }

        if     ($Message.ID)           {$msgID     = $Message.ID}
        elseif ($Message -is [string]) {$msgID     = $Message }
        else   {Write-Warning 'Could not determine the ID for the message.'; return}


        if ($PassThru -and $Message -is [MicrosoftGraphChatMessage]) {$Message}
        Get-ChannelMessagesByURI -URI "$GraphUri/teams/$teamid/channels/$channelid/Messages/$msgID/replies"
    }
}

function Add-GraphWikiTab {
    <#
      .Synopsis
        Adds a wiki tab to a channel in teams
      .Example
        >New-GraphWikiTab -Channel $Channel -TabLabel Wiki
        Channel contains an object representing a channel in teams,
        this adds a Wiki to it. The Wiki will need to be initialized
        when the tab is first opened
    #>
    [CmdletBinding(SupportsShouldprocess=$true)]
    param(
        #An ID or Channel object which may contain the team ID
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab
        $TabLabel = "Wiki",
        #If specified the tab will be added without prompting for confirmation
        [switch]$Force
    )

    ContextHas -WorkOrSchoolAccount -BreakIfNot
    if     ($Channel.Team)            {$teamID  = $Channel.Team }
    elseif ($Team.id)                 {$teamID  = $Team.id      }
    elseif ($Team -is [string] -and
            $Team -match $GUIDRegex)  {$teamID    = $Team}
    elseif ($Team -is [string])       {
            $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
    }       #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams
    if     ($Channel.id)           {$channelID = $Channel.id }
    elseif ($Channel -is [string] -and $Channel -notmatch '@thread') {
                $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
    }
    elseif ($Channel -is [string]) {$channelID = $channel  }

    if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                $channelID -is [string] -and $channelID -match '@thread'))  {
        #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
        Write-Warning -Message 'Could not determine the team and channel IDs'; return
    }
    $webparams = @{'Method'          = 'Post'
                   'Uri'             = "$graphuri/teams/$teamID/channels/$channelID/tabs"
                   'ContentType'     = 'application/json'
                   'AsType'          =  ([MicrosoftGraphTeamsTab])
                   'ExcludeProperty' = '@odata.context'
    }
    $webparams['Body'] = ConvertTo-Json ([ordered]@{
        'displayname'         = $TabLabel
        'teamsApp@odata.bind' = "$GraphURI/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"}
    )

    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.Shouldprocess($TabLabel,"Create wiki tab")) {
       Invoke-GraphRequest @webparams | #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -PassThru -MemberType ScriptProperty -Name teamsAppName -Value {$this.teamsApp.displayName}
    }
}
# Adding tab https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
# https://products.office.com/en-us/microsoft-teams/appDefinitions.xml

function Add-GraphPlannerTab     {
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
    #region get IDs needed
    ContextHas -WorkOrSchoolAccount -BreakIfNot
    if     ($Channel.Team)            {$teamID  = $Channel.Team }
    elseif ($Team.id)                 {$teamID  = $Team.id      }
    elseif ($Team -is [string] -and
            $Team -match $GUIDRegex)  {$teamID    = $Team}
    elseif ($Team -is [string])       {
            $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
    } #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams
    if     ($Channel.id)           {$channelID = $Channel.id }
    elseif ($Channel -is [string] -and $Channel -notmatch '@thread') {
                $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
    }
    elseif ($Channel -is [string]) {$channelID = $channel  }

    if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                $channelID -is [string] -and $channelID -match '@thread'))  {
        #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
        Write-Warning -Message 'Could not determine the team and channel IDs'; return
    }
    #endregion
    if ((-not $TabLabel) -and $Plan.Title) {
        Write-Verbose -Message "ADD-GRAPHPLANNERTAB: No Tab label was specified, using the Plan title '$($Plan.Title)'"
        $TabLabel = $Plan.Title
    }
    #If Plan and/or channel were objects with IDs use the ID
    if       ($Channel.id) {$Channel = $Channel.id}
    if       ($Plan.id)    {$Plan    = $Plan.id}
    $tabURI = "https://tasks.office.com/{0}/Home/PlannerFrame?page=7&planId={1}" -f $global:GraphUser  , $Plan

    $webparams = @{'Method'          = 'Post'
                   'Uri'             = "$graphuri/teams/$teamID/channels/$channelID/tabs"
                   'ContentType'     = 'application/json'
                   'AsType'          =  ([MicrosoftGraphTeamsTab])
                   'ExcludeProperty' = '@odata.context'
    }

    $webparams['body'] = ConvertTo-Json ([ordered]@{
        'displayname'         = $TabLabel
        'teamsApp@odata.bind' = "$GraphURI/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
        'configuration' = [ordered]@{
                   'entityId'   = $plan
                   'contentUrl' = $tabURI
                   'websiteUrl' = $tabURI
                   'removeUrl'  = $tabURI
        }
    })
    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
       Invoke-GraphRequest @webparams | #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -PassThru -MemberType ScriptProperty -Name teamsAppName -Value {$this.teamsApp.displayName}
    }
}

function Add-GraphOneNoteTab     {
    <#
      .Synopsis
        Adds a tab in a Teams channel for a OneNote section or Notebook
      .Description
        This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
        which requires consent to use the Group.ReadWrite.All scope.
        The Notebook Parameter has an alias of 'Section' and will accept either
        a OneNote Notebook object (or its 'Self' URI - which requires the tab name to be
        set explicitly) or a Section object. If the notebook is specified it opens at the
        first section.
      .Example
        >
        > $section = Get-GraphTeam -ByName accounts -Notebooks | Select-Object -ExpandProperty sections  | where displayname -like "FY-19*"
        > $channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
        > Add-GraphOneNoteTab  $section $channel -TabLabel "FY-19 Notes"

        The first command gets the Notebook for the Accounts team and finds the "FY-19 Year End" section
        The second command gets the channels for the same team and finds the "Year end" channel
        The Third command creates a tab in the channel named 'FY-19 Notes' which opens the team notebook
        at its 'FY-19 Year End' section.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #The Notebook or Section to associate with the tab
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Section')]
        $Notebook,
        #An ID or Channel object which may contain the team ID; the tab will be created in this channel
        [Parameter(Mandatory=$true, Position=1)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab, if left blank the name of the Notebook or Section will be sued
        $TabLabel,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        [Alias('PT')]
        [switch]$PassThru,
        #If Specified the tab will be added without pausing for confirmation, this is the default unless $ConfirmPreference has been set.
        $Force
    )
    ContextHas -scopes 'Group.ReadWrite.All' -BreakIfNot
    if     ($Channel.Team)           {$teamID     = $Channel.Team }
    elseif ($Team -is [string] -and
            $Team -match $GUIDRegex)  {$teamID    = $Team}
    elseif ($Team -is [string])       {
            $teamID = (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=startswith(displayname,'$Team')" -ValueOnly).id
    }       #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams
    if     ($Channel.id)           {$channelID = $Channel.id }
    elseif ($Channel -is [string] -and
            $Channel -match '@thread') {$channelID = $channel  }
    elseif ($Channel -is [string])    {
            $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
    }
    if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
              $channelID -is [string] -and $channelID -match '@thread'))  {
        #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
        Write-Warning -Message 'Could not determine the team and channel IDs'; return
    }
    if       (-not $TabLabel -and
                $notebook.displayName) {$TabLabel = $Notebook.displayName}
    elseif   (-not $TabLabel)          {Write-warning 'Unable to determin a name for the tab, please specify one explicitly'; return}

    $webparams = @{'Method'          = 'Post'
                   'Uri'             = "$graphuri/teams/$teamID/channels/$channelID/tabs"
                   'ContentType'     = 'application/json'
                   'AsType'          =  ([MicrosoftGraphTeamsTab])
                   'ExcludeProperty' = '@odata.context'
    }
    #This bit had to be reverse engineered, from a beta version of the API, so if it works past next week, be happy.
    #If the "Notebook" object is actually a section, and it was fetched by one of the module commands (get-GraphTeam -notebook, or get-graphNotebook -section)
    #then $Notebook it will have a a parentNotebook ID. This IF..Else is to make sure we have the real notebook ID, and catch a sectionID if there is one.
    if   ($Notebook.parentNotebook.id) {
                    $ParamsPt2    = '&notebookSource=PickSection&sectionId='+ $Notebook.id
                    $NotebookID   = $Notebook.parentNotebook.id
          }
    else  {         $ParamsPt2    = '&notebookSource=New'
                    $NotebookID   = $Notebook.id }

    #if $Notebook is a section its url will end ?wd=(something). We need to split this off the URL and re-use it. The () need to be unescapted too,
    if ($Notebook.links.oneNoteWebUrl.href -match '\?(wd=.*$)') {
                $ParamsPt2       += '&' + ( $Matches[1] -replace '%28','(' -replace '%29',')' )
                $OnenoteWebUrl    = $Notebook.links.oneNoteWebUrl.href  -replace  '\?wd=.*$', ''
    }
    else      { $OnenoteWebUrl    = $Notebook.links.oneNoteWebUrl.href}

    #We need the teamsite URL for the team who owns this channel, and the URL to the the Notebook. Both need to be escaped.
    $OnenoteWebUrl  = $OnenoteWebUrl                             -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'
    $siteUrl        = (Get-GraphTeam -Team $Teamid -Site).webUrl -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'

    #Now we need to build up the mother and father of all URIs It contains the ID and URL for the notebook (not section). The Name, the teamsite. And Section specifics if applicable.
    $URIParams      = "?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&ui={locale}&tenantId={tid}"+
                      "&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$Team%2Fnotes%2Fnotebooks%2F"+ $NotebookID   +
                      "&oneNoteWebUrl=" + $oneNoteWebUrl +
                      "&notebookName="  + [uri]::EscapeDataString( $notebook.displayName ) +
                      "&siteUrl="       + $SiteUrl +
                      $ParamsPt2

    #Now we can create the JSON. Such information as there is can be found at https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
    $json = ConvertTo-Json ([ordered]@{
                'teamsApp@odata.bind' = "$GraphURI/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
                'displayname'         = $TabLabel
                'configuration'       = [ordered]@{
                    'entityId'        = ((New-Guid).tostring() + "_" +  $Notebook.ID)
                    'contentUrl'      = "https://www.onenote.com/teams/TabContent" + $URIParams
                    'removeUrl'       = "https://www.onenote.com/teams/TabRemove"  + $URIParams
                    'websiteUrl'      = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"
                }})
    $webparams['body']= $json  -replace "\\u0026","&"
    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
         Invoke-GraphRequest @webparams | #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -PassThru -MemberType ScriptProperty -Name teamsAppName -Value {$this.teamsApp.displayName}
    }
}
