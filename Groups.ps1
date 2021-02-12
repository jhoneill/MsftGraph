using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Globalization


$Script:GraphUri  = "https://graph.microsoft.com/v1.0"
$Script:GUIDRegex = "^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$"

class GroupCompleter : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result = [System.Collections.Generic.List[CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        if ($wordToComplete) {$uri =  $script:GraphUri +  ("/Groups/?&`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $wordToComplete)}
        else                 {$uri = "$script:GraphUri/Groups/?&`$Top=20"}

        Invoke-GraphRequest -Uri $uri -ValueOnly | ForEach-Object displayname | Sort-Object | ForEach-Object {
                $result.Add(( New-Object -TypeName CompletionResult -ArgumentList "'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
        }

        return $result
    }
}

function Get-GraphGroupList      {
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

function Get-GraphGroup          {
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
                else {throw $_  }
            }
        }
    }
    end     {
        foreach( $g in $UsersAndGroups.where({$_.'@odata.type' -match 'group$'})) {
            $displayname = $g.GroupName
            $g.Remove('GroupName')
            $g.remove('@odata.type')
            $g.remove('@odata.context')
            $g.remove('creationOptions')
            New-Object -Property  $g -TypeName MicrosoftGraphGroup |
                Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
        }
        foreach( $u in $UsersAndGroups.where({$_.'@odata.type' -match 'user$'})) {
            $displayname = $u.GroupName
            $u.Remove('GroupName')
            $u.Remove('@odata.type')
            $u.Remove('@odata.context')
            New-Object -Property $u -TypeName MicrosoftGraphUser |
                Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
        }
    }
}

function New-GraphGroup          {
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

function Set-GraphGroup          {
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
    param (
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

function Set-GraphTeam           {
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

function Remove-GraphGroup       {
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
                Write-Verbose "Removed Group $($g.displayname)"
            }
        }
    }
<#
    Groups in the recycle bin (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group").value
   DELETE /directory/deletedItems/{id}                permanent delete
   POST /directory/deletedItems/{id}/restore          restore item
#>
}

function Add-GraphGroupMember    {
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
                    Write-Verbose "Added $($m.displayname) to group $($g.displayname)"
                }
            }
        }
    }
}

function Remove-GraphGroupMember {
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
                        Write-Verbose "Removed $($m.displayname) from group $($g.displayname)"
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

function Export-GraphGroupMember {
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

Function Import-GraphGroupMember {
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

Function Import-GraphGroup       {
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

function New-GraphTeamPlan       {
    <#
      .Synopsis
        Creates new a plan for a team.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
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
