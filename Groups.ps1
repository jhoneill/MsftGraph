using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Management.Automation

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
        >(Get-GraphGroupList -Name consult* | Get-GraphTeam -Site).weburl
        Gets any group whose name begins "Consult" , finds its sharepoint site, and returns the site's URL
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup])]
    param   (
        #if specified limits the groups returned to those with names begining...
        [Parameter(parameterSetName='FilterByName', Position=0)]
        [string]$Name = "*",
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
        #A field to sort by - sorting is applied on the client side because filter and selct cannot be combined with server-side sort
        [ValidateSet('allowExternalSenders', 'assignedLabels', 'assignedLicenses', 'autoSubscribeNewMembers', 'classification',
                'createdDateTime', 'deletedDateTime', 'description', 'displayName', 'expirationDateTime', 'groupTypes',
                'hasMembersWithLicenseErrors', 'hideFromAddressLists', 'hideFromOutlookClients', 'id', 'isArchived',
                'isSubscribedByMail', 'licenseProcessingState', 'mail', 'mailEnabled', 'mailNickname', 'membershipRule',
                'membershipRuleProcessingState', 'onPremisesDomainName', 'onPremisesLastSyncDateTime', 'onPremisesNetBiosName',
                'onPremisesProvisioningErrors', 'onPremisesSamAccountName', 'onPremisesSecurityIdentifier',
                'onPremisesSyncEnabled', 'preferredDataLocation', 'preferredLanguage', 'proxyAddresses', 'renewedDateTime',
                'securityEnabled', 'securityIdentifier', 'theme', 'unseenCount', 'visibility')]
        [string]$OrderBy = 'displayName',
        [Parameter(parameterSetName='Sort')]
        [Switch]$Descending,
        #An oData filter string; there is a graph limitation that you can't filter by description or Visibility.
        [Parameter(Mandatory=$true, parameterSetName='FilterByString')]
        [string]$Filter
    )
    process {
        #xxxx to do: investigate "groupTypes/any(c:  c eq 'Unified')"  -filter "groupTypes/any(x: x eq 'DynamicMembership')"
        # check access to scopes  Group.Read.All

        if     ($Select)  {
            if ("id"          -notin $select) {$select += 'id'}
            if ("displayName" -notin $select) {$select += 'displayName'}
            $uri =  $GraphUri + '/Groups/?$select='  + ($Select -join ',')
        }
        elseif ($Filter)        {$uri =  $GraphUri + '/Groups/?$filter='  + $Filter }
        elseif ($Name -and $name -match  '\*.*\*|^\*.+')   {  # ie. *xxx* or xxx* but not "*"  alone or xxx*
                                 $uri =  $GraphUri + '/Groups/?$search="displayname:'  + ($Name -replace '\*','') +'"'
        }
        elseif ($Name -ne '*')  {$uri =  $GraphUri + '/Groups/?$filter='  +(FilterString $Name)}
        elseif ($orderby)       {$uri =  $GraphUri + '/Groups/?$OrderBy=' + $OrderBy }
        else                    {$uri =  $GraphUri + '/Groups/' }
        Write-Progress -Activity "Finding Groups"
        Invoke-GraphRequest -Uri $uri -AllValues -ExcludeProperty 'creationOptions' -AsType ([MicrosoftGraphGroup]) -Headers @{'consistencyLevel'='eventual'} |
            Where-Object        {$_.displayname -like $Name -or $_.displayname -like [WildcardPattern]::Escape($Name)} |
                Sort-Object -Property $OrderBy    -Descending:$Descending
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
         Note that because we are refering to "Teams" the command is called using its alias of Get-GraphTeam.
         The last plan is selected and details of the plan are fetched, showing the result as a table.
      .Example
        >(Get-GraphGroup -Site).lists | where name -match document
        If no Group/Team is provided the command gets those associated with the current user;
        it this case it returns their associated site(s).
        Site objects include a lists property, which holds a collection of lists
        This command will fiter the lists down to those where name matches "document",
        giving the "Shared Documents" list for each team
      .Example
        >Get-GraphGroup -Drive  | Get-GraphDrive -Subfolders | Select  name, weburl, id,@{n="drive";e={$_.parentReference.driveId}}
        As with the previous example gets this command gets Groups/Teams for current user,
        in this case the command returns their associated drive(s)
        It is possible to refer to the drive's 'root' property, and the root's 'childre'n property
        which contains files and folder objects, and filter to objects with a folder property but
        for ease of reading this  pipeline passes the drive to Get-GraphDrive to get subfolders.
        It then returns the  name, WebURl and the item ID and Drive ID needed to access each folder.
      .Example
        >Get-GraphGroup 'Consultants' -Drive  | Set-GraphHomeDrive
        Sets the drive for the consultants group to be the default graph drive for the PowerShell session.
     .Example
        >Get-GraphGroup -Notebooks | select -ExpandProperty sections | where "Displayname" -eq "General_Notes"
        Again gets Groups/Teams for the current user and returns their associated notebooks(s)
        Notebook objects include a Sections property, which holds a collection of OneNote sections in the notebook;
        This command gets returns any section in a team notebook which has the name "General_Notes"
      .Example
        > Get-GraphTeam -threads | where LastDeliveredDateTime -gt ([datetime]::Now.AddDays(-7))
        Gets the teams conversation threads which have been updated in the last 7 days.
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeam","ggg")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',  Justification='Write-warning could be used, but the is informational non-output.')]
    param   (
        #The name of a team.
        #One more Team IDs or team objects containing and ID. If omitted the current user's teams will be used.
        [Parameter(ValueFromPipeline=$true, Position=0)]
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
        #If specified gets the Team's OneDrive to see contents of the root of the drive you can refer to the drives .root.children property
        [Parameter(Mandatory=$true, parameterSetName='Drive')]
        [switch]$Drive,
        #If specified returns the members of the team
        [Parameter(Mandatory=$true, parameterSetName='Members')]
        [switch]$Members,
        #If specified returns the transitive members of the team
        [Parameter(Mandatory=$true, parameterSetName='TransitiveMembers')]
        [switch]$TransitiveMembers,
        #If specified returns the groups this group is directly a member of
        [Parameter(Mandatory=$true, parameterSetName='Memberof')]
        [switch]$MemberOf,
        #If specified returns the groups this group is nested into transitively
        [Parameter(Mandatory=$true, parameterSetName='TransitiveMemberof')]
        [switch]$TransitiveMemberOf,
        #If specified returns the Owners of the team
        [Parameter(Mandatory=$true, parameterSetName='Owners')]
        [switch]$Owners,
         #Field(s) to select for the members or owners of the group : ID and displayname are always included
        [Parameter(Mandatory=$false, parameterSetName='Members')]
        [Parameter(Mandatory=$false, parameterSetName='TransitiveMembers')]
        [parameter(Mandatory=$false, parameterSetName="Owners")]
        [ValidateSet  ('aboutMe', 'accountEnabled' , 'activities', 'ageGroup', 'agreementAcceptances' , 'appRoleAssignments',
                       'assignedLicenses', 'assignedPlans', 'authentication', 'birthday', 'businessPhones',
                       'calendar', 'calendarGroups', 'calendars', 'calendarView', 'city', 'companyName', 'consentProvidedForMinor',
                       'contactFolders', 'contacts', 'country', 'createdDateTime', 'createdObjects', 'creationType' ,
                       'deletedDateTime', 'department', 'directReports', 'displayName', 'drive', 'drives',
                       'employeeHireDate', 'employeeId', 'employeeOrgData', 'employeeType', 'events', 'extensions',
                       'externalUserState', 'externalUserStateChangeDateTime', 'faxNumber', 'followedSites', 'givenName', 'hireDate',
                       'id', 'identities', 'imAddresses', 'inferenceClassification', 'insights', 'interests', 'isResourceAccount', 'jobTitle', 'joinedTeams',
                       'lastPasswordChangeDateTime', 'legalAgeGroupClassification', 'licenseAssignmentStates', 'licenseDetails' ,
                       'mail' , 'mailFolders' , 'mailNickname', 'managedAppRegistrations', 'managedDevices', 'manager' , 'memberOf' , 'messages', 'mobilePhone', 'mySite',
                       'oauth2PermissionGrants' ,  'officeLocation', 'onenote', 'onlineMeetings' , 'onPremisesDistinguishedName', 'onPremisesDomainName',
                       'onPremisesExtensionAttributes', 'onPremisesImmutableId', 'onPremisesLastSyncDateTime',   'onPremisesProvisioningErrors', 'onPremisesSamAccountName',
                       'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName', 'otherMails', 'outlook', 'ownedDevices', 'ownedObjects',
                       'passwordPolicies', 'passwordProfile', 'pastProjects', 'people', 'photo', 'photos', 'planner', 'postalCode', 'preferredLanguage',
                       'preferredName', 'presence', 'provisionedPlans', 'proxyAddresses',  'registeredDevices', 'responsibilities',
                       'schools', 'scopedRoleMemberOf', 'settings', 'showInAddressList', 'signInSessionsValidFromDateTime', 'skills', 'state', 'streetAddress', 'surname',
                       'teamwork', 'todo', 'transitiveMemberOf', 'usageLocation', 'userPrincipalName', 'userType')]
        [String[]]$UserProperties =  $Script:DefaultUserProperties ,
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
        [Parameter(parameterSetName='Apps') ]
        [String]$AppName,
        #limits searches for channels by name. Other items can't be filtered by name ...  perhaps notebooks can but a group only has one.
        [Parameter(parameterSetName='Channels')]
        [ArgumentCompleter([ChannelCompleter])]
        [String]$ChannelName,
         #Field(s) to select for the group: ID and displayname are always included
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
        [switch]$Mine,
        [Parameter(Mandatory=$true, parameterSetName='BareGroups')]
        [switch]$NoTeamInfo
    )
    begin   {
        function sendUsersAndGroups {
            param (
                [parameter(valueFromPipeline=$true)]
                $ug ,
                $DisplayName
            )
            process {
                if     ($ug.'@odata.type' -match 'group$') {
                        $null =  $ug.Remove('@odata.type'),  $ug.Remove('@odata.context'),  $ug.remove('@odata.id'),  $ug.remove('creationOptions')
                        New-Object -Property  $ug -TypeName MicrosoftGraphGroup |
                            Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
                }
                elseif ($ug.'@odata.type' -match 'user$') {
                        $null =  $ug.Remove('@odata.type'),  $ug.Remove('@odata.context')
                        New-Object -Property $ug -TypeName MicrosoftGraphUser |
                            Add-Member -PassThru -NotePropertyName GroupName  -NotePropertyValue $displayname
                }
            }
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx toDo check scopes - Scopes Group.Read.All, Files.Read, Sites.Read.All, Notes.Create, Notes.Read, depending on params passed.
        # if we didn't get passed a group but we did get asked for something about a group or groups then  get the current user's groups,
        if      ($PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters}) -and -not $ID)  {
                       $ID = Get-GraphUser -Current -Groups
        }
        # If we got nothing return the list,
        elseif  (-not  $ID) { Get-GraphGroupList ; return  }
        # if we got a single string that looks like a name (not a GUID) resolve it - it may be a wildcard
        # The proceed as if ID had been passed as one or more groups.
        elseif  ($ID -is [string] -and  $ID -notmatch $guidregex)   {
                       $ID = Get-GraphGroupList -Name $id
                       if (-not $id) {Write-Warning "'$($PSBoundParameters['id'])' did not match any groups.";  return}
        }
        # We'll loop through an array (or single object) with either GUIDs or objects.
        foreach ($i in $ID) {
            <# not all teams have team set in resource provisioning options
                if  ($i.ResourceProvisioningOptions -is [array] -and
                 $i.ResourceProvisioningOptions -notcontains "Team" -and
                ($Channels -or $ChannelName -or $Apps)) {
                Write-Verbose "$($i.DisplayName) is a group but not a team"
                continue
            }#>
            if       ($i.id -and -not $i.DisplayName) {
                      $i = Invoke-GraphRequest "$GraphUri/groups/$($i.id)"
            }
            elseif   ($i -is [string] -and  $i -match $guidregex)   {
                      $i = Invoke-GraphRequest "$GraphUri/groups/$i" }
            elseif   ($i -is [string]) {
                      $i = Get-GraphGroupList -Name $i}
            if (-not ($i.DisplayName -and $i.id)) {
                Write-Warning 'Could not resoulve the group' ; continue
            }
            if ($Owners -or $Members -or $TransitiveMembers) {
                if ('id'          -notin $UserProperties) {$UserProperties += 'id'}
                if ('displayName' -notin $UserProperties) {$UserProperties += 'displayName'}
                $UserPropertiesSelect = '?$Select='  +   ($UserProperties -join ',')
            }
            $displayname  =  $i.DisplayName
            $groupid      =  $teamid = $i.id
            $groupURI     = "$GraphUri/groups/$groupid"
            $teamURI      = "$GraphUri/teams/$teamid"
            try   {
                #For each of the switches get the data from /groups{id}/whatever or /teams/{id}.whatever
                #Add a type to PS Type names so we can format it, and add any properties we expect to want later.
                Write-Progress -Activity 'Getting Group Information' -CurrentOperation $displayname
                if     ($Site)               {
                    $uri = ("$groupURI/sites/root?expand=drives,sites,lists(expand=columns,contenttypes,drive)")
                    $result  =  Invoke-GraphRequest -Uri $uri -ExcludeProperty 'sites@odata.context', '@odata.context', 'drives@odata.context', 'lists@odata.context' -AsType ([MicrosoftGraphSite]) |
                            Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    foreach ($siteObj in $result) {
                        foreach ($l in $siteObj.lists) {
                            Add-Member -InputObject $l -NotePropertyName SiteID    -NotePropertyValue  $siteObj.id
                            Add-Member -InputObject $l -NotePropertyName ParentUrl -NotePropertyValue  $siteObj.weburl
                        }
                        $siteobj
                    }
                    continue
                }
                elseif ($Calendar)           {
                    Invoke-GraphRequest -Uri  "$groupURI/calendar" -ExcludeProperty "@odata.context" -AsType ([MicrosoftGraphCalendar]) |
                        Add-Member -PassThru -NotePropertyName GroupID      -NotePropertyValue $groupid   |
                        Add-Member -PassThru -NotePropertyName CalendarPath -NotePropertyValue "groups/$groupid/Calendar" |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    continue
                }
                elseif ($Drive)              {
                    $uri = ("$groupURI/drive" + '?$expand=root($expand=children)' )
                    Invoke-GraphRequest  -Uri $uri -ExcludeProperty "@odata.context", "root@odata.context" -AsType ([MicrosoftGraphDrive]) |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    continue
                }
                elseif ($Owners)             {
                    Invoke-GraphRequest -Uri  "$groupURI/Owners$UserPropertiesSelect"  -AllValues             | sendUsersAndGroups -Displayname $displayname
                    continue
                }
                elseif ($Members)            { #can do group ?$expand=Memebers, the others don't expand
                    Invoke-GraphRequest  -Uri "$groupURI/members$UserPropertiesSelect"  -AllValues            |  sendUsersAndGroups -Displayname $displayname
                    continue
                }
                elseif ($TransitiveMembers)  {
                    Invoke-GraphRequest  -Uri  "$groupURI/TransitiveMembers?$UserPropertiesSelect" -AllValues |  sendUsersAndGroups -Displayname $displayname
                    continue
                }
                #xxxx Allow group properties for memeber of ?
                elseif ($MemberOf)           {
                    Invoke-GraphRequest  -Uri  "$groupURI/memberof"  -AllValues                               |  sendUsersAndGroups -Displayname $displayname
                    continue
                }
                elseif ($TransitiveMemberOf) {
                    $usersAndGroups += Invoke-GraphRequest  -Uri  "$groupURI/TransitiveMemberof"  -AllValues  |  sendUsersAndGroups -Displayname $displayname
                    continue
                }
                elseif ($Notebooks)          {
                    #if groups can have more than one book , then add if name ... uri = blah + "?`$expand=sections&`$filter=startswith(tolower(displayname),'$name')"
                    $uri = $groupURI + '/onenote/notebooks?$expand=sections'
                    $response = Invoke-GraphRequest  -Uri $uri -ValueOnly -ExcludeProperty 'sections@odata.context'  -AsType ([MicrosoftGraphNotebook]) |
                        Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    foreach ($bookobj in $response) {
                        #Section fetched this way won't have parentNotebook, so make sure it is available when needed
                        foreach ($s in $bookobj.sections) {$s.ParentNotebook = $bookobj}
                        $bookobj
                    }
                    continue
                }
                elseIf ($Plans)              {
                    #would like to have expand details here but it only works with a single plan.
                    try {
                        $result  = Invoke-GraphRequest  -Uri  "$groupURI/planner/plans"  -AllValues -ExcludeProperty  "@odata.etag" -AsType ([MicrosoftGraphPlannerPlan]) |
                            Add-Member -PassThru -NotePropertyName GroupName    -NotePropertyValue $displayname
                    }
                    catch             { Write-Warning "Could not get plans for $($ID.DisplayName)." ;   continue}
                    if (-not $result) { Write-Host "The team $($ID.DisplayName) has not created any plans" ;   continue}
                    $dirObjectsHash = @{}
                    if ($i.displayName) {$dirObjectsHash[$teamId] = $i.displayName}
                    @() + $result.owner + $result.createdby.user.id  |ForEach-Object  {
                        if (-not $dirObjectsHash[$_]) {
                            $dirObjectsHash[$_] = (Invoke-GraphRequest  -Uri "$GraphUri/directoryobjects/$_").displayname
                        }
                    }
                    foreach ($r in $result) {
                        Add-Member -PassThru  -InputObject $r  -NotePropertyName OwnerName   -NotePropertyValue $dirObjectsHash[$r.owner] |
                        Add-Member -PassThru                   -NotePropertyName CreatorName -NotePropertyValue $dirObjectsHash[$r.createdBy.user.id]
                    }
                    continue
                }
                elseif ($Threads)            {
                    Invoke-GraphRequest  -Uri  "$groupURI/threads"  -AllValues -AsType ([MicrosoftGraphConversationThread]) |
                        Add-Member -PassThru -NotePropertyName Group       -NotePropertyValue $groupid  |
                        Add-Member -PassThru -NotePropertyName GroupName   -NotePropertyValue $displayname
                    continue
                }
                elseif ($Conversations)      {
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
                        $ChannelName)
                                             {
                    if ($ChannelName)   { $uri =  "$teamURI/channels?`$filter=startswith(tolower(displayname), '$($ChannelName.ToLower())')"}
                    else                { $uri =  "$teamURI/channels"}
                    Invoke-GraphRequest  -Uri $uri -ValueOnly -As ([MicrosoftGraphChannel]) -ExcludeProperty "tenantId" |
                        Add-Member -PassThru -NotePropertyName Team      -NotePropertyValue $teamid |
                        Add-Member -PassThru -NotePropertyName TeamName  -NotePropertyValue $displayname
                    continue
                }
                elseif ($Apps -or
                        $AppName)
                                             {
                    $uri = $teamURI + '/installedApps?$expand=teamsAppDefinition'
                    if ($AppName) { $uri = $URI + '&$filter=' +
                                    "startswith(tolower(teamsappdefinition/displayname),'$($AppName.ToLower())')"
                    }
                    Invoke-GraphRequest  -Uri $uri -ValueOnly  -As ([MicrosoftGraphTeamsAppDefinition]) |
                        Add-Member -PassThru -NotePropertyName Team      -NotePropertyValue $teamid |
                        Add-Member -PassThru -NotePropertyName TeamName  -NotePropertyValue $displayname
                    continue
                }
                elseif ($Select)             {
                    $SelectList = (@('id','displayName') + $Select ) -join','
                    Invoke-GraphRequest      -Uri "$groupuri`?`$Select=$SelectList" -As ([MicrosoftGraphGroup]) -ExcludeProperty '@odata.context'
                }
                else                         {
                    $g =  Invoke-GraphRequest    -Uri $groupuri -As ([MicrosoftGraphGroup]) -ExcludeProperty '@odata.context','creationOptions'
                    #removed $expand=Members as it only returns the first 20.
                    if ($g.resourceProvisioningOptions -notcontains 'Team' -or
                        $MyInvocation.InvocationName -ne 'Get-GraphTeam' -or $NoTeamInfo) { $g }
                    else {
                        $t = Invoke-GraphRequest -Uri  "$teamURI"                   -As ([MicrosoftGraphTeam])  -ExcludeProperty '@odata.context'
                      # $t.members = $g.Members
                        $t
                    }
                }
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
        Write-Progress -Activity 'Getting Group/Team information' -Completed
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
    param   (
        #The name of the Group / Team
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Name,

        #Unless specified, groups will be mail enabled "unfied" / Microsoft365 groups
        #The Graph API doesn't allow mail-enabled & security-enabled,  or mail-disabled & unified
        #Only unified groups can be made into teams. Unified groups can only contain users,
        #Security groups can contain other security principals
        [parameter(ParameterSetName='Security',Mandatory=$true)]
        [Switch]$AsSecurity,

        #If specified allows Azure AD roles can be assigned to the group. This forces visibility to be private, and can't be changed.
        [parameter(ParameterSetName='Security')]
        [Switch]$AsAssignableToRole,

        #New-GraphGroup only enables teams functonality if -AsTeam is specified. Calling as New-GraphTeam defaults AsTeam to true
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


        #Some settings can only be set when the group is created By default any user in the organization can post conversations to the group,
        #groups are visible and discoverable in Outlook, New Group-Members are not subscribed to conversations and a welcome mail is sent.
        [validateSet('AllowOnlyMembersToPost', 'HideGroupInOutlook', 'SubscribeNewGroupMembers', 'WelcomeEmailDisabled')]
        [string[]]$BehaviorOptions,

        #if specified group will be added without prompting
        [Switch]$Force
    )
    ContextHas -WorkOrSchoolAccount -BreakIfNot

    if (Invoke-GraphRequest -Uri "$GraphUri/groups?`$filter=displayname eq '$Name'" -ValueOnly) {
        throw "There is already a group with the display name '$Name'." ; return
    }
    #Server-side is case-sensitive for [most] JSON so make sure hashtable names and constants have the right case!
    if (-not $MailNickName) {$MailNickName    = $Name -replace "\W",'' }
    $settings = @{  'displayName'             = $Name
                    'mailNickname'            = $MailNickName
                    'mailEnabled'             = -not $AsSecurity
                    'securityEnabled'         = $AsSecurity -as [bool]
                    'visibility'              = $Visibility.ToLower()
                    'groupTypes'              = @()
    }
    if (-not $AsSecurity ) {
          $settings.groupTypes               += 'Unified'
          if ($MyInvocation.InvocationName  -eq 'New-GraphTeam' -and -not $PSBoundParameters.ContainsKey('AsTeam')) {
              $AsTeam = $true
          }
    }
    elseif ($AsAssignableToRole) {
          $settings['isAssignableToRole']     = $true
          $settings['visibility']             ='Private'
    }
    if ($BehaviorOptions) {
        $settings["resourceBehaviorOptions"]  = $BehaviorOptions
    }
    if ($Description) {
          $settings['description']            = $Description
    }
    #if we got owners or users with no ID, fix them at the end, if they have an ID add them now
    if ($Members) {
        $settings['members@odata.bind']       = @();
        foreach ($m in $Members) {
            if  ($m.id) {$settings['members@odata.bind'] += "$GraphUri/users/$($m.id)"}
            else        {$settings['members@odata.bind'] += "$GraphUri/users/$m"}
        }
    }
    #If we make someone else the owner of the group, we can't make it a team,
    #so parameter sets should ensure we can't get owners here if we are making a team.
    if ($Owners) {
        $settings['owners@odata.bind']        = @()
        foreach    ($o in $Owners)  {
            if     ($o.id) {$settings['owners@odata.bind']  += "$GraphUri/users/$($o.id)"}
            else{           $settings['owners@odata.bind']  += "$GraphUri/users/$o"}
        }
    }
    $webparams = @{
        Method      = 'Post'
        Uri         = "$GraphUri/groups"
        Body        = (ConvertTo-Json $settings)
        ContentType = 'application/json'
    }
    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.shouldprocess($Name,"Add new Group")) {
        Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Adding Group $Name"
        $group = Invoke-GraphRequest @webparams -As ([MicrosoftGraphGroup]) -Exclude "@odata.context","creationOptions"
        if     ($Owners) { $
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Setting Group ownership on $Name"
            Owners | Add-GraphGroupMember -Group $group -AsOwner -Force
        }
        if     (-not $AsTeam) {
            Write-Progress -Activity 'Creating Group/Team' -Completed
            return $group
        }
        elseif ($Group.GroupTypes) {
            Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name"
            $webparams.Uri   +=  "/$($group.id)/team"
            $webparams.Method = 'Put'
            $webparams.Body   = '{ }'
            $TimeToStop = [datetime]::Now.AddMinutes(2)
            $retries = 0
            do {
                try   {
                    $team     = Invoke-GraphRequest @webparams -Exclude '@odata.context' -As ([MicrosoftGraphTeam]) |
                                    Add-Member -PassThru -NotePropertyName Mail -NotePropertyValue $group.Mail
                }
                catch {
                    $retries ++
                    Write-Progress -Activity 'Creating Group/Team' -CurrentOperation "Team-enabling Group $Name" -status "Retries $retries"
                    Start-Sleep -Seconds 5
                }
            }
            until   ($team -or [datetime]::now -gt $TimeToStop)
            if (-not $team ) {
                Write-Warning "Group was created, but could not elevate it to a team."
                return $group
            }
            $team.Description = $group.description
            $team.Members     = $group.members     #Check that all users are returned if more than 20 added on creation.
            $team.visibility  = $group.visibility

            Write-Progress -Activity 'Creating Group/Team' -Completed
            $team
        }
    }
}

function idfromteam                 {
    <#
    .synopsis
        Helper function to return a team ID - filters out not teams-enabled groups
    .Description
        if $team is null or emptry, returns nothing.
        if it has an ID property returns the ID with no further checks
        if it is a string holding a GUID, it  it returns the string with no further checks
        if it is any other string searches for any group with a matching display name and returns the result(s)
    #>
    param   (
             $Team
    )
    if (-not $Team)                   {return}
    if      ($Team.id)                {return $Team.id}
    if      ($Team -is [string] -and
             $Team -match $GUIDRegex) {return $Team}
    elseif  ($Team -is [string])      {
        Invoke-GraphRequest  ("$GraphUri/groups?`$select=id,resourceProvisioningOptions,displayname&`$filter=" + (FilterString $Team)) -ValueOnly |
            ForEach-Object  {if ("Team" -in $_.resourceProvisioningOptions) {$_.id}}
    }
}

function idfromgroup                {
    <#
    .synopsis
        Helper function to return a Group ID - filters out not Groups-enabled groups
    .Description
        if $Group is null or emptry, returns nothing.
        if it has an ID property returns the ID with no further checks
        if it is a string holding a GUID, it  it returns the string with no further checks
        if it is any other string searches for any group with a matching display name and returns the result(s)
    #>
    param   (
             $Group
    )
    foreach ($g in $Group) {
        if      ($g.id)                {$g.id}
        if      ($g -is [string] -and
                 $g -match $GUIDRegex) {$g}
        elseif  ($g -is [string])      {
            Invoke-GraphRequest  ("$GraphUri/groups?`$select=id,resourceProvisioningOptions,displayname&`$filter=" + (FilterString $g)) -ValueOnly |
                ForEach-Object  {$_.id}
        }
    }
}

function Set-GraphDefaultGroup      {
    <#
    .Synopsis
        Sets the default paramater for group or team in most functions which take one.
    .Description
        Takes a group as a parameter or via the pipeline.
        If a string is passed it will try to get a matching group from Get-GraphGroup,
        a string may be a wildcard for a group name - provided that it only finds one matching group.
        If the group has been provisioned as a team then it will be the default for commands which take a -Team parameter.
        The primary purpose is to avoid specifying a Group/Team when working with messages, calendar / planner / team channels,
        but working with the group itself or its membership it is safer not to default the selection, so no defaults
        are set for for Set-Team, Set-Group, Get- Remove-Group Remove-GroupMember or Add-GroupMember or Import and Export
    .Example
        > Set-GraphDefaultGroup Accounts
        >  Get-GraphChannel
        Display Name description
        ----------- -----------
        General      The Accounts Department
        Mccaw        For anything about project Mccaw

        The first command sets the default group - because "Accounts" has been provisioned as a team,
        it becomes the default team for Get-GraphChannel
    #>
    [Alias('Set-GraphDefaultTeam')]
    Param (
        #The group to set as the default for other commands
        [Alias('Team')]
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
        [AllowEmptyString()]
        [AllowNull()]
        [ArgumentCompleter([GroupCompleter])]
        $Group
    )
    @('*-GraphChannel*:Team','Add-Graph*Tab:Team','*New-GraphTeamPlan:Team',
      '*-GraphEvent:Group','*-GraphGroupThread:Group','Get-GraphGroupConversation:Group') | ForEach-Object {
       $null = $Global:PSDefaultParameterValues.Remove($_)
    }
    if ($Group -is [string]) {$Group = Get-GraphGroup $Group -NoTeamInfo}
    if (-not $Group -or ($group.Count -gt 1)){
        Write-Warning 'Could not resolve the information provided to a single group.'
    }
    elseif ($Group.securityEnabled) {
        Write-Warning "Only commands for unified groups accept a default, but $($Group.DisplayName) is a security group."
    }
    else {
        $Global:PSDefaultParameterValues[              '*-GraphEvent:Group'] = $Group
        $Global:PSDefaultParameterValues[        '*-GraphGroupThread:Group'] = $Group
        $Global:PSDefaultParameterValues[      'Send-GraphGroupReply:Group'] = $Group
        $Global:PSDefaultParameterValues['Get-GraphGroupConversation:Group'] = $Group
        if ($Group.ResourceProvisioningOptions -contains 'team' -or $Group -is [MicrosoftGraphTeam]){
            $Global:PSDefaultParameterValues[       '*-GraphChannel*:Team']  = $Group
            $Global:PSDefaultParameterValues[         'Add-Graph*Tab:Team']  = $Group
            $Global:PSDefaultParameterValues[    '*New-GraphTeamPlan:Team']  = $Group
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
       Get-GraphGroupList -Name consult* | Set-GraphGroup -Description "People who do consulting work" -Force
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
                  if     ($g.id -and $g.displayname) {$g}
                  else   {Get-GraphGroup $g  -NoTeamInfo -ErrorAction Stop  }
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
                 Method      = 'Patch'
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
        >Get-GraphTeam accounts* | Set-GraphTeam -AllowGiphy:$false
        Gets a the team(s) with a name that begins with accounts, and turns off Giphy content
        Note the use of -SwitchName:$false.
    #>
    [Cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
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
    begin   {
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
              $webparams  = @{
               'Method'      = 'PATCH'
               'ContentType' = 'application/json'
               'Body'        = ConvertTo-Json $settings -Depth 10
              }
              Write-Debug $webparams['Body']
        }
        else {Write-Warning -Message "Nothing to set"}
    }
    process {
        if (-not $webparams) {return}
        Write-Progress -Activity "Updating Team" -Status "Checking team is valid"
        if     ($Team.id -and $Team.DisplayName -and
                ($Team -is [MicrosoftGraphTeam] -or
                $Team.resourceProvisioningOptions -contains 'Team')) {
            $group = $team
        }
        elseif ($Team.id) {
            $group =  Invoke-GraphRequest -method get "$GraphUri/groups/$($Team.id)"
        }
        elseif ($Team -is [string] -and $team -match $GuidRegex ) {
            $group =  Invoke-GraphRequest -method get "$GraphUri/groups/$Team"
        }
        elseif ($Team -is [string]  ) {
            $group =  Get-GraphGroupList -Name $Team
        }
        if ( $group.id -and $group.displayName -and
            ($group.resourceProvisioningOptions -contains 'Team' -or
            $group -is [MicrosoftGraphTeam] )) {
            $webparams['Uri'] = "$GraphUri/teams/$($group.id)"
        }
        else   {
            Write-Progress -Activity "Updating Team" -Completed
            Write-Warning -Message 'Could not resolve the team';
            return
        }
        if ($PSCmdlet.ShouldProcess($group.displayName,'Update Team settings')) {
            Write-Progress -Activity "Updating Team" -CurrentOperation $group.displayName -Status "Committing changes"
            Invoke-GraphRequest @webparams
            Write-Progress -Activity "Updating Team" -Completed
        }
    }
}

function Remove-GraphGroup          {
    <#
      .Synopsis
        Removes a group/team
      .Description
        Requires consent to use the Group.ReadWrite.All scope.
        The group may remain visible for a short time.
        Deleted groups can be recovered using Get-GraphDeletedObject and Restore-GraphDeletedObject
    #>
    [Cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    [Alias("Remove-GraphTeam")]
    param   (
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
                  if     ($g.id -and $g.displayname) {$g}
                  else   {Get-GraphGroup $g  -NoTeamInfo -ErrorAction Stop  }
        }
        foreach ($g in $Group){
            if ($Force -or $PSCmdlet.Shouldprocess("$($g.displayname)","Delete Group")) {
                Write-Progress -Activity "Deleting Group" -CurrentOperation $g.displayname
                Invoke-GraphRequest -Method Delete  -Uri "$GraphUri/groups/$($g.id)/"
                Write-Verbose "REMOVED GROUP $($g.displayname)"
            }
        }
        Write-Progress -Activity "Deleting Group" -Completed
    }
}

function Add-GraphGroupMember       {
    <#
      .Synopsis
        Adds a user (or group) to a group/team as either a member or owner.
      .Description
        Because the group may be a team the this command has alias of Add-GraphTeamMember.
        it requires consent to use the Group.ReadWrite.All, Directory.ReadWrite.All, or
        Directory.AccessAsUser.All scope.
      .Example
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
        [ArgumentCompleter([UPNCompleter])]
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
                  if     ($g.id -and $g.displayname) {$g}
                  else   {Get-GraphGroup $g  -NoTeamInfo -ErrorAction Stop  }
            }
        }
        $memberHash = @{}
        foreach ($g in  $Group) {
            $memberHash[$g.id] = Get-GraphGroup -ID $g.id -Members | Select-Object -ExpandProperty id
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        foreach ($g in $Group) {
            #group(s) resolved in begin block so should have an ID and display name.
            if ($AsOwner) {$uri   = "$GraphUri/groups/$($g.ID)/owners/`$ref" }
            else          {$uri   = "$GraphUri/groups/$($g.ID)/members/`$ref"}
            #There is more efficient way to add many users, but that isn't the main use case so using one call per user.
            #That optimization would be to collect piped user in the process block and make one call per group in the end block
            foreach ($m in $member) {
                #if we weren't passed as a user as a an object, resolve what we did get ...
                if   (-not ($m.id -and $g.displayname))  {
                    try   {$m     = Get-GraphUser -User $m -Select displayname}
                    catch {throw "Could not get a user matching $m"; return }
                    if (-not $m) {throw "Could not get a member ID"; return }
                }
                #Getting a group gets the members but we can't expand members AND owners.
                if ($m.id -in $memberHash[$g.id] -and -not $AsOwner) {
                    Write-Warning "'$($m.displayName)' is already a member of the group '$($g.displayname)'."
                    continue
                }
                $body = ConvertTo-Json @{'@odata.id' = "$GraphUri/directoryObjects/$($m.id)"   }
                Write-Debug $body
                if ($Force -or $PSCmdlet.shouldprocess($m.displayname,"Add to Group '$($g.displayname)'")) {
                    try   {  Invoke-GraphRequest -Method post -Uri $uri -Body $body -ContentType 'application/json'  }
                    catch { #if the group is was a variable, the member list may not be current, and we don't validate new owners
                        if ($_.Exception.Response.StatusCode.value__ -eq 400) {
                            Write-Warning "Adding to group $($g.displayname) returned 'Bad Request' - $($m.displayname) may be assigned to the group already."
                        }
                        else {throw $_}
                    }
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
        Removes a user from the owners of a group without prompting for confirmation.
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
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #A group object with an ID field, or a user object, user ID or UPN
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [ArgumentCompleter([UPNCompleter])]
        $Member,
        #If specified the member will be removed from the owners rather than members
        [switch]$FromOwners,
        #If specified the member will be removed without prompting for confirmation
        [switch]$Force
    )
    begin   {
        #ensure we have an ID for the group(s) we were passed. If we got a GUID in a string, we'll confirm it's a group and get the display name.
        $Group = foreach ($g in $Group) {
              # at least this user must be in the group.
              if     ($g.id -and $g.displayname ) {$g}
              else   {Get-GraphGroup $g  -NoTeamInfo -ErrorAction Stop  }
        }
        $memberHash = @{}
        foreach ($g in  $Group) {
            $memberHash[$g.id] = Get-GraphGroup -ID $g.id -Members | Select-Object -ExpandProperty ID
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

                if ($FromOwners)                       {$uri = "$GraphUri/groups/$($g.ID)/owners/$($m.id)/`$ref" }
                elseif ($m.id -in  $memberHash[$g.id]) {$uri = "$GraphUri/groups/$($g.ID)/members/$($m.id)/`$ref"}
                else {Write-Warning "'$($m.displayName)' is not a member of the group '$($g.displayname)'." ; continue }
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
    param   (
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        #One or more group(s) to export
        [ArgumentCompleter([GroupCompleter])]
        $Group,
        #Destination for CSV output
        $Path,
        #If specified , output will be in Group name order (default is User name.)
        [switch]$OrderByGroup
    )
    begin   {
    $list = @()
    }
    process {
        foreach ($g in $group) {
            $list += Get-GraphGroup $g -Members |
                    Select-Object -Property @{n='Action';  e={'Add'}} ,
                                            @{n='MemberOf';e={$_.groupName}},
                                            UserPrincipalName,
                                            Displayname
        }
    }
    end     {
        if ($OrderByGroup) {$list = $list | Sort-Object -Property Memberof, UserPrincipalName }
        else               {$list = $list | Sort-Object -Property UserPrincipalName, Memberof }
        if (-not $path) {return $list}
        else   {$list | Export-Csv -Path $Path -NoTypeInformation }
    }
}

function Import-GraphGroupMember    {
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
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Usually the command will prompt for confirmation -Force disables this primpt
        [switch]$Force
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
        $groups = ($List | Group-Object -NoElement -Property memberof).Name
        foreach ($g in $groups) {
            $w = $Null
            $Members = (Get-GraphGroup $g -Members -WarningAction SilentlyContinue -WarningVariable W).UserPrincipalName
            if ($w) {Write-Warning "Skipping '$g' it did not match a group." ; continue}
            foreach    ($member in $list.where({$_.memberof -eq $g}) ) {
                $upn =  $member.UserPrincipalName
                if    (($member.Action -eq 'Add' -and $upn -notin $Members) -and
                       ($force -or $PSCmdlet.ShouldProcess($upn,"Add user to group '$g'"))) {
                        Add-GraphGroupMember -Force -Group $g -Member $upn
                        Write-Verbose "Added $UPN user to group'$g'"
                }
                elseif (($member.Action -eq 'Remove' -and $upn -in $Members) -and
                        ($force -or $PSCmdlet.ShouldProcess($upn,"Remove member from group '$g'"))){
                        Remove-GraphGroupMember -Force -Group $g -Member $upn
                        Write-Verbose "Removed $UPN user from group'$g'"
                }
                else   {Write-Verbose -Message "No action needed for $g / $upn"}
            }
        }
    }
}

function Import-GraphGroup          {
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
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Disables any prompt for confirmation
        [switch]$Force
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
        $existingGroups = Get-GraphGroupList
        $existingNames  = $existingGroups.DisplayName
        foreach ($group in $list) {
            $displayName = $group.DisplayName
            if (($group.Action -eq 'Remove' -and $displayname -in $existingNames) -and
                ($force -or $PSCmdlet.ShouldProcess($displayname,"Remove group "))){
                        Remove-GraphGroup -Force -Group $displayname
                        Write-Host "Removed group'$displayname'"
            }
            elseif (($group.Action -eq 'Add' -and $displayname -notin $existingNames) -and
                ($force -or $PSCmdlet.ShouldProcess($displayname,"Add new group"))){
                    $params = @{Force=$true; Name=$displayName}
                    if ($group.Type -match 'Security') {$params['AsSecurity'] = $true}
                    if ($group.Type -match 'Team')     {$params['AsTeam'] = $true}
                    if ($group.Visibility)             {$params['Visibility'] = $group.Visibility}
                    if ($group.Description)            {$params['Description'] = $group.Description}
                    $g = New-GraphGroup @params
                    if ($g -isnot [MicrosoftGraphTeam]) {$g}
                    else {
                        $stoptime = [datetime]::Now.AddMinutes(2);
                        do    {$g = Get-GraphGroup $g.id}
                        until ($g.ResourceProvisioningOptions.Count -or [datetime]::now -gt $stoptime -or (start-sleep -Seconds 5))
                        $g
                    }
                    Write-Verbose "Added group'$displayName'"
            }
            else {  Write-Verbose "No action taken for group '$displayName'"}
        }
    }
}

function New-GraphTeamPlan          {
    <#
      .Synopsis
        Creates new a plan (in the planner app) for a team.
    #>
    [Alias('New-GraphGroupPlan')]
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
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        $settings =  @{owner = (idfromgroup $team ) }

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
        [Parameter(ParameterSetName='InTeam',Position=0)]
        [Parameter(ParameterSetName='OneConversation',Position=0)]
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
        if     ($Conversation.Group)  {
                $groupID = $Conversation.Group
        }
        else   {$groupID = idfromgroup $Group}
        if     (-not $groupID -or $groupID.count -gt 1) {
                Write-Warning -Message 'Could not resolve group ID'; return
        }
        if     ($Conversation.id) {$Conversation = $Conversation.id}
        if     ($Threads) {
            $uri    = "$GraphUri/groups/$groupID/conversations/$conversation/Threads"
            Invoke-GraphRequest  -Uri $uri -ValueOnly -AsType ([MicrosoftGraphConversationThread]) |
                Add-Member -PassThru -NotePropertyName Group        -NotePropertyValue $GroupID   |
                Add-Member -PassThru -NotePropertyName Conversation -NotePropertyValue $Conversation
        }
        else   {
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
        [Parameter(ParameterSetName='SingleThread', Position=1, ValueFromPipeline=$true, Mandatory=$true)]
        $Thread,
        #The group holding the thread (s), if thread is either not passed or is just the ID of a thread.
        [Alias("Team")]
        [Parameter(ParameterSetName='GroupThreads', Position=0)]
        [Parameter(ParameterSetName='SingleThread', Position=0)]
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
        else   {$groupID = idfromgroup $Group}
        if     (-not $groupID -or $groupID.count -gt 1) {
                Write-Warning -Message 'Could not resolve group ID'; return
        }

        if     ($Thread.id)                    {$threadID = $Thread.id}
        elseif ($Thread -is [string] -and
                $Thread -match '\S{100}')      {$threadID = $Thread}
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
    param   (
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
    $groupID = idfromgroup $Group
    if     (-not $groupID -or $groupID.count -gt 1) {
                Write-Warning -Message 'Could not resolve group ID'; return
    }
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
        Get-GraphGroup consultants -Threads | where topic -eq "Today's tests..."  | Remove-GraphGroupThread
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
        else   {$groupID = idfromgroup $Group}
        if     (-not $groupID -or $groupID.count -gt 1) {
                Write-Warning -Message 'Could not resolve group ID'; return
        }

        if     ($Thread.ID)           {$threadid = $Thread.id  }
        elseif ($Thread -is [string]) {$threadid = $Thread     }
        else   {Write-Warning 'Could not resolve the Thread ID' ; return}

        $uri    =  "$GraphUri/groups/$GroupID/threads/$threadID"
        Write-Progress -Activity "Deleting thread" -Status "Checking existing thread"
        try   {$thread  = Invoke-GraphRequest -Method Get $uri }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-warning 'Thread not found, either the ID was wrong or it may have been deleted already'
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
            Write-Progress -Activity "Deleting thread" -Status "Sending delete instruction" -CurrentOperation $thread.topic
            Invoke-GraphRequest -Method Delete $uri
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
        > Set-GraphDefaultGroup 'Consultants'
        > ...
        > $post = Get-GraphGroupThread -Topic  "Today's tests..."  -Posts | select -last 1
        >Send-GraphGroupReply $post -Content "Please join a celebration of the successful test at 4PM"
        This example finds threads for the consultants group, Isolates the one with the topic of
        "Today's Tests..." and finds the last post in the thread. It then posts a reply with the content as plain text.
        This example stores the Post between the two commands but they could be piped together as in the previous example
      .link
        Add-GraphGroupThread
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param   (
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

    if     ($Post.Group)          {$groupID  = $Post.group  }
    elseif ($Thread.Group)        {$groupID  = $Thread.group}
    else                          {$groupID = idfromgroup $Group}
    if     (-not $groupID -or $groupID.count -gt 1) {
                Write-Warning -Message 'Could not resolve group ID'; return
    }
    if     ($Post.Thread)         {$threadID = $Post.Thread}
    elseif ($Thread.ID)           {$threadID = $Thread.id  }
    elseif ($Thread -is [String]) {$threadID = $Thread.id  }
    else   {Write-warning -Message 'Could not resolve the Thread ID.' ; return}

    if     ($Post.Topic)          {$Topic= $Post.Topic}
    elseif ($Thread.Topic)        {$Topic= $Thread.Topic}
    else                          {$Topic= 'Unknown Topic'}

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

    if ($Force -or $PSCmdlet.Shouldprocess($topic,"Reply to thread")) {
        $uri     += "/Reply"
        Write-Progress -Activity 'Posting reply to thread' -Status 'sending reply'
        Invoke-GraphRequest -Method Post -Uri $URI  -Body $Json -ContentType "application/json"
        Write-Progress -Activity 'Posting reply to thread' -Completed
    }
}

function Get-ChannelMessagesByURI   {
    <#
      .synopsis
        Helper function to add get and expand messages or replies to messages
    #>
    param   (
        [parameter(Position=0,ValueFromPipeline=$true)]
        $URI,
        $Top = 20,
        $channelID = "",
        $teamID = ""
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
            ForEach-Object {
                $msgObj = New-Object -TypeName MicrosoftGraphChatMessage -Property $_
                if ($channelID -and -not $msgobj.ChannelIdentity.ChannelId) {
                    $msgobj.ChannelIdentity.ChannelId = $channelID
                }
                if ($Team -and -not $msgobj.ChannelIdentity.TeamId  ) {
                    $msgobj.ChannelIdentity.TeamId = $teamID
                }
                $msgObj
            } |  Sort-Object -Property createdDateTime -Descending | Select-Object -First $top
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
        }#>
    }
}

function Get-GraphChannel           {
    <#
      .Synopsis
        Gets details of a channel, or its Tabs or messages shown in Teams
      .Example
        >Get-GraphTeam  consultants -ChannelName general | Get-GraphChannel -Tabs
        Gets channels for the team(s) with a name beginning 'Consultants' and selects channel(s)
        with a name beginning "general"; then gets the tabs shown in Teams for this channel
      .Example
        > Set-GraphDefaultGroup 'Consultants'
        > ...
        > Get-GraphChannel 'General' -Messages
        If the default group is set to a suitable team, it is possible to tab complete the channel name
        and ther is no need specify the team
      .Example
        >Get-GraphChannel -Team accounts -channel general -Messages
        This specifies a non-default team, and gets messages from the teams general channel.
      .Example
        >Get-GraphChannel -Team $t
       Gets the basic channel information for team.
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeamChannel")]
    param   (
        #The ID of the team if it is not in the channel object. If not specified the current users teams are tried
        [Parameter()]
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #The channel either as a name, an ID or as a channel object (which may contain the team as a property)
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0, parameterSetName="CHMsgs")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0, parameterSetName="CHTabs")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0, parameterSetName="CHFolder")]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0, parameterSetName="CHFiles")]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
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
        #We want team to tab complete on an empty line, but than can mean if channel is the only thing on the command line it goes into $team
        if ($Team -is [MicrosoftGraphChannel]) {
            $Channel = $Team
            $Team    = $Null
        }
        if (-not $Channel) {
            $Channel = Get-GraphTeam -Team $Team -Channels
        }
        #Support -channel being given an array of channels.
        foreach ($ch in $channel) {
            if     ($ch.Team)               {$teamID    = $ch.team        }
            elseif ($team)                  {$teamid    = idfromteam $Team}
            if (-not $teamId -or $teamid.Count -gt 1) {
                Write-Warning -Message 'Could not resolve the Team ID'; continue
            }
            if     ($ch.id  )               {$channelID = $ch.ID          }
            elseif ($ch -is [string] -and
                    $ch -match '@thread\.') {$channelID = $ch             }
            elseif ($ch -is [string])       {
               $uri = "$graphuri/teams/$teamID/channels?`$select=id,displayname&`$filter=" +  (FilterString $ch -ToLower)
               $channelID = (Invoke-GraphRequest $uri -ValueOnly).id
            }
            if (-not ($channelID -or $channelID.Count -gt 1)) {Write-warning -Message "Could not resolve the Channel ID"; return}
            elseif   ($Messages -or $Top)              {
                $uri      =  "$graphuri/teams/$teamID/channels/$channelID/messages"
                if ($top) {Get-ChannelMessagesByURI -channelid $channelID -TeamID $teamID -URI $uri -Top $Top}
                else      {Get-ChannelMessagesByURI -channelid $channelID -TeamID $teamID -URI $uri}
                return
            }
            elseif   ($Tabs)                           {
                if ($ch.DisplayName) {Write-Progress -Activity 'Getting Tab information' -CurrentOperation $ch.DisplayName}
                else                 {Write-Progress -Activity 'Getting Tab information' }
                $uri = "$GraphUri/teams/$teamID/channels/$channelID/tabs?`$expand=teamsApp"
                Invoke-GraphRequest -Uri $uri  -ValueOnly -astype ([MicrosoftGraphTeamsTab]) |
                    ForEach-Object { #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                        $_.TeamsAppID = $_.TeamsApp.ID
                        $_
                    }
                Write-Progress -Activity 'Getting Tab information' -Completed
            }
            elseif   ($Folder -or $Files)              {
                if ($ch.DisplayName) {Write-Progress -Activity 'Getting Folder information' -CurrentOperation $ch.DisplayName}
                else                 {Write-Progress -Activity 'Getting Folder information' }
                $uri = "$GraphUri/teams/$teamID/channels/$channelID/filesFolder"
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
    param   (
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
    begin   {  $webparams = @{Method = "POST";   ContentType = "application/json"  } }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        $teamid = idfromteam $Team
        if (-not $teamid -or $teamid.count -gt 1) {
               Write-Warning "Could not resolve $team to team-enabled group." ; return
        }
        else { $webparams['uri'] = "$GraphUri/teams/$teamid/channels"}

        foreach ($n in $Name) {
            $checkUri = $webparams.uri + '?$filter=' + (FilterString $n)
            if (Invoke-GraphRequest $checkUri -ValueOnly) {
                Write-Warning -Message "Channel '$n' already exists in team '$($team.tostring())''."
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
        >Get-GraphTeam Developers -ChannelName "Project Firebird" | Remove-GraphChannel
        Finds a channel by name from a named team , and removes it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param   (
        #A team object or the ID of the team, if it can't be derived from the channel.
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #The channel to delete; either as an ID, or a channel object
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
        #if Specified the channel will be deleted without prompting
        [switch]$Force
    )
    process {

        if ($Channel.Team) { $teamid  = $Channel.team }
        elseif  ($Team)    { $teamid  = idfromteam $Team}
        if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
        }
        if     ($Channel -is [MicrosoftGraphChannel]) { $c = $Channel }
        elseif ($Channel -is [string]) {
            if ($Channel -match '@thread\.' ) {
                   $checkUri = "$GraphUri/teams/$teamid/channels?`$filter=id eq '$Channel'"
            }
            else { $checkUri = "$GraphUri/teams/$teamid/channels?`$filter=$(FilterString $Channel -ToLower)"}
            $c =  Invoke-GraphRequest  $checkUri -ValueOnly -AsType ([MicrosoftGraphChannel])
         }
        if (-not $c -or $c.count -gt 1)  {Write-Warning "Could not resolve the channel" ; return}
        elseif ($force -or $PSCmdlet.Shouldprocess($c.displayname, "Delete Channel")) {
            Invoke-GraphRequest -Method "Delete" -Uri "$GraphUri/teams/$teamid/channels/$($c.id)"
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
    param   (
        #A team object or the ID of the team, if it can't be derived from the channel.
        [ArgumentCompleter([TeamCompleter])]
        [Parameter()]
        $Team,
        #The channel to post to either as an ID or a channel object.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true,Position=1)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        if      ($Channel.Team)            {$teamID  = $Channel.Team }
        elseif  ($Team)                    { $teamid  = idfromteam $Team}
            if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
        }

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
            $result
        }
    }
}

function New-GraphChannelReply      {
    <#
      .Synopsis
        Posts a reply to a message in a Teams channel
    #>
    [Cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The Message to reply to as an ID or a message object
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Message,
        #If the message or channel parameters don't include the team ID, the team either as an ID or an object containing the ID
        [ArgumentCompleter([TeamCompleter])]
        [Parameter()]
        $Team,
        #If Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        [ArgumentCompleter([ChannelCompleter])]
        [Parameter()]
        $Channel,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Position=1,Mandatory=$true)]
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
    if      ($message.ChannelIdentity.TeamId)    {$teamId    = $message.ChannelIdentity.TeamId }
    elseif  ($Message.team)                      {$teamid    = $Message.team}
    elseif  ($Channel.team)                      {$teamid    = $Channel.team}
    elseif  ($Team)                              {$teamid    = idfromteam $Team}
    if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
    }

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

    if ($force -or $PSCmdlet.Shouldprocess("Post Reply")) {
        Try {Invoke-GraphRequest @webparams}
        catch {Write-Warning "A bug in the API means an error is returned even when the reply succeeds. An error was returned here but the reply may have been posted"}
    }
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
    param   (
        #The Message to reply to as an ID or a message object containing an ID (and possibly the team and channel ID)
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Message,
        #If the message or channel parameters don't include the team ID, the team either as an ID or an object containing the ID
        [ArgumentCompleter([TeamCompleter])]
        [Parameter()]
        $Team,
        #If Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        [ArgumentCompleter([ChannelCompleter])]
        [Parameter()]
        $Channel,
        #If specified returns the message, followed by its replies. (Otherwise , only the replies are returned)
        [switch]$PassThru
    )
    process {
        ContextHas -scopes 'ChannelMessage.Read.All' -BreakIfNot
        #region convert the information from the message (and optionally channel and team) into a URI to post to
        if      ($message.ChannelIdentity.TeamId)    {$teamId    = $message.ChannelIdentity.TeamId }
        elseif  ($Message.team)                      {$teamid    = $Message.team}
        elseif  ($Channel.Team)                      {$teamId    = $Channel.Team }
        elseif  ($Team)                              {$teamid    = idfromteam $Team}
        if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
    }

        if      ($Message.ChannelIdentity.ChannelId) {$channelid = $Message.ChannelIdentity.ChannelId}
        elseif  ($Message.channel)                   {$channelid = $Message.channel}
        elseif  ($Channel.id)                        {$channelid = $channel.id}
        elseif  ($Channel -is [string] -and
                 $Channel -match '@thread')          {$channelID = $channel}
        elseif  ($Channel -is [string]) {
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
        Get-ChannelMessagesByURI -URI "$GraphUri/teams/$teamid/channels/$channelid/Messages/$msgID/replies" -Channelid $channelid -teamID $teamID
     }
}

function Add-GraphWikiTab           {
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
    param   (
        #A team ID, or a team object if the team can't be found from the the channel
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #An ID or Channel object which may contain the team ID
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
        #The label for the tab
        $TabLabel = "Wiki",
        #If specified the tab will be added without prompting for confirmation
        [switch]$Force
    )

    ContextHas -WorkOrSchoolAccount -BreakIfNot
    if      ($Channel.Team)            {$teamID  = $Channel.Team }
    elseif  ($Team)                    { $teamid  = idfromteam $Team}
    if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
    }

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
                   'Uri'             = "$GraphUri/teams/$teamID/channels/$channelID/tabs"
                   'ContentType'     = 'application/json'
                   'AsType'          =  ([MicrosoftGraphTeamsTab])
                   'ExcludeProperty' = '@odata.context'
    }
    $webparams['Body'] = ConvertTo-Json ([ordered]@{
        'displayname'         = $TabLabel
        'teamsApp@odata.bind' = "$GraphUri/appCatalogs/teamsApps/com.microsoft.teamspace.tab.wiki"}
    )

    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.Shouldprocess($TabLabel,"Create wiki tab")) {Invoke-GraphRequest @webparams}
}

function Add-GraphPlannerTab        {
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
    param   (
        #An ID or Plan object for a plan within the team
        [Parameter(Mandatory=$true,Position=0)]
        $Plan,
        #An ID or Channel object for a channel (which may contain the team ID)
        [Parameter(Mandatory=$true,Position=1)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
        #A team ID, or a team object, if not specified as part of the channel
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #The label for the tab.
        $TabLabel,
        #If Specified the tab will be added without confirming
        $Force
    )
    #region get IDs needed
    ContextHas -WorkOrSchoolAccount -BreakIfNot
    if      ($Channel.Team)                   {$teamID  = $Channel.Team }
    elseif  ($Team)                           {$teamid  = idfromteam $Team}
    if (-not $teamid -or $teamid.Count -gt 1) {
            Write-Warning "Could not resolve the channel's team from the information given" ; return
    }
    if       ($Channel.id)                    {$channelID = $Channel.id}
    elseif   ($Channel   -is [string] -and
              $Channel -notmatch '@thread')   {$channelID = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id}
    elseif   ($Channel   -is [string])        {$channelID = $channel}

    if (-not ($teamID    -is [string]      -and
              $teamId    -match $GUIDRegex -and
              $channelID -is [string]      -and
              $channelID -match '@thread'))   {
        #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
        Write-Warning -Message 'Could not determine the team and channel IDs'; return
    }
    if       ($Plan.id)                       {$Plan      = $Plan.id}
    #endregion
    if ((-not $TabLabel) -and $Plan.Title)    {
        Write-Verbose -Message "ADD-GRAPHPLANNERTAB: No Tab label was specified, using the Plan title '$($Plan.Title)'"
        $TabLabel = $Plan.Title
    }

    $tabURI = "https://tasks.office.com/{0}/Home/PlannerFrame?page=7&planId={1}" -f $Global:GraphUser, $Plan

    $webparams = @{'Method'          = 'Post'
                   'Uri'             = "$GraphUri/teams/$teamID/channels/$channelID/tabs"
                   'ContentType'     = 'application/json'
                   'AsType'          =  ([MicrosoftGraphTeamsTab])
                   'ExcludeProperty' = '@odata.context'
    }

    $webparams['body'] = ConvertTo-Json ([ordered]@{
        'displayname'         = $TabLabel
        'teamsApp@odata.bind' = "$GraphUri/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
        'configuration'       = [ordered]@{
                   'entityId'   = $plan
                   'contentUrl' = $tabURI
                   'websiteUrl' = $tabURI
                   'removeUrl'  = $tabURI
        }
    })
    Write-Debug $webparams.body
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {Invoke-GraphRequest @webparams}
}

function Add-GraphOneNoteTab        {
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
    param   (
        #The Notebook or Section to associate with the tab
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Alias('Section')]
        $Notebook,

        #An ID or Channel object which may contain the team ID; the tab will be created in this channel
        [Parameter(Mandatory=$true,Position=1)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
         #A team ID, or a team object if the team can't be found from the the channel
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #The label for the tab, if left blank the name of the Notebook or Section will be sued
        $TabLabel,
        #If Specified the tab will be added without pausing for confirmation, this is the default unless $ConfirmPreference has been set.
        $Force
    )
    process {
        ContextHas -scopes 'Group.ReadWrite.All' -BreakIfNot
        #region ensure we have the team , channel and notebook IDs, and a label for the tab
        if      ($Channel.Team)                   {$teamID  = $Channel.Team }
        elseif  ($Team)                           {$teamid  = idfromteam $Team}
        if (-not $teamid -or $teamid.Count -gt 1) {
                Write-Warning "Could not resolve the channel's team from the information given" ; return
        }
        if       ($Channel.id)                    {$channelID = $Channel.id }
        elseif   ($Channel -is [string] -and
                $Channel -match '@thread')        {$channelID = $channel  }
        elseif   ($Channel -is [string])          {
                $Channelid = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id
        }
        if (-not ($teamID  -is [string] -and   $teamId    -match $GUIDRegex -and
                $channelID -is [string] -and $channelID -match '@thread'))  {
            #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
            Write-Warning -Message 'Could not determine the team and channel IDs'; return
        }
        if       (-not $TabLabel -and
                       $Notebook.displayName) {$TabLabel = $Notebook.displayName}
        elseif   (-not $TabLabel)             {
            Write-warning 'Unable to determin a name for the tab, please specify one explicitly'; return
        }
        #endregion
        if       (-not $Notebook.Id -or
                      ($Notebook -is [MicrosoftGraphOnenoteSection] -and -not
                       $Notebook.ParentNotebook.Id  )) {
            Write-Warning 'Could not determine the notebook ID.'; return
        }
        $webparams = @{
            'Method'          = 'Post'
            'Uri'             = "$GraphUri/teams/$teamID/channels/$channelID/tabs"
            'ContentType'     = 'application/json'
            'AsType'          =  ([MicrosoftGraphTeamsTab])
            'ExcludeProperty' = '@odata.context'
        }
        #This had to be reverse engineered, from a beta version of the API, so if it works past next week, be happy.
        #If the "Notebook" object is actually a section, and it was fetched by one of the module commands (get-GraphTeam -notebook, or get-graphNotebook -section)
        #then $Notebook it will have a a parentNotebook ID. This IF..Else is to make sure we have the real notebook ID, and catch a sectionID if there is one.
        if       ($Notebook.parentNotebook.id) {
                  $ParamsPt2     = '&notebookSource=PickSection&sectionId='+ $Notebook.id
                  $NotebookID    = $Notebook.parentNotebook.id
        }
        else  {   $ParamsPt2     = '&notebookSource=New'
                  $NotebookID    = $Notebook.id }

        #if $Notebook is a section its url will end ?wd=(something). We need to split this off the URL and re-use it. The () need to be unescapted too,
        if       ($Notebook.links.oneNoteWebUrl.href -match '\?(wd=.*$)') {
                  $ParamsPt2    += '&' + ( $Matches[1] -replace '%28','(' -replace '%29',')' )
                  $OnenoteWebUrl = $Notebook.links.oneNoteWebUrl.href  -replace  '\?wd=.*$', ''
        }
        else     {$OnenoteWebUrl = $Notebook.links.oneNoteWebUrl.href}

        #We need the teamsite URL for the team who owns this channel, and the URL to the the Notebook. Both need to be escaped.
        $OnenoteWebUrl  = $OnenoteWebUrl                             -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'
        $siteUrl        = (Get-GraphTeam -Team $Teamid -Site).webUrl -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'

        #Now we need to build up the mother and father of all URIs It contains the ID and URL for the notebook (not section). The Name, the teamsite. And Section specifics if applicable.
        $URIParams      = "?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&ui={locale}&tenantId={tid}"+
                          "&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$TeamID%2Fnotes%2Fnotebooks%2F"+ $NotebookID   +
                          "&oneNoteWebUrl=" + $oneNoteWebUrl +
                          "&notebookName="  + [uri]::EscapeDataString( $notebook.displayName ) +
                          "&siteUrl="       + $SiteUrl + $ParamsPt2

        #Now we can create the JSON. Such information as there is can be found at https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
        $json = ConvertTo-Json ([ordered]@{
                'teamsApp@odata.bind' = "$GraphUri/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
                'displayname'         = $TabLabel
                'configuration'       = [ordered]@{
                    'entityId'        = ((New-Guid).tostring() + "_" +  $Notebook.ID)
                    'contentUrl'      = "https://www.onenote.com/teams/TabContent" + $URIParams
                    'removeUrl'       = "https://www.onenote.com/teams/TabRemove"  + $URIParams
                    'websiteUrl'      = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"
                }})
        $webparams['body'] = $json  -replace "\\u0026","&"
        Write-Debug $webparams.body
        if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {Invoke-GraphRequest @webparams }
    }
}

function Add-GraphSharePointTab     {
    <#
      .Synopsis
        Adds a planner tab to a team-channel for sharepoint deurl
      .Description
        This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
        which requires consent to use the Group.ReadWrite.All scope.
      .Example

    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param   (
        #An ID or Plan object for a plan within the team
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $WebUrl,
        #The label for the tab by default the displayname for of the list
        [Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true)]
        [Alias('DisplayName')]
        $TabLabel,
        #The label for the tab.
        [Parameter(Position=2,ValueFromPipelineByPropertyName=$true)]
        #Either a genericList (default) or a documentLibrary
        $Template = 'genericList',
        #An ID or Channel object for a channel (which may contain the team ID)
        [Parameter(Mandatory=$true,Position=3)]
        [ArgumentCompleter([ChannelCompleter])]
        $Channel,
        #A team ID, or a team object, if not specified as part of the channel
        [ArgumentCompleter([TeamCompleter])]
        $Team,
        #If Specified the tab will be added without confirming
        $Force
    )
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #region get IDs needed
        if       ($Channel.Team)                   {$teamID  = $Channel.Team }
        elseif   ($Team)                           {$teamid  = idfromteam $Team}
        if (-not  $teamid -or $teamid.Count -gt 1) {
                Write-Warning "Could not resolve the channel's team from the information given" ; return
        }

        if       ($Channel.id)                     {$channelID = $Channel.id }
        elseif   ($Channel   -is [string] -and
                  $Channel -notmatch '@thread')    {$channelID = (Get-GraphTeam -Team $teamID -Channels -ChannelName $channel).id}
        elseif   ($Channel   -is [string])         {$channelID = $Channel  }

        if (-not ($teamID    -is [string] -and $teamId    -match $GUIDRegex -and
                  $channelID -is [string] -and $channelID -match '@thread'))  {
            #we got zero matches or more than one for a team/channel name, or we got an object without an ID, or an object where the ID wasn't a guid
            Write-Warning -Message 'Could not determine the team and channel IDs'; return
        }
        #endregion

        $webparams = @{
            'Method'          = 'Post'
            'Uri'             = "$GraphUri/teams/$teamID/channels/$channelID/tabs"
            'ContentType'     = 'application/json'
            'AsType'          =  ([MicrosoftGraphTeamsTab])
            'ExcludeProperty' = '@odata.context'
        }
        if ($Template = 'genericList') {
            $webparams['body'] = ConvertTo-Json ([ordered]@{
                'displayname'         = $TabLabel
                'teamsApp@odata.bind' = "$GraphUri/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88"
                'configuration'       = [ordered]@{
                            'entityId'   = ""
                            'contentUrl' =  ($WebUrl -replace '(.*/sites/[^/]+/).*$','$1_layouts/15/teamslogon.aspx?spfx=true&dest=') + [uri]::EscapeDataString($WebUrl)
                            'websiteUrl' = $WebUrl
                            'removeUrl'  = $null
                }
            })
        }
        elseif ($Template = 'genericList') {
            $webparams['body'] = ConvertTo-Json ([ordered]@{
                'displayname'         = $TabLabel
                'teamsApp@odata.bind' = "$GraphUri/appCatalogs/teamsApps/com.microsoft.teamspace.tab.files.sharepoint"
                'configuration'       = [ordered]@{
                            'entityId'   = ""
                            'contentUrl' = $WebUrl
                            'websiteUrl' = $null
                            'removeUrl'  = $null
                }
            })
        }
        else {
            Write-Warning "Cannot handle the template type of '$Template'." ; return
        }
        Write-Debug $webparams.body
        if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {Invoke-GraphRequest @webparams}
    }
}

# Adding tab https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
# Get-GraphTeamsApp will get the apps but we don't get the ability to configure them
#-  often some other things will get called as part of setup and need be reverse engineered. e.g. whiteboard calls another service to get a new whiteboard GUID
function Get-GraphTeamsApp          {
    <#
      .synopsis
        Returns apps from the teams app catalog
    #>
    param   (
        [string]$App
    )
    process {
        ContextHas -scopes 'AppCatalog.Submit', 'AppCatalog.Read.All', 'AppCatalog.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All' -BreakIfNot
        $uri = "$GraphUri/appcatalogs/teamsApps"
        if ($App -match $guidRegex) {
            Invoke-GraphRequest "$uri/$App`?`$expand=appdefinitions" -ExcludeProperty '@odata.context'  -AsType ([MicrosoftGraphTeamsApp])
        }
        elseif ($App) {
            $uri += '?$filter=startswith(tolower(displayname),''{0}'')' -f $App.toLower()
            Invoke-GraphRequest $uri  -ValueOnly -AsType ([MicrosoftGraphTeamsApp])  | Sort-Object -Property Displayname
        }
        else   {
            Invoke-GraphRequest $uri   -ValueOnly -AsType ([MicrosoftGraphTeamsApp]) -Headers @{'ConsistencyLevel'='eventual'} | Sort-Object -Property Displayname
        }
    }
}
