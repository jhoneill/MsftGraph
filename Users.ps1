using namespace Microsoft.Graph.PowerShell.Models

function ConvertTo-GraphUser      {
    <#
      .Synopsis
        Helper function (not exported) to expand users' manager or direct reports when converting results to userObjects
    #>
    param (
        #The dictionary /hash table object returned by the REST API
        [Parameter(ValueFromPipeline=$true,Mandatory=$true,Position=0)]
        $RawUser
    )
    process {
        foreach ($r in $RawUser) {
            #We expand manager by default, or might be told to expand direct reports. Make either into users.
            if (-not $r.manager) { $mgr = $null}
            else {
                    $disallowedProperties = $r.manager.keys.where({$_ -notin $script:UserProperties})
                    foreach ($p in $disallowedProperties) {$null = $r.manager.remove($p)}
                    $mgr  = New-Object -TypeName MicrosoftGraphUser -Property $r.manager
                    $null = $r.remove('manager')
            }
            if (-not $r.directReports) {$directs = $null}
            else {
                $directs = @()
                foreach ($d in $r.directReports) {
                    $disallowedProperties = $d.keys.where({$_ -notin $script:UserProperties})
                    foreach ($p in $disallowedProperties) {$null = $d.remove($p)}
                    $directs += New-Object -TypeName MicrosoftGraphUser -Property $d
                }
                $null = $r.remove('directReports')
            }
            $disallowedProperties = $r.keys.where({$_ -notin $script:UserProperties})
            foreach ($p in $disallowedProperties) {$null = $r.remove($p)}
            $user =  New-Object -TypeName MicrosoftGraphUser -Property $r
            if ($mgr)      {$user.manager      = $mgr}
            if ($directs)  {$user.DirectReports= $directs}
            $user
        }
    }
}

function Get-GraphUserList        {
    <#
      .Synopsis
        Returns a list of Azure active directory users for the current tennant.
      .Example
        Get-GraphUserList -filter "Department eq 'Accounts'"
        Gets the list with a custom filter this is typically fieldname eq 'value' for equals or
        startswith(fieldname,'value') clauses can be joined with and / or.
    #>
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser])]
    [cmdletbinding(DefaultparameterSetName="None")]
    param   (
        #If specified searches for users whose first name, surname, displayname, mail address or UPN start with that name.
        [parameter(Mandatory=$true, parameterSetName='FilterByName', Position=0,ValueFromPipeline=$true )]
        [ArgumentCompleter([UPNCompleter])]
        [string[]]$Name,

        #Names of the fields to return for each user.Note that some properties - aboutMe, Birthday etc, are only available when getting a single user, not a list.
        #The  API defaults to :  businessPhones, displayName, givenName, id, jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName
        #The module adds to this set - the default list can be set with Set-GraphOption -DefaultUserProperties which has the master copy of the validate set used here
        [validateSet('accountEnabled', 'ageGroup', 'assignedLicenses', 'assignedPlans', 'businessPhones',
                     'city', 'companyName', 'consentProvidedForMinor', 'country', 'createdDateTime', 'creationType',
                     'deletedDateTime', 'department',  'displayName',
                     'employeeHireDate', 'employeeID', 'employeeOrgData', 'employeeType', 'externalUserState', 'externalUserStateChangeDateTime',
                     'givenName', 'id', 'identities', 'imAddresses', 'isResourceAccount','jobTitle', 'legalAgeGroupClassification',
                     'mail', 'mailNickname', 'mobilePhone',
                     'officeLocation', 'onPremisesDistinguishedName', 'onPremisesDomainName', 'onPremisesExtensionAttributes',
                     'onPremisesImmutableId', 'onPremisesLastSyncDateTime', 'onPremisesProvisioningErrors', 'onPremisesSamAccountName', 'otherMails',
                     'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName',
                     'passwordPolicies', 'passwordProfile', 'postalCode', 'preferredDataLocation',
                     'preferredLanguage', 'provisionedPlans', 'proxyAddresses',
                     'showInAddressList','state', 'streetAddress', 'surname', 'usageLocation', 'userPrincipalName', 'userType')]
        [Alias('Property')]
        [string[]]$Select = $Script:DefaultUserProperties  ,

        #The default is to get all
        $Top ,

        #Order by clause for the query - most fields result in an error and it can't be combined with some other query values.
        [parameter(Mandatory=$true, parameterSetName='Sorted')]
        [ValidateSet('displayName', 'userPrincipalName')]
        [Alias('OrderBy')]
        [string]$Sort,

        #Filter clause for the query for example "startswith(displayname,'Bob') or startswith(displayname,'Robert')"
        [parameter(Mandatory=$true, parameterSetName='FilterByString')]
        [string]$Filter,

        #Adds a filter clause "userType eq 'Member'"
        [parameter(Mandatory=$true, parameterSetName='FilterToMembers')]
        [switch]$MembersOnly,

        #Adds a filter clause "userType eq 'Guest'"
        [parameter(Mandatory=$true, parameterSetName='FilterToGuests')]
        [switch]$GuestsOnly,

        [validateSet('directReports', 'manager', 'memberOf', 'ownedDevices', 'ownedObjects', 'registeredDevices', 'transitiveMemberOf',  'extensions','')]
        [string]$ExpandProperty = 'manager',

        # The URI for the proxy server to use
        [Parameter(DontShow)]
        [System.Uri]
        $Proxy,

        # Credentials for a proxy server to use for the remote call
        [Parameter(DontShow)]
        [ValidateNotNull()]
        [PSCredential]$ProxyCredential,

        # Use the default credentials for the proxygit
        [Parameter(DontShow)]
        [Switch]$ProxyUseDefaultCredentials
    )
    process {
        Write-Progress "Getting the list of Users"
        $webParams =  @{ValueOnly = $true }

        if     ($MembersOnly) {$Filter = "userType eq 'Member'"}
        elseif ($GuestsOnly)  {$Filter = "userType eq 'Guest'"}

        if     ($Filter -or $Sort -or $Name) {
                $webParams['Headers'] = @{'ConsistencyLevel'='eventual'}
        }
        #Ensure at least ID, UPN and displayname are selected - and we always have something in $select
        foreach ($s in @('ID', 'userPrincipalName', 'displayName')){
             if ($s -notin $Select) {$Select += $s }
        }
        $uri = "$GraphUri/users?`$select="  + ($Select -join ',')

        if (-not $Top)       {$webParams['AllValues'] = $true               }
        else                 {$uri = $uri + '&$top='     + $Top             }
        if ($Filter)         {$uri = $uri + '&$Filter='  + $Filter          }
        if ($Sort)           {$uri = $uri + '&$orderby=' + $Sort            }
        if ($ExpandProperty) {$uri = $uri + '&expand='   + $ExpandProperty  }

        if (-not $Name)      {Invoke-GraphRequest -Uri $uri @webParams | ConvertTo-GraphUser}
        else {
            foreach ($n in $Name) {
                $filter = '&$Filter=' + (FilterString $n -ExtraFields 'userPrincipalName','givenName','surname','mail')
                Invoke-GraphRequest -Uri ($uri + $filter) @webParams | ConvertTo-GraphUser
            }
        }
        Write-Progress "Getting the list of Users" -Completed
    }
}

function Get-GraphUser            {
    <#
      .Synopsis
        Gets information from the MS-Graph API about the a user (current user by default)
      .Description
        Queries https://graph.microsoft.com/v1.0/me or https://graph.microsoft.com/v1.0/name@domain
        or https://graph.microsoft.com/v1.0/<<guid>> for information about a user.
        Getting a user returns a default set of properties only (businessPhones, displayName, givenName,
        id, jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName).
        Use -select to get the other properties.
        Most options need consent to use the Directory.Read.All or Directory.AccessAsUser.All scopes.
        Some options will also work with user.read; and the following need consent which is task specific
        Calendars needs Calendars.Read, OutLookCategries needs MailboxSettings.Read, PlannerTasks needs
        Group.Read.All, Drive needs Files.Read (or better), Notebooks needs either Notes.Create or
        Notes.Read (or better).
      .Example
        Get-GraphUser -MemberOf | ft displayname, description, mail, id
        Shows the name description, email address and internal ID for the groups this user is a direct member of
      .Example
        (get-graphuser -Drive).root.children.name
        Gets the user's one drive. The drive object has a .root property which is represents its
        root-directory, and this has a .children property which is a collection of the objects
        in the root directory. So this command shows the names of files and folders in the root directory. To just see sub folders it is possible to use
        get-graphuser -Drive | Get-GraphDrive -subfolders
    #>
    [cmdletbinding(DefaultparameterSetName="None")]
    [Alias('ggu')]
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser])]
    param   (
        #UserID as a guid or User Principal name. If not specified, it will assume "Current user" if other paraneters are given, or "All users" otherwise.
        [parameter(Position=0,valueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [alias('id')]
        [ArgumentCompleter([UPNCompleter])]
        $UserID,
        #Get the user's Calendar(s)
        [parameter(Mandatory=$true, parameterSetName="Calendars")]
        [switch]$Calendars,
        #Select people who have the user as their manager
        [parameter(Mandatory=$true, parameterSetName="DirectReports")]
        [switch]$DirectReports,
        #Get the user's one drive
        [parameter(Mandatory=$true, parameterSetName="Drive")]
        [switch]$Drive,
        #Get user's license Details
        [parameter(Mandatory=$true, parameterSetName="LicenseDetails")]
        [switch]$LicenseDetails,
        #Get the user's Mailbox Settings
        [parameter(Mandatory=$true, parameterSetName="MailboxSettings")]
        [switch]$MailboxSettings,
        #Get the users Outlook-categories (by default, 6 color names)
        [parameter(Mandatory=$true, parameterSetName="OutlookCategories")]
        [switch]$OutlookCategories,
        #Get the user's manager
        [parameter(Mandatory=$true, parameterSetName="Manager")]
        [switch]$Manager,
        #Get the user's teams
        [parameter(Mandatory=$true, parameterSetName="Teams")]
        [switch]$Teams,
        #Get the user's Groups
        [parameter(Mandatory=$true, parameterSetName="Groups")]
        [switch]$Groups,
        [parameter(Mandatory=$false, parameterSetName="Groups")]
        [parameter(Mandatory=$true, parameterSetName="SecurityGroups")]
        [switch]$SecurityGroups,
        #Get the Directory-Roles and Groups the user belongs to; -Groups or -Teams only return one type of object.
        [parameter(Mandatory=$true, parameterSetName="MemberOf")]
        [switch]$MemberOf,
        #Get the Directory-Roles and Groups the user belongs to; -Groups or -Teams only return one type of object.
        [parameter(Mandatory=$true, parameterSetName="TransitiveMemberOf")]
        [switch]$TransitiveMemberOf,
        #Get the user's Notebook(s)
        [parameter(Mandatory=$true, parameterSetName="Notebooks")]
        [switch]$Notebooks,
        #Get the user's photo
        [parameter(Mandatory=$true, parameterSetName="Photo")]
        [switch]$Photo,
        #Get the user's assigned tasks in planner.
        [parameter(Mandatory=$true, parameterSetName="PlannerTasks")]
        [Alias('AssignedTasks')]
        [switch]$PlannerTasks,
        #Get the plans owned by the user in planner.
        [parameter(Mandatory=$true, parameterSetName="PlannerPlans")]
        [switch]$Plans,
        #Get the users presence in Teams
        [parameter(Mandatory=$true, parameterSetName="Presence")]
        [switch]
        $Presence,
        #Get the user's MySite in SharePoint
        [parameter(Mandatory=$true, parameterSetName="Site")]
        [switch]$Site,
        #Get the user's To-do lists
        [parameter(Mandatory=$true, parameterSetName="ToDoLists")]
        [switch]$ToDoLists,

        #specifies which properties of the user object should be returned Additional options are available when selecting individual users
        #The API documents list deviceEnrollmentLimit, deviceManagementTroubleshootingEvents , mailboxSettings which cause errors
        [parameter(Mandatory=$true,parameterSetName="Select")]
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
        [String[]]$Select =  $Script:DefaultUserProperties ,

        #Used to explicitly say "Current user" and will over-ride UserID if one is given.
        [switch]$Current

    )
    process {
        $result       = @()
        if ((ContextHas -Not -WorkOrSchoolAccount) -and ($MailboxSettings -or $Manager -or $Photo -or $DirectReports -or $LicenseDetails -or $MemberOf -or $Teams -or $PlannerTasks -or $Devices ))  {
            Write-Warning   -Message "Only the -Drive, -Calendars and -Notebooks options work when you are logged in with this kind of account." ; return
            #to do check scopes.
            # Most options need consent to use the Directory.Read.All or Directory.AccessAsUser.All scopes.
            # Some options will also work with user.read; and the following need consent which is task specific
            # Calendars needs Calendars.Read, OutLookCategries needs MailboxSettings.Read, PlannerTasks needs
            # Group.Read.All, Drive needs Files.Read (or better), Notebooks needs either Notes.Create or  Notes.Read (or better).
        }
        #region resolve User name(s) to IDs,
        #If we got -Current use the "me" path - otherwise if we didn't get an ID return the list. So Get-Graphuser = list; Get-Graphuser -current = me ; Get-graphuser -memberof = memberships for me; otherwise we get a name or ID
        if     ($Current -or ($PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters}) -and -not $UserID)  )  {$UserID = "me"}
        elseif (-not $UserID) {
            Get-GraphUserList ;
            return
        }

        #if we got a user object use its ID, if we got an array and it contains names (not guid or UPN or "me") and also contains Guids we can't unravel that.
        if ($UserID -is [array] -and $UserID -notmatch "$GuidRegex|\w@\w|^me`$" -and
                                     $UserID -match     $GuidRegex ) {
            Write-Warning   -Message 'If you pass an array of values they cannot be names. You can pipe names or pass and array of IDs/UPNs' ; return
        }
        #if it is a string and not a guid or UPN - or an array where at least some members are not GUIDs/UPN/me try to resolve it
        elseif (($UserID -is [string] -or $UserID -is [array]) -and
                 $UserID -notmatch "$GuidRegex|\w@\w|^me`$" ) {
                 $UserID = Get-GraphUserList -Name $UserID
        }
        #endregion

        [void]$PSBoundParameters.Remove('UserID')

        foreach ($u in $UserID) {
            #region set up the user part of the URI that we will call
            if ($u -is [MicrosoftGraphUser] -and -not  ($PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters})  )) {
                $u
                continue
            }
            if     ($u.id)         { $id = $u.Id}
            else                   { $id = $u   }
            if ($id -notmatch "^me$|$guidRegex|\w@\w") {
                Write-Warning "User ID '$id' does not look right"
            }
            Write-Progress -Activity 'Getting user information' -CurrentOperation "User = $id"
            if     ($id -eq 'me') { $Uri = "$GraphUri/me"  }
            else                  { $Uri = "$GraphUri/users/$id" }

            # -Teams requires a GUID, photo doesn't work for "me"
            if (  (($Teams -or $Presence) -and $id -notmatch $GuidRegex ) -or
                  ($Photo -and $id -eq 'me')        ) {
                                    $id  =   (Invoke-GraphRequest -Method GET -Uri $uri).id
                                    $Uri = "$GraphUri/users/$id"
            }
            #endregion
            #region add the data-specific part of the URI, make the rest call and convert the result to the desired objects
            <#available:  but not implemented in this command (some in other commands )
                managedAppRegistrations, appRoleAssignments,
                activities &  activities/recent, needs UserActivity.ReadWrite.CreatedByApp permission
                calendarGroups, calendarView, contactFolders, contacts, mailFolders,  messages,
                createdObjects, ownedObjects,
                managedDevices, registeredDevices, deviceManagementTroubleshootingEvents,
                events, extensions,
                followedSites,
                inferenceClassification,
                insights/used" /trending or /stored.
                oauth2PermissionGrants,
                onlineMeetings,
                photos,
                presence,
                scopedRoleMemberOf,
                (content discovery) settings,
                teamwork (apps),
            "https://graph.microsoft.com/v1.0/me/getmemberobjects"  -body '{"securityEnabledOnly": false}'  ).value
            #>
            try   {
                if     ($Drive -and (ContextHas -WorkOrSchoolAccount)) {
                    Invoke-GraphRequest -Uri ($uri + '/Drive?$expand=root($expand=children)') -PropertyNotMatch '@odata'                -As ([MicrosoftGraphDrive])    }
                elseif ($Drive              ) {
                    Invoke-GraphRequest -Uri ($uri + '/Drive')                                                                          -As ([MicrosoftGraphDrive])    }
                elseif ($LicenseDetails     ) {
                    Invoke-GraphRequest -Uri ($uri + '/licenseDetails')           -All                                                  -As ([MicrosoftGraphLicenseDetails]) }
                elseif ($MailboxSettings    ) {
                    Invoke-GraphRequest -Uri ($uri + '/MailboxSettings')                -Exclude '@odata.context'                       -As ([MicrosoftGraphMailboxSettings])}
                elseif ($OutlookCategories  ) {
                    Invoke-GraphRequest -Uri ($uri + '/Outlook/MasterCategories') -All                                                  -As ([MicrosoftGraphOutlookCategory]) }
                elseif ($Photo              ) {
                    $response = Invoke-GraphRequest -Uri ($uri + '/Photo')
                    if ($response.'@odata.context' -match "#users\('(.*)'\)/") {
                        $picUserId = $Matches[1]
                    }
                    else {$picUserId = $null}
                    if ($response.'@odata.mediaContentType' -and
                        $PSVersionTable.Platform -like "win*" -and
                        (Test-Path "HKLM:\SOFTWARE\Classes\MIME\Database\Content Type\$($response.'@odata.mediaContentType')")) {
                        $picExtension = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\MIME\Database\Content Type\$($response.'@odata.mediaContentType')").Extension
                    }
                    else {$picExtension = $null}
                    $Null = $response.Remove('@odata.mediaEtag'), $response.Remove('@odata.context'),  $response.Remove('@odata.id'), $response.Remove('@odata.mediaContentType')
                    New-Object -TypeName MicrosoftGraphProfilePhoto -Property $response |
                        Add-Member -PassThru -NotePropertyName UserID -NotePropertyValue  $picUserId |
                        Add-Member -PassThru -NotePropertyName Ext    -NotePropertyValue  $picExtension |
                        Add-Member -PassThru -NotePropertyName URI    -NotePropertyValue ($uri + '/Photo/$value') |
                        Add-Member -PassThru -MemberType ScriptMethod -Name Download -value{
                            param ($Path="$Pwd\userPhoto", [switch]$Passthru)
                            if ($path -notlike "*$($this.ext)") {$Path += $this.ext}
                            if ($this.uri) {Invoke-GraphRequest $this.uri -OutputFilePath $path}
                            if ($Passthru) {Get-Item $path}
                        }
                }
                elseif ($PlannerTasks       ) {
                    Invoke-GraphRequest -Uri ($uri + '/planner/tasks')            -All  -Exclude '@odata.etag'                          -As ([MicrosoftGraphPlannerTask])}
                elseif ($Plans              ) {
                    Invoke-GraphRequest -Uri ($uri + '/planner/plans')            -All  -Exclude "@odata.etag"                          -As ([MicrosoftGraphPlannerPlan])}
                elseif ($Presence           ) {
                    if ($u.DisplayName)         {$displayName = $u.DisplayName}  else {$displayName=$null}
                    if ($u.UserPrincipalName)   {$upn = $u.UserPrincipalName}
                    elseif ($u -match '\w@\w')  {$upn = $u}
                    else                        {$upn = $null}
                    #can also use GET https://graph.microsoft.com/v1.0/communications/presences/<<id>>>
                    #see https://docs.microsoft.com/en-us/graph/api/cloudcommunications-getpresencesbyuserid for getting bulk presence
                    Invoke-GraphRequest -Uri ($uri + '/presence')                       -Exclude "@odata.context"                       -As ([MicrosoftGraphPresence]) |
                      Add-Member -PassThru -NotePropertyName DisplayName       -NotePropertyValue $displayName |
                      Add-Member -PassThru -NotePropertyName UserPrincipalName -NotePropertyValue $upn
                }
                elseif ($Teams              ) {
                    Invoke-GraphRequest -Uri ($uri + '/joinedTeams')              -All                                                  -As ([MicrosoftGraphTeam])}
                elseif ($ToDoLists          ) {
                    Invoke-GraphRequest -Uri ($uri + '/todo/lists')               -All  -Exclude "@odata.etag"                          -As ([MicrosoftGraphTodoTaskList]) |
                      Add-Member -PassThru -NotePropertyName UserId -NotePropertyValue $id
                }
                # Calendar wants a property added so we can find it again
                elseif ($Calendars          ) {
                    Invoke-GraphRequest -Uri ($uri + '/Calendars?$orderby=Name' ) -All                                                  -As ([MicrosoftGraphCalendar]) |
                        ForEach-Object {
                            if ($id -eq 'me') {$calpath = "me/Calendars/$($_.id)"}
                            else              {$calpath = "users/$id/calendars/$($_.id)"
                                               Add-Member -InputObject $_ -NotePropertyName User -NotePropertyValue $id
                            }
                            Add-Member -PassThru -InputObject $_ -NotePropertyName CalendarPath -NotePropertyValue $calpath
                        }
                }
                elseif ($Notebooks          ) {
                    $response = Invoke-GraphRequest -Uri ($uri +
                                          '/onenote/notebooks?$expand=sections' ) -All  -Exclude 'sections@odata.context'               -As ([MicrosoftGraphNotebook])
                    #Section fetched this way won't have parentNotebook, so make sure it is available when needed
                    foreach ($bookobj in $response) {
                        foreach ($s in $bookobj.Sections) {$s.parentNotebook = $bookobj }
                        $bookobj
                    }
                }
                # for site, get the user's MySite. Convert it into a graph URL and get that, expand drives subSites and lists, and add formatting types
                elseif ($Site               ) {
                        $response  = Invoke-GraphRequest -Uri ($uri + '?$select=mysite')
                        $uri       = $GraphUri + ($response.mysite -replace '^https://(.*?)/(.*)$', '/sites/$1:/$2?expand=drives,lists,sites')
                        $siteObj    = Invoke-GraphRequest $Uri                          -Exclude '@odata.context', 'drives@odata.context',
                                                                                           'lists@odata.context', 'sites@odata.context' -As ([MicrosoftGraphSite])
                        foreach ($l in $siteObj.lists) {
                            Add-Member -InputObject $l -MemberType NoteProperty   -Name SiteID   -Value  $siteObj.id
                        }
                        $siteObj
                    }
                elseif ($Groups -or
                        $SecurityGroups     ) {
                    if  ($SecurityGroups)   {$body = '{  "securityEnabledOnly": true  }'}
                    else                    {$body = '{  "securityEnabledOnly": false }'}
                    $response         = Invoke-GraphRequest -Uri ($uri  + '/getMemberGroups') -Method POST  -Body $body -ContentType 'application/json'
                    foreach ($r in $response.value) {
                        $result     += Invoke-GraphRequest  -Uri "$GraphUri/directoryObjects/$r"
                    }
                }
                elseif ($Manager            ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/Manager') }
                elseif ($DirectReports      ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/directReports')       -All}
                elseif ($MemberOf           ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/MemberOf')            -All}
                elseif ($TransitiveMemberOf ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/TransitiveMemberOf')  -All}
                else                          {
                    foreach ($s in @('ID', 'userPrincipalName', 'displayName')){if ($s -notin $Select) {$Select += $s }}
                    $result += Invoke-GraphRequest -Uri ($uri + '?$expand=manager&$select=' + ($Select -join ','))
                }
            }
            #if we get a not found error that's propably OK - bail for any other error.
            catch {
                if     ($_.exception.response.statuscode.value__ -eq 404) {
                    Write-Warning -Message "'Not found' error while getting data for user '$($u.ToString())'"
                }
                elseif ($_.exception.response.statuscode.value__ -eq 403) {
                    Write-Warning -Message "'Forbidden' error while getting data for user '$($u.ToString())'. Do you have access to the correct scope?"
                }
                else {
                    Write-Progress -Activity 'Getting user information' -Completed
                    throw $_ ; return
                }
            }
             #endregion

        }
        foreach ($r in ($result )) {
            if     ($r.'@odata.type' -match 'directoryRole$') {
                    #This is a hack so we get role memberships and group memberships laying nicely
                    $disallowedProperties = $r.keys.where({$_ -notin $script:GroupProperties})
                    foreach ($p in $disallowedProperties) {$null = $r.remove($p)}
                    [void]$r.add('GroupTypes','DirectoryRole')
                    New-Object -Property $r -TypeName ([MicrosoftGraphGroup])
            }
            elseif ($r.'@odata.type' -match 'group$') {
                    $disallowedProperties = $r.keys.where({$_ -notin $script:GroupProperties})
                    foreach ($p in $disallowedProperties) {$null = $r.remove($p)}
                    New-Object -Property $r -TypeName ([MicrosoftGraphGroup])
            }
            elseif ($r.'@odata.type' -match 'user$' -or $PSCmdlet.parameterSetName -eq 'None' -or $Select) {
                    ConvertTo-GraphUser -RawUser $r
            }
            else    {$r}
        }
    }
    end     {
        Write-Progress -Activity 'Getting user information' -Completed
    }
}

function Set-GraphUser            {
    <#
      .Synopsis
        Sets properties of  a user (the current user by default)
      .Example
        Set-GraphUser -Birthday "31 march 1965"  -Aboutme "Lots to say" -PastProjects "Phoenix","Excalibur" -interests "Photography","F1" -Skills "PowerShell","Active Directory","Networking","Clustering","Excel","SQL","Devops","Server builds","Windows Server","Office 365" -Responsibilities "Design","Implementation","Audit"
        Sets the current user, giving lists for projects, interests and skills
      .Description
        Needs consent to use the User.ReadWrite, User.ReadWrite.All, Directory.ReadWrite.All,
        or Directory.AccessAsUser.All scope.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSPossibleIncorrectComparisonWithNull', '', Justification='In this case we want exactly that behaviour')]
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser])]
    [cmdletbinding(SupportsShouldprocess=$true)]
    param   (
        #ID for the user if not the current user
        [parameter(Position=0,ValueFromPipeline=$true)]
        [ArgumentCompleter([UPNCompleter])]
        $UserID = "me",
        #A freeform text entry field for the user to describe themselves.
        [String]$AboutMe,
        #The SMTP address for the user, for example, 'Alex@contoso.onmicrosoft.com'
        [String]$Mail,
        #A list of additional email addresses for the user; for example: ['bob@contoso.com', 'Robert@fabrikam.com'].
        [String[]]$OtherMails,
        #User's mobile phone number
        [String]$MobilePhone,
        #The telephone numbers for the user. NOTE: Although this is a string collection, only one number can be set for this property
        [String[]]$BusinessPhones,
        #Url for user's personal site.
        [String]$MySite,
        #A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries.  Examples include: 'US', 'JP', and 'GB'
        [ValidateNotNullOrEmpty()]
        [UpperCaseTransformAttribute()]
        [ValidateCountryAttribute()]
        [string]$UsageLocation,
        #The name displayed in the address book for the user. This is usually the combination of the user''s first name, middle initial and last name. This property is required when a user is created and it cannot be cleared during updates.
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,
        #The given name (first name) of the user.
        [Alias('FirstName')]
        [string]$GivenName,
        #User's last / family name
        [Alias('LastName')]
        [string]$Surname,
        #The user's job title
        [string]$JobTitle,
        #The name for the department in which the user works.
        [string]$Department,
        #The office location in the user's place of business.
        [string]$OfficeLocation,
        # The company name which the user is associated. This property can be useful for describing the company that an external user comes from. The maximum length of the company name is 64 chararcters.
        $CompanyName,
        #ID or UserPrincipalName of the user's manager
        [ArgumentCompleter([UPNCompleter])]
        [string]$Manager,
        #The employee identifier assigned to the user by the organization
        [string]$EmployeeID,
        #Captures enterprise worker type: Employee, Contractor, Consultant, Vendor, etc.
        [string]$EmployeeType,
        #The date and time when the user was hired or will start work in case of a future hire
        [datetime]$EmployeeHireDate,
        #For an external user invited to the tenant using the invitation API, this property represents the invited user's invitation status. For invited users, the state can be PendingAcceptance or Accepted, or null for all other users.
        $ExternalUserState,
        #The street address of the user's place of business.
        $StreetAddress,
        #The city in which the user is located.
        $City,
        #The state, province or county in the user's address.
        $State,
        #The country/region in which the user is located; for example, 'US' or 'UK'
        $Country,
        #The postal code for the user's postal address, specific to the user's country/region. In the United States of America, this attribute contains the ZIP code.
        $PostalCode,
        #User's birthday as a date. If passing a string it can be "March 31 1965", "31 March 1965", "1965/03/31" or  "3/31/1965" - this layout will always be read as US format.
        [DateTime]$Birthday,
        #List of user's interests
        [String[]]$Interests,
        #List of user's past projects
        [String[]]$PastProjects,
        #Path to a .jpg file holding the users photos
        [String]$Photo,
        #List of user's responsibilities
        [String[]]$Responsibilities,
        #List of user's Schools
        [String[]]$Schools,
        #List of user's skills
        [String[]]$Skills,
        #Set to disable the user account, to re-enable an account use $AccountDisabled:$false
        [switch]$AccountDisabled,
        #If specified the modified user will be returned
        [switch]$PassThru,
        #Supresses any confirmation prompt
        [Switch]$Force
    )
    begin   {
        #things we don't want to put in the JSON body when we send the changes.
        $excludedParams = [Cmdlet]::CommonParameters + [Cmdlet]::OptionalCommonParameters + @('Force', 'PassThru', 'UserID', 'AccountDisabled', 'Photo', 'Manager')
        $settings       = @{}
        $returnProps    = $Script:DefaultUserProperties
        if ($userid -ne $me -and $userid -ne $global:GraphUser -and
            $PSBoundparameters['aboutMe', 'birthday', 'hireDate', 'interests', 'mySite', 'pastProjects',
                              'preferredName', 'responsibilities', 'schools', 'skills'] -ne $null) {
            Write-Warning "One or more of the selected properties can only be set by user '$($UserID.ToString())'."
            return
        }
        foreach ($p in $PSBoundparameters.Keys.where({$_ -notin $excludedParams})) {
            #turn "Param" into "param" make dates suitable text, and switches booleans
            $key   = $p.toLower()[0] + $p.Substring(1)
            if ($key -notin $returnProps) {$returnProps += $key}
            $value = $PSBoundparameters[$p]
            if ($value -is [datetime]) {$value = $value.ToString("yyyy-MM-ddT00:00:00Z")}  # 'o' for ISO date time may work here
            if ($value -is [switch])   {$value = $value -as [bool]}
            $settings[$key] = $value
        }
        if ($PSBoundparameters.ContainsKey('AccountDisabled')) {#allows -accountDisabled:$false
            $settings['accountEnabled'] = -not $AccountDisabled
            if ($returnProps -notcontains 'accountEnabled') {$returnProps += 'accountEnabled'}
        }
        if  ($settings.count -eq 0 -and -not $Photo -and -not $Manager) {
            Write-Warning -Message "Nothing to set"
        }
        else {
            #Don't put the body into webparams which will be used for multiple things
            $json = (ConvertTo-Json $settings) -replace '""' , 'null'
            Write-Debug  $json
        }
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx todo check scopes  User.ReadWrite, User.ReadWrite.All, Directory.ReadWrite.All,        or Directory.AccessAsUser.All scope.

        if ($UserID -is [string] -and
            $UserID -notmatch "^me$|\w@\w|$GUIDRegex" ) {
            $UserID = Get-GraphUser $UserID
        }

        #allow an array of users to be passed.
        foreach ($id in $UserID ) {
            #region configure the web parameters for changing the user. Allow for filtered objects with an ID or a UPN
            $webparams = @{
                    'Method'            = 'PATCH'
                    'Contenttype'       = 'application/json'
            }
            if     ($id -is [string] -and
                    $id -match "\w@\w|$GUIDRegex" ){
                                            $webparams['uri'] = "$GraphUri/users/$id/" }
            elseif ($id -eq "me")          {$webparams['uri'] = "$GraphUri/me/"            }
            elseif ($id.id)                {$webparams['uri'] = "$GraphUri/users/$($id.id)/"}
            elseif ($id.UserPrincipalName) {$webparams['uri'] = "$GraphUri/users/$($id.UserPrincipalName)/"}
            else                           {Write-Warning "$id does not look like a valid user"; continue}

            if ($id -is [string])          {$promptName = $id}
            elseif ($id.DisplayName)       {$promptName = $id.DisplayName}
            elseif ($id.UserPrincipalName) {$promptName = $id.UserPrincipalName}
            else                           {$promptName = $id.ToString() }
            #endregion
            if ($json -and ($Force -or $Pscmdlet.Shouldprocess($promptName ,'Update User'))){
                $null = Invoke-GraphRequest  @webparams -Body $json
            }
            if ($Photo)     {
                if (-not (Test-Path $Photo) -or $photo -notlike "*.jpg" ) {
                    Write-Warning "$photo doesn't look like the path to a .jpg file" ; return
                }
                else {$photoPath = (Resolve-Path $Photo).Path }
                $baseUri                    =  $webparams['uri']
                $webparams['uri']           =  $webparams['uri'] + 'photo/$value'
                $webparams['Method']        = 'Put'
                $webparams['Contenttype']   = 'image/jpeg'
                Write-Debug "Uploading Photo: '$photoPath'"
                if ($Force -or $Pscmdlet.Shouldprocess($userID ,'Update User')) {
                    $null = Invoke-GraphRequest  @webparams -InputFilePath $photoPath
                }
                $webparams['uri'] = $baseUri
            }
            if ($Manager)   {
                if ($Manager -is [string] -and $Manager -notmatch "$GUIDRegex|\w@\w") {
                                 $Manager = Get-GraphUser $manager}
                if ($Manger.id) {$Manager = $Manager.id}
                if ($Manager -isnot [string] -or $Manager -notmatch "$GUIDRegex|\w@\w" ) {
                    Write-Warning "Could not resolve the manager"
                }
                else {
                    $json = ConvertTo-Json @{ '@odata.id' =  "$GraphUri/users/$manager" }
                    Write-Debug  $json
                    $baseUri                    =  $webparams['uri']
                    $webparams['uri']           =  $webparams['uri'] + 'manager/$ref'
                    $webparams['Method']        = 'Put'
                    $webparams['Contenttype']   = 'application/json'
                    if ($Force -or $Pscmdlet.Shouldprocess($userID ,'Update User')) {
                        $null = Invoke-GraphRequest  @webparams -Body $json
                    }
                    $webparams['uri'] = $baseUri
                }
            }
            if ($PassThru)  {
               Invoke-GraphRequest ($webparams.uri + '?$expand=manager&$select=' + ($returnProps -join ',')) | ConvertTo-GraphUser
            }
        }
    }
}

function New-GraphUser            {
    <#
        .synopsis
            Creates a new user in Azure Active directory

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification="False positive and need to support plain text here")]
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #User principal name for the new user. If not specified it can be built by specifying Mail nickname and domain name.
        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNDisplay',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [alias("UPN")]
        [string]$UserPrincipalName,

        #Mail nickname for the new user. If not specified the part of the UPN before the @sign will be used, or using the displayname or first/last name
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='UPNFromDomainDisplay',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [Parameter(ParameterSetName='DomainFromUPNDisplay')]
        [ValidateNotNullOrEmpty()]
        [Alias("Nickname")]
        [string]$MailNickName,

        #Domain for the new user - used to create UPN name if the UPN paramater is not provided
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='UPNFromDomainDisplay')]
        [ValidateNotNullOrEmpty()]
        [ArgumentCompleter([DomainCompleter])]
        [string]$Domain,

        #The name displayed in the address book for the user. This is usually the combination of the user''s first name, middle initial and last name. This property is required when a user is created and it cannot be cleared during updates.
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [Parameter(ParameterSetName='UPNFromDomainDisplay',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNDisplay',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,

        #The given name (first name) of the user.
        [Parameter(ParameterSetName='UPNFromDomainLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Alias('FirstName')]
        [string]$GivenName,

        #User's last / family name
        [Parameter(ParameterSetName='UPNFromDomainLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Alias('LastName')]
        [string]$Surname,

        #ID or UserPrincipalName of the user's manager
        [ArgumentCompleter([UPNCompleter])]
        [string]$Manager,

        #A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries.  Examples include: 'US', 'JP', and 'GB'
        [ValidateNotNullOrEmpty()]
        [UpperCaseTransformAttribute()]
        [ValidateCountryAttribute()]
        [string]$UsageLocation = $Script:DefaultUsageLocation,

        [ArgumentCompleter([GroupCompleter])]
        $Groups,

        [ArgumentCompleter([RoleCompleter])]
        $Roles,

        [ArgumentCompleter([SkuCompleter])]
        $Licenses,

        #The initial password for the user. If none is specified one will be generated and output by the command
        [string]$Initialpassword,

        #If specified the user will not have to change their password on first logon
        [switch]$NoPasswordChange,

        #If specified the user will need to use Multi-factor authentication when changing their password.
        [switch]$ForceMFAPasswordChange,


        #Specifies built-in password policies to apply to the user
        [ValidateSet('DisableStrongPassword','DisablePasswordExpiration')]
        [string[]]$PasswordPolicies,

        #A hash table of properties which can be passed as parameters to Set-GraphUser command after the account is created
        [hashtable]$SettableProperties,

        #A script block specifying how the displayname should be built, by default it is {"$GivenName $Surname"};
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$DisplayNameRule = {"$GivenName $Surname"},

        #A script block specifying how the mailnickname should be built, by default it is $GivenName.$Surname with punctuation removed;
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$NickNameRule    = {($GivenName -replace '\W','') +'.' + ($Surname -replace '\W','')},

        #A script block specifying how to create a password, by default a date between 1800 and 2199 like 10Oct2126 - easy to type and meets complexity rules.
        [scriptblock]$PasswordRule    = {([datetime]"1/1/1800").AddDays((Get-Random 146000)).tostring("ddMMMyyyy")},

        #If specified prevents any confirmation dialog from appearing
        [switch]$Force
    )
    #region we allow the names to be passed flexibly make sure we have what we need
    # Accept upn and display name -split upn to make a mailnickname, leave givenname/surname blank
    #        upn, display name, first and last
    #        mailnickname, domain, display name [first & last] - create a UPN
    #        domain, first & last - create a display name, and mail nickname, use the nickname in upn
    #re-create any scriptblock passed as a parameter, otherwise variables in this function are out of its scope.
    if ($NickNameRule)            {$NickNameRule      = [scriptblock]::create( $NickNameRule )   }
    if ($DisplayNameRule)         {$DisplayNameRule   = [scriptblock]::create( $DisplayNameRule) }
    if ($PasswordRule)            {$PasswordRule      = [scriptblock]::create( $PasswordRule)    }
     #if we didn't get a UPN or a mail nickname, make the nickname first, then add the domain to make the UPN
    if (-not $UserPrincipalName -and
        -not $MailNickName  )     {$MailNickName      = Invoke-Command -ScriptBlock $NickNameRule
    }
    #if got a UPN but no nickname, split at @ to get one
    elseif ($UserPrincipalName -and
              -not $MailNickName) {$MailNickName      = $UserPrincipalName -replace '@.*$','' }
    #If we didn't get a UPN we should have a domain and a nickname, combine them
    if ($MailNickName -and
         -not $UserPrincipalName) {
         if (-not $Domain) {
             $Domain = (Invoke-GraphRequest "$GraphUri/domains?`$select=id,isDefault" -ValueOnly -AsType ([psobject]) |
                            Where-Object {$_.isdefault}  #filter doesn't work in the rest call :-(
                       ).id
        }
         $UserPrincipalName = "$MailNickName@$Domain"    }

    #if we didn't get a display name build it
    if (-not $DisplayName)        {$DisplayName       = Invoke-Command -ScriptBlock $DisplayNameRule}

    #We should have all 3 by now
    if (-not ($DisplayName -and $MailNickName -and $UserPrincipalName -and $UserPrincipalName -match "\w+@\w+")) {
        throw "couldn't make sense of those parameters"
    }
    #A simple way to create one in 100K temporary passwords. You might get 10Oct2126 - easy to type and meets complexity rules.
    if (-not $Initialpassword)    {$Initialpassword   = Invoke-Command -ScriptBlock $PasswordRule
                                   [pscustomobject]@{'DisplayName'       = $DisplayName
                                                     'UserPrincipalName' = $UserPrincipalName
                                                     'Initialpassword'   = $Initialpassword}
    }
    $settings = @{
        'accountEnabled'    = $true
        'displayName'       = $DisplayName
        'mailNickname'      = $MailNickName
        'userPrincipalName' = $UserPrincipalName
        'usageLocation'     = $UsageLocation
        'passwordProfile'   =  @{
            'forceChangePasswordNextSignIn' = -not $NoPasswordChange
            'password' = $Initialpassword
        }
    }
    if ($ForceMFAPasswordChange) {$settings.passwordProfile['forceChangePasswordNextSignInWithMfa'] = $true}
    if ($PasswordPolicies)       {$settings['passwordPolicies'] = $PasswordPolicies -join ', '}
    if ($GivenName)              {$settings['givenName']        = $GivenName }
    if ($Surname)                {$settings['surname']          = $Surname }

    $webparams = @{
        'Method'            = 'POST'
        'Uri'               = "$GraphUri/users"
        'Contenttype'       = 'application/json'
        'Body'              = (ConvertTo-Json $settings -Depth 5)
        'AsType'            = [MicrosoftGraphUser]
        'ExcludeProperty'   = '@odata.context'
    }
    Write-Debug $webparams.Body
    if ($force -or $pscmdlet.ShouldProcess($displayname, 'Create New User')){
        try {
            $u = Invoke-GraphRequest @webparams
            if ($SettableProperties) {
                Set-GraphUser -UserID $u.id @SettableProperties -Force
            }
            if ($manager) {
                Set-GraphUser -UserID $u.id -Manager $manager -Force
            }
            if ($Groups)   {
                Add-GraphGroupMember -Group $groups -Member $u
            }
            if ($Roles)    {
                Grant-GraphDirectoryRole -Role $Roles -Member $u
            }
            if ($Licenses) {
                Grant-GraphLicense -SKUID $Licenses -UserID $u
            }
            if ($PSBoundParameters['Initialpassword'] ) {return $u }
        }
        catch {
        # xxxx Todo figure out what errors need to be handled (illegal name, duplicate user)
        $_
        }
    }
}

function Reset-GraphUserPassword  {
    <#
        .synopsis
            Administrative reset to a given our auto-generated password, defaulting to 'reset at next logon'

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification="False positive and need to support plain text here")]
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
        #User principal name for the user.
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [alias("UPN")]
        [ArgumentCompleter([UPNCompleter])]
        [string]$UserPrincipalName,

        #The replacement password for the user. If none is specified one will be generated and output by the command
        [string]$Initialpassword,

        #If specified the user will not have to change their password at their next logon
        [switch]$NoPasswordChange,

        #If Specified prevents any confirmation dialog from appearing
        [switch]$Force
    )

    if ($UserPrincipalName -notmatch "$Guidregex|\w@\w")  {
        Write-Warning "$UserPrincipalName does not look like an ID or UPN." ; return
    }
    if (-not $Initialpassword)    {
             $Initialpassword   = ([datetime]"1/1/1800").AddDays((Get-Random 146000)).tostring("ddMMMyyyy")
    }
    $webparams = @{
        'Method'            = 'PATCH'
        'Uri'               = "$GraphUri/users/$UserPrincipalName/"
        'Contenttype'       = 'application/json'
        'Body'              = (ConvertTo-Json @{'passwordProfile' =  @{
                                        'password'                      =      $Initialpassword
                                        'forceChangePasswordNextSignIn' = -not $NoPasswordChange}})
    }

    Write-Debug $webparams.Body
    if ($force -or $pscmdlet.ShouldProcess($UserPrincipalName, 'Reset password for user')){
        Write-Output "$UserPrincipalName, $Initialpassword"
        Invoke-GraphRequest @webparams
    }
}

function Remove-GraphUser         {
    <#
      .Synopsis
        Deletes a user from Azure Active directory
    #>
    [cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    param   (
        #ID for the user
        [parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        [ArgumentCompleter([UPNCompleter])]
        $UserID,
        #If specified the user is deleted without a confirmation prompt.
        [Switch]$Force
    )
    process{
       ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx todo check scopes
        if ($userid -is [string] -and $UserID -notmatch "\w@\w|$guidregex") {
            $userId = Get-GraphUser $UserID
        }
        #allow an array of users to be passed.
        foreach ($u in $UserID ) {
            if     ($u.displayName)       {$displayname = $u.displayname}
            elseif ($u.UserPrincipalName) {$displayName = $u.UserPrincipalName}
            else                          {$displayName = $u}
            if     ($u.id)                {$u = $U.id}
            elseif ($u.UserPrincipalName) {$u = $U.UserPrincipalName}
            if ($Force -or $pscmdlet.ShouldProcess($displayname,"Delete User")) {
                try   {
                    Remove-MgUser_Delete -UserId $u -ErrorAction Stop
                }
                catch {
                    if ($_.exception.statuscode.value__ -eq 404) {
                          Write-Warning -Message "'Not found' error while trying to delete '$displayname'."
                    }
                    else {throw $_}
                }
            }
        }
    }
}

function Find-GraphPeople         {
    <#
      .Synopsis
        Searches people in your inbox / contacts / directory
     .Example
        Find-GraphPeople -Topic timesheet -First 6
        Returns the top 6 results for people you have discussed timesheets with.
      .Description
        Requires consent to use either the People.Read or the People.Read.All scope
    #>
    [cmdletbinding(DefaultparameterSetName='Default')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="Person would be incorrect")]
    param   (
        #Text to use in a 'Topic' Search. Topics are not pre-defined, but inferred using machine learning based on your conversation history (!)
        [parameter(ValueFromPipeline=$true,Position=0,parameterSetName='Default',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $Topic,
        #Text to use in a search on name and email address
        [parameter(ValueFromPipeline=$true,parameterSetName='Fuzzy',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $SearchTerm,
        #Number of results to return (10 by default)
        [ValidateRange(1,1000)]
        [int]$First = 10
    )
    process {
    #xxxx todo check scopes    Requires consent to use either the People.Read or the People.Read.All scope
        if     ($Topic) {
            $uri = $GraphUri +'/me/people?$search="topic:{0}"&$top={1}' -f $Topic, $First
        }
        elseif ($SearchTerm) {
            $uri = $GraphUri + '/me/people?$search="{0}"&$top={1}' -f $SearchTerm, $First
        }

        Invoke-GraphRequest $uri -ValueOnly -As ([MicrosoftGraphPerson])
    }
}

function Import-GraphUser         {
<#
    .synopsis
       Imports a list of users from a CSV file
    .description
        Takes a list of CSV files and looks for xxxx columns
        * Action is either Add, Remove or Set - other values will cause the row to be ignored
        * DisplayName

#>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
        #One or more files to read for input.
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Disables any prompt for confirmation
        [switch]$Force,
        #Supresses output of Added, Removed, or No action messages for each row in the file.
        [switch]$Quiet,
        #Fields which are lists will be split at , or ; by default but a replacement split expression may be given
        [String]$ListSeparator = '\s*,\s*|\s*;\s*'
    )
    begin   {
        Test-GraphSession
        $list = @()
    }
    process {
        foreach ($p in $path) {
            if (Test-Path $p) {$list += Import-Csv -Path $p}
            else { Write-Warning -Message "Cannot find $p" }
        }
    }
    end     {
        foreach ($user in $list) {
            $upn = $user.UserPrincipalName
            if     (-not $upn) {
                    Write-Warning "User was missing a UPN"
                    continue
            }
            else   {
                    Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Checking for existing user'
                    $exists =  (Invoke-GraphRequest "$GraphUri/users?`$Filter=userprincipalName eq '$upn'" -ValueOnly) -as [bool]}

            if     ($user.Action -eq 'Remove' -and (-not $exists)) {
                    Write-Warning "User '$upn' was marked for removal, but no matching user was found."
                    continue
            }
            elseif ($user.Action -eq 'Remove' -and
                   ($force -or $PSCmdlet.ShouldProcess($upn,"Remove user "))){
                    Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Removing existing user'
                    Remove-Graphuser -Force -user $user
                    Write-Verbose "Removed user '$upn'."
                    continue
            }

            if     ($user.Action -eq 'Add'    -and $exists) {
                    Write-Warning  "User '$upn' was marked for addition, but that name already exists."
                    continue
            }
            elseif ($user.Action -eq 'Add'    -and
                   ($force -or $PSCmdlet.ShouldProcess($upn,"Create new user"))){
                    Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Creating user'
                    $params = @{Force=$true}
                    foreach ($p in @('DisplayName','UserPrincipalName', 'MailNickName','GivenName',
                                     'Surname', 'Initialpassword','UsageLocation').where({$user.$_})) {
                        $params[$p] = $user.$p
                    }
                    if ($user.PasswordPolicies)  {
                        $params['PasswordPolicies'] = $user.PasswordPolicies -split $ListSeparator
                    }
                    if ($user.NoPasswordChange -in @("Yes","True","1") ) {
                        $params['NoPasswordChange'] = $true
                    }
                    if ($user.ForceMFAPasswordChange -in @("Yes","True","1") ) {
                        $params['ForceMFAPasswordChange'] = $true
                    }
                    New-GraphUser @params
                    Write-Verbose "Added user '$($user.DisplayName)' as '$upn'"
                    $exists      = $true
                    $user = $user | Select-Object -Property * -ExcludeProperty 'DisplayName', 'MailNickName','GivenName', 'Surname','UsageLocation'
                    $user.Action = "Set"
                    Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Checking new account is available'
                    $stopTime = [datetime]::now.AddMinutes(2)
                    do    {$newUser = Get-graphuser $upn }
                    until ($newUser -or [datetime]::now -gt $stopTime -or (Start-Sleep -Seconds 5))
            }
            if     ($user.Action -eq 'Set'    -and (-not $exists)) {
                    Write-Warning "User '$upn' was marked for update, but no matching user was found."
                    continue
            }
            if     ($user.Action -eq 'Set' -and
                   ($force -or $PSCmdlet.ShouldProcess($upn,"Set properties of user"))){
                    $params = @{'UserId' = $upn ; 'Force'= $true}
                    $Setparameters = (Get-Command Set-GraphUser ).Parameters.Values |
                        Where-Object name -notin ([Cmdlet]::CommonParameters + [Cmdlet]::OptionalCommonParameters  )

                    foreach ($p in $setparameters.where({$user.($_.name)}) ) {
                        $pName = $p.name
                        if      ($p.parameterType -eq [string[]] ) {$params[$pName] = $user.$pName -split $ListSeparator  }
                        elseif  ($p.switchParameter)               {$params[$pName] = $user.$pName -in @("Yes","True","1")  }
                        else                                       {$params[$pName] = $user.$pName}
                    }
                    if ($params.count -gt 2) {
                        Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Setting user properties'
                        Set-GraphUser @params
                    }

                    if ($user.Groups)   {
                        Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Adding user to group(s)'
                        Add-GraphGroupMember   -Group ($user.groups -split $ListSeparator) -Member $upn -Force
                    }
                    if ($user.Licenses) {
                        Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Granting license(s)'
                        Grant-GraphLicense   -SKUID ($user.Licenses -split $ListSeparator) -UserID $upn -Force
                    }
                    if ($user.Roles)    {
                        Write-Progress -Activity 'Configuring users' -CurrentOperation $upn -Status 'Adding user to role(s)'
                        Grant-GraphDirectoryRole -Role ($user.Roles -split $ListSeparator) -Member $upn -Force
                    }
                    Write-Verbose "Updated properties of user '$upn'"
            }
        }
        Write-Progress -Activity 'Configuring users' -Completed
    }
}

function Export-GraphUser         {
<#
    .synopsis
       Exports a list of users to a CSV file
#>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #Destination for CSV output
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Filter clause for the query for example "department eq 'accounts'"
        $Filter,
        #String to insert between parts of multi-part items.
        $ListSeparator = "; ",
        # Fields to export for each user the values here are the same as the ones in Get-GraphUserList but here we don't use the default set.
        [validateSet('accountEnabled', 'ageGroup', 'assignedLicenses', 'assignedPlans', 'businessPhones',
                     'city', 'companyName', 'consentProvidedForMinor', 'country', 'createdDateTime', 'creationType',
                     'deletedDateTime', 'department',  'displayName',
                     'employeeHireDate', 'employeeID', 'employeeOrgData', 'employeeType', 'externalUserState', 'externalUserStateChangeDateTime',
                     'givenName', 'id', 'identities', 'imAddresses', 'isResourceAccount','jobTitle', 'legalAgeGroupClassification',
                     'mail', 'mailNickname', 'mobilePhone',
                     'officeLocation', 'onPremisesDistinguishedName', 'onPremisesDomainName', 'onPremisesExtensionAttributes',
                     'onPremisesImmutableId', 'onPremisesLastSyncDateTime', 'onPremisesProvisioningErrors', 'onPremisesSamAccountName', 'otherMails',
                     'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName',
                     'passwordPolicies', 'passwordProfile', 'postalCode', 'preferredDataLocation',
                     'preferredLanguage', 'provisionedPlans', 'proxyAddresses',
                     'showInAddressList','state', 'streetAddress', 'surname', 'usageLocation', 'userPrincipalName', 'userType')]
        [Alias('Property')]
        [string[]]$Select =  @('UserPrincipalName', 'MailNickName',  'mail', 'GivenName', 'Surname',  'DisplayName', 'UsageLocation',
                           'PasswordPolicies',  'MobilePhone',   'BusinessPhones',    'JobTitle', 'Department',  'OfficeLocation',
                           'CompanyName',       'StreetAddress', 'City', 'State',     'Country',  'PostalCode',  'accountEnabled'),

        [validateSet('directReports', 'manager', 'memberOf', 'ownedDevices', 'ownedObjects', 'registeredDevices', 'transitiveMemberOf',  'extensions','')]
        [string]$ExpandProperty = 'manager',

        $OutputProperty =  @(   'UserPrincipalName', 'MailNickName',   'GivenName', 'Surname',  'DisplayName', 'UsageLocation',
                                @{n='AccountDisabled';e={-not $_.accountEnabled}} ,
                                'PasswordPolicies', 'Mail',  'MobilePhone',
                                @{n='BusinessPhones' ;e={$_.'BusinessPhones' -join $ListSeparator }},
                                @{n='Manager';e={$_.manager.AdditionalProperties.userPrincipalName}},
                                'JobTitle',  'Department', 'OfficeLocation', 'CompanyName',
                                'StreetAddress', 'City', 'State', 'Country', 'PostalCode')
    )
    $listParams = @{
        Select         = $select
        ExpandProperty = $ExpandProperty
    }
    if ($Filter) {$listParams['Filter'] = $Filter}
    $progressCount = 0
    Get-GraphUserList @listParams |
        ForEach-Object {if (-not $progressCount %1000) {Write-Progress -Activity 'Exporting' -Status $progressCount } ; $progressCount ++ ; $_ } |
            Select-Object  $exportFields | Export-Csv -Path $Path -NoTypeInformation
}

#region MailBox commands: these only depend on the user module from the SDK so go in the same file as user commands
function New-GraphMailAddress     {
    <#
      .synopsis
        Helper function to create a email addresses
    #>
    param   (
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        [Alias('Mail')]
        [String]$Address,
        #The displayname for the recipient
        [Alias('DisplayName')]
        $Name
    )
    @{name=$name;Address=$Address}
}

function New-GraphRecipient       {
    <#
      .Synopsis
        Creats a new meeting attendee, with a mail address and the type of attendance.
    #>
    param   (
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        $Mail,
        #The displayname for the recipient
        [Parameter(Position=2)]
        $DisplayName
    )
    @{ 'emailAddress' =  @{'address'=$mail; name=$DisplayName }}
}

function Get-GraphMailFolder      {
    <#
      .Synopsis
        Get the user's Mailbox folders
      .Example
        Get-GraphMailFolderList -Name inbox
        Gets the current users inbox folder
    #>
    [cmdletbinding(DefaultParameterSetName="FilterByName")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphMailFolder])]
    param   (
        #Filter the folders returned by a name
        [Parameter(ParameterSetName='FilterByName',Position=0)]
        [ArgumentCompleter([MailFolderCompleter])]
        [string]$Name,
        $ParentFolder,
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$User,
        #Select the first n folders.
        [validaterange(1,1000)]
        [int]$Top = 100,
        #fields to select in the query - will add a validate set later
        [string[]]$Select  ,
        #String with orderby clause e.g. "name", "lastmodifiedDate desc"
        [Parameter(ParameterSetName='Sorted')]
        [ValidateSet('childFolderCount', 'childFolderCount desc', 'displayName', 'displayName desc',
                    'totalItemCount', 'totalItemCount desc', 'unreadItemCount', 'unreadItemCount desc')]
        [string]$OrderBy = 'displayname',
        #A custom filter clause.
        [Parameter(ParameterSetName='FilterByString')]
        [string]$Filter,
        [switch]$ChildItems
    )
    #region set-up URI . If we got a user ID, use it other otherwise use the current user, add select, orderby, filter & top parameters as needed
    if     ($User.UserPrincipalName)  {$baseUri = "$GraphUri/users/$($User.UserPrincipalName)/mailFolders" }
    elseif ($User)                    {$baseUri = "$GraphUri/users/$User/mailFolders" }
    else                              {$baseUri = "$GraphUri/me/mailFolders" }

    if     ($Name -match $WellKnownMailFolderRegex -or $Name -match '\S{100}') {
                                         $uri     = $baseUri + '/{1}?$top={0}' -f $top, ($Name -replace '[/\\]','') }
    else {
        if     ($Name -match  "^[/\\]?(\w.*)[/\\](\w.*?$)") {
            $Name         = $Matches[2]
            $ParentFolder = Get-GraphMailFolder -Name $Matches[1] -User $User
            if (-not $ParentFolder -or $ParentFolder.count -gt 1) {
                Write-Warning "$($parentfolder.count)Could not resolve $($matches[1]) as a folder path" ; return
            }
        }
        if     ($ParentFolder.id) {      $Uri     = $baseUri + '/{1}/childfolders?$top={0}' -f $top, $parentfolder.id }
        elseif ($ParentFolder)    {      $Uri     = $baseUri + '/{1}/childfolders?$top={0}' -f $top, $ParentFolder    }
        else                      {      $Uri     = $baseUri + '?$top={0}'                  -f $top }
        if     ($Name)            {      $filter  = FilterString $Name }
    }
    if     ($Select)                     {$uri    = $uri + '&$select=' + ($Select -join ',') }
    if     ($Filter)                     {$uri    = $uri + $JoinChar + '&$Filter='  + $Filter  }
    #The API order by DOES NOT WORK :-( -always by display name.
   #if     ($OrderBy)                    {$uri    = $uri + $JoinChar + '&$orderby=' + $OrderBy }
    #endregion

    #region get the data, to keep the size attribute which is missing from SDK object, we will handle paging and converting to an object locally.
    $folderList              = @()
    $response                = Invoke-GraphRequest -Uri $uri
    if ($response.Keys -notcontains 'value') { #Value may be empty.
           $null = $response.remove('@odata.context'), $response.remove('@odata.id')
           $folderList      += $response
    }
    else  {$folderList      += $response.value
        while ($response.'@odata.nextLink' -and $folderList.count -lt $Top) {
               $response     = Invoke-GraphRequest -Uri  $response.'@odata.nextLink' ;
               $folderList  += $response.value
        }
    }
    if ($ChildItems)   {$result = foreach ($f in $folderlist) {Get-GraphMailFolder -ParentFolder $f.id -User $User }}
    else               {$result = foreach ($f in $folderList) {
        $size = $f.sizeInBytes
        [void]$f.remove('sizeInBytes')
        New-object -TypeName MicrosoftGraphMailFolder -Property $f |
            Add-Member -PassThru -NotePropertyName SizeInBytes -NotePropertyValue  $size |
            Add-Member -PassThru -NotePropertyName Path        -NotePropertyValue "$baseUri/$($f.id)"
    }}
    if (-not $OrderBy) {$result}
    elseif  ($OrderBy -match 'Desc') {
                        $result | Sort-Object -Property ($OrderBy -replace '\s*desc\s*$','') -Descending }
    else               {$result | Sort-Object -Property  $OrderBy }
    #endregion
}

function Get-GraphMailItem        {
    <#
      .Synopsis
        Get items in a mail folder
      .Example
        >Get-GraphMailItem -top 5
        Gets the top 5 items in the current users Inbox
      .Example
        >Get-GraphMailItem -Mailfolder "sentitems" -top 5
        Gets the top 5 items in the current users sent items folder
      .Example
        >Get-GraphMailFolderList -Name sent | Get-GraphMailItem -top 5
        This has the same result as before but could find any folder
      .Example
        >Get-GraphMailItem -Search 'criminal'
        Searches the default folder (inbox) for 'Criminal' in any field
      .Example
        >Get-GraphMailItem -Search 'criminal' -Mailfolder ''
        Searches for 'Criminal' in any field but this time searches the whole mailbox
      .Example
        >Get-GraphMailItem -Search 'subject:criminal' -Mailfolder ''
        This time limits the search to just the subject line. from:, to: etc can be
        used in the same way as they can in a search in outlook.
      .Example
        >Get-GraphMailItem -filter "from/emailAddress/address eq 'alex@contoso.com'"
        Instead of a free text search this applies a filter on email address, looking at the inbox.
      .Example
        Get-GraphMailItem -Filter "(hasattachments eq true) and startswith(from/emailAddress/name, 'alex')"
        This shows a filter based on two conditions.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphMessage])]
    param   (
        #A folder objet or the ID of a folder, or one of the well known folder names 'archive', 'clutter', 'conflicts', 'conversationhistory', 'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures', 'msgfolderroot', 'outbox', 'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'sentitems', 'serverfailures', 'syncissues'
        [Parameter(ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([MailFolderCompleter])]
        $Mailfolder = "Inbox",
        #UserID as a guid or User Principal name, if it can't be discovered from the mailfolder. If not specified defaults to "me"
        [string]$User,
        #Selects only unread mail (equivalent to isread:no in Outlook)
        [switch]$Unread,
        #Searches based on the subject field (equivalent to subject: in Outlook)
        [string]$Subject,
        #Searches based on the from field (equivalent to from: in Outlook)
        [string]$From,
        #Searches based on the to field (equivalent to to: in Outlook)
        [string]$To,
        #Selects only mail with attachments (equivalent to hasAttachments:yes in Outlook). Note this does not combine well with date based searches
        [switch]$HasAttachments,
        #Selects only mail marked as important (equivalent to importance:high in Outlook).
        [switch]$Important,
        #Selects only mail from today (equivalent to received:today in Outlook).
        [switch]$Today,
        #Selects only mail from today (equivalent to received:yesterday in Outlook).
        [switch]$Yesterday,
        #Selects only mail from before a given date
        [datetime]$Before,
        #Selects only mail from after a given date
        [datetime]$After,
        #A term to do a free text search for in the mail box (see examples)
        [string]$Search,
        #If specified returns the top X items, defaults to 100
        [int]$Top = 100 ,
        #Sorting option, defaults to sorting by SentDateTime with newest first. Searches are not sorted.
        [string]$OrderBy ='SentdateTime desc',
        #Select particular mail fields , ignored if -ChildFolders is specified; defaults to From, Subject, SentDateTime, BodyPreview, and Weblink
        [ValidateSet('bccRecipients', 'body', 'bodyPreview', 'categories', 'ccRecipients', 'changeKey', 'conversationId', 'createdDateTime',
        'flag', 'from', 'hasAttachments', 'id', 'importance', 'inferenceClassification', 'internetMessageHeaders', 'internetMessageId',
        'isDeliveryReceiptRequested', 'isDraft', 'isRead', 'isReadReceiptRequested', 'lastModifiedDateTime', 'parentFolderId',
        'receivedDateTime', 'replyTo', 'sender', 'sentDateTime', 'subject', 'toRecipients', 'uniqueBody', 'webLink' )]
        [string[]]$Select = @('From', 'Subject', 'SentDatetime', 'hasAttachments', 'BodyPreview', 'weblink'),
        #A Custom filter string; for example "importance eq high" - the examples have more cases
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )
    process {
       # $wellKnownMailFolderRegex = '^[/\\]?(archive|clutter|conflicts|conversationhistory|deleteditems|drafts|inbox|junkemail|localfailures|msgfolderroot|outbox|recoverableitemsdeletions|scheduled|searchfolders|sentitems|serverfailures|syncissues)[/\\]?$'

        #if mailfolder is a path (not a well known name or 120 chars of ID) get the folder.
        if ($Mailfolder -is [string] -and $Mailfolder -notmatch $WellKnownMailFolderRegex -and $Mailfolder -Notmatch "\S{100}") {
            $Mailfolder = Get-GraphMailFolder -User $User -Name $Mailfolder
        }
        #if mailfolder was a folder object with a path to start or we just got one,  know where to look.
        if ($Mailfolder.path)   {$baseUri = $Mailfolder.path}
        else {  #build a path for a string holding a well known name or ID  or use the ID if the folder was fetched without adding path.
            if     ($Mailfolder.id)             { $MailPath = 'mailfolders/' +  $Mailfolder.id}
            elseif ($Mailfolder -is [string] )  { $MailPath = 'mailfolders/' + ($Mailfolder -replace '^/','' -replace '/$','')}
            else {Write-Warning 'Could not make sense of the the folder provided'}

            if ($User.id) {$baseUri   = "$GraphUri/users/$($user.id)/$MailPath"}
            if ($User)    {$baseUri   = "$GraphUri/users/$user/$MailPath" }
            else          {$baseUri   = "$GraphUri/me/$MailPath" }
        }
        #baseURI should be something like https://graph.microsoft.com/v1.0/users/{some-user-id}/mailfolders/inbox
        #                              or https://graph.microsoft.com/v1.0/me/mailfolders/{somefolderID}
        $webparams = @{
            'Headers'        = @{'Prefer'          ='outlook.body-content-type="text"'
                                    'ConsistencyLevel'='eventual'}
            'ValueOnly'      = $true
            'Uri'            = $baseUri + '/messages?$select='  + ($Select -join ',') +
                                '&$expand=attachments($select=id,name,size,contenttype)&$top=' + $Top
        }
        #Get-GraphMailitem -Search "hasattachments:yes from:tom" -top 3
        if ($HasAttachments)      {$Search = $search + ' hasattachments:yes'}
        if ($Important)           {$Search = $search + ' importance:high'}
        if ($Unread)              {$Search = $search + ' isread:no'}
        if ($Subject)             {$Search = $search + " subject:$subject"}
        if ($To)                  {$Search = $search + " to:$to"}
        if ($From)                {$Search = $search + " from:$from"}
        if ($Before -and $After)  {$Search = $search + ' received>={0:MM-dd-yyyy} AND received<={1:MM-dd-yyyy}' -f $After,$Before }
        if              ($After)  {$Search = $search + ' received>={0:MM-dd-yyyy}' -f $After}
        elseif         ($Before)  {$Search = $search + ' received<={0:MM-dd-yyyy}' -f $Before }
        elseif          ($Today)  {$Search = $search + ' received:today'}
        elseif      ($Yesterday)  {$Search = $search + ' received:yesterday'}

        if     ($Search)     {$webparams.Uri +=  '&$search="' + $Search + '"' }
        elseif ($Filter)     {$webparams.Uri +=  '&$filter='  + $Filter  }
        else                 {$webparams.Uri +=  '&$orderby=' + $OrderBy }
        Write-Debug $webparams.uri

        #we need to handle attachments here.
        $results = Invoke-GraphRequest @webparams
        foreach ($msg in $results) {
            $null = $msg.Remove('@odata.etag') ,$msg.Remove('@odata.id') , $msg.Remove("@odata.type")
            $msgpath =  "$baseUri/messages/$($msg.id)"
            #$msg won't convert unless we convert the attachments first.
            foreach ($a in $msg.attachments) {
                        $null = $a.Remove("@odata.type"),  $a.Remove("@odata.id"), $a.Remove("@odata.mediacontenttype")
                        $a = New-Object MicrosoftGraphAttachment -Property $a
            }
            $newMsg = New-object MicrosoftGraphMessage -Property $msg |
                Add-Member -PassThru -NotePropertyName Path -NotePropertyValue $msgpath
            #Converting a message to to an object will strip extra members off the attachments, so do attachments in 2 parts
            foreach ($a in $newMsg.Attachments) {
                Add-Member -InputObject $a -NotePropertyName Path -NotePropertyValue "$msgpath/attachments/$($a.id)"
            }
            $newMsg
        }
    }
}

function Move-GraphMailItem       {
    param   (
            #The mail item to move. If can be a message object or the ID of a message
            [parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
            $Item,
            #The destination folder. It can be a folder object, a folder ID or a well known folder name like "DeletedItems" or "Inbox"
            [parameter(Mandatory=$true,Position=2)]
            [ArgumentCompleter([MailFolderCompleter])]
            $Destination,
            #Specifies the user if it cannot be discovered from the item and is not "me"
            $User
        )
    begin   {
        #if mailfolder is a path (not a well known name or 120 chars of ID) get the folder.
        if     ($Destination -is [string] -and $Destination -notmatch $WellKnownMailFolderRegex -and $Mailfolder -Notmatch "\S{100}") {
                $Destination = Get-GraphMailFolder -User $User -Name $Destination
        }
        if     ($Destination.Id)           {$body = @{'destinationId' = $Destination.Id}}
        elseif ($Destination -is [string]) {$body = @{'destinationId' = $Destination}}
        else {Write-Warning 'Could not get the destination.' ; break}
        $webparms = @{
            'ContentType'     = 'application/json'
            'Body'            = (ConvertTo-Json $body)
            'Method'          = 'Post'
        }
        Write-Debug $webparms.Body
    }
    process {
        foreach ($i in $item){
            if  ($i.path) {$Uri = $i.path +"/move"}
            else {
                if     ($i.id)             {$mailPath = "messages/$($i.id)/move"}
                elseif ($i -is [string])   {$mailPath = "messages/$i/move"}
                else {Write-warning "Could Not make sense of that item"}
                if ($User.id) {$Uri   = "$GraphUri/users/$($user.id)/$MailPath"}
                if ($User)    {$Uri   = "$GraphUri/users/$user/$MailPath" }
                else          {$Uri   = "$GraphUri/me/$MailPath" }
            }
            Write-Debug $uri
            $null = Invoke-GraphRequest @webparms -uri $uri
        }
    }
}
# only supporting move to deleted items however the DELETE method "really" deletes e.g  DELETE /users/{id | userPrincipalName}/messages/{id} DELETE /me/mailFolders/{id}/messages/{id}

function Save-GraphMailAttachment {
    param   (
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Attachment,
        #if Destination is a folder the file saved will use the name of the attachment. A file name can be specified.
        [Parameter(Position=2)]
        $Destination = (Get-Location),
        #if specfied the downloaded item(s) will be returned as a file
        [Alias('PT')]
        [switch]$PassThru
    )
    process {
        foreach ($a in $Attachment) {
            if        ($a -is [string] -and $a -match '/messages/.*/attachments/') {
                $uri = $a -replace '/\$value$','' }
            elseif    ($a.path) {
                $uri = $a.path
                if    ($a.Name) {$Filename = $a.Name}
            }
            else {Write-Warning 'Could not make sense of attachment provided'}
            if (Test-Path $Destination -PathType Container) {
                    if (-not $filename) {
                            $filename = (Invoke-GraphRequest "$uri`?`$select=Name").name
                    }
                    $outfile = Join-Path $Destination $filename
            }
            else  { $outfile = $Destination}
            Invoke-GraphRequest -OutputFilePath $outfile -Uri "$uri/`$value"
            if ($PassThru) {Get-Item $outfile}
        }
    }
}

function Send-GraphMailMessage {
    <#
    .Synopsis
    Sends Mail using the Graph API from the current user's mailbox. Requires "Mail.Send" permission.

    .PARAMETER User
    me or UserID as ID or User Principal name, whose calendar should be fetched If not specified defaults to "me", Requires "Mail.Send" API permission or delegated permission "Mail.Send if you want to use /me
    Default value: me --> delegated permissions needed

    .PARAMETER To
    Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.

    .PARAMETER CC
    Recipient(s) on the "CC" line

    .PARAMETER BCC
    Recipient(s) on the "Bcc line" line

    .PARAMETER Subject
    The subject of the message. A message must have a subject and/or body and/or attachments. If the subject is left blank it will be sent as "No Subject"

    .PARAMETER Body
    The type of the body  content. Possible values are Text and HTML.

    .PARAMETER BodyType
    The type of the body  content. Possible values are Text and HTML.

    .PARAMETER Importance
    The importance of the message: Low, Normal or High

    .PARAMETER Attachments
    Path to file(s) to send as attachments

    .PARAMETER Receipt
    If Specified, requests a receipt.

    .PARAMETER SaveDraftOnly
    If specified leaves the message in the drafts folder without sending it and returns a link to open the message.

    .PARAMETER NoSave
    If specified specifies that a copy of the mail should not be saved

    .Example
    >Send-GraphMail -User "jane@contoso.com" -To "chris@contoso.com" -subject "You left your keys behind[nt]"
    Sends a mail with a subject but no body or attachments

    .Example
    >Send-GraphMail -User "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -To "chris@contoso.com" -subject "You left your keys behind[nt]"
    Sends a mail with a subject but no body or attachments

    .Example
    >Send-GraphMail -To "chris@contoso.com" -subject "You left your keys behind[nt]"
    Sends a mail with a subject but no body or attachments
    .Example
    >Send-GraphMail -To "chris@contoso.com" -body "Keys are with reception" -NoSave
    Sends a mail but thi time the subject will read "No subject" and the test will be in the body.
    -NoSave means that no copy of this message will be kept in sent items
    .Example
    >Send-GraphMail -To "chris@contoso.com" -Subject "Screen shot" -body "How does this look ?" -Attachments .\Logon.png -Receipt
    #This message has an attachement and requests a read receipt.
    .Example
    >$body"<h1>New dialog</h1><br /><img src='cid:Logon.png' -alt='Look at that'><br/>what do you think"
    >$link = Send-GraphMail -To "jhoneill@waitrose.com" -Subject "Login Sreen" -body $body -BodyType HTML  -NoSave -Attachments .\Logon.png -SaveDraftOnly
    This creates an HTML body, the attached picture can be referenced in an <img> tag with cid:fileName.ext
    this time the mail is not sent but left in the user's drafts folder for review.
    #>
    [Cmdletbinding(DefaultParameterSetName = 'None')]
    param   (
        [parameter( HelpMessage = "`"me`" or UserID as ID or User Principal name, whose calendar should be fetched If not specified defaults to `"me`", Requires `"Mail.Send`" API permission or delegated permission `"Mail.Send if you want to use /me")]
        $User = "me",
        [parameter(Mandatory = $true, Position = 0, HelpMessage = "Recipient(s) on the `"to`" line, each is either created with New-MailRecipient (a hash table), or a string holding an address")]
        $To,
        [parameter(HelpMessage = "Recipient(s) on the `"CC`" line")]
        $CC,
        [parameter(HelpMessage = "Recipient(s) on the `"Bcc line`" line, not visible for other recipients")]
        $BCC,
        [parameter(HelpMessage = "The subject of the message. A message must have a subject and/or body and/or attachments. If the subject is left blank it will be sent as `"No Subject`"")]
        [String]$Subject,
        [parameter(HelpMessage = "The content of the message; assumed to be plain text, but HTML can be specified with -BodyType")]
        [String]$Body,
        [ValidateSet("Text", "HTML")]
        [parameter(HelpMessage = "The type of the body  content. Possible values are Text and HTML.")]
        $BodyType = "Text",
        [ValidateSet('Low', 'Normal', 'High')]
        [parameter(HelpMessage = "The importance of the message: Low, Normal or High")]
        $Importance = 'Normal',
        [parameter(HelpMessage = "Path to file(s) to send as attachments.")]
        $Attachments,
        [parameter(HelpMessage = "If Specified, requests a receipt.")]
        [switch]$Receipt,
        [parameter(ParameterSetName = 'SaveDraftOnly', Mandatory = $true, HelpMessage = "If specified leaves the message in the drafts folder without sending it and returns a link to open the message")]
        [switch]$SaveDraftOnly,
        [parameter(ParameterSetName = 'NoSave', Mandatory = $true, HelpMessage = "If specified specifies that a copy of the mail should not be saved")]
        [switch]$NoSave
    )

    #Do we post a message, or do we create a draft ? We need to check attacment sizes to be sure...
    $asDraft = [bool]$SaveDraftOnly
    if ($Attachments) {
        $AttachmentItems = Get-item $Attachments -ErrorAction SilentlyContinue
        if (-not $AttachmentItems) {
            Write-Warning (($Attachments -join ", ") + "Gave no items. Message sending will continue")
        }
        else {
            if ($Attachments.Where({ $_.length -gt 2.85mb })) {
                #The Maximum size for a POST is 4MB.
                #Attachments are base 64 encoded so 3MB of attachements become 4MB. Don't try closer than 95% of that
                throw ("Attachment would exceed maximum size for a POST. Maximum file size is ~ 2, 900, 000 bytes")
                return
            }
            elseif (-not $asDraft -and ($Attachments | Measure-Object -Sum length).sum -gt 2.7mb) {
                #If all the attaments add up to more than 90% of the possible message size, we need to
                #create a a draft and add each on its own.  BUT this method does not support "No save to sent items"
                if ($NoSave) {
                    throw ("The total size of attachments would result in an HTTP Post which greater than 4MB. Individual uploads are not possible when SaveToSentItems is disabled.")
                    return
                }
                else {
                    Write-Verbose -Message "SEND-GRAPHMAILMESSAGE After BASE64 encoding attacments, message may exceed 4MB. Using Draft and sequential attachment method"
                    $asDraft = $true
                }
            }
            else { Write-Verbose -Message "SEND-GRAPHMAILMESSAGE $($Attachments).count attachment(s); small enough to send in a single operation" }
        }
    }
    elseif (-not $Subject -and -not $Body) {
        Write-Warning -Message "Nothing to send" ; return
    }
    elseif (-not $Subject) { $Subject = "No subject" }
    $prefix = ""
    if ($User -ne "me") {
        $prefix = "/users" #needed for other users than /me
    }
    if ($asDraft) { $Uri = "$GraphUri$prefix/$User/Messages" }
    else { $Uri = "$GraphUri$prefix/$User/sendmail" }

    #Build a hash table with the parts of the message, this will be coverted into JSON
    #BEWARE names are case sensitive. if you create $msgSettings.Body instead of $msgSettings.body
    #the capital B will cause a 400 bad request error.
    #My personal coding style is to use inital CAPS for parameters and inital lower case for variables (though Powershell doesn't care)
    #so the parameter is $Body and the hash table key name and JSON label is body.

    $msgSettings = @{   'body' = @{
            'contentType' = $BodyType;
            'content'     = $Body
        }
        'subject'              = $Subject
        'importance'           = $Importance
        'toRecipients'         = @()
    }
    foreach ($recip in $To ) {
        if ($recip -is [string] ) { $msgSettings[ 'toRecipients'] += New-GraphRecipient $recip }
        else { $msgSettings[ 'toRecipients'] += $recip }
    }
    if ($CC) {
        $msgSettings['ccRecipients'] = @()
        foreach ($recip in $cc ) {
            if ($recip -is [string] ) { $msgSettings[ 'ccRecipients'] += New-GraphRecipient $recip }
            else { $msgSettings[ 'ccRecipients'] += $recip }
        }
    }
    if ($BCC) {
        $msgSettings['bccRecipients'] = @()
        foreach ($recip in $bcc ) {
            if ($recip -is [string] ) { $msgSettings['bccRecipients'] += New-GraphRecipient $recip }
            else { $msgSettings['bccRecipients'] += $recip }
        }
    }
    if ($Receipt) { $msgSettings['isDeliveryReceiptRequested'] = $true }

    #If we are creating a draft, save it now; if sending-in-one be ready for attachments
    if ($asDraft) {
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading draft"
        $json = ConvertTo-Json $msgSettings -Depth 5 #default depth isn't enough !
        try { $msg = Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch { throw "There was an error creating the draft message."; return }
        if (-not $msg) { throw "The draft message was not created as expected" ; return }
        else {
            Write-Verbose -Message "SEND-GRAPHMAILMESSAGE  Message created with id '$($msg.id)'"
            $uri = $uri + "/" + $msg.id
        }
    }
    elseif ($AttachmentItems) {
        $msgSettings["attachments"] = @()
    }

    foreach ($f in $AttachmentItems) {
        $Filesettings = @{
            '@odata.type' = '#microsoft.graph.fileAttachment';
            name          = $f.Name ;
            contentId     = $f.name ;
            contentBytes  = [convert]::ToBase64String( [system.io.file]::readallbytes($f.FullName))

        }
        if ($asDraft) {
            Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading $($f.Name)"
            try {
                $null = Invoke-GraphRequest -Method post  -uri "$uri/attachments"  -Body (ConvertTo-Json $Filesettings) -ContentType "application/json" -ErrorAction Stop
            }
            catch {
                Write-warning -Message "Error occured uploading file $($f.name) - will attempt to delete the draft message"
                Invoke-GraphRequest -Method Delete  -Uri "$uri"
                throw "Failure during attachment upload"
                return
            }
        }
        else {
            $msgSettings["attachments"] += $Filesettings
        }
    }

    if ($SaveDraftOnly) {
        Write-Progress -Activity "Sending Message" -Completed
        return $msg.webLink
    }
    elseif ($asDraft) {
        Write-Progress -Activity "Sending Message" -CurrentOperation "Sending Draft"
        try {
            Invoke-GraphRequest  -Method post  -uri "$uri/send" -Body " " # underlying stuff requires -body, but server ignores it.
            Write-Progress -Activity "Sending Message" -Completed
        }
        catch { throw "There was an error sending the draft message; it remains in the drafts folder" }
    }
    else {
        $mail = @{Message = $msgSettings }
        if ($NoSave) {
            $mail['saveToSentItems'] = $false
        }
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading and sending"

        $json = ConvertTo-Json $mail -Depth 10
        Write-Debug $Json
        try { Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch { throw "There was an error sending message."; return }
        Write-Progress -Activity "Sending Message" -Completed
    }
}

function Send-GraphMailReply      {
    <#
      .synopsis
        Replies to a mail message.
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param   (
        #Either a message ID or a Message object with an ID.
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        $Message,
        #Comment to attach when repling to the message - blank replies aren't allowed.
        [parameter(Mandatory=$true,Position=1)]
        $Comment,
        #If specified changes reply mode from reply [to sender] to Reply-to-all
        [Alias('All')]
        [switch]$ReplyAll
    )
    $msgSettings =  @{'comment' = $Comment }
    if ($Message.id) {$uri =  "$GraphUri/me/Messages/$($Message.id)/"}
    else             {$uri =  "$GraphUri/me/Messages/$Message"}
    if ($ReplyAll)   {$uri += '/replyAll' }
    else             {$uri += '/reply' }

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Debug $Json
    Invoke-GraphRequest -Method post -Uri $uri -ContentType 'application/json' -Body $json
}

function Send-GraphMailForward    {
    <#
      .synopsis
        Forwards a mail message.
      .example
      >
      > $alex = New-GraphRecipient Alex@contoso.com -DisplayName "Alex B."
      > Get-GraphMailItem -top 1 | Send-GraphMailForward -to $Alex -Comment "FYI :-)"
      Creates a recipient , and forwards the top mail in the users inbox to that recipent
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param   (
        #Either a message ID or a Message object with an ID.
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        $Message,
        #Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.
        [parameter(Mandatory=$true,Position=1)]
        $To ,
        #Comment to attach when forwarding the message.
        $Comment
    )
    $msgSettings   =  @{     toRecipients = @() }
    foreach ($recip in $To ) {
        if     ($recip  -is [string] ) { $msgSettings[ 'toRecipients'] += New-GraphRecipient $recip}
        else                           { $msgSettings[ 'toRecipients'] += $recip}
    }
    if ($Comment)                      { $msgSettings[ 'comment'] = $Comment}
    if ($Message.id) {$uri = "$GraphUri/me/Messages/$($Message.id)/forward"}
    else             {$uri = "$GraphUri/me/Messages/$Message/forward"}

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Debug $Json
    Invoke-GraphRequest -Method post -Uri $uri -ContentType 'application/json' -Body $json
}
#endregion

#region Outlook calendar -  only needs items found in the user module, so we don't give it it's own PS1 file
function New-GraphAttendee        {
    <#
      .Synopsis
        Helper function to create a new meeting attendee, with a mail address and the type of attendance.
    #>
    [cmdletbinding(DefaultParameterSetName='Default')]
    [alias('New-EventAttendee')]
    [outputType([system.collections.hashtable])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state.')]
    param   (
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true,ParameterSetName='Default',Mandatory=$true)]
        [Alias('Mail')]
        [String]$Address,
        #The displayname for the recipient
        [Parameter(Position=1, ValueFromPipelineByPropertyName=$true,ParameterSetName='Default')]
        [Alias('DisplayName')]
        $Name,
        #Is the attendee required or optional or a resource (such as a room). Defaults to required
        [ValidateSet('required', 'optional', 'resource')]
        $AttendeeType = 'required',
        [Parameter(ValueFromPipeline=$true,ParameterSetName='PipedStrings',Mandatory=$true)]
        $InputObject
    )
    #$EmailAddress = New-GraphMailAddress -Address $Address -DisplayName $Name
    # New-Object -TypeName MicrosoftGraphAttendee -Property @{emailaddress=$EmailAddress ; Type=$AttendeeType}
    process {
        if ($Address) {
            if (-not $Name) {$Name = $Address}
            @{ 'type'= $AttendeeType ; 'emailAddress' =  @{'address'=$Address; name=$Name }}
        }
    }
}

function New-GraphRecurrence      {
<#
    .synopsis
        Helper function to create the patterned recurrence for a task or event
    .links
        https://docs.microsoft.com/en-us/graph/api/resources/patternedrecurrence?view=graph-rest-1.0
#>
    [alias('New-RecurrencePattern')]
    param   (
        #The day of the month on which the event occurs. Required if type is absoluteMonthly or absoluteYearly.
        [ValidateRange(1,31)]
        [int]$DayOfMonth = 1,

        #Required if type is weekly, relativeMonthly, or relativeYearly. A collection of the days of the week on
        # which the event occurs. If type is relativeMonthly or relativeYearly,
        # and daysOfWeek specifies more than one day, the event falls on the first day that satisfies the pattern.
        [ValidateSet('sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday')]
        [String[]]$DaysOfWeek = @(),

        #The first day of the week. Default is sunday. Required if type is weekly.
        [ValidateSet('sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday')]
        [String]$FirstDayOfWeek ='sunday',

        #Specifies on which instance of the allowed days specified in daysOfsWeek the event occurs, counted from the first instance in the month.
        #Default is first. Optional and used if type is relativeMonthly or relativeYearly.
        [ValidateSet('first', 'second', 'third', 'fourth', 'last')]
        [string]$Index  ="first",

        #The number of units between occurrences, where units can be in days, weeks, months, or years, depending on the type. Defaults to 1
        [int]$Interval = 1,

        #The month in which the event occurs. This is a number from 1 to 12.
        [ValidateRange(1,12)]
        [int]$Month,

        #The recurrence pattern type:  daily = repeats based on the number of days specified by interval between occurrences.;
        #Weekly = repeats on the same day or days of the week, based on the number of weeks between each set of occurrences.
        #absoluteMonthly = Event repeats on the specified day of the month, based on the number of months between occurrences.
        #relativeMonthly = Event repeats on the specified day or days of the week, in the same relative position in the month, based on the number of months between occurrences.
        #absoluteYearly	Event repeats on the specified day and month, based on the number of years between occurrences.
        #relativeYearly	Event repeats on the specified day or days of the week, in the same relative position in a specific month of the year
        [validateSet('daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly')]
        [string]$Type = 'daily',

        #The number of times to repeat the event. Required and must be positive if type is numbered.
        $NumberOfOccurrences = 0 ,

        #The date to start applying the recurrence pattern. The first occurrence of the meeting may be this date or later, depending on the recurrence pattern of the event.
        #Must be the same value as the start property of the recurring event. Required
        [DateTime]$startDate = ([datetime]::now),

        #The date to stop applying the recurrence pattern. Depending on the recurrence pattern of the event, the last occurrence of the meeting may not be this date.
        [DateTime]$EndDate,

        # 'Time zone for the startDate and endDate properties. Optional. If not specified, the time zone of the event is used.'
        [string]$RecurrenceTimeZone
    )
    if ($endDate) {
        $range =   @{
            numberOfOccurrences = $numberOfOccurrences
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            endDate             = ($EndDate.ToString(  'yyyy-MM-dd') )
            recurrenceTimeZone  = $RecurrenceTimeZone
            type                = 'endDate'
        }
    }
    elseif ($numberOfOccurrences) {
        $range =  @{
            numberOfOccurrences = $numberOfOccurrences
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            type                = 'numbered'
        }
    }
    else {
        $range =  @{
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            type                = 'noEnd'
        }
    }
    $pattern =  @{
        dayOfMonth     = $DayOfMonth
        daysOfWeek     = $DaysOfWeek
        firstDayOfWeek = $FirstDayOfWeek
        index          = $Index
        interval       = $Interval
        month          = $month
        type           = $Type
    }
    return @{
            pattern   = $pattern
            range     = $range
    }
}

function Get-GraphCalendarPath    {
    param   (
        $Calendar,
        $Group,
        $User
    )
    #if we already have the path, just return it; if we got no parameters assume current user's default calendar.
    if ((-not $User -or $user -eq 'me')  -and
        (-not $Group) -and
        (-not $Calendar)   )                         {return '/me/calendar' }  #for the default calendar you can also use me/events or me/calendarView?params without "Calendar"
    elseif   ($Calendar -and $Calendar.CalendarPath) {return $Calendar.CalendarPath}
    elseif   ($Group)                                { # get the [only] calendar for a group
             $groupId = idfromGroup $group
             if (-not $groupId -or $groupId.count -gt 1 ) {
                throw ([System.Management.Automation.ParameterBindingException]::new('Cannot resolve the group information provided to a single group.'))
            }
            else {return "groups/$groupId/calendar" }  #for the default calendar you can also use groups/{id}/events or groups/calendarView?param without "Calendar"
    }
    if       ($Calendar -and -not $user)             { #if we got a calendar without a path or user assume it is current user's
         if  ($Calendar.id) {
              $path = "me/calendars/$($Calendar.id)"
              Add-Member -Force -InputObject $Calendar -NotePropertyName CalendarPath -NotePropertyValue $path
              return $path
          }
          else {return "me/calendars/$Calendar"}
    }

    #we must have a user with or without a calendar at this point - so resolve the user
    if       ($User -and $User.ID)               {$User = $User.ID}
    elseif   ($User -and $User -is [string]  -and
              $User -notmatch "\w@\w|$GUIDRegex")    {#Resolve name to UPN
            $User =  Invoke-GraphRequest  -ValueOnly ($GraphUri + '/users/?$Filter=' +
                            (FilterString $user  -ExtraFields 'userPrincipalName','givenName','surname','mail')) | ForEach-Object userPrincipalName
    }
    if       ($User.count -gt 1 -or -not $User)  {
             throw ([System.Management.Automation.ParameterBindingException]::new('Cannot resolve the user information provided to an account.'))
    }

    if (-not $Calendar)    {return "users/$user/calendar"} #get the default calendar for a specific user
    elseif  ($Calendar.id) {
             $path = "users/$User/calendars/$($Calendar.id)"
             Add-Member -Force -InputObject $Calendar -NotePropertyName CalendarPath -NotePropertyValue $path
             return $path
    }
    else                   {return "users/$user/calendars/$Calendar"}
    #for the default calendar you can also use users/{id}/events or users//calendarView?param without "Calendar"
}

function Get-GraphEvent           {
    <#
      .Synopsis
        Get the  events in a calendar
      .Description
        Depending on the parameters the events my come from
           * A specified calendar (retrieved by get-graphGroup or Get-GraphUser)
           * The default calendar for a group, (if only -group is provided)
           * The default calendar for a specific user, if only user is specified
           * The default calendar for the current user if no user, group, or calendar is specified.
           The request can specify the first n events in the calendar, or a number of days into
           the future, or specify the subject line or a custom filter.
      .Example
        >
        >Get-GraphEvent -Team consultants
        Finds the team (group) named "Consultants", and gets events in the team's calendar.
        Note that the because "team" and "group" are used interchangably the parameter is
        named "Group" with an alias of "Team"
      .Example
        >
        >get-graphuser -Calendars | where name -match "holidays" |
             get-graphevent -days 365 -order "start/datetime desc" -select start,end, subject |
                format-table subject, when
        Gets the user's calendars and selects the national holidays one;
        gets the events from this calendar for the next 365 days, sorting them to
        soonest last and selecting only the dates and subject; 'when' is calculated from
        start and end, so it is available to the format table command at the end of the pipeline.
      .Example
        >Get-GraphEvent -user alex@contoso.com -filter "isorganizer eq false"
        Gets events from the specified user's calendar where they are not the organizer;
        this requires access to have been granted access to the calendar by its owener.
      .Example
        >Get-GraphEvent  -filter "isorganizer eq false" -OrderBy start/datetime
        This uses the same filter as the previous example but sorts the results at the
        server before they are returned. Note that some fields like 'start' are record types,
        and one of their properties may need to be specified to perform a sort, as in this case,
        and the syntax is property/ChildProperty.
      .Example
        >
        >$userTimezone = (Get-GraphUser -MailboxSettings).timezone
        >Get-GraphEvent -Days 150 -TimeZone $userTimezone -Filter "showas eq 'free'"
        The first command gets the current user's preferred time zone, which may not
        match the local computer, and the second requests items for the next 150 days,
        where the time is shown as Free, displaying using that time zone
      .Example
        >Get-graphEvent -filter "start/dateTime ge '2019-04-01T08:00'"   | ft
        Gets the events in the signed-in user's default calendar which start after April 1 2019
        format-table will pick up the default display properties (Subject, When, Where and ShowAs)
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param   (
        #UserID as a guid or User Principal name, whose calendar should be fetched. "me" can be used as a shortcut for current user
        [Parameter( ParameterSetName="User",           ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter( ParameterSetName="UserAndSubject", ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter( ParameterSetName="UserAndFilter",  ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [ArgumentCompleter([UPNcompleter])]
        [string]$User,

        #A sepecific calendar
        [Parameter( ParameterSetName="Cal",            ValueFromPipelineByPropertyName=$true, Mandatory=$true, ValueFromPipeline=$true)]
        [Parameter( ParameterSetName="CalAndSubject",  ValueFromPipelineByPropertyName=$true, Mandatory=$true, ValueFromPipeline=$true)]
        [Parameter( ParameterSetName="CalAndFilter",   ValueFromPipelineByPropertyName=$true, Mandatory=$true, ValueFromPipeline=$true)]
        [Parameter( ParameterSetName="User",           ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Parameter( ParameterSetName="UserAndSubject", ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [Parameter( ParameterSetName="UserAndFilter",  ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        $Calendar,

        #Group ID or a Group object with an ID, whose calendar should be fetched
        [Parameter(ParameterSetName="GroupID",         ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="GroupAndSubject", ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="GroupAndFilter",  ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Alias("Team")]
        [ArgumentCompleter([GroupCompleter])]
        $Group,

        #Time zone to rennder event times. By default the time zone of the local machine will me use
        $Timezone = $(tzutil.exe /g),

        #Number of days of calendar to fetch from today
        [int]$Days,

        #The neumber of events to fetch. Must be greater than zero, and capped at 1000
        [ValidateRange(1,1000)]
        [int]$Top,

        #Fields to select
        [ValidateSet('attendees', 'body', 'bodyPreview', 'categories', 'changeKey', 'createdDateTime', 'end', 'hasAttachments',
                     'iCalUId', 'id', 'importance', 'isAllDay', 'isCancelled', 'isOrganizer', 'isReminderOn', 'lastModifiedDateTime',
                     'location', 'locations', 'onlineMeetingUrl', 'organizer', 'originalEndTimeZone', 'originalStart',
                     'originalStartTimeZone', 'recurrence', 'reminderMinutesBeforeStart', 'responseRequested', 'responseStatus',
                    'sensitivity', 'seriesMasterId', 'showAs', 'start', 'subject', 'type', 'webLink' )]
        [string[]]$Select,

        #An order-by clause to sort the events
        [string]$OrderBy,

        #If specified, fetch events where the subject line begins with
        [Parameter(ParameterSetName='CalAndSubject',   ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="UserAndSubject",  ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="GroupAndSubject", ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [string]$Subject,

        #A custom selection filter
        [Parameter(ParameterSetName="CurrentFilter",   ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="CalAndFilter",    ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="UserAndFilter",   ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [Parameter(ParameterSetName="GroupAndFilter",  ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
        [string]$Filter
    )

    begin   {
        $webParams = @{
            'AllValues'       = $true
            'ValueOnly'       = $true
            'ExcludeProperty' = 'icaluid','@odata.etag','calendar@odata.navigationLink','calendar@odata.associationLink'
            'AsType'          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
        }
        if ($TimeZone) {$webParams['Headers'] =@{"Prefer"="Outlook.timezone=""$TimeZone"""}}
    }
    process {
        $CalendarPath = Get-GraphCalendarPath -Calendar $Calendar -Group $Group -User $User
        $uri          = "$GraphUri/$CalendarPath"
        #region apply the selection criteria. If -days is specified use calendar view, otherwise use events and add filter, orderby, select and top as needed
        if  ($days)      {
                           $start = [datetime]::Today.ToString("yyyy-MM-dd't'HH:mm:ss")       # 'o' for ISO format time may work here.
                           $end   = [datetime]::Today.AddDays($days).tostring("yyyy-MM-dd't'HH:mm:ss")
                           $uri  += "/calendarview?`$expand=calendar&startdatetime=$start&enddatetime=$end"
        }
        else             { $uri  +=  '/events?$expand=calendar'}

        if ($Select)     { $uri  +=  '&$select=' + ($Select -join ',') }

        if ($Subject)    { $uri  += ('&$filter=startswith(subject,''{0}'')' -f ($subject -replace "'","''") ) }
        elseif ($Filter) { $uri  +=  '&$Filter='  + $Filter }

        if ($OrderBy)    { $uri  +=  '&$orderby=' + $orderby }

        if ($Top)        { $uri  +=  '&$top='     + $top  }
        #endregion
        #region get the data.
        Invoke-GraphRequest @webParams -Uri $uri |
            Add-Member -PassThru -NotePropertyName CalendarPath -NotePropertyValue $CalendarPath
        #endregion
    }
}

function Add-GraphEvent           {
    <#
      .Synopsis
        Adds an event to a calendar
      .link
        Get-GraphEvent
      .Example
        >
        >$rec = New-RecurrencePattern -Weekly Friday -EndDate 2019-04-01
        >Add-GraphEvent -Start "2019-01-23 15:30:00" -subject "Enter time sheet" -Recurrence $rec
        Creates a recurring event. The first sets up a weekly schedule for Fridays until April 1st.
        The second sets the time  (if no end is given, it is set for 30 minutes after the start),
        the subject, and the recurrence pattern
      .Example
        >
        >$Chris = New-Attendee -Mail Chris@Contoso.com
        >Add-GraphEvent -subject "Requirements for Basingstoke project" -Start "2019-02-02 10:00" -End "2019-02-02 11:00" -Attendees $chris
        Creates a meeting with a second person. The first command creates an attendee - by default the attendee is 'required'
        The second creates the appointment, adding the attendee and sending a meeting request.
      .Example
        >
        >$Chris = New-Attendee -Mail Chris@Contoso.com -display 'Chris Cross' optional
        >$Phil  = New-Attendee -Mail Phil@Contoso.com
        >Add-GraphEvent -subject "Phase II planning" -Start "2019-02-02 14:00" -End "2019-02-02 14:30" -Attendees $chris,$phil
        Creates a meeting with a second additonal attendee. The first command creates an optional attendee with a display name
        the second creates an attendee with no displayed name and the default 'required' type
        Finally the meeting is created.
    #>
    [cmdletbinding()]
    param   (
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [ArgumentCompleter([UPNCompleter])]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
        [ArgumentCompleter([GroupCompleter])]
        $Group,

        #Subject for the appointment
        [string]$Subject ,

        #Start time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter(Mandatory=$true)]
        [datetime]$Start,

        #End Time - if -Timezone is not used this will be the in local machine's times zone
        [datetime]$End,

        #Timezone - by default the local machine's time zone is used
        $Timezone = $(tzutil.exe /g),

        #Creates the event as all day.
        [Switch]$AllDay,

        #Location for the appointment
        $Location,

        #People or resources involved in the event.
        $Attendees,
        #Sets the task to appear as Free, Tenatative, Off-of-facility etc
        [ValidateSet('busy','free','oof','tentative','workingElsewhere')]
        [string]$ShowAs,

        #Unless -Reminder on is specified no reminder will sound before the meeting
        [switch]$ReminderOn,
        #Time in Minutes, before the start time, that the reminder should appear. It will be set even if -ReminderOn is omitted
        $ReminderTime,
        #Body text - if using HTML set the body type to HTML
        $Body  ,
        #Type of text used for the body, Text or HTML
        [ValidateSet('Text','HTML')]
        $BodyType = 'Text',
        #Priority setting , high , normal or low.
        [ValidateSet('low','normal','high')]
        [String]$Importance ,
        #Privacy setting - normal or Private
        [ValidateSet('normal','private')]
        [String]$Sensitivity,

        #Recurrence pattern build with New-recurrencePattern
        $Recurrence,

        [alias('PT')]
        [switch]$PassThru
        # for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
        # and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
        # Attendees is one. link says this also sends the invite

    )
    begin   {
        $webParams = @{
                    'Method'          = 'Post'
                    'ExcludeProperty' = 'icaluid','@odata.etag','@odata.context'
                    'AsType'          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
                    'Contenttype'     = 'application/json'
                    'Headers'         = @{Prefer        = "Outlook.timezone=""$TimeZone"""}
        }
    }
    process {
        $CalendarPath     = Get-GraphCalendarPath -Calendar $Calendar -Group $Group -User $User
        $webParams['uri'] = $GraphUri + "/" + $CalendarPath + '/events'

        #region assemble the body needed to create the event
        $settings = @{      'subject'=  $Subject;     'isReminderOn' = [bool]$ReminderOn}
        if ($Location)     {$settings['location']                    = @{'displayName'=$Location} }
        if ($Body)         {$settings['body']                        = @{'contentType'=$BodyType ; 'Content'=$Body}}
        if ($ShowAs)       {$settings['showAs']                      = $ShowAs}
        if ($ReminderTime) {$settings['reminderMinutesBeforeStart']  = $ReminderTime}
        if ($Importance)   {$settings['importance']                  = $Importance}
        if ($Sensitivity)  {$settings['sensitivity']                 = $Sensitivity}
        if ($AllDay)       {$settings['isAllDay']                    = $true
                            $Start = $Start.Date
                            if (-not $End)        { $End             = $Start.AddDays(1)}
                            else                  { $End             = $End.Date        }
                            if ($End -eq $Start)  { $End             = $End.AddDays(1)  }
        }
        elseif (-not $End) {$End = $Start.AddMinutes(30)    }
        $settings['start' ] = @{'timeZone'=$Timezone; 'dateTime'    = $Start.ToString("yyyy-MM-dd'T'HH:mm:ss")} ;
        $settings['end'   ] = @{'timeZone'=$Timezone; 'dateTime'    = $End.ToString(  "yyyy-MM-dd'T'HH:mm:ss")};

        if ($Recurrence)   {$settings['recurrence']                 = $Recurrence
                            $settings.recurrence.range['startDate'] = $Start.ToString('yyyy-MM-dd');
        }
        if ($Attendees)    {$settings['attendees'] = @() + $Attendees }
        $json =  (ConvertTo-Json $settings -Depth 10)
        Write-Debug $json
        #endregion

        $result = Invoke-GraphRequest @webParams -Body $json
        if ($PassThru) {$result | Add-Member -PassThru -NotePropertyName CalendarPath -NotePropertyValue $CalendarPath}
    }
}

function Set-GraphEvent           {
    <#
      .Synopsis
        Modifies an event on a calendar
      .link
        Get-GraphEvent
    #>
    [cmdletbinding(SupportsShouldProcess=$true,DefaultParameterSetName='None')]
    param   (
        #The event to be updateds either as an ID or as an event object containing an ID.
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)]
        $Event,

        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [ArgumentCompleter([UPNCompleter])]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
        [ArgumentCompleter([GroupCompleter])]
        $Group,

        #Subject for the appointment
        [string]$Subject ,

        #Start time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter( ParameterSetName="Group" )]
        [Parameter( ParameterSetName="None"  )]
        [Parameter( ParameterSetName="User"   )]
        [Parameter( ParameterSetName="AllDay", Mandatory=$true )]
        [Nullable[datetime]]$Start,

        #End Time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter( ParameterSetName="Group"  )]
        [Parameter( ParameterSetName="None"   )]
        [Parameter( ParameterSetName="User"   )]
        [Parameter( ParameterSetName="AllDay", Mandatory=$true )]
        [Nullable[datetime]]$End,

        #Creates the event as all day - you must also set the start and end time.
        [Parameter(Mandatory=$true, ParameterSetName="AllDay")]
        [switch]$AllDay,

        #Timezone - by default the local machine's time zone is used
        $Timezone = $(tzutil.exe /g),
        #Location for the appointment
        $Location,
        #Body text - if using HTML set the body type to HTML
        $Body  ,
        #Type of text used for the body, Text or HTML
        [ValidateSet('Text','HTML')]
        $BodyType = 'Text',
        #Unless -Reminder on is specified no reminder will sound before the meeting
        [switch]$ReminderOn,
        #Time in Minutes, before the start time, that the reminder should appear. It will be set even if -ReminderOn is omitted
        $ReminderTime,
        #Sets the task to appear as Free, Tenatative, Off-of-facility etc
        [ValidateSet('busy','free','oof','tentative','workingElsewhere')]
        [string]$ShowAs,
        #Priority setting , high , normal or low.
        [ValidateSet('low','normal','high')]
        [String]$Importance ,
        #Privacy setting - normal or Private
        [ValidateSet('normal','private')]
        [String]$Sensitivity,

        #Recurrence pattern build with New-recurrencePattern
        $Recurrence,
        #If specified the update will be performed without prompting for confirmation (this is the default)
        [switch]$Force,
        # for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
        # and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
        # Attendees is one. link says this also sends the invite
        [switch]$PassThru
    )

    $webParams = @{
                   'Method'          = 'Patch'
                   'ExcludeProperty' = 'icaluid','@odata.etag','@odata.context'
                   'AsType'          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
                   'Contenttype'     = 'application/json'
                   'Headers'         = @{Prefer        = "Outlook.timezone=""$TimeZone"""}
    }

    if   ($Event.calendarPath) {$CalendarPath = $Event.calendarPath}
    else {$CalendarPath = Get-GraphCalendarPath -Calendar $Calendar -Group $Group -User $User}

    if ($Event.id) {$webParams['uri'] += $GraphUri + $CalendarPath + "/events/$($Event.id)"}
    else           {$webParams['uri'] += $GraphUri + $CalendarPath + "/events/$Event"      }
    #region assemble the body needed to update the event
    $settings  =   @{ }
    if ($PSBoundParameters.ContainsKey('ReminderOn')) {
                           $settings['isReminderOn']    = [bool]$ReminderOn}
    if ($PSBoundParameters.ContainsKey('AllDay'))     {
                           $settings['isAllDay']        = [bool]$AllDay  }
    if ($Subject)         {$settings['subject']         = $subject };
    if ($Location)        {$settings['location']        = @{'displayName'=$Location} }
    if ($Body)            {$settings['body']            = @{'contentType'=$BodyType ; 'Content'=$Body}}
    if ($ShowAs)          {$settings['showAs']          = $ShowAs}
    if ($Importance)      {$settings['importance']      = $Importance}
    if ($Sensitivity)     {$settings['sensitivity']     = $Sensitivity}
    if ($ReminderTime)    {$settings['reminderMinutesBeforeStart'] = $ReminderTime}
    if ($Start)           {$settings['start'] = @{'timeZone'       = $Timezone; 'dateTime' = $Start.ToString("yyyy-MM-dd'T'HH:mm:ss")} }
    if ($End)             {$settings['end'  ] = @{'timeZone'       = $Timezone; 'dateTime' = $End.ToString(  "yyyy-MM-dd'T'HH:mm:ss")} }
    if ($Recurrence)      {$settings['recurrence']                 = $Recurrence
                           $settings.recurrence.range['startDate'] = $Start.ToString('yyyy-MM-dd');
    }
    $json =  (ConvertTo-Json $settings -Depth 10)
    Write-Debug $json
    #endregion

    if ($Force -or $PSCmdlet.ShouldProcess($Event.subject,'Update calendar event')) {
        $result = Invoke-GraphRequest @webParams -Body $json
        if ($PassThru) {$result | Add-Member -PassThru -NotePropertyName CalendarPath -NotePropertyValue $CalendarPath }
     }
}

function Remove-GraphEvent        {
    <#
      .Synopsis
        Deletes an item from the calendar
      .Description
        Deletes items from the calendar. If other people have beeen invited to a meeting,
        they will reveive a cancellation message.
    #>
    [cmdletbinding(DefaultParameterSetName="None",SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (

        #The event to be removed either as an ID or as an event object containing an ID.
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        $Event,

        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",Mandatory=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",Mandatory=$true)]
        $Calendar,

        #Group ID or a Group object with an ID, whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="GroupID")]
        [Alias("Team")]
        $Group,

        #if Sepcified the event will be deleted without prompting for confirmation
        [switch]$Force
    )
    process {
        if     ($event.calendarpath) {$CalendarPath = $Event.calendarPath}
        else   {$CalendarPath = Get-GraphCalendarPath -Calendar $Calendar -Group $Group -User $User}

        if ($Force -or $PSCmdlet.ShouldProcess($Event.Subject ,'Delete from calendar')) {
            if ($Event.ID) { Invoke-GraphRequest -Method Delete -Uri  ($GraphUri + $calendarPath  + "/Events/$($Event.ID)") }
            else           { Invoke-GraphRequest -Method Delete -Uri  ($GraphUri + $calendarPath  + "/Events/$($Event.ID)") }
        }
    }
}
#endregion

#region to-do-list functions are here because they are in the User module  -- they require the Tasks.ReadWrite  scope
function ConvertTo-GraphDateTimeTimeZone {
    <#
        .synopsis
            Converts a datetime object to dateTimeTimezone object with the current time zone.
    #>
    param   (
        [dateTime]$d
    )
    New-object MicrosoftGraphDateTimeZone -Property @{
                Datetime = $d.ToString('yyyy-MM-ddTHH:mm:ss')
                Timezone = (Get-TimeZone).id
    }
}

function Get-GraphToDoList        {
    <#
      .Synopsis
        Gets information about lists used in the To Do app.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param   (
        #The ID of the plan or a plan object with an ID property. if omitted the current users planner will be assumed.
        [Parameter( ValueFromPipeline=$true,Position=0)]
        [alias('id')]
        $ToDoList = 'defaultList',

        #The User ID (GUID or UPN) of the list owner. Defaults to the current user.
        $UserId,

        #If specified returns the tasks in the list.
        [switch]$Tasks
    )
    process {
        contexthas -WorkOrSchoolAccount -BreakIfNot
        if ($UserId) {$uri    = "$GraphUri/users/$userid/todo/lists"}
        else         {$uri    = "$GraphUri/me/todo/lists"
                      $UserId =  $Global:GraphUser
        }
        if ( $ToDoList -is [string] -and  $ToDoList -match "\w{100}" ) {
            try {
                 $ToDoList  = Invoke-GraphRequest -Uri "$uri/$ToDoList" -ExcludeProperty  "@odata.etag", "@odata.context" -AsType ([MicrosoftGraphTodoTaskList]) |
                                Add-Member -PassThru -NotePropertyName UserId -NotePropertyValue $UserId
            }
            catch {
                $ToDoList = $PSBoundParameters['ToDoList']
                Write-Warning 'To Do list parameter looks like a ID but did not return a list '
            }
        }
        if ($ToDoList -is [String]) { #including if the last step tried and failed.
             $ToDoList = Invoke-GraphRequest -Uri $uri -ValueOnly -ExcludeProperty "@odata.etag" -AsType ([MicrosoftGraphTodoTaskList]) |
                            Where-Object {$_.displayname  -like $ToDoList -or $_.WellknownListName -like $ToDoList}
        }
        if (-not ($ToDoList.displayName -and $ToDoList.id)) {
            Write-Warning "Could not get a To Do list from the information provided" ; return
        }
        if     (-not $Tasks) {return $ToDoList}
        else  {
            if (-not $UserId) {$UserId = $Global:GraphUser }
            Invoke-GraphRequest  -Method get  -uri "$uri/$($ToDoList.id)/tasks" -ValueOnly -ExcludeProperty "@odata.etag" -AsType ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphTodoTask]) |
                Add-Member -PassThru -NotePropertyName UserId   -NotePropertyValue $userID |
                Add-Member -PassThru -NotePropertyName ListID   -NotePropertyValue $ToDoList.Id |
                Add-Member -PassThru -NotePropertyName ListName -NotePropertyValue $ToDoList.DisplayName
        }
    }
}

function New-GraphToDoList        {
<#
    .synopsis
        Creates a new list for the To-Do app
#>
[cmdletBinding(SupportsShouldProcess=$true)]
param   (
    [parameter(Mandatory=$true,Position=0)]
    #The name for the list
    [string]$Displayname    ,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user,
    $UserId =  $Global:GraphUser,

    #If specified the the list will be created as a shared list
    [switch]$IsShared,

    #If specified any confirmation will be supressed
    [switch]$Force
)
    if ($Force -or $pscmdlet.ShouldProcess($Displayname,"Create new To-Do list")){
        Test-GraphSession
        Microsoft.Graph.Users.private\New-MgUserTodoList_CreateExpanded1 -UserId $UserId -DisplayName $displayname -IsShared:$IsShared -Confirm:$false |
                Add-Member -PassThru -NotePropertyName UserId -NotePropertyValue $UserId
    }
}

function New-GraphToDoTask        {
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (

        #A To-do list object or the ID of a To-do list
        [Parameter()]
        [alias('TodoTaskListId')]
        $ToDoList,

        #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on theToDo list object
        [Parameter()]
        [string]
        $UserId =  $Global:GraphUser,

        # A brief description of the task.
        [Parameter(mandatory=$true, position=0)]
        [string]$Title,

        #The text or HTML content of the task body
        [string]$BodyText,

        #The type of the content. Possible values are text and html, defaults to Text
        [ValidateSet('text', 'html')]
        [string]$BodyType = 'text',

        #The importance of the task.
        [ValidateSet('low', 'normal', 'high')]
        [string]$Importance = 'normal',

        #The date/time in the current time zone that the task is to be finished.
        [datetime]$DueDateTime,

        #Indicates the state or progress of the task.
        [ValidateSet('notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred')]
        [string]$Status = 'notStarted',

        #The date/time in the current time zone that the task was finished.
        [datetime]$CompletedDateTime,

        #The date and time for a reminder alert of the task to occur.
        [datetime]$ReminderDateTime,

        #The recurrence pattern for the task. - May be created with Get-GraphRecurrence
        $Recurrence,

        #If specified any confirmation will be supressed
        [switch]$Force
    )

    if ($userID -and -not $ToDoList) {$ToDoList = Get-GraphToDoList}
    if ($ToDoList.userID)  {$userID   = $ToDoList.userId}
    if ($ToDoList.ID)      {$ToDoList = $ToDoList.Id}

    $Params =  @{
        'UserId'          = $UserId
        'TodoTaskListId'  = $ToDoList
        'Title'           = $Title
        'Body'            = (New-Object -TypeName MicrosoftGraphItemBody -Property @{content=$BodyText; contentType=$BodyType} )
        'Importance'      = $Importance
        'Status'          = $status
        'IsReminderOn'    = $ReminderDateTime -as [bool]
    }
    if ($Recurrence)        {$Params['Recurrence']        = $Recurrence
                             if (-not $DueDateTime) {$DueDateTime = [datetime]::Today.AddDays(1)}
    }
    if ($ReminderDateTime)  {$Params['ReminderDateTime']  = (ConvertTo-GraphDateTimeTimeZone $ReminderDateTime)}
    if ($CompletedDateTime) {$Params['CompletedDateTime'] = (ConvertTo-GraphDateTimeTimeZone $CompletedDateTime)}
    if ($DueDateTime)       {$Params['DueDateTime']       = (ConvertTo-GraphDateTimeTimeZone $DueDateTime)}

    if ($Force -or $PSCmdlet.ShouldProcess($title,"Add NewTask")) {
        Test-GraphSession
        Microsoft.Graph.Users.private\New-MgUserTodoListTask_CreateExpanded1 @params |
            Add-Member -PassThru -NotePropertyName UserId   -NotePropertyValue $userID |
            Add-Member -PassThru -NotePropertyName ListID   -NotePropertyValue $ToDoList
    }
}

function Update-GraphToDoTask     {
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
    #A Task object or the ID of a task.
    [Parameter(mandatory=$true,ValueFromPipelineByPropertyName =$true, ValueFromPipeline=$true)]
    [alias('ID')]
    $Task,

    #A To-do list object or the ID of a To-do list, may be found on the task object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [alias('TodoTaskListId','ListID')]
    $ToDoList,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on the task or the ToDo list object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$UserId =  $Global:GraphUser,

    # A brief description of the task.
    [Parameter(position=0)]
    [string]$Title,

    #The text or HTML content of the task body
    [string]$BodyText,

    #The type of the content. Possible values are text and html, defaults to Text
    [ValidateSet('text', 'html')]
    [string]$BodyType = 'text',

    #The importance of the task.
    [ValidateSet('low', 'normal', 'high')]
    [string]$Importance,

    #The date/time in the current time zone that the task is to be finished.
    [datetime]$DueDateTime,

    #Indicates the state or progress of the task.
    [ValidateSet('notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred')]
    [string]$Status ,

    #The date/time in the current time zone that the task was finished
    [datetime]$CompletedDateTime,

    #The date and time for a reminder alert of the task to occur.
    [datetime]$ReminderDateTime,

    #Turns off any alert which has been set to remind the user of the task.
    [switch]$ReminderOff,

    #The recurrence pattern for the task. - May be created with Get-GraphRecurrence
    $Recurrence,

    #If specified, no confirmation will be displayed before updating the task
    [switch]$Force
    )
    process {
        if (-not $Task.ListID -and -not $ToDoList) {
            Write-Warning "Could not obtain Task list ID" ; return
        }
        elseif (-not $ToDoList) {$ToDoList = $task.ListID}

        if ((-not $userID)-and -not ($Task.userid -or $ToDoList.userID )) {
            Write-Warning "Could not obtain Task list ID" ; return
        }
        elseif ((-not $PSBoundParameters['userID']) -and $Task.userid )     {$userID   = $Task.userId}
        elseif ((-not $PSBoundParameters['userID']) -and $ToDoList.userid ) {$userID   = $ToDoList.userId}

        if ($ToDoList.ID)      {$ToDoList = $ToDoList.Id}
        if ($Task.id)          {$Task     = $Task.id}

        $Params =  @{
            TodoTaskId      = $Task
            UserId          = $UserId
            TodoTaskListId  = $ToDoList
            IsReminderOn    = (-not $ReminderOff)
        }
        if ($Title)      {      $Params['Title']              = $Title}
        if ($BodyText)   {      $Params['Body']               = New-Object -TypeName MicrosoftGraphItemBody -Property @{content=$BodyText; contentType=$BodyType} }
        if ($Importance) {      $Params['Importance']         = $Importance}
        if ($Status)     {      $Params['Status']             = $Status}
        if ($Recurrence) {      $Params['Recurrence']         = $Recurrence ;  if (-not $DueDateTime) {
                                $DueDateTime                  =  [datetime]::Today.AddDays(1)}
        }
        if ($ReminderDateTime)  {$Params['ReminderDateTime']  = (ConvertTo-GraphDateTimeTimeZone $ReminderDateTime)}
        if ($CompletedDateTime) {$Params['CompletedDateTime'] = (ConvertTo-GraphDateTimeTimeZone $CompletedDateTime)}
        if ($DueDateTime)       {$Params['DueDateTime']       = (ConvertTo-GraphDateTimeTimeZone $DueDateTime)}
        if ($force -or $pscmdlet.ShouldProcess("Task update")) {
            Test-GraphSession
            Microsoft.Graph.Users.private\Update-MgUserTodoListTask_UpdateExpanded1 @params
        }
    }
}

function Remove-GraphToDoTask     {
    <#
        .synopsis
            Removes a task from the To Do app
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
    #A Task object or the ID of a task.
    [Parameter(mandatory=$true,ValueFromPipelineByPropertyName =$true, ValueFromPipeline=$true)]
    [alias('ID')]
    $Task,

    #A To-do list object or the ID of a To-do list, may be found on the task object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [alias('TodoTaskListId','ListID')]
    $ToDoList,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on the task or the ToDo list object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$UserId =  $Global:GraphUser,

    #If specified, no confirmation will be displayed before deleting the task
    [switch]$Force
    )
    process {
        if (-not $Task.ListID -and -not $ToDoList) {
            Write-Warning "Could not obtain Task list ID" ; return
        }
        elseif (-not $ToDoList) {$ToDoList = $Task.ListID}

        if ((-not $userID)-and -not ($ToDoList.userID -or $Task.userid)) {
            Write-Warning "Could not obtain Task list ID" ; return
        }
        elseif ((-not $PSBoundParameters['userID'])  -and $Task.userid )     {$userID   = $Task.userId}
        elseif ((-not $PSBoundParameters['userID'])  -and $ToDoList.userid ) {$userID   = $ToDoList.userId}

        if ($ToDoList.ID)      {$ToDoList = $ToDoList.Id}
        if ($Task.Title)       {$Title    = $Task.title}
        if ($Task.id)          {$Task     = $Task.id}

        $Params =  @{
            TodoTaskId      = $Task
            UserId          = $UserId
            TodoTaskListId  = $ToDoList
        }
        if ($force -or $pscmdlet.ShouldProcess($Title,'Task deletion')) {
            Test-GraphSession
            Microsoft.Graph.Users.private\Remove-MgUserTodoListTask_Delete1 @Params
        }
    }
}

function Remove-GraphToDoList     {
    <#
        .synopsis
            Removes a list from the To Do app, including any task in contains.
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
    #A To-do list object or the ID of a To-do list
    [Parameter(mandatory=$true,ValueFromPipelineByPropertyName =$true, ValueFromPipeline=$true)]
    [alias('TodoTaskListId','ListID')]
    $ToDoList,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on the ToDo list object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$UserId =  $Global:GraphUser,

    #If specified, no confirmation will be displayed before deleting the list
    [switch]$Force
    )
    process {
        if ((-not $userID)-and -not ($ToDoList.userID )) {
            Write-Warning "Could not obtain Task list ID" ; return
        }
        elseif ((-not $PSBoundParameters['userID']) -and $ToDoList.userid )     {$userID   = $ToDoList.userId}
        if ($ToDoList.Displayname)  {$Title    = $ToDoList.Displayname}
        if ($ToDoList.ID)           {$ToDoList = $ToDoList.Id}

        $Params =  @{
            UserId          = $UserId
            TodoTaskListId  = $ToDoList
        }
        if ($force -or $pscmdlet.ShouldProcess($Title,'Delete whole to-do list')) {
                Microsoft.Graph.Users.private\Remove-MgUserTodoList_Delete1 @Params
        }
    }
}
#endregion