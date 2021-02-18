using namespace Microsoft.Graph.PowerShell.Models

function ConvertTo-GraphDateTimeTimeZone {
    <#
        .synopsis
            Converts a datetime object to dateTimeTimezone object with the current time zone.
    #>
    param (
        [dateTime]$d
    )
    New-object MicrosoftGraphDateTimeZone -Property @{
                Datetime = $d.ToString('yyyy-MM-ddTHH:mm:ss')
                Timezone = (Get-TimeZone).id
    }
}

function Get-GraphUserList     {
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
    param(
        #If specified searches for users whose first name, surname, displayname, mail address or UPN start with that name.
        [parameter(Mandatory=$true, parameterSetName='FilterByName', Position=1,ValueFromPipeline=$true )]
        [string[]]$Name,

        #Names of the fields to return for each user.
        [validateSet('accountEnabled', 'ageGroup', 'assignedLicenses', 'assignedPlans', 'businessPhones', 'city',
                    'companyName', 'consentProvidedForMinor', 'country', 'createdDateTime', 'department',
                    'displayName', 'givenName', 'id', 'imAddresses', 'jobTitle', 'legalAgeGroupClassification',
                    'mail','mailboxSettings', 'mailNickname', 'mobilePhone', 'officeLocation',
                    'onPremisesDomainName', 'onPremisesExtensionAttributes', 'onPremisesImmutableId',
                    'onPremisesLastSyncDateTime', 'onPremisesProvisioningErrors', 'onPremisesSamAccountName',
                    'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName',
                    'passwordPolicies', 'passwordProfile', 'postalCode', 'preferredDataLocation',
                    'preferredLanguage', 'provisionedPlans', 'proxyAddresses', 'state', 'streetAddress',
                    'surname', 'usageLocation', 'userPrincipalName', 'userType')]
        [Alias('Select')]
        [string[]]$Property,

        #Order by clause for the query - most fields result in an error and it can't be combined with some other query values.
        [parameter(Mandatory=$true, parameterSetName='Sorted')]
        [ValidateSet('displayName', 'userPrincipalName')]
        [Alias('OrderBy')]
        [string]$Sort,

        #Filter clause for the query for example "startswith(displayname,'Bob') or startswith(displayname,'Robert')"
        [parameter(Mandatory=$true, parameterSetName='FilterByString')]
        [string]$Filter,

        [validateSet('directReports', 'manager', 'memberOf', 'ownedDevices', 'ownedObjects', 'registeredDevices', 'transitiveMemberOf',  'extensions')]
        [string]$ExpandProperty,

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
        Write-Progress "Getting the List of users"
        if (-not $Name) {
            Microsoft.Graph.Users.private\Get-MgUser_List  -ConsistencyLevel eventual -All @PSBoundParameters
        }
        else {
            [void]$PSBoundParameters.Remove('Name')
            foreach ($n in $Name) {
                $PSBoundParameters['Filter'] = ("startswith(displayName,'{0}') or startswith(givenName,'{0}') or startswith(surname,'{0}') or startswith(mail,'{0}') or startswith(userPrincipalName,'{0}')" -f $n )
                Microsoft.Graph.Users.private\Get-MgUser_List  -ConsistencyLevel eventual -All @PSBoundParameters
            }
    }
    Write-Progress "Getting the List of users" -Completed
    }
}

function Get-GraphUser         {
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
        in the root directory. So this command shows the names of the objects in the root directory.
    #>
    [cmdletbinding(DefaultparameterSetName="None")]
    param   (
        #UserID as a guid or User Principal name. If not specified, it will assume "Current user" if other paraneters are given, or "All users" otherwise.
        [parameter(Position=0,valueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [alias('id')]
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

        #specifies which properties of the user object should be returned ( aboutMe, birthday, deviceEnrollmentLimit, hireDate,interests,mailboxSettings,mySite,pastProjects,preferredName,responsibilities,schools and skills are not available)
        [parameter(Mandatory=$true,parameterSetName="Select")]
        [ValidateSet  (
        'accountEnabled', 'activities', 'ageGroup', 'appRoleAssignments', 'assignedLicenses', 'assignedPlans',  'businessPhones',
        'calendar', 'calendarGroups', 'calendars', 'calendarView', 'city', 'companyName', 'consentProvidedForMinor', 'contactFolders', 'contacts', 'country', 'createdDateTime', 'createdObjects', 'creationType', 'department',
        'deviceManagementTroubleshootingEvents', 'directReports',
        'displayName', 'drive', 'drives', 'employeeHireDate', 'employeeId', 'employeeOrgData', 'employeeType', 'events', 'extensions', 'externalUserState',
        'externalUserStateChangeDateTime', 'faxNumber', 'followedSites', 'givenName',  'ID', 'identities', 'imAddresses', 'inferenceClassification',
        'insights', 'isResourceAccount', 'jobTitle', 'joinedTeams', 'lastPasswordChangeDateTime', 'legalAgeGroupClassification', 'licenseAssignmentStates',
        'licenseDetails', 'mail', 'mailFolders', 'mailNickname', 'managedAppRegistrations', 'managedDevices', 'manager', 'memberOf', 'messages',
        'mobilePhone', 'oauth2PermissionGrants', 'officeLocation', 'onenote', 'onlineMeetings', 'onPremisesDistinguishedName',
        'onPremisesDomainName', 'onPremisesExtensionAttributes', 'onPremisesImmutableId', 'onPremisesLastSyncDateTime', 'onPremisesProvisioningErrors',
        'onPremisesSamAccountName', 'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName', 'otherMails', 'outlook',
        'ownedDevices', 'ownedObjects', 'passwordPolicies', 'passwordProfile',  'people', 'photo', 'photos', 'planner', 'postalCode',
        'preferredLanguage', 'presence', 'provisionedPlans', 'proxyAddresses', 'registeredDevices', 'scopedRoleMemberOf', 'settings', 'showInAddressList',
        'signInSessionsValidFromDateTime',   'state', 'streetAddress', 'surname',
         'teamwork', 'todo', 'transitiveMemberOf', 'usageLocation', 'userPrincipalName', 'userType')]
        [String[]]$Select,

        #Used to explicitly say "Current user" and will over-ride UserID if one is given.
        [switch]$Current

    )
    begin   {
        $result       = @()
    }
    process {
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
        if ($UserID -is [array] -and $UserID -notmatch "$GuidRegex|\w@\w|me" -and
                                         $UserID -match     $GuidRegex ) {
            Write-Warning   -Message 'If you pass an array of values they cannot be names. You can pipe names or pass and array of IDs/UPNs' ; return
        }
        #if it is a string and not a guid or UPN - or an array where at least some members are not GUIDs/UPN/me try to resolve it
        elseif ($UserID -notmatch "$GuidRegex|\w@\w|me" ) {
                $UserID = Get-GraphUserList -Name $UserID
        }
        #endregion
        #if select is in use ensure we get ID, UPN and Display-name.
        if ($Select) {
            foreach ($s in @('ID','userPrincipalName','displayName')) {
                 if ($s -notin $select) {$select += $s }
            }
        }
        [void]$PSBoundParameters.Remove('UserID')
        foreach ($id in $UserID) {
            #region set up the user part of the URI that we will call
            if ($id -is [MicrosoftGraphUser] -and -not  ($PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters})  )) {
                $id
                continue
            }
            if     ($id.id)       { $ID = $id.Id}
            Write-Progress -Activity 'Getting user information' -CurrentOperation "User = $ID"
            if     ($id -eq 'me') { $Uri = "$GraphUri/me"  }
            else                  { $Uri = "$GraphUri/users/$id" }

            # -Teams requires a GUID, photo doesn't work for "me"
            if (  ($Teams -and $id -notmatch $GuidRegex ) -or
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
                transitiveMemberOf
            "https://graph.microsoft.com/v1.0/me/getmemberobjects"  -body '{"securityEnabledOnly": false}'  ).value
            #>
            try   {
                if     ($Drive -and (ContextHas -WorkOrSchoolAccount)) {
                    Invoke-GraphRequest -Uri (
                                         $uri + '/Drive?$expand=root($expand=children)') -Exclude '@odata.context','root@odata.context' -As ([MicrosoftGraphDrive])}
                elseif ($Drive             ) {
                    Invoke-GraphRequest -Uri ($uri + '/Drive')                           -Exclude '@odata.context','root@odata.context' -As ([MicrosoftGraphDrive])}
                elseif ($LicenseDetails    ) {
                    Invoke-GraphRequest -Uri ($uri + '/licenseDetails')           -All                                                  -As ([MicrosoftGraphLicenseDetails]) }
                elseif ($MailboxSettings   ) {
                    Invoke-GraphRequest -Uri ($uri + '/MailboxSettings')                -Exclude '@odata.context'                       -As ([MicrosoftGraphMailboxSettings])}
                elseif ($OutlookCategories ) {
                    Invoke-GraphRequest -Uri ($uri + '/Outlook/MasterCategories') -All                                                  -As ([MicrosoftGraphOutlookCategory]) }
                elseif ($Photo             ) {
                    Invoke-GraphRequest -Uri ($uri + '/Photo')                          -Exclude '@odata.mediaEtag', '@odata.context',
                                                                                                              '@odata.mediaContentType' -As ([MicrosoftGraphProfilePhoto])}
                elseif ($PlannerTasks      ) {
                    Invoke-GraphRequest -Uri ($uri + '/planner/tasks')            -All  -Exclude '@odata.etag'                          -As ([MicrosoftGraphPlannerTask])}
                elseif ($Plans             ) {
                    Invoke-GraphRequest -Uri ($uri + '/planner/plans')            -All  -Exclude "@odata.etag"                          -As ([MicrosoftGraphPlannerPlan])}
                elseif ($Presence          )  {
                    Invoke-GraphRequest -Uri ($uri + '/presence')                       -Exclude "@odata.context"                       -As ([MicrosoftGraphPresence])}
                elseif ($Teams             ) {
                    Invoke-GraphRequest -Uri ($uri + '/joinedTeams')              -All                                                  -As ([MicrosoftGraphTeam])}
                elseif ($ToDoLists         ) {
                    Invoke-GraphRequest -Uri ($uri + '/todo/lists')               -All  -Exclude "@odata.etag"                          -As ([MicrosoftGraphTodoTaskList]) |
                      Add-Member -PassThru -NotePropertyName UserId -NotePropertyValue $id
                    }
                # Calendar wants a property added so we can find it again
                elseif ($Calendars         ) {
                    Invoke-GraphRequest -Uri ($uri + '/Calendars?$orderby=Name' ) -All                                                  -As ([MicrosoftGraphCalendar]) |
                        ForEach-Object {
                            if ($id -eq 'me') {$calpath = "me/Calendars/$($_.id)"}
                            else              {$calpath = "users/$id/calendars/$($_.id)"
                                               Add-Member -InputObject $_ -NotePropertyName User -NotePropertyValue $id
                            }
                            Add-Member -PassThru -InputObject $_ -NotePropertyName CalendarPath -NotePropertyValue $calpath |
                            Add-Member -PassThru -MemberType AliasProperty -Name   Calendar -Value ID
                        }
                }
                elseif ($Notebooks         ) {
                    $result += Invoke-GraphRequest -Uri ($uri +
                                          '/onenote/notebooks?$expand=sections' ) -All  -Exclude 'sections@odata.context'               -As ([MicrosoftGraphNotebook])
                    #Section fetched this way won't have parentNotebook, so make sure it is available when needed
                    foreach ($bookobj in $result) {
                        foreach ($s in $b.Sections) {
                                $s.parentNotebook.id          = $b.id
                                $s.parentNotebook.displayname = $b.displayname
                                $s.parentNotebook.self        = $b.self
                        }
                        $bookobj
                    }
                }
                # for site, get the user's MySite. Convert it into a graph URL and get that, expand drives subSites and lists, and add formatting types
                elseif ($Site              ) {
                        $response  = Invoke-GraphRequest -Uri ($uri + '?$select=mysite')
                        $uri       = $GraphUri + ($response.mysite -replace '^https://(.*?)/(.*)$', '/sites/$1:/$2?expand=drives,lists,sites')
                        $siteObj    = Invoke-GraphRequest $Uri                          -Exclude '@odata.context', 'drives@odata.context',
                                                                                           'lists@odata.context', 'sites@odata.context' -As ([MicrosoftGraphSite])
                        foreach ($l in $siteObj.lists) {
                            Add-Member -InputObject $l -MemberType NoteProperty   -Name SiteID   -Value  $siteObj.id
                            Add-Member -InputObject $l -MemberType ScriptProperty -Name Template -Value {$this.list.template}
                        }
                        $siteObj
                    }
                elseif ($Groups -or
                        $SecurityGroups   ) {
                    if  ($SecurityGroups)   {$body = '{  "securityEnabledOnly": true  }'}
                    else                    {$body = '{  "securityEnabledOnly": false }'}
                    $response         = Invoke-GraphRequest -Uri ($uri  + '/getMemberGroups') -Method POST  -Body $body -ContentType 'application/json'
                    foreach ($r in $response.value) {
                        $result     += Invoke-GraphRequest  -Uri "$GraphUri/directoryObjects/$r"
                    }
                }
                elseif ($DirectReports            ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/directReports')  -All       }
                elseif ($Manager                  ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/Manager') }
                elseif ($MemberOf                 ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '/MemberOf')  -All       }
                elseif ($Select                   ) {
                    $result += Invoke-GraphRequest -Uri ($uri + '?$select=' + ($Select -join ','))}
                else                                {
                    $result += Invoke-GraphRequest -Uri $uri  }
            }
            #if we get a not found error that's propably OK - bail for any other error.
            catch {
                if ($_.exception.response.statuscode.value__ -eq 404) {
                    Write-Warning -Message "'Not found' error while getting data for user '$userid'"
                }
                if ($_.exception.response.statuscode.value__ -eq 403) {
                    Write-Warning -Message "'Forbidden' error while getting data for user '$userid'. Do you have access to the correct scope?"
                }
                else {
                    Write-Progress -Activity 'Getting user information' -Completed
                    throw $_ ; return
                }
            }
             #endregion
        }
    }
    end     {
        Write-Progress -Activity 'Getting user information' -Completed
        foreach ($r in $result) {
           #if     ($r.'@odata.type' -match 'directoryRole$')  { $r.pstypenames.Add('GraphDirectoryRole') }
           #elseif ($r.'@odata.type' -match 'device$')         { $r.pstypenames.Add('GraphDevice')        }
           #else
            if     ($r.'@odata.type' -match 'group$') {
                    $r.remove('@odata.type')
                    $r.remove('@odata.context')
                    $r.remove('creationOptions')
                    New-Object -Property $r -TypeName ([MicrosoftGraphGroup])
            }
            elseif ($r.'@odata.type' -match 'user$' -or $PSCmdlet.parameterSetName -eq 'None' -or $Select) {
                    $r.Remove('@odata.type')
                    $r.Remove('@odata.context')
                    New-Object -Property $r -TypeName ([MicrosoftGraphUser])
            }
            else    {$r}
        }
    }
}

function Set-GraphUser         {
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
    [cmdletbinding(SupportsShouldprocess=$true)]
    param   (
        #ID for the user if not the current user
        [parameter(Position=1,ValueFromPipeline=$true)]
        $UserID = "me",
        #A freeform text entry field for the user to describe themselves.
        [String]$AboutMe,
        #The SMTP address for the user, for example, 'jeff@contoso.onmicrosoft.com'
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
        [Switch]$Force
    )
    begin   {
        #things we don't want to put in the JSON body when we send the changes.
        $excludedParams = [Cmdlet]::CommonParameters +  @('Photo','UserID','AccountDisabled', 'UsageLocation', 'Manager')
    }
    process {
        ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx todo check scopes  User.ReadWrite, User.ReadWrite.All, Directory.ReadWrite.All,        or Directory.AccessAsUser.All scope.

        #allow an array of users to be passed.
        foreach ($u in $UserID ) {
            #region configure the web parameters for changing the user. Allow for user objects with an ID or a UP
            $webparams = @{
                    'Method'            = 'PATCH'
                    'Contenttype'       = 'application/json'
            }
            if ($U -eq "me") {
                    $webparams['uri']   = "$Graphuri/me/"
            }
            elseif ($U.id)  {
                    $webparams['uri']   = "$Graphuri/users/$($U.id)/"
            }
            elseif ($U.UserPrincipalName) {
                    $webparams['uri']   = "$Graphuri/users/$($U.UserPrincipalName)/"
            }
            else {  $webparams['uri']   = "$Graphuri/users/$U/" }
            #endregion
            #region Convert Settings other than manager and Photo into a block of JSON and send it as a request body
            $settings = @{}
            foreach ($p in $PSBoundparameters.Keys.where({$_ -notin $excludedParams})) {
                #turn "Param" to "param" make dates suitable text, and switches booleans
                $key   = $p.toLower()[0] + $p.Substring(1)
                $value = $PSBoundparameters[$p]
                if ($value -is [datetime]) {$value = $value.ToString("yyyy-MM-ddT00:00:00Z")}  # 'o' for ISO date time may work here
                if ($value -is [switch])   {$value = $value -as [bool]}
                $settings[$key] = $value
            }
            if ($PSBoundparameters['AccountDisabled']) {$settings['accountEnabled'] = -not $AccountDisabled} #allows -accountDisabled:$false
         # if ($PSBoundparameters['UsageLocation'])   {$settings['usageLocation']  = $UsageLocation.ToUpper() } #Case matters now do this with a transformer attribute.
            if ($settings.count -eq 0 -and -not $Photo -and -not $Manager) {
                Write-Warning -Message "Nothing to set" ; continue
            }
            elseif ($settings.count -gt 0)  {
                $json = (ConvertTo-Json $settings) -replace '""' , 'null'
                Write-Debug  $json
                if ($Force -or $Pscmdlet.Shouldprocess($userID ,'Update User')) {Invoke-GraphRequest  @webparams -Body $json }
            }
            #endregion
            if ($Photo)   {
                if (-not (Test-Path $Photo) -or $photo -notlike "*.jpg" ) {
                    Write-Warning "$photo doesn't look like the path to a .jpg file" ; return
                }
                else {$photoPath = (Resolve-Path $Photo).Path }
                $BaseURI                    =  $webparams['uri']
                $webparams['uri']           =  $webparams['uri'] + 'photo/$value'
                $webparams['Method']        = 'Put'
                $webparams['Contenttype']   = 'image/jpeg'
                $webparams['InputFilePath'] =  $photoPath
                Write-Debug "Uploading Photo: '$photoPath'"
                if ($Force -or $Pscmdlet.Shouldprocess($userID ,'Update User')) {Invoke-GraphRequest  @webparams}
                $webparams['uri'] = $BaseURI
            }
            if ($Manager) {
                $BaseURI                    =  $webparams['uri']
                $webparams['uri']           =  $webparams['uri'] + 'manager/$ref'
                $webparams['Method']        = 'Put'
                $webparams['Contenttype']   = 'application/json'
                $json = ConvertTo-Json @{ '@odata.id' =  "$GraphUri/users/$manager" }
                Write-Debug  $json
                if ($Force -or $Pscmdlet.Shouldprocess($userID ,'Update User')) {Invoke-GraphRequest  @webparams -Body $json}
                $webparams['uri'] = $BaseURI
            }
        }
    }
}

function New-GraphUser         {
    <#
        .synopsis
            Creates a new user in Azure Active directory

    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification="False positive and need to support plain text here")]
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param (
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
        [Parameter(ParameterSetName='UPNFromDomainLast',Mandatory=$true)]
        [Parameter(ParameterSetName='UPNFromDomainDisplay',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ArgumentCompleter([DomainCompleter])]
        [string]$Domain,

        #The name displayed in the address book for the user. This is usually the combination of the user''s first name, middle initial and last name. This property is required when a user is created and it cannot be cleared during updates.
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [Parameter(ParameterSetName='UPNFromDomainDisplay',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNDisplay',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,

        #The given name (first name) of the user.
        [Parameter(ParameterSetName='UPNFromDomainDisplay')]
        [Parameter(ParameterSetName='DomainFromUPNDisplay')]
        [Parameter(ParameterSetName='UPNFromDomainLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Alias('FirstName')]
        [string]$GivenName,

        #User's last / family name
        [Parameter(ParameterSetName='UPNFromDomainDisplay')]
        [Parameter(ParameterSetName='DomainFromUPNDisplay')]
        [Parameter(ParameterSetName='UPNFromDomainLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Alias('LastName')]
        [string]$Surname,

        #A script block specifying how the displayname should be built, by default if is {"$GivenName $Surname"};
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$DisplayNameRule = {"$GivenName $Surname"},

        #A script block specifying how the mailnickname should be built, by default if is {"$GivenName.$Surname"};
        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$NickNameRule    = {"$GivenName.$Surname"},

        #A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries.  Examples include: 'US', 'JP', and 'GB'
        [ValidateNotNullOrEmpty()]
        [UpperCaseTransformAttribute()]
        [ValidateCountryAttribute()]
        [string]$UsageLocation = 'GB',

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
        [hashtable]$SetableProperties,

        #If Specified prevents any confirmation dialog from appearing
        [switch]$Force,

        #Unless passthru is specified, only passwords created when running the command are returned. When specified user objects are returned.
        [Alias('Pt')]
        [switch]$Passthru
    )
    #region we allow the names to be passed flexibly make sure we have what we need
    # Accept upn and display name -split upn to make a mailnickname, leave givenname/surname blank
    #        upn, display name, first and last
    #        mailnickname, domain, display name [first & last] - create a UPN
    #        domain, first & last - create a display name, and mail nickname, use the nickname in upn
    #re-create any scriptblock passed as a parameter, otherwise variables in this function are out of its scope.
    if ($NickNameRule)            {$NickNameRule      = [scriptblock]::create( $NickNameRule )   }
    if ($DisplayNameRule)         {$DisplayNameRule   = [scriptblock]::create( $DisplayNameRule) }
    #if we didn't get a display name build it
    if (-not $DisplayName)        {$DisplayName       = Invoke-Command -ScriptBlock $DisplayNameRule}
    #if we didn't get a UPN or a mail nickname, make the nickname first, then add the domain to make the UPN
    if (-not $UserPrincipalName -and
        -not $MailNickName  )     {$MailNickName      = Invoke-Command -ScriptBlock $NickNameRule
    }
    #if got a UPN but no nickname, split at @ to get one
    elseif ($UserPrincipalName -and
              -not $MailNickName) {$MailNickName      = $UserPrincipalName -replace '@.*$','' }
    #If we didn't get a UPN we should have a domain and a nickname, combine them
    if (($MailNickName -and $Domain) -and
         -not $UserPrincipalName) {$UserPrincipalName = "$MailNickName@$Domain"    }

    #We should have all 3 by now
    if (-not ($DisplayName -and $MailNickName -and $UserPrincipalName)) {
        throw "couldn't make sense of those parameters"
    }
    #A simple way to create one in 100K temporaty passwords. You might get 10Oct2126 - easy to type and meets complexity rules.
    if (-not $Initialpassword)    {
             $Initialpassword   = ([datetime]"1/1/1800").AddDays((Get-Random 146000)).tostring("ddMMMyyyy")
             Write-Output "$UserPrincipalName, $Initialpassword"
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
            if ($Passthru ) {return $u }
        }
        catch {
        # xxxx Todo figure out what errors need to be handled (illegal name, duplicate user)
        $_
        }
    }
}

function Remove-GraphUser      {
    <#
      .Synopsis
        Deletes a user from Azure Active directory
    #>
    [cmdletbinding(SupportsShouldprocess=$true,ConfirmImpact='High')]
    param (
        #ID for the user
        [parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        $UserID,
        #If specified the user is deleted without a confirmation prompt.
        [Switch]$Force
    )
    process{
       ContextHas -WorkOrSchoolAccount -BreakIfNot
        #xxxx todo check scopes

        #allow an array of users to be passed.
        foreach ($u in $UserID ) {
            if     ($u.displayName)       {$displayname = $u.displayname}
            elseif ($u.UserPrincipalName) {$displayName = $u.UserPrincipalName}
            else                          {$displayName = $u}
            if     ($u.id)                {$u =$U.id}
            elseif ($u.UserPrincipalName) {$u = $U.UserPrincipalName}
            if ($Force -or $pscmdlet.ShouldProcess($displayname,"Delete User")) {
                try {
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

function Find-GraphPeople      {
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
    param (
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
    begin {
    }
    process {
    #xxxx todo check scopes    Requires consent to use either the People.Read or the People.Read.All scope
        if ($Topic) {
            $uri = $GraphURI +'/me/people?$search="topic:{0}"&$top={1}' -f $Topic, $First
        }
        elseif ($SearchTerm) {
            $uri = $GraphURI + '/me/people?$search="{0}"&$top={1}' -f $SearchTerm, $First
        }

        Invoke-GraphRequest $uri -ValueOnly -As ([MicrosoftGraphPerson]) |
            Add-Member -PassThru -MemberType ScriptProperty -Name mobilephone    -Value {$This.phones.where({$_.type -eq 'mobile'}).number -join ', '} |
            Add-Member -PassThru -MemberType ScriptProperty -Name businessphones -Value {$This.phones.where({$_.type -eq 'business'}).number }         |
            Add-Member -PassThru -MemberType ScriptProperty -Name Score          -Value {$This.scoredEmailAddresses[0].relevanceScore }                |
            Add-Member -PassThru -MemberType AliasProperty  -Name emailaddresses -Value scoredEmailAddresses
    }
}

Function Import-GraphUser      {
<#
    .synopsis
       Imports a list of users from a CSV file
    .description
        Takes a list of CSV files and looks for xxxx columns
        * Action is either Add, Remove or Set - other values will cause the row to be ignored
        * DisplayName

#>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #One or more files to read for input.
        [Parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Disables any prompt for confirmation
        [switch]$Force,
        #Supresses output of Added, Removed, or No action messages for each row in the file.
        [switch]$Quiet,
        #Fields which are lists will be split at , or ; by default but a replacement split expression may be given
        [String]$ListSeparator = '\s*,\s*|\s*;\s*'
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

        foreach ($user in $list) {
            $upn = $user.UserPrincipalName
            if (-not $upn) {
                Write-Warning "User was missing a UPN"
                continue
            }
            else {
                 $exists = (Microsoft.Graph.Users.private\Get-MgUser_List -Filter "userprincipalName eq '$upn'") -as [bool]
            }

            if     ($user.Action -eq 'Remove' -and (-not $exists)) {
                Write-Warning "User '$upn' was marked for removal, but no matching user was found."
                continue
            }
            elseif ($user.Action -eq 'Remove' -and
                   ($force -or $PSCmdlet.ShouldProcess($upn,"Remove user "))){
                Remove-Graphuser -Force -user $user
                Write-Information "Removed user'$upn'"
                continue
            }

            if     ($user.Action -eq 'Add'    -and $exists) {
                    Write-Warning  "User '$upn' was marked for addition, but that name already exists."
                    continue
            }
            elseif ($user.Action -eq 'Add'    -and (-not $user.DisplayName) ) {
                Write-Warning "User was missing a UPN"
                continue
            }
            elseif ($user.Action -eq 'Add'    -and
                ($force -or $PSCmdlet.ShouldProcess($upn,"Add new user"))){
                $params = @{Force=$true; DisplayName=$user.DisplayName; UserPrincipalName= $user.UserPrincipalName;   }
                if ($user.MailNickName)      {$params['MailNickName']    = $user.MailNickName   }
                if ($user.GivenName)         {$params['GivenName']       = $user.GivenName      }
                if ($user.Surname)           {$params['Surname']         = $user.Surname        }
                if ($user.Initialpassword)   {$params['Initialpassword'] = $user.Initialpassword}
                if ($user.PasswordPolicies)  {$params['Initialpassword'] = $user.PasswordPolicies -split $ListSeparator}
                if ($user.NoPasswordChange -in @("Yes","True","1") ) {
                            {$params['NoPasswordChange'] = $true}
                }
                if ($user.ForceMFAPasswordChange -in @("Yes","True","1") ) {
                            {$params['ForceMFAPasswordChange'] = $true}
                }
                New-GraphUser @params
                Write-Information "Added user '$($user.DisplayName)' as '$upn'"
                $exists = $true
                $user.Action = "Set"
            }

            if     ($user.Action -eq 'Set'    -and (-not $exists)) {
                Write-Warning "User '$upn' was marked for update, but no matching user was found."
                continue
            }
            if     ($user.Action -eq 'Set') {
                $params = @{'UserId' = $upn}
                $Setparameters = (Get-Command Set-GraphUser ).Parameters.Values |
                    Where-Object name -notin (Cmdlet]::CommonParameters + [Cmdlet]::OptionalCommonParameters  )

                foreach ($p in $setparameters) {
                    $pName = $p.name
                    if     ($user.$pname -and ($p.parameterType -eq [string[]] )) {$params[$pName] = $user.$pName -split $ListSeparator  }
                    elseif ($user.$pname -and ($p.switchParameter))               {$params[$pName] = $user.$pName -in @("Yes","True","1")  }
                    elseif ($user.$pname)                                         {$params[$pName] = $user.$pName}
                }
                Set-GraphUser @params
                Write-Information "Updated properties of user '$upn'"
            }
        }
    }
}

Function Export-GraphUser      {
<#
    .synopsis
       Exports a list of users to a CSV file
#>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param (
        #Destination for CSV output
        [Parameter(Position=1,ValueFromPipeline=$true,Mandatory=$true)]
        $Path,
        #Filter clause for the query for example "department eq 'accounts'"
        $Filter,
        #String to insert between parts of multi-part items.
        $ListSeparator = "; "
    )
    Microsoft.Graph.Users.private\Get-MgUser_List -filter $Filter -ExpandProperty manager -Select 'UserPrincipalName',
                        'MailNickName','GivenName', 'Surname', 'DisplayName', 'UsageLocation', 'accountEnabled',
                        'PasswordPolicies', 'Mail',  'MobilePhone', 'BusinessPhones', 'JobTitle',  'Department',
                        'OfficeLocation', 'CompanyName','StreetAddress', 'City', 'State', 'Country',
                        'PostalCode' |
        Select-Object   'UserPrincipalName', 'MailNickName',   'GivenName', 'Surname',  'DisplayName', 'UsageLocation',
                        @{n='AccountDisabled';e={-not 'accountEnabled'}} , 'PasswordPolicies', 'Mail',  'MobilePhone',
                        @{n='BusinessPhones';e={$_.'BusinessPhones' -join $ListSeparator }},
                        @{n='Manager';e={$_.manager.AdditionalProperties.userPrincipalName}},
                        'JobTitle',  'Department', 'OfficeLocation', 'CompanyName',
                        'StreetAddress', 'City', 'State', 'Country', 'PostalCode' |
            Export-Csv -NoTypeInformation -Path $Path
}

#MailBox commands: these only depend on the user module from the SDK so go in the same file as user commands
function New-GraphMailAddress  {
    <#
      .synopsis
        Helper function to create a email addresses
    #>
    param (
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

function New-GraphRecipient    {
    <#
      .Synopsis
        Creats a new meeting attendee, with a mail address and the type of attendance.
    #>
    param(
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=1, ValueFromPipeline=$true)]
        $Mail,
        #The displayname for the recipient
        [Parameter(Position=2)]
        $DisplayName
    )
    @{ 'emailAddress' =  @{'address'=$mail; name=$DisplayName }}
}

function New-GraphAttendee     {
    <#
      .Synopsis
        Helper function to create a new meeting attendee, with a mail address and the type of attendance.
    #>
    [cmdletbinding(DefaultParameterSetName='Default')]
    [outputType([system.collections.hashtable])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state.')]
    param(
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
    #$EmailAddress = New-GraphMailAddress -Address $Address -DisplayName $DisplayName
    # New-Object -TypeName MicrosoftGraphAttendee -Property @{emailaddress=$EmailAddress ; Type=$AttendeeType}

    @{ 'type'= $AttendeeType ; 'emailAddress' =  @{'address'=$mail; name=$DisplayName }}
}

function New-GraphPhysicalAddress {
    <#
      .synopsis
        Builds a street / postal / physical address to use in the contact commands
       .Example
        >$fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
        Creates an address - if the -Street, City,  State, Postalcode country are not explictly
        specified they will be assigned in that order. Quotes are desireable but only necessary
        when a value contains spaces.
        It can then be used like this. Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
    #>
    [cmdletbinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state.')]
    param (
        #Street address. This can contain carriage returns for a district, e.g. "101 London Road`r`nBotley"
        [String]$Street,
        #City, or town as people outside the US tend to call it
        [String]$City,
        #State, Province, County, the administrative level below country
        [String]$State,
        #Postal code. Even it parses as a number, as with US ZIP codes, it will be converted to a string
        [String]$PostalCode,
        #Usually a country but could be some other geographical entity
        [String]$CountryOrRegion
    )
    $Address = @{}
    foreach ($P in $PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters})) {
        $Address[$p] + $PSBoundParameters[$p]
    }
    $Address
}

function New-GraphRecurrence   {
<#
    .synopsis
        Helper function to create the patterned recurrence for a task or event
    .links
        https://docs.microsoft.com/en-us/graph/api/resources/patternedrecurrence?view=graph-rest-1.0
#>
    param(
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
        $range =  New-Object -TypeName MicrosoftGraphRecurrenceRange -Property @{
            numberOfOccurrences = $numberOfOccurrences
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            endDate             = ($EndDate.ToString(  'yyyy-MM-dd') )
            recurrenceTimeZone  = $RecurrenceTimeZone
            type                = 'endDate'
        }
    }
    elseif ($numberOfOccurrences) {
        $range =  New-Object -TypeName MicrosoftGraphRecurrenceRange -Property @{
            numberOfOccurrences = $numberOfOccurrences
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            type                = 'numbered'
        }
    }
    else {
        $range =  New-Object -TypeName MicrosoftGraphRecurrenceRange -Property @{
            startDate           = ($startDate.ToString('yyyy-MM-dd') )
            type                = 'noEnd'
        }
    }
    $pattern = New-Object -TypeName MicrosoftGraphRecurrencePattern -Property @{
        dayOfMonth     = $DayOfMonth
        daysOfWeek     = $DaysOfWeek
        firstDayOfWeek = $FirstDayOfWeek
        index          = $Index
        interval       = $Interval
        month          = $month
        type           = $type
    }
    New-object -TypeName MicrosoftGraphPatternedRecurrence -Property @{
            pattern   = $pattern
            range     = $range
    }
}

function Expand-GraphEvent     {
    param   (
        [Parameter(Position=1,ValueFromPipeline=$true)]
        $Event,
        $CalendarPath

    )
    begin   {
        $whensb = {
            $s = [convert]::ToDateTime($this.Start.datetime)
            $e = [convert]::ToDateTime($this.end.datetime)
            if ($s.AddDays(1) -eq $e -and
                $s.hour -eq 0 -and $s.minute -eq 0 ) {
                $s.ToShortDateString() + ' All day'
            }
            else {$s.ToString("g") + ' to ' +  $e.ToString("g") + $this.End.timezone}
        }
    }
    process {
        if ($CalendarPath) {
            Add-Member -NotePropertyName CalendarPath   -NotePropertyValue $CalendarPath    -InputObject $Event
        }
        Add-Member -PassThru -MemberType ScriptProperty -Name When          -Value $whenSB  -InputObject $Event |
        Add-Member -PassThru -MemberType ScriptProperty -Name StartDateTime -Value {[convert]::ToDateTime($this.start.dateTime)}       |
        Add-Member -PassThru -MemberType ScriptProperty -Name EndDateTime   -Value {[convert]::ToDateTime($this.end.dateTime)}         |
        Add-Member -PassThru -MemberType ScriptProperty -Name Where         -Value {$this.location.displayname}
    }
}

function Get-GraphCalendarPath {
    param (
        $Calendar,
        $Group,
        $User
    )
    if     ($Calendar -and $Calendar.CalendarPath) {
        retrun $Calendar.CalendarPath
    }
    elseif ($Calendar -and  $User)      { #get a specific calendar for a specific user
        if ($User.ID)     {$user = $User.ID}
        if ($Calendar.id) {
            $path = "users/$User/calendars/$($Calendar.id)"
            Add-Member -Force -InputObject $Calendar -NotePropertyName CalendarPath -NotePropertyValue $path
            return $path
        }
        else {return "users/$user/calendars/$Calendar"}
    }
    elseif ($Calendar -and -not $group) { #if we got a calendar without a path or a user or group assume it is current user's
        if ($Calendar.id) {
            $path = "me/calendars/$($Calendar.id)"
            Add-Member -Force -InputObject $Calendar -NotePropertyName CalendarPath -NotePropertyValue $path
            return $path
        }
        else {return "me/calendars/$Calendar"}

    }
    elseif ($User)    {  # get the default calendar for a specific user
        if ($User.ID)    {return "users/$($user.ID)/calendar"}
        else             {return "users/$user/calendar"}
        #for the default calendar you can also use users/{id}/events or users//calendarView?param without "Calendar"
    }
    elseif ($Group)   {  # get the [only] calendar for a group
        if ($Group.ID)   {return "groups/$($Group.id)/calendar"}
        else             {return "groups/$Group/calendar" }  #for the default calendar you can also use groups/{id}/events or groups/calendarView?param without "Calendar"
    }
    else  {  #no User, group or calendar specified get the current users default calendar.
            return '/me/calendar'   #for the default calendar you can also use me/events or me/calendarView?params without "Calendar"
    }
}

function Get-GraphMailFolder   {
    <#
      .Synopsis
        Get the user's Mailbox folders
      .Example
        Get-GraphMailFolderList -Name inbox
        Gets the current users inbox folder
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$UserID,
        #Select the first n folders.
        [validaterange(1,1000)]
        [int]$Top,
        #fields to select in the query - will add a validate set later
        [string[]]$Select  ,
        #String with orderby clause e.g. "name", "lastmodifiedDate desc"
        [string]$OrderBy,
        #filter the folders returned by a name
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
        [string]$Name,
        #A custom filter clause.
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )

    #region set-up URI . If we got a user ID, use it other otherwise use the current user, add select, orderby, filter & top parameters as needed
    if ($UserID)  {$uri = "$GraphUri/users/$userID/mailFolders" }
    else          {$uri = "$GraphUri/me/mailFolders" }
    $JoinChar = "?"  #Will the next parameter be joined onto the URI with a "?"" or with "&"  ?
    if ($Select)  {$uri = $uri + '?$select=' + ($Select -join ',') ;                                 $JoinChar = "&"}
    if ($Name)    {$uri = $uri + $JoinChar + ("`$filter=startswith(displayName,'{0}') " -f $Name ) ; $JoinChar = "&"}
    if ($Filter)  {$uri = $uri + $JoinChar + '$Filter='  +$Filter                                  ; $JoinChar = "&"}
    if ($OrderBy) {$uri = $uri + $JoinChar + '$orderby='  +$Filter                                 ; $JoinChar = "&"}
    if ($Top)     {$uri = $uri + $JoinChar + '$top=' + $top }
    #endregion

    #region get the data, to keep the size attribute we will handle paging and converting to an object locally.
    $folderList    = @()
    $result       = Invoke-GraphRequest -Uri $uri
    $folderList   += $result.value
    while ($result.'@odata.nextLink') {
        $result          = Invoke-GraphRequest -Uri  $result.'@odata.nextLink' ;
        $folderList += $result.value
    }

    foreach ($f in $folderList) {
        $size = $f.sizeInBytes
        $f.remove('sizeInBytes')
        New-object -TypeName MicrosoftGraphMailFolder -Property $f |
            Add-Member -PassThru -NotePropertyName SizeInBytes -NotePropertyValue $size
        }
    #endregion\\
}

function Get-GraphMailItem     {
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
    param(
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$User ,
        #The ID of a folder, or one of the well known folder names 'archive', 'clutter', 'conflicts', 'conversationhistory', 'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures', 'msgfolderroot', 'outbox', 'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'sentitems', 'serverfailures', 'syncissues'
        [Parameter(ValueFromPipeline=$true)]
        $Mailfolder = "Inbox",
        #if specified the command will return child folders instead of messages
        [switch]$ChildFolders,
        #A term to do a free text search for in the mail box (see examples)
        [string]$Search,
        #If specified returns the top X items
        [int]$Top,
        #Sorting option, defaults to sorting by SentDateTime with newest first. Searches are not sorted.
        [string]$OrderBy ='SentdateTime desc',
        #Select particular mail fields , ignored if -ChildFolders is specified; defaults to From, Subject, SentDateTime, BodyPreview, and Weblink
        [ValidateSet('bccRecipients', 'body', 'bodyPreview', 'categories', 'ccRecipients', 'changeKey', 'conversationId', 'createdDateTime',
        'flag', 'from', 'hasAttachments', 'id', 'importance', 'inferenceClassification', 'internetMessageHeaders', 'internetMessageId',
        'isDeliveryReceiptRequested', 'isDraft', 'isRead', 'isReadReceiptRequested', 'lastModifiedDateTime', 'parentFolderId',
        'receivedDateTime', 'replyTo', 'sender', 'sentDateTime', 'subject', 'toRecipients', 'uniqueBody', 'webLink' )]
        [string[]]$Select = @('From', 'Subject', 'SentDatetime', 'BodyPreview', 'weblink'),
        #A Custom filter string; for example "importance eq high" - the examples have more cases
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )
    process {
        if     ($Mailfolder.id) {$MailPath = 'mailfolders/' +  $Mailfolder.id}
        elseif ($Mailfolder)    {$MailPath = 'mailfolders/' + ($Mailfolder -replace '^/','')}
        else                    {$MailPath = ''}

        if ($User.id) {$User  = $User.id}
        if ($User)    {$uri   = "$GraphUri/users/$user/$MailPath" }
        else          {$uri   = "$GraphUri/me/$MailPath" }

        if ($ChildFolders -and '' -ne $MailPath)    {
            Invoke-GraphRequest -Uri "$uri/childfolders" -ValueOnly -AsType ([MicrosoftGraphMailFolder])
        }
        elseif ($ChildFolders) {
            Write-Warning -Message 'You need to specify a folder when requesting child folders.'
        }
        else {
            $uri =  $uri + '/messages?$select='  + ($Select -join ',')
            if     ($Top)    {$uri = $uri + '&$top='     + $Top              }
            if     ($Search) {$Uri = $uri + '&$search="' + $Search + '"'     }
            elseif ($Filter) {$Uri = $uri + '&$filter='  + $Filter + ''      }
            else             {$uri = $uri + '&$orderby=' + $OrderBy          }


            Invoke-GraphRequest -Uri $uri -Headers @{'Prefer' ='outlook.body-content-type="text"'} -ValueOnly -AsType ([MicrosoftGraphMessage]) -ExcludeProperty '@odata.etag' |
                Add-Member -PassThru -MemberType ScriptProperty -Name "fromName"    -Value {$this.from.emailAddress.name} |
                Add-Member -PassThru -MemberType ScriptProperty -Name "fromAddress" -Value {$this.from.emailAddress.address} |
                Add-Member -PassThru -MemberType ScriptProperty -Name "bodyText"    -Value {$this.body.content}
        }
    }
}

function Send-GraphMailMessage {
    <#
      .Synopsis
        Sends Mail using the Graph API from the current user's mailbox.
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
    [Cmdletbinding(DefaultParameterSetName='None')]
    param (
        #Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.
        [parameter(Mandatory=$true,Position=0)]
        $To ,
        #Recipient(s) on the "CC" line,
        $CC  ,
        #Recipient(s) on the "Bcc line" line,
        $BCC,
        #The subject of the message. A message must have a subject and/or body and/or attachments. If the subject is left blank it will be sent as "No Subject"
        [String]$Subject,
        #The content of the message; assumed to be plain text, but HTML can be specified with -BodyType
        [String]$Body    ,
        #The type of the body  content. Possible values are Text and HTML.
        [ValidateSet("Text","HTML")]
        $BodyType = "Text",
        #The importance of the message: Low, Normal or High
        [ValidateSet('Low','Normal', 'High')]
        $Importance = 'Normal' ,
        #Path to file(s) to send as attachments
        $Attachments,
        #If Specified, requests a receipt.
        [switch]$Receipt,
        #If specified leaves the message in the drafts folder without sending it and returns a link to open the message.
        [parameter(ParameterSetName='SaveDraftOnly',Mandatory=$true)]
        [switch]$SaveDraftOnly,
        #If specified specifies that a copy of the mail should not be saved
        [parameter(ParameterSetName='NoSave',Mandatory=$true)]
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
            if ($Attachments.Where({$_.length -gt 2.85mb}))  {
                #The Maximum size for a POST is 4MB.
                #Attachments are base 64 encoded so 3MB of attachements become 4MB. Don't try closer than 95% of that
                throw ("Attachment would exceed maximum size for a POST. Maximum file size is ~ 2,900,000 bytes")
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
                    $asDraft= $true
                }
            }
            else { Write-Verbose -Message "SEND-GRAPHMAILMESSAGE $($Attachments).count attachment(s); small enough to send in a single operation"}
         }
    }
    elseif (-not $Subject -and -not $Body) {
        Write-Warning -Message "Nothing to send" ; return
    }
    elseif (-not $Subject) {$Subject = "No subject"}

    if ($asDraft) {$Uri = "$GraphUri/me/Messages"}
    else          {$Uri = "$GraphUri/me/sendmail"}

    #Build a hash table with the parts of the message, this will be coverted into JSON
    #BEWARE names are case sensitive. if you create $msgSettings.Body instead of $msgSettings.body
    #the capital B will cause a 400 bad request error.
    #My personal coding style is to use inital CAPS for parameters and inital lower case for variables (though Powershell doesn't care)
    #so the parameter is $Body and the hash table key name and JSON label is body.

    $msgSettings   =  @{   'body' = @{
            'contentType'  = $BodyType;
                'content'  = $Body}
                'subject'  = $Subject
             'importance'  = $Importance
            'toRecipients' = @()
    }
    foreach ($recip in $To ) {
            if     ($recip  -is [string] ) { $msgSettings[ 'toRecipients'] += New-GraphRecipient $recip}
            else                           { $msgSettings[ 'toRecipients'] += $recip}
    }
    if ($CC) {
        $msgSettings['ccRecipients']      = @()
        foreach ($recip in $cc ) {
            if     ($recip  -is [string] ) { $msgSettings[ 'ccRecipients'] += New-GraphRecipient $recip}
            else                           { $msgSettings[ 'ccRecipients'] += $recip}}
    }
    if ($BCC) {
        $msgSettings['bccRecipients']      = @()
        foreach ($recip in $bcc ) {
            if     ($recip  -is [string] ) { $msgSettings['bccRecipients'] += New-GraphRecipient $recip}
            else                           { $msgSettings['bccRecipients'] += $recip}}
    }
    if ($Receipt)                          { $msgSettings['isDeliveryReceiptRequested'] = $true }

    #If we are creating a draft, save it now; if sending-in-one be ready for attachments
    if     ($asDraft) {
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading draft"
        $json = ConvertTo-Json $msgSettings -Depth 5 #default depth isn't enough !
        try            {$msg  = Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch          {throw "There was an error creating the draft message."; return }
        if (-not $msg) {throw "The draft message was not created as expected" ; return }
        else           {
            Write-Verbose -Message "SEND-GRAPHMAILMESSAGE  Message created with id '$($msg.id)'"
            $uri = $uri + "/" + $msg.id
        }
    }
    elseif ($AttachmentItems) {
        $msgSettings["attachments"]= @()
    }

    foreach ($f in $AttachmentItems) {
        $Filesettings = @{
            '@odata.type' = '#microsoft.graph.fileAttachment';
            name          = $f.Name ;
            contentId     = $f.name ;
            contentBytes  =  [convert]::ToBase64String( [system.io.file]::readallbytes($f.FullName))

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
            catch {throw "There was an error sending the draft message; it remains in the drafts folder"}
    }
    else {
        $mail = @{Message=$msgSettings}
        if ($NoSave) {
           $mail['saveToSentItems'] = $false
        }
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading and sending"

        $json = ConvertTo-Json $mail -Depth 10
        Write-Debug $Json
        try            {Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch          {throw "There was an error sending message."; return }
        Write-Progress -Activity "Sending Message" -Completed
    }
}

function Send-GraphMailReply   {
    <#
      .synopsis
        Replies to a mail message.
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param (
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

function Send-GraphMailForward {
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
    param (
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

function Get-GraphContact      {
    <#
      .Synopsis
        Get the user's contacts
      .Example
        get-graphContacts -name "o'neill" | ft displayname, mobilephone
        Gets contacts where the display name, given name, surname, file-as name, or email begins with
        O'Neill - note the function handles apostrophe, - a single one would normal cause an error with the query.
        The results are displayed as table with display name and mobile number
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param(
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$UserID,
        #If specified selects the first n contacts
        [int]$Top,
        #A custom set of contact properties to select
        [ValidateSet('assistantName', 'birthday', 'businessAddress', 'businessHomePage', 'businessPhones',
                     'categories', 'changeKey', 'children', 'companyName', 'createdDateTime', 'department',
                     'displayName', 'emailAddresses', 'fileAs', 'generation', 'givenName', 'homeAddress',
                     'homePhones', 'id', 'imAddresses', 'initials', 'jobTitle', 'lastModifiedDateTime',
                     'manager', 'middleName', 'mobilePhone',  'nickName', 'officeLocation', 'otherAddress',
                     'parentFolderId', 'personalNotes', 'profession', 'spouseName', 'surname', 'title',
                     'yomiCompanyName', 'yomiGivenName', 'yomiSurname')]
        [string[]]$Select,
        #A custom OData Sort string.
        [string]$OrderBy,
        #If specified looks for contacts where the display name, file-as Name, given name or surname beging with ...
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
        [string]$Name,
        #A custom OData Filter String
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )

    #region build the URI - if we got a user ID, use it, add select, filter, orderby and/or top as needed
    if     ($UserID.id) {$uri = "$GraphUri/users/$($userID.id)/contacts"}
    elseif ($UserID)    {$uri = "$GraphUri/users/$userID/contacts" }
    else                {$uri = "$GraphUri/me/contacts" }

    $JoinChar = "?" #will next parameter be added to the URI with a "?" or a "&" ?
    if ($Select)  { $uri = $uri + '?$select=' + ($Select -join ',')    ; $JoinChar = "&" }
    if ($Name)    { $uri = $uri + $JoinChar + ("`$filter=startswith(displayName,'{0}') or startswith(givenName,'{0}') or startswith(surname,'{0}')  or startswith(fileAs,'{0}')" -f ($Name -replace "'","''" )  )
                  $JoinChar = "&"
    }
    if ($Filter)  { $uri = $uri + $JoinChar + '$Filter='  + $Filter  ; $JoinChar = "&" }
    if ($OrderBy) { $uri = $uri + $JoinChar + '$orderby=' + $orderby ; $JoinChar = "&" }
    if ($Top)     { $uri = $uri + $JoinChar + '$top='     + $top}
    #endregion

    #region get the data - cope with it being paged - add a type for fomatting, and return it
    $defaultProperties = @('displayname','jobtitle','companyname','mail','mobile','business','home')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    Invoke-GraphRequest -Uri  $uri -ValueOnly -AllValues -AsType ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact]) -ExcludeProperty "@odata.etag" |
        Add-Member -PassThru -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers        |
        Add-Member -PassThru -MemberType AliasProperty  -Name mobile            -Value 'mobilephone'             |
        Add-Member -PassThru -MemberType ScriptProperty -Name business          -Value {$this.businessPhones[0]} |
        Add-Member -PassThru -MemberType ScriptProperty -Name home              -Value {$this.HomePhones[0]}
    #endregion
}

function New-GraphContact      {
    <#
      .Synopsis
        Adds an entry to the current users Outlook contacts
      .Description
        Almost all the paramters can be accepted form a piped object to make import easier.
       .Example
       >New-GraphContact -GivenName Pavel -Surname Bansky -Email pavelb@fabrikam.onmicrosoft.com -BusinessPhones  "+1 732 555 0102"
       Creates a new contact; if no displayname is given, one will be decided using given name and suranme;
       .Example
       >
       >$PavelMail = New-GraphRecipient -DisplayName "Pavel Bansky [Fabikam]" -Mail  pavelb@fabrikam.onmicrosoft.com
       >New-GraphContact -GivenName Pavel -Surname Bansky -Email $pavelmail  -BusinessPhones  "+1 732 555 0102"
        This creates the same contanct but sets up their email with a display name.
        New recipient creates a hash table
        @{'emailaddress' = @ {
                'name' = 'Pavel Bansky [Fabikam]'
                'address' = 'pavelb@fabrikam.onmicrosoft.com'
            }
        }
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param   (
        [Parameter(ValueFromPipelineByPropertyName)]
        $GivenName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $MiddleName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Initials ,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Surname,
        [Parameter(ValueFromPipelineByPropertyName)]
        $NickName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $FileAs,
        [Parameter(ValueFromPipelineByPropertyName)]
        $DisplayName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $CompanyName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $JobTitle,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Department,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Manager,
        #One or more instant messaging addresses, as a single string with semi colons between addresses or as an array of strings or MailAddress objects created with New-GraphMailAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $Email,
        #One or more instant messaging addresses, as an array or as a single string with semi colons between addresses
        [Parameter(ValueFromPipelineByPropertyName)]
        $IM,
        #A single mobile phone number
        [Parameter(ValueFromPipelineByPropertyName)]
        $MobilePhone,
        #One or more Business phones either as an array or as single string with semi colons between numbers
        [Parameter(ValueFromPipelineByPropertyName)]
        $BusinessPhones,
        #One or more home phones either as an array or as single string with semi colons between numbers
        [Parameter(ValueFromPipelineByPropertyName)]
        $HomePhones,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $Homeaddress,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $BusinessAddress,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $OtherAddress,
        #One or more categories either as an array or as single string with semi colons between them.
        [Parameter(ValueFromPipelineByPropertyName)]
        $Categories,
        #The contact's Birthday as a date
        [Parameter(ValueFromPipelineByPropertyName)]
        [dateTime]$Birthday ,
        [Parameter(ValueFromPipelineByPropertyName)]
        $PersonalNotes,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Profession,
        [Parameter(ValueFromPipelineByPropertyName)]
        $AssistantName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Children,
        [Parameter(ValueFromPipelineByPropertyName)]
        $SpouseName,
        #If sepcified the contact will be created without prompting for confirmation. This is the default state but can change with the setting of confirmPreference.
        [Switch]$Force
    )

    process {
        Set-GraphContact @PSBoundParameters -IsNew
    }
}

function Set-GraphContact      {
    <#
      .Synopsis
        Modifies or adds an entry in the current users Outlook contacts
      .Example
        >
        > $pavel = Get-GraphContact -Name pavel
        > Set-GraphContact $pavel -CompanyName "Fabrikam" -Birthday "1974-07-22"
        The first line gets the Contact which was added in the 'New-GraphContact" example
        and the second adds Birthday and Company-name attributes to the contact.
       .Example
        >
        > $fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
        > Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
        This continues from the previous example, creating an address in the first line
        and adding it to the contact in the second.

    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param   (
    #The contact to be updated either as an ID or as contact object containing an ID.
    [Parameter(ValueFromPipeline=$true,ParameterSetName='UpdateContact',Mandatory=$true, Position=0 )]
    $Contact,
    #If specified, instead of providing a contact, instructs the command to create a contact instead of updating one.
    [Parameter(ParameterSetName='NewContact',Mandatory=$true )]
    [switch]$IsNew,
    [Parameter(ValueFromPipelineByPropertyName)]
    $GivenName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $MiddleName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Initials ,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Surname,
    [Parameter(ValueFromPipelineByPropertyName)]
    $NickName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $FileAs,
    #If not specified a display name will be generated, so updates without the display name may result in overwriting an existing one
    [Parameter(ValueFromPipelineByPropertyName)]
    $DisplayName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $CompanyName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $JobTitle,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Department,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Manager,
    #One or more mail addresses, as a single string with semi colons between addresses or as an array of strings or MailAddress objects created with New-GraphMailAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $Email,
    #One or more instant messaging addresses, as an array or as a single string with semi colons between addresses
    [Parameter(ValueFromPipelineByPropertyName)]
    $IM,
    #A single mobile phone number
    [Parameter(ValueFromPipelineByPropertyName)]
    $MobilePhone,
    #One or more Business phones either as an array or as single string with semi colons between numbers
    [Parameter(ValueFromPipelineByPropertyName)]
    $BusinessPhones,
    #One or more home phones either as an array or as single string with semi colons between numbers
    [Parameter(ValueFromPipelineByPropertyName)]
    $HomePhones,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $Homeaddress,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $BusinessAddress,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $OtherAddress,
    #One or more categories either as an array or as single string with semi colons between them.
    [Parameter(ValueFromPipelineByPropertyName)]
    $Categories,
    #The contact's Birthday as a date
    [Parameter(ValueFromPipelineByPropertyName)]
    [nullable[dateTime]]$Birthday ,
    [Parameter(ValueFromPipelineByPropertyName)]
    $PersonalNotes,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Profession,
    [Parameter(ValueFromPipelineByPropertyName)]
    $AssistantName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Children,
    [Parameter(ValueFromPipelineByPropertyName)]
    $SpouseName,
    #If sepcified the contact will be created without prompting for confirmation. This is the default state but can change with the setting of confirmPreference.
    [Switch]$Force
    )
    begin   {
        $webParams = @{
            'ContentType'    = 'application/json'
            'URI'             = "$GraphUri/me/contacts"
            'AsType'          =  ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])
            'ExcludeProperty' = @('@odata.etag', '@odata.context' )
        }
        $defaultProperties = @('displayname','jobtitle','companyname','mail','mobile','business','home')
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
        $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    }
    process {
        $contactSettings = @{  }
        if ($Email)                           {$contactSettings['emailAddresses'] = @() }
        if ($Email -is [string])              {$Email = $Email -split '\s*;\s*'}
        foreach ($e in $Email) {
            if     ($e.emailAddress)          {$contactSettings.emailAddresses    += $e.emailAddress   }
            elseif ($e -is [string])          {$contactSettings.emailAddresses    += @{'address' = $e} }
            else                              {$contactSettings.emailAddresses    += $e  }
        }
        if     ($IM             -is [string]) {$contactSettings['imAddresses']     = @() + $IM             -split '\s*;\s*'}
        elseif ($IM                         ) {$contactSettings['imAddresses']     =       $IM}
        if     ($Categories     -is [string]) {$contactSettings['categories']      = @() + $Categories     -split '\s*;\s*'}
        elseif ($Categories                 ) {$contactSettings['categories']      =       $Categories}
        if     ($Children       -is [string]) {$contactSettings['children']        = @() + $Children       -split '\s*;\s*'}
        elseif ($Children                   ) {$contactSettings['children']        =       $Children}
        if     ($BusinessPhones -is [string]) {$contactSettings['businessPhones']  = @() + $BusinessPhones -split '\s*;\s*'}
        elseif ($BusinessPhones             ) {$contactSettings['businessPhones']  =       $BusinessPhones}
        if     ($HomePhones     -is [string]) {$contactSettings['homePhones']      = @() + $HomePhones     -split '\s*;\s*'}
        elseif ($HomePhones                 ) {$contactSettings['homePhones']      =       $HomePhones  }
        if     ($MobilePhone                ) {$contactSettings['mobilePhone']     =       $MobilePhone}
        if     ($GivenName                  ) {$contactSettings['givenName']       =       $GivenName}
        if     ($MiddleName                 ) {$contactSettings['middleName']      =       $MiddleName}
        if     ($Initials                   ) {$contactSettings['initials']        =       $Initials}
        if     ($Surname                    ) {$contactSettings['surname']         =       $Surname}
        if     ($NickName                   ) {$contactSettings['nickName']        =       $NickName}
        if     ($FileAs                     ) {$contactSettings['fileAs']          =       $FileAs}
        if     ($DisplayName                ) {$contactSettings['displayName']     =       $DisplayName}
        if     ($Manager                    ) {$contactSettings['manager']         =       $Manager}
        if     ($JobTitle                   ) {$contactSettings['jobTitle']        =       $JobTitle}
        if     ($Department                 ) {$contactSettings['department']      =       $Department}
        if     ($CompanyName                ) {$contactSettings['companyName']      =      $CompanyName}
        if     ($PersonalNotes              ) {$contactSettings['personalNotes']   =       $PersonalNotes}
        if     ($Profession                 ) {$contactSettings['profession']      =       $Profession}
        if     ($AssistantName              ) {$contactSettings['assistantName']   =       $AssistantName}
        if     ($Children                   ) {$contactSettings['children']        =       $Children}
        if     ($SpouseName                 ) {$contactSettings['spouseName']      =       $spouseName}
        if     ($Homeaddress                ) {$contactSettings['homeaddress']     =       $Homeaddress}
        if     ($BusinessAddress            ) {$contactSettings['businessAddress'] =       $BusinessAddress}
        if     ($OtherAddress               ) {$contactSettings['otherAddress']    =       $OtherAddress}
        if     ($Birthday                   ) {$contactSettings['birthday']        =       $Birthday.tostring('yyyy-MM-dd')} #note this is a different date format to most things !

        $json = ConvertTo-Json $contactSettings
        Write-Debug $json
        if ($IsNew) {
            if ($force -or $PSCmdlet.ShouldProcess($DisplayName,'Create Contact')) {
                Invoke-GraphRequest @webParams -method Post  -Body $json  |
                    Add-Member -PassThru -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers        |
                    Add-Member -PassThru -MemberType AliasProperty  -Name mobile            -Value 'mobilephone'             |
                    Add-Member -PassThru -MemberType ScriptProperty -Name business          -Value {$this.businessPhones[0]} |
                    Add-Member -PassThru -MemberType ScriptProperty -Name home              -Value {$this.HomePhones[0]}
            }
        }
        else {#if ContactPassed
            if ($force -or $PSCmdlet.ShouldProcess($Contact.DisplayName,'Update Contact')) {
                if ($Contact.id) {$webParams.uri += '/' + $Contact.ID}
                else             {$webParams.uri += '/' + $Contact }
                Invoke-GraphRequest @webParams -method Patch -Body $json |
                    Add-Member -PassThru -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers        |
                    Add-Member -PassThru -MemberType AliasProperty  -Name mobile            -Value 'mobilephone'             |
                    Add-Member -PassThru -MemberType ScriptProperty -Name business          -Value {$this.businessPhones[0]} |
                    Add-Member -PassThru -MemberType ScriptProperty -Name home              -Value {$this.HomePhones[0]}
            }
        }
    }
}

function Remove-GraphContact   {
    <#
      .synopsis
         Deletes a contact from the default user's contacts
      .Example
        > Get-GraphContact -Name pavel | Remove-GraphContact
        Finds and removes any user whose given name, surname, email or display name
        matches Pavel*. This might return unexpected users, fortunately there is a prompt
        before deleting - the prompt it can be supressed by using the -Force switch if you
        are confident you have the right contact selected.
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param (
        #The contact to remove, as an ID or as a contact object containing an ID
        [parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true )]
        $Contact,
        #If specified the contact will be removed without prompting for confirmation
        $Force
    )
    begin {

    }
    process {
        if ($force -or $pscmdlet.ShouldProcess($Contact.DisplayName, 'Delete contact')) {
            if ($Contact.id) {$Contact = $Contact.id}
            Invoke-GraphRequest -Method Delete -uri "$GraphUri/me/contacts/$Contact"
        }
    }
}

#Outlook calendar - also only needs items found in the user module, so we don't give it it's own PS1 file
function Get-GraphEvent        {
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
        >$team = Get-GraphTeam -ByName consultants
        >Get-GraphEvent -Team $team
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
        and one of their propperties, as in this case, may need to be specified to perform
        a sort, and the syntax is property/ChildProperty.
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
    param (
        #UserID as a guid or User Principal name, whose calendar should be fetched.
        [Parameter( Mandatory=$true, ParameterSetName="User"          ,ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="UserAndSubject",ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="UserAndFilter" ,ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar
        [Parameter( Mandatory=$true, ParameterSetName="Cal",           ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="CalAndSubject", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="CalAndFilter",  ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( ParameterSetName="User",          ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( ParameterSetName="UserAndSubject",ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( ParameterSetName="UserAndFilter", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID, whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="GroupID"        ,ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndSubject",ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndFilter" ,ValueFromPipelineByPropertyName=$true)]
        [Alias("Team")]
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
        [Parameter(Mandatory=$true, ParameterSetName='CalAndSubject',  ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="UserAndSubject", ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndSubject",ValueFromPipelineByPropertyName=$true)]
        [string]$Subject,

        #A custom selection filter
        [Parameter(Mandatory=$true, ParameterSetName="CurrentFilter", ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="CalAndFilter",  ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="UserAndFilter", ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndFilter",ValueFromPipelineByPropertyName=$true)]
        [string]$Filter
    )

    begin {
        $webParams = @{
            'AllValues'       = $true
            'ValueOnly'       = $true
            'ExcludeProperty' = 'icaluid','@odata.etag','calendar@odata.navigationLink','calendar@odata.associationLink'
            'AsType'          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
        }
        if ($TimeZone) {$webParams['Headers'] =@{"Prefer"="Outlook.timezone=""$TimeZone"""}}
    }
    Process {
        $CalendarPath = Get-GraphCalendarPath -Calendar $Calendar -Group $Group -User $User
        $uri          = "$GraphUri/$CalendarPath"
        #region apply the selection criteria. If -days is specified use calendar view, otherwise use events and add filter, orderby, select and top as needed
        if  ($days)    {
                        $start = [datetime]::Today.ToString("yyyy-MM-dd't'HH:mm:ss")       # 'o' for ISO format time may work here.
                        $end   = [datetime]::Today.AddDays($days).tostring("yyyy-MM-dd't'HH:mm:ss")
                        $uri  += "/calendarview?`$expand=calendar&startdatetime=$start&enddatetime=$end"
        }
        else {             $uri  +=  '/events?$expand=calendar'}

        if ($Select)     { $uri  +=  '&$select=' + ($Select -join ',') }

        if ($Subject)    { $uri  += ('&$filter=startswith(subject,''{0}'')' -f $subject ) }
        elseif ($Filter) { $uri  +=  '&$Filter='  + $Filter }

        if ($OrderBy)    { $uri  +=  '&$orderby=' + $orderby }

        if ($Top)        { $uri  +=  '&$top='     + $top  }
        #endregion
        #region get the data.
        Invoke-GraphRequest @webParams -Uri $uri | Expand-GraphEvent -CalendarPath $CalendarPath
        #endregion
    }
}

function Add-GraphEvent        {
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
        Creates a meeting with a two additonal attendee. The first command creates an optional attendee with a display name
        the second creates an attendee with no displayed name and the default 'required' type
        Finally the meeting is created.
    #>
    [cmdletbinding()]
    param (
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
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
    begin {
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
        $webParams['uri'] = $GraphUri + $CalendarPath + '/events'

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
        if ($PassThru) {$result | Expand-GraphEvent -CalendarPath $CalendarPath }
    }
}

function Set-GraphEvent        {
    <#
      .Synopsis
        Modifies an event on a calendar
      .link
        Get-GraphEvent
      .Example
        TBC
    #>
    [cmdletbinding(SupportsShouldProcess=$true,DefaultParameterSetName='None')]
    param (
        #The event to be updateds either as an ID or as an event object containing an ID.
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)]
        $Event,

        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
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
        if ($PassThru) {$result |  Expand-GraphEvent -CalendarPath $CalendarPath}
     }
}

function Remove-GraphEvent     {
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

#To-do-list functions are here because they are in the Users.private module, not a module of their own
# they require the Tasks.ReadWrite  scope

function Get-GraphToDoList     {
    <#
      .Synopsis
        Gets information about lists used in the To Do app.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param   (
        #The ID of the plan or a plan object with an ID property. if omitted the current users planner will be assumed.
        [Parameter( ValueFromPipeline=$true,Position=0)]
        [alias('id')]
        $TodoTaskListId = 'defaultList',

        #The User ID (GUID or UPN) of the list owner. Defaults to the current user.
        $UserId,

        #If specified returns the tasks in the list.
        [switch]$Tasks
    )
    process {
        contexthas -WorkOrSchoolAccount -BreakIfNot
        if ($UserId) {$uri    = "$GraphUri/users/$userid/todo/lists"}
        else         {$uri    = "$GraphUri/me/todo/lists"
                      $UserId =  $global:GraphUser
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
            if (-not $UserId) {$UserId = $global:GraphUser }
            Invoke-GraphRequest  -Method get  -uri "$uri/$($ToDoList.id)/tasks" -ValueOnly -ExcludeProperty "@odata.etag" -AsType ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphTodoTask]) |
                Add-Member -PassThru -NotePropertyName UserId   -NotePropertyValue $userID |
                Add-Member -PassThru -NotePropertyName ListID   -NotePropertyValue $ToDoList.Id |
                Add-Member -PassThru -NotePropertyName ListName -NotePropertyValue $ToDoList.DisplayName
        }
    }
}

function New-GraphToDoList     {
<#
    .synopsis
        Creates a new list for the To-Do app
#>
[cmdletBinding(SupportsShouldProcess=$true)]
Param(
    [parameter(Mandatory=$true,Position=1)]
    #The name for the list
    [string]$Displayname    ,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user,
    $UserId =  $global:GraphUser,

    #If specified the the list will be created as a shared list
    [switch]$IsShared,

    #If specified any confirmation will be supressed
    [switch]$Force
)
    if ($Force -or $pscmdlet.ShouldProcess($Displayname,"Create new To-Do list")){
        Microsoft.Graph.Users.private\New-MgUserTodoList_CreateExpanded -UserId $UserId -DisplayName $displayname -IsShared:$IsShared -Confirm:$false |
                Add-Member -PassThru -NotePropertyName UserId -NotePropertyValue $UserId
    }
}

function New-GraphToDoTask     {
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param (

        #A To-do list object or the ID of a To-do list
        [Parameter()]
        [alias('TodoTaskListId')]
        $ToDoList,

        #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on theToDo list object
        [Parameter()]
        [string]
        $UserId =  $global:GraphUser,

        # A brief description of the task.
        [Parameter(mandatory=$true, position=1)]
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
        Microsoft.Graph.Users.private\New-MgUserTodoListTask_CreateExpanded @params |
            Add-Member -PassThru -NotePropertyName UserId   -NotePropertyValue $userID |
            Add-Member -PassThru -NotePropertyName ListID   -NotePropertyValue $ToDoList
    }
}

function Update-GraphToDoTask  {
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param (
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
    [string]$UserId =  $global:GraphUser,

    # A brief description of the task.
    [Parameter(position=1)]
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
            Microsoft.Graph.Users.private\Update-MgUserTodoListTask_UpdateExpanded @params
        }
    }
}

function Remove-GraphToDoTask  {
    <#
        .synopsis
            Removes a task from the To Do app
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    Param (
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
    [string]$UserId =  $global:GraphUser,

    #If specified, no confirmation will be displayed before deleting the task
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
        if ($Task.Title)       {$Title    = $Task.title}
        if ($Task.id)          {$Task     = $Task.id}

        $Params =  @{
            TodoTaskId      = $Task
            UserId          = $UserId
            TodoTaskListId  = $ToDoList
        }
        if ($force -or $pscmdlet.ShouldProcess($Title,'Task deletion')) {
                Microsoft.Graph.Users.private\Remove-MgUserTodoListTask_Delete @Params
        }
    }
}

function Remove-GraphToDoList  {
    <#
        .synopsis
            Removes a list from the To Do app, including any task in contains.
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    Param (
    #A To-do list object or the ID of a To-do list
    [Parameter(mandatory=$true,ValueFromPipelineByPropertyName =$true, ValueFromPipeline=$true)]
    [alias('TodoTaskListId','ListID')]
    $ToDoList,

    #The User ID (GUID or UPN) of the list owner. Defaults to the current user, and may be found on the ToDo list object
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$UserId =  $global:GraphUser,

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
                Microsoft.Graph.Users.private\Remove-MgUserTodoList_Delete @Params
        }
    }
}
