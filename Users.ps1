using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Globalization

class UpperCaseTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute  {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$EngineIntrinsics, [object] $InputData) {
        if ($inputData -is [string]) {return $Inputdata.toUpper()}
        else                         {return ($InputData) }
    }
}

class ValidateCountryAttribute : ValidateArgumentsAttribute {
    [void]Validate([object]$Argument, [EngineIntrinsics]$EngineIntrinsics)  {
        if ($Argument -notin [cultureInfo]::GetCultures("SpecificCultures").foreach({
                                New-Object -TypeName RegionInfo -ArgumentList $_.name
                             }).TwoLetterIsoRegionName) {
            Throw [ParameterBindingException]::new("'$Argument' is not an ISO 3166 country Code")
        }
    }
}

function New-GraphRecurrence             {
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

function New-GraphMailAddress            {
    param (
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        [Alias('Mail')]
        [String]$Address,
        #The displayname for the recipient
        [Alias('DisplayName')]
        $Name
    )
    New-Object -TypeName MicrosoftGraphEmailAddress -Property $PSBoundParameters
}

function New-GraphAttendee               {
    <#
      .Synopsis
        Creats a new meeting attendee, with a mail address and the type of attendance.
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
    $EmailAddress = New-MailAddress -Address $Address -DisplayName $DisplayName
    New-Object -TypeName MicrosoftGraphAttendee -Property @{emailaddress=$EmailAddress ; Type=$AttendeeType}

    @{ 'type'= $AttendeeType ; 'emailAddress' = (New-MailAddress -Mail:$mail -DisplayName:$DisplayName )}
}

function Expand-GraphEvent {
    param (
        [Parameter(Position=1,ValueFromPipeline=$true)]
        $Event,
        $CalendarPath

    )
    begin {
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
    param (
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
    begin {

        #things we don't want to put in the JSON body when we send the changes.
        $excludedParams = [Cmdlet]::CommonParameters +  @('Photo','UserID','AccountDisabled', 'UsageLocation', 'Manager')
    }

    Process {
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
                $key   = $p.toLower()[0] + $p.Substring(1)
                $value = $PSBoundparameters[$p]
                if ($value -is [datetime]) {$value = $value.ToString("yyyy-MM-ddT00:00:00Z")}  # 'o' for ISO date time may work here
                if ($value -is [switch])   {$value = $value -as [bool]}
                $settings[$key] = $value
            }
            if ($PSBoundparameters['AccountDisabled']) {$settings['accountEnabled'] = -not $AccountDisabled} #allows -accountDisabled:$false
            if ($PSBoundparameters['UsageLocation'])   {$settings['usageLocation']  = $UsageLocation.ToUpper() } #Case matters I should have a transformer attribute.
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
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification="False positive and need to support plain text here")]
    [cmdletbinding(SupportsShouldProcess=$true)]
    Param (

        [Parameter(ParameterSetName='DomainFromUPNLast',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNDisplay',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [alias("UPN")]
        [string]$UserPrincipalName,

        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='UPNFromDomainDisplay',Mandatory=$true)]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [Parameter(ParameterSetName='DomainFromUPNDisplay')]
        [ValidateNotNullOrEmpty()]
        [Alias("Nickname")]
        [string]$MailNickName,

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

        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$DisplayNameRule = {"$GivenName $Surname"},

        [Parameter(ParameterSetName='UPNFromDomainLast')]
        [Parameter(ParameterSetName='DomainFromUPNLast')]
        [scriptblock]$NickNameRule    = {"$GivenName.$Surname"},

        #A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries.  Examples include: 'US', 'JP', and 'GB'
        [ValidateNotNullOrEmpty()]
        [UpperCaseTransformAttribute()]
        [ValidateCountryAttribute()]
        [string]$UsageLocation = 'GB',

        [string]$Initialpassword,
        [switch]$NoPasswordChange,
        [switch]$ForceMFAPasswordChange,

        [ValidateSet('DisableStrongPassword','DisablePasswordExpiration')]
        [string[]]$PasswordPolicies,
        [hashtable]$SetableProperties,
        [switch]$Force,
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
    #A simple way to create one in 100K temporaty passwords. You might get 10Oct2126 Easy to type and meets complexity rules.
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

function New-GraphInvitation {
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        #The email address of the user being invited.
        #The characters  #~ ! $ %  ^  & * ( [ { < > } ] ) +  = \ /  | ; : " " ? , are not permitted
        #A  . or - is permitted except at the beginning or end of the name. A _  is permitted anywhere.
        [Parameter(Position=1,ValueFromPipeline=$true)]
        [string]$EmailAddress,
        #The display name of the user being invited.
        [string]$DisplayName,
        #The userType of the user being invited. By default, this is Guest. You can invite as Member if you are a company administrator.'
        [string]$UserType,
        #The URL the user should be redirected to once the invitation is redeemed. Required.
        [string]$RedirectUrl  = 'https://mysignins.microsoft.com/',
        #Indicates whether an email should be sent to the user being invited or not.
        [switch]$SendInvitationMessage
    )

    ContextHas -WorkOrSchoolAccount -BreakIfNot
    $settings = @{
        'invitedUserEmailAddress'    = $EmailAddress
        'sendInvitationMessage'      = $SendInvitationMessage -as [bool]
        'inviteRedirectUrl'          = $RedirectUrl
    }
    if ($DisplayName) {$settings['invitedUserDisplayName']  = $DisplayName}
    if ($UserType)    {$settings['invitedUserType']         = $UserType}

    $webparams = @{
        'Method'            = 'POST'
        'Uri'               = "$GraphUri/invitations"
        'Contenttype'       = 'application/json'
        'Body'              = (ConvertTo-Json $settings -Depth 5)
        'AsType'            = [MicrosoftGraphInvitation]
        'ExcludeProperty'   = '@odata.context'
    }
    Write-Debug $webparams.Body
    if ($force -or $pscmdlet.ShouldProcess($EmailAddress, 'Invite User')){
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
        Deletes a user
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
            $uri = $GraphUI + '/me/people?$search="{0}"&$top={1}' -f $SearchTerm, $First
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
    param (
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
    begin {
        $list = @()
    }
    process {
        foreach ($p in $path) {
            if (Test-Path $p) {$list += Import-Csv -Path $p}
            else { Write-Warning -Message "Cannot find $p" }
        }
    }
    end {
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
       Imports a list of users from a CSV file
    .description
        Takes a list of CSV files and looks for xxxx columns
        * Action is either Add, Remove or Set - other values will cause the row to be ignored
        * DisplayName

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
Microsoft.Graph.Users.private\Get-MgUser_List -Select 'UserPrincipalName',  'MailNickName','GivenName', 'Surname', 'DisplayName', 'UsageLocation',
                        'accountEnabled', 'PasswordPolicies', 'Mail',  'MobilePhone', 'BusinessPhones',
                        'JobTitle',  'Department',  'OfficeLocation', 'CompanyName',
                        'StreetAddress', 'City', 'State', 'Country', 'PostalCode' -ExpandProperty manager -filter $Filter |
    Select-Object      'UserPrincipalName', 'MailNickName',   'GivenName', 'Surname',  'DisplayName', 'UsageLocation',
                        @{n='AccountDisabled';e={-not 'accountEnabled'}} , 'PasswordPolicies', 'Mail',  'MobilePhone',
                        @{n='BusinessPhones';e={$_.'BusinessPhones' -join $ListSeparator }},
                        @{n='Manager';e={$_.manager.AdditionalProperties.userPrincipalName}},
                        'JobTitle',  'Department', 'OfficeLocation', 'CompanyName',
                        'StreetAddress', 'City', 'State', 'Country', 'PostalCode' | Export-Csv -NoTypeInformation -Path $Path
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
        UserId          = $UserId
        TodoTaskListId  = $ToDoList
        Title           = $Title
        Body            = (New-Object -TypeName MicrosoftGraphItemBody -Property @{content=$BodyText; contentType=$BodyType} )
        Importance      = $Importance
        Status          = $status
        IsReminderOn    = $ReminderDateTime -as [bool]
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
        $webParams = @{ Method          = 'Get'
                        AllValues       = $true
                        ValueOnly       = $true
                        ExcludeProperty = 'icaluid','@odata.etag','calendar@odata.navigationLink','calendar@odata.associationLink'
                        AsType          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
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
        $webParams = @{Method          = 'Post'
                    ExcludeProperty = 'icaluid','@odata.etag','@odata.context'
                    AsType          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
                    Contenttype     = 'application/json'
                    Headers         = @{Prefer        = "Outlook.timezone=""$TimeZone"""}
        }
        $whensb = {
            if ($this.Start.datetime.AddDays(1) -eq  $this.End.datetime -and
                $this.Start.datetime.hour -eq 0 -and $this.Start.datetime.minute -eq 0 ) {
                $this.Start.datetime.ToShortDateString() + ' All day'
            }
            else {$this.Start.datetime.ToString("g") + ' to ' +  $this.End.datetime.ToString("g") + $this.End.timezone}
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

function Set-GraphEvent          {
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

    $webParams = @{Method          = 'Patch'
                   ExcludeProperty = 'icaluid','@odata.etag','@odata.context'
                   AsType          = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphEvent])
                   Contenttype     = 'application/json'
                   Headers         = @{Prefer        = "Outlook.timezone=""$TimeZone"""}
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

function Remove-GraphEvent       {
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
