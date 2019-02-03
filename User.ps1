#See also Msonline\Get-MsolUser
Function Get-GraphUserList{
    <#
      .Synopsis
        Returns a list of Azure active directory users for the current tennant.
      .Example
        Get-GraphUserList - filter "Department eq 'Finance'""
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
   Param(
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
        #Names of the fields to return for each user.
        [string[]]$Select,
        #Order by clause for the query
        [string]$OrderBy,
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
         #If specified searches for users whose first name, surname, displayname, mail address or UPN start with that name.
        [string]$Name,
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        #Filter clause for the query
        [string]$Filter
    )

    Connect-MSGraph
    $webParams = @{Method = "Get"
                  Headers = $Script:DefaultHeader
    }
    $uri = "https://graph.microsoft.com/v1.0/users"
    #order by and filter do work for the user list (unlike the descendants of a single user. )
    $JoinChar = "?"
    if ($Select)   {
      $uri = $uri + '?$select=' + ($Select -join ',')
      $JoinChar = "&"
    }
    if ($Name)     {
      $uri = $uri + $JoinChar + ("`$filter=startswith(displayName,'{0}') or startswith(givenName,'{0}') or startswith(surname,'{0}') or startswith(mail,'{0}') or startswith(userPrincipalName,'{0}')" -f $Name )
      $JoinChar = "&"
    }
    if ($OrderBy)  {
      $uri = $uri + $JoinChar + '$OrderBy=' + $OrderBy
      $JoinChar = "&"
    }
    if ($Filter)   {
      $uri = $uri + $JoinChar + '$Filter='  +$Filter
    #s  $JoinChar = "&"
    }
    Write-Progress "Getting the List of users"
    $result  =  ( Invoke-RestMethod @webParams -Uri $uri)
    $users   =  $result.value
    while      ($result.'@odata.nextLink') {
            $result   =  Invoke-RestMethod @webParams -Uri $result.'@odata.nextLink'
            $users   += $result.value
    }
    foreach ($u in $users) {$u.pstypenames.Add("GraphUser") }
    Write-Progress "Getting the List of users" -Completed
    return $users
}

Function Get-GraphUser {
    <#
      .Synopsis
        Gets information from the MS-Graph API about the a user (current user by default)
      .Example
        get-graphuser -MemberOf | ft displayname, description, mail, id
        Shows the name description, email address and internal ID for the groups this user is a direct member of
      .Example
        (get-graphuser -Drive).root.children.name
        Gets the user's one drive. The drive object has a root property which is represents the drives root
        directory, and this has a children property which is a collection of the objects in the root directory.
        This command shows the names of the objects in the root directory.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param   (
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [parameter(Position=0)]
        [string]$UserID,
        #Get the user's Calendar(s)
        [Parameter(Mandatory=$true, ParameterSetName="Calendars")]
        [switch]$Calendars,
        #Gets the user's Owned devices (this API is still in Beta)
        [Parameter(Mandatory=$true, ParameterSetName="Devices")]
        [switch]$Devices,
        #Select people who have the user as their manager
        [Parameter(Mandatory=$true, ParameterSetName="DirectReports")]
        [switch]$DirectReports,
        #Get the user's one drive
        [Parameter(Mandatory=$true, ParameterSetName="Drive")]
        [switch]$Drive,
        #Get users license Details
        [Parameter(Mandatory=$true, ParameterSetName="LicenseDetails")]
        [switch]$LicenseDetails,
        #Get the user's Mailbox Settings
        [Parameter(Mandatory=$true, ParameterSetName="MailboxSettings")]
        [switch]$MailboxSettings,
        #Get the users Outlook-categories (by default, 6 color names)
        [Parameter(Mandatory=$true, ParameterSetName="OutlookCategories")]
        [switch]$OutlookCategories,
        #Get the user's manager
        [Parameter(Mandatory=$true, ParameterSetName="Manager")]
        [switch]$Manager,
        #Get the user's teams
        [Parameter(Mandatory=$true, ParameterSetName="Teams")]
        [switch]$Teams,
        #Get the user's Groups
        [Parameter(Mandatory=$true, ParameterSetName="Groups")]
        [switch]$Groups,
        [Parameter(Mandatory=$false, ParameterSetName="Groups")]
        [Parameter(Mandatory=$true, ParameterSetName="SecurityGroups")]
        [switch]$SecurityGroups,
        #Get the Directory-Roles and Groups the user belongs to; -Groups or -Teams only return one type of object.
        [Parameter(Mandatory=$true, ParameterSetName="MemberOf")]
        [switch]$MemberOf,
        #Get the user's Notebook(s)
        [Parameter(Mandatory=$true, ParameterSetName="Notebooks")]
        [switch]$Notebooks,
        #Get the user's photo
        [Parameter(Mandatory=$true, ParameterSetName="Photo")]
        [switch]$Photo,
        #Get the user's planners
        [Parameter(Mandatory=$true, ParameterSetName="Planner")]
        [switch]$PlannerTasks,
        #Get the user's MySite in SharePoint
        [Parameter(Mandatory=$true, ParameterSetName="Site")]
        [switch]$Site,
        #specifies which properties of the user object should be returned
        [Parameter(Mandatory=$true,ParameterSetName="Select")]
        [ValidateSet  ("aboutMe", "accountEnabled", "ageGroup", "assignedLicenses", "assignedPlans", "birthday", "businessPhones",
        "city", "companyName", "consentProvidedForMinor", "country", "createdDateTime", "department", "displayName", "givenName",
        "hireDate", "id", "imAddresses", "interests", "jobTitle", "legalAgeGroupClassification", "mail", "mailboxSettings",
        "mailNickname", "mobilePhone", "mySite", "officeLocation", "onPremisesDomainName", "onPremisesExtensionAttributes",
        "onPremisesImmutableId", "onPremisesLastSyncDateTime", "onPremisesProvisioningErrors", "onPremisesSamAccountName",
        "onPremisesSecurityIdentifier", "onPremisesSyncEnabled", "onPremisesUserPrincipalName", "passwordPolicies",
        "passwordProfile", "pastProjects", "postalCode", "preferredDataLocation", "preferredLanguage", "preferredName",
        "provisionedPlans", "proxyAddresses", "responsibilities", "schools", "skills", "state", "streetAddress",
        "surname", "usageLocation", "userPrincipalName", "userType")]
        [String[]]$Select
    )
    Begin   {
        Connect-MSGraph
    }
    Process {
        if ($UserID) {$userID = "users/$userID"} else {$userid = "me"}
        $webParams = @{Method = "Get"
                    Headers = $Script:DefaultHeader
        }
        if (-not $Script:WorkOrSchool -and ($MailboxSettings -or $Manager -or $Photo -or $DirectReports -or $LicenseDetails -or $MemberOf -or $Teams -or $PlannerTasks -or $Devices ))  {
            Write-Warning   -Message "Only the -Drive, -Calendars and -Notebooks options work when you are logged in with this kind of account." ; return
        }
        #available:  but not implemented
        #   https://graph.microsoft.com/beta/me/transitiveMemberOf
        #   https://graph.microsoft.com/beta/me/insights/used" /trending or /stored.
        #   Https://graph.microsoft.com/beta/me/Activities"         needs UserActivity.ReadWrite.CreatedByApp permission
        #   https://graph.microsoft.com/v1.0/me/activities/recent
        #   https://graph.microsoft.com/v1.0/me/createdobjects
        #(Invoke-RestMethod -Method POST -Headers @{Authorization = "Bearer $script:AccessToken"} -Uri "https://graph.microsoft.com/v1.0/me/getmemberobjects"  -body '{"securityEnabledOnly": false}' -  ).value

        #It would be nice if we could apply filter and orderby to some of these, but for some they are ignored and for others they cause errors.

        #for everything Except -Site we can define a URI and either return the Value Propety of the result, or the whole result.

        # Site needs special handling. Get the user's MySite. Convert it into a graph URL and get that, expand drives subSites and lists, and add formatting types
        Write-Progress -Activity 'Getting user information'
        if     ($Site) {
            $uri    = "https://graph.microsoft.com/v1.0/$userID`?`$select=mysite "
            $result = Invoke-RestMethod @webParams -Uri $uri
            $uri    = $result.mysite -replace "^https://(.*?)/(.*)$", 'https://graph.microsoft.com/v1.0/sites/$1:/$2?expand=drives,lists,sites'
            $result = Invoke-RestMethod @webParams -Uri $uri
            $result.pstypenames.Add("GraphSite") 
            foreach ($l in $result.lists) {
                $l.pstypenames.Add("GraphList") 
                Add-Member -InputObject $l -MemberType NoteProperty   -Name SiteID   -Value  $result.id
                Add-Member -InputObject $l -MemberType ScriptProperty -Name Template -Value {$this.list.template}
            }
            Write-Progress -Activity 'Getting user information' -Completed
            return $result
        }        
        elseif ($Devices          ) { $uri = "https://graph.microsoft.com/beta/$userID/owneddevices"    ; $returnTheValue = $true }
        elseif ($DirectReports    ) { $uri = "https://graph.microsoft.com/v1.0/$userID/directReports"   ; $returnTheValue = $true }
        elseif ($LicenseDetails   ) { $uri = "https://graph.microsoft.com/v1.0/$userID/licenseDetails"  ; $returnTheValue = $true }
        elseif ($MemberOf         ) { $uri = "https://graph.microsoft.com/v1.0/$userID/MemberOf"        ; $returnTheValue = $true }
        elseif ($Teams            ) { $uri = "https://graph.microsoft.com/v1.0/$userID/joinedTeams"     ; $returnTheValue = $true }
        elseif ($PlannerTasks     ) { $uri = "https://graph.microsoft.com/v1.0/$userID/Planner/tasks"   ; $returnTheValue = $true }
        elseif ($Photo            ) { $uri = "https://graph.microsoft.com/v1.0/$userID/Photo"           ; $returnTheValue = $false}
        elseif ($MailboxSettings  ) { $uri = "https://graph.microsoft.com/v1.0/$userID/MailboxSettings" ; $returnTheValue = $false}
        elseif ($Manager          ) { $uri = "https://graph.microsoft.com/v1.0/$userID/Manager"         ; $returnTheValue = $false}
        elseif ($Drive            ) { $uri = "https://graph.microsoft.com/v1.0/$userID/Drive"           ; $returnTheValue = $false
                                      if ($WorkOrSchool) {$uri += '?$expand=root($expand=children)'}                              }
        elseif ($Groups -or 
                $SecurityGroups   ) { $uri = "https://graph.microsoft.com/v1.0/$userID/getMemberGroups"} #special handler no need for $return the value
            
        elseif ($OutlookCategories) { $uri = "https://graph.microsoft.com/v1.0/$userID/Outlook"   +
                                                                            '/MasterCategories'         ; $returnTheValue = $true }
        elseif ($Calendars        ) { $uri = "https://graph.microsoft.com/v1.0/$userID/Calendars" +
                                                                                 '?$orderby=Name'       ; $returnTheValue = $true }
        elseif ($Notebooks        ) { $uri = "https://graph.microsoft.com/v1.0/$userID/onenote/"  +
                                                                    'notebooks?$expand=sections'        ; $returnTheValue = $true }
        else                        { $uri = "https://graph.microsoft.com/v1.0/$userID"                 ; $returnTheValue = $false
                                      if ($select) {$uri = $uri + '?$select=' + ($Select -join ",") }                             }

        try   {
            if ($Groups -or $SecurityGroups) {
                $uri          = "https://graph.microsoft.com/v1.0/$userID/getMemberGroups"
                if  ($SecurityGroups) {$body = '{  "securityEnabledOnly": true  }'}
                else                  {$body = '{  "securityEnabledOnly": false }'}
                
                $result       = Invoke-RestMethod  -Uri $uri -Method Post -Headers $Script:DefaultHeader -Body $body -ContentType 'application/json'
                $results      = @()
                foreach ($r in $result.value) {
                    $uri = "https://graph.microsoft.com/v1.0/directoryObjects/$r"
                    $results += Invoke-RestMethod -Uri $uri -Method Get  -Headers $Script:DefaultHeader
                }
            }
            elseif (-not $returnTheValue) {
                    $results = Invoke-RestMethod -Uri $uri @webParams  
            }
            else {
                    $result  = Invoke-RestMethod -Uri $uri @webParams  
                    $results = $result.value
                    while      ($result.'@odata.nextLink') {
                        $result   =  Invoke-RestMethod @webParams -Uri $result.'@odata.nextLink'
                        $results += $result.value
                    }
            }
        }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Progress -Activity 'Getting user information' -Completed
                Write-Warning -Message "Not found error while getting data for user '$userid'" ; return
            }
            else {throw $_ ; return}
        }

        foreach ($r in $results) {
                if     ($r.'@odata.type' -match 'directoryRole$') 
                                           { $r.pstypenames.Add('GraphDirectoryRole')}
                elseif (($r.'@odata.type' -match 'user$' -or 
                         $PSCmdlet.ParameterSetName -eq 'None') -and
                        (-not $Select ))   { $r.pstypenames.Add('GraphUser') }
                elseif ($r.'@odata.type' -match 'group$') 
                                           { $r.pstypenames.Add('GraphGroup') }
                elseif ($r.'@odata.type' -match 'device$') 
                                           { $r.pstypenames.Add('GraphDevice') }
                elseif ($MailboxSettings ) { $r.pstypenames.Add('GraphMailboxSettings')}
                elseif ($Photo           ) { $r.pstypenames.Add('GraphPhoto')}
                elseif ($Drive           ) { $r.pstypenames.Add('GraphDrive')}
                elseif ($Calendars       ) { $r.pstypenames.Add('GraphCalendar')}
                elseif ($LicenseDetails  ) { $r.pstypenames.Add('GraphLicense')}
                elseif ($PlannerTasks    ) { $r.pstypenames.Add('GraphTask')}
                elseif ($Notebooks      )  { 
                    $r.pstypenames.Add('GraphOneNoteBook')
                    foreach ($s in $r.sections) {
                            Add-Member -InputObject $s -MemberType NoteProperty -Name ParentNotebookID -Value $r.id
                            $s.pstypeNames.add("GraphOneNoteSection")
                    }
                }
                elseif ($Teams           ) {
                    $defaultProperties = @('displayName','description','isArchived')
                    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
                    $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                    Add-Member -InputObject $r -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers
                }
        }
        Write-Progress -Activity 'Getting user information' -Completed
        return $results
    }
}

Function Set-GraphUser{
    <#
      .Synopsis
        Sets properties of  a user (the current user by default)
      .Example
        Set-GraphUser -Birthday "31 march 1965"  -Aboutme "Lots to say" -PastProjects "Phoenix","Excalibur" -interests "Photography","F1" -Skills "PowerShell","Active Directory","Networking","Clustering","Excel","SQL","Devops","Server builds","Windows Server","Office 365" -Responsibilities "Design","Implementation","Audit"
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param (
        #ID for the user if not the current user
        $userID = "me",
        #Text for the user's 'about me' text
        [String]$AboutMe,
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
        [Switch]$Force

    )
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

    $webParams = @{ 'Method'      = 'PATCH'
                    'Headers'     = $Script:DefaultHeader
                    'Contenttype' = 'application/json'
    }
    if ($UserID -eq "me") {
              $webParams['uri']   = "https://graph.microsoft.com/v1.0/me/"
    }
    else   {  $webParams['uri']   = "https://graph.microsoft.com/v1.0/users/$UserID/" }


    $settings = @{}
    foreach ($p in $PSBoundParameters.Keys.where({$_ -notin @('Photo','UserID')})) {
        $key   = $p.toLower()[0] + $p.Substring(1)
        $value = $PSBoundParameters[$p]
        if ($value -is [datetime]) {$value = $value.ToString("yyyy-MM-ddT00:00:00Z")}  # 'o' for ISO date time may work here
        $settings[$key] = $value
    }

    if ($Settings.count) {
        $json = (ConvertTo-Json $settings)
        Write-Debug  $json
        if ($Force -or $Pscmdlet.ShouldProcess($userID ,'Update User')) {Invoke-RestMethod @webParams -Body $json }
    }
    elseif (-not $Photo) {Write-Warning -Message "Nothing to set"}
    if ($photo) {
        if (-not (Test-Path $Photo) -or $photo -notlike "*.jpg" ) {
            Write-Warning "$photo doesn't look like the path to a .jpg file" ; return
        }
        $webParams = @{'Method'      = 'Put'
                       'URI'         = 'https://graph.microsoft.com/v1.0/me/photo/$value'
                       'Headers'     = $Script:DefaultHeader
                       'Contenttype' = 'image/jpeg'
                       'infile'      = $Photo
        }
        Invoke-RestMethod @webParams
  }
}

Function Find-GraphPeople {
    <#
       .Synopsis
          Searches people in your inbox / contancts / directory
       .Example
          Find-GraphPeople -Topic timesheet -First 6
          Returns the top 6 results for people you have discussed timesheets with.
    #>
    [cmdletbinding(DefaultParameterSetName='Default')]
    Param (
        #Text to use in a 'Topic' Search. Topics are not pre-defined, butinferred using machine learning based on your conversation history (!)
        [parameter(ValueFromPipeline=$true,Position=0,ParameterSetName='Default',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $Topic,
        #Text to use in a search on name and email address
        [parameter(ValueFromPipeline=$true,ParameterSetName='Fuzzy',Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $SearchTerm,
        #Number of results to return (10 by default)
        [ValidateRange(1,1000)]
        [int]$First = 10
    )
    begin {
        Connect-MSGraph
        $webParams = $webParams = @{Method = "Get"
                    Headers = $Script:DefaultHeader
        }
    }
    process {
        if ($Topic) {
            $webParams['uri'] = 'https://graph.microsoft.com/v1.0/me/people?$search="topic:{0}"&$top={1}' -f $Topic, $First
        }
        elseif ($SearchTerm) {
            $webParams['uri'] = 'https://graph.microsoft.com/v1.0/me/people?$search="{0}"&$top={1}' -f $SearchTerm, $First
        }

        $result = Invoke-RestMethod @webParams

        foreach ($response in $result.value) {
            $response.pstypenames.add('GraphContact')
            Add-Member -InputObject $response -MemberType ScriptProperty -Name mobilephone    -Value {$This.phones.where({$_.type -eq 'mobile'}).number -join ', '}
            Add-Member -InputObject $response -MemberType ScriptProperty -Name businessphones -Value {$This.phones.where({$_.type -eq 'business'}).number }
            Add-Member -InputObject $response -MemberType ScriptProperty -Name Score          -Value {$This.scoredEmailAddresses[0].relevanceScore }
            Add-Member -InputObject $response -MemberType AliasProperty  -Name emailaddresses -Value scoredEmailAddresses -PassThru
        }
    }
}

<#
PUT https://graph.microsoft.com/v1.0/users/{id}/manager/$ref   Content-type: application/json
    {   "@odata.id": "https://graph.microsoft.com/v1.0/users/{id}" }
#>

<#
POST https://graph.microsoft.com/beta/me/assignLicense
Content-type: application/json
Content-length: 185

{
  "addLicenses": [
    {
      "disabledPlans": [ "11b0131d-43c8-4bbb-b2c8-e80f9a50834a" ],
      "skuId": "skuId-value-1"
    },
    {
      "disabledPlans": [ "a571ebcc-fqe0-4ca2-8c8c-7a284fd6c235" ],
      "skuId": "skuId-value-2"
    }
  ],
  "removeLicenses": []
}
#>