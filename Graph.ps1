#Most functions for this are in files based on the application where they would surface
# OneDrive, OneNote, Outlook-Calendar, Outlook-Contacts, Outlook-Mail, Planner, SharePoint, Teams.
# Those in this file don't belong to an application.
Write-Host -ForegroundColor Red "Using the default / sample app ID. You should edit the .PSM1 file and either replace the ID with your own, or remove this message"
$Script:ClientID  = "bf546ecc-067d-4030-9edd-7b0d74913411"  #You can also try  "1950a258-227b-4e31-a9cf-717495945fc2"  # Well known client ID for PowerShell
#$script:Tenant    = Guid-for-your-tennant if the Client ID is set up as below

<#
    You can create an app in Azure AD or at https://apps.dev.microsoft.com/
    I created mine as a native app, with a re-direct URI of https://login.microsoftonline.com/common/oauth2/nativeclient and
    gave it a set of Microsoft graph permissions in Azure AD

    You must use this route if you want people outside your Azure AD (Microsoft accounts) to use the app
    If you ONLY Want to work with accounts in Azure AD you can set up your app with these instructions which I lifted from
    https://msunified.net/2018/12/12/post-at-microsoftteams-channel-chat-message-from-powershell-using-graph-api/
    1.  Log on to https://portal.azure.com with a GA administrator
    2.  Navigate to Azure Active Directory
    3   Go to App registration (Preview)
    4.  Click + New registration
    5.  Call it PowerShellMSGraphAPI
    6.  Leave Redirect URI blank
    7.  Go to Authentication and under Redirect URIs choose urn:ietf:wg:oauth:2.0:oob
    8.  Click Save
    9.  Go to API permissions to grant the required group read and write permissions
    10. Click + Add a permission
    11. Choose Microsoft Graph, Delegated permissions and choose Group.Read.All and ReadWrite.All (remember you need to expand Group)
    12. Click Grant admin Consent from  and click Yes
    13. You now have admin consent granted for your tenant
    14  Navigate to Overview
    15 Copy the Application (client) ID    Paste it into the script as the value for $Script:ClientID;
    Also copy the tenant ID or domain name and make it the value for $script:Tenant
#>

#To prevent tokens being saved, remove the savePath or Set SaveCreds to FALSE
$Script:SavePath    = Join-Path -Path (Split-Path -Path $profile  -Parent) -ChildPath "graph.xml"
$Script:SaveCreds   = $true

#The scopes requested. You can shorten this of you don't need all things phovided in the module
$script:RequestedScope          = @(
      'Directory.AccessAsUser.All', #Grant same rights to the directory as the user has
           'User.ReadWrite.all',    # Read write users and groups may not be needed
          'Group.ReadWrite.All',    # if Directory is granted
      'Calendars.ReadWrite',
      'Calendars.ReadWrite.Shared'
       'Contacts.ReadWrite',
       'Contacts.ReadWrite.Shared',
          'Files.ReadWrite.All',
           'Mail.ReadWrite',
'MailboxSettings.ReadWrite',
          'Notes.ReadWrite',
          'Notes.Create',
         'People.Read.All',
        'Reports.Read.All',
          'Sites.ReadWrite.All',
          'Sites.Manage.All',       #Needed to create lists.
         'openid',
        'profile',
 'offline_access'
)

#Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
$global:drivecache  = @{}

if (-not $script:Tenant) {
    #if we are not working against a known Azure AD tennant,
    #we will need windows forms to the display the logon, which limits where we can work
    Add-Type -AssemblyName System.Windows.Forms
}

Function Connect-MSGraph {
    <#
      .Synopsis
        Connects to the Microsoft graph API; supporting Microsoft accounts (Live) and Office 365 (Azure AD)
      .Example
        >Connect-Msgraph
        Gets a new access token for the graph API if there isn't a current one.
        If a refresh token is avaialble from that will be used to re-establish a session, otherwise a
        logon dialog will be presented.
      .Example
        >Connect-Msgraph -forceNew
        Discards existing credentials and displays a logon dialog.
      .Example
        >Connect-Msgraph -CheckOnly
        Returns True or False as an answer to the question "Is there a current session"
      .Example
        >$AccessToken = Connect-MSGraph -PassThru -Verbose
        >Invoke-RestMethod -Method get -Uri "https://graph.microsoft.com/v1.0/me" -headers @{Authorization = "bearer $AccessToken" }
        Returns the accesstoken, and uses it to call the graph API to get the current users details
    #>
    [cmdletbinding()]
    param (
        #If Specified disposes of any existing connection and creates a new one
        [Switch]$ForceNew,
        #If Specified returns the access token
        [Alias('PT')]
        [Switch]$PassThru,
        #If Specified returns true if we have a current session and false if not
        [Switch]$CheckOnly,
        #If specified, and the tenant is fixed, logs on as a temporary Session using that credential
        [pscredential]$Credential,
        #If specified creates prompts for a login for a temporary session (a login which doesn't save the token or use an existing saved one)
        [Switch]$Temp
    )

    Function Convert-AuthResponse {
        <#
        .Synopsis
            Sets Global variables from REST Response to an OAuth login
        .INPUTS
            Response
        #>
        [CmdletBinding()]
        Param (# The response either read from a file, or fetched from the web service
            [Parameter(Mandatory=$true,ValueFromPipeline=$true)] $Response,
            # A message to send to verbose to say what we did
            [String]$Action = "Processed Response for",
            # Dump the info to an XML file
            [switch]$Save,
            # If the info is fresh from the web, set the expiry time, otherwise don't
            [switch]$SetExpiry
        )
        if ($Save -and $Script:SavePath) {
              Export-Clixml -Path $Script:SavePath  -InputObject $Response -Depth 5
              Write-Verbose -Message "Saving response to $Script:SavePath"
        }
        else {Write-Verbose -Message "Coverting but not saving"}
        if ($Response.access_token) {
            $Script:AccessToken     = $Response.access_token
            $Script:AuthHeader      = 'Bearer ' + $Response.access_token
            $Script:DefaultHeader   = @{Authorization = $Script:AuthHeader}
            $Script:AuthorizedScope = $Response.scope -split " "
        }
        $Script:RefreshToken  = $Response.refresh_token
        if ($setExpiry) { #if we're reading from a file: life in seconds is meaningless, don't set expiry or get the username either
            $Script:TokenExpiry = (Get-Date).AddSeconds([int]$Response.expires_in -60 )
            if ($Response.access_token) {
                Write-Progress -Activity "Authenticating" -Status "Getting Token from Server" -PercentComplete 50
                $Script:GraphUser = (Invoke-RestMethod -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/me/")
                $Script:GraphUser.pstypenames.Add('GraphUser')
                Write-Verbose ($action + $Script:GraphUser.UserPrincipalName)
                $Organization     = (Invoke-RestMethod -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/organization/").Value
                if ($Organization.id) {
                    $script:TenantId     = $Organization.ID
                    $script:TenantName   = $Organization.DisplayName
                    $Script:WorkOrSchool = $true
                    Write-Verbose -Message "Account is from $($Organization.DisplayName)"
                }
                else {
                    $script:TenantId     = $Null
                    $script:TenantName   = $Null
                    $Script:WorkOrSchool = $false
                    Write-Verbose -Message "Account is from Windows live"
                }
            }
        }
    }

    if ($script:Tenant) {
        $tokenUri   = "https://login.microsoft.com/$script:Tenant/oauth2/token"
    }
    else                {
        #The URIs and parameters are set out at https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow
        $CallBackUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"   # windows live used "https://login.live.com/oauth20_desktop.srf"
        $tokenUri    = "https://login.microsoftonline.com/common/oauth2/v2.0/token"     # windows live used "https://login.live.com/oauth20_token.srf"
    }
    if ($ForceNew)      {
        Write-Verbose -Message "ForceNew Specified; removing any existing logon info"
        Remove-item -Path $Script:SavePath -ErrorAction SilentlyContinue
        $Script:RefreshToken = $Script:AccessToken = $Script:TokenExpiry =  $Script:GraphUser = $null
    }
    if ($Temp -or $Credential) {
        $Script:RefreshToken = $Script:AccessToken = $Script:TokenExpiry =  $Script:GraphUser = $null
    }
    $Script:SaveCreds = $saveCreds -and (-not ($Temp -or $Credential))
    <# Scenarios
        A. We have a current access token. Hooray! If called with checkonly return true, if called with passthrough return the token, if called with neither just return
        B. We don't. If called with check only return false. Otherwise we need to logon and that breaks down to
           1. We haven't logged on in this session (no refresh token) but there is a saved refresh token - load it
           2. The access token has expired (or we never had one) but we do have a refersh token (from 1 or from an earlier login) - Make a refresh call
           3  We don't have a current access token, or a referesh token. So we put up a dialog box for the user, and save the token so it can be used in (1).
              As a simple way to prevent tokens being saved, only save if $SavePath contains a path.
           The URLs and logic are described here  https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow
    #>
    #scenario A.  we have a  current token.
    if ($Script:AccessToken -and ((Get-Date) -lt $Script:TokenExpiry) ) { #  if
        Write-Verbose  -Message ('Existing Access Token is good for {0:N0} seconds.' -f ($Script:TokenExpiry.Subtract([datetime]::Now).totalseconds))
        if     ($PassThru)  {return $script:AccessToken}
        elseif ($CheckOnly) {return $true}
        else                {return}
    }
    elseif     ($CheckOnly) {return $false}

    #Scenario B, 1 we have any tokens but there is a saved one, and we're not logging in with a temporary login (so $Script:SaveCreds is true)
    if ((-not $Script:RefreshToken) -and $Script:SavePath -and $Script:SaveCreds -and (Test-Path -Path $Script:SavePath) ) {
        Write-Verbose -Message "No Refresh token, loading data from $Script:SavePath"
        Import-Clixml -Path $Script:SavePath | Convert-AuthResponse
    }
    #Scenario B, 2 We have a refresh token (might have just loaded it or we had it already. So Refresh)
    if   ($Script:RefreshToken) {
        Write-Progress -Activity "Authenticating" -Status "Getting Token from Server"
        $tokenBody = @{'grant_type' = 'refresh_token'; 'refresh_token' = $Script:RefreshToken ; 'client_id' = $Script:ClientID;}
        if  ($Tenant)  {$tokenBody['resource']     = 'https://graph.microsoft.com'}
        else           {$tokenBody['redirect_uri'] = $CallBackUri}      #some cases need &client_secret=xxxyyyyzzz - but not here
        Invoke-RestMethod  -Method Post -Uri $tokenUri -Body $tokenBody  |  Convert-AuthResponse -Save -SetExpiry -Action "Refreshed token for "
        Write-Progress -Activity "Authenticating" -Status "Getting Token from Server" -Completed
    }
    else { #Scenario B, 3 we need to log on
        if ($script:Tenant) { #if we don't have the tennant we need to display the web UI. If do, we can just prompt for creds
            Write-Verbose -Message "Using a fixed tennant"
            if (-not $Credential) {$Credential = Get-Credential -Message "Please enter your credentials for Office 365" }
            if (-not $Credential) {Write-Warning -Message "Can't login without a credential !"; return}
            Write-Progress -Activity "Authenticating" -Status "Getting Token from Server"
            Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                     'grant_type' = 'password'; 'username' = $Credential.username; 'password' = $Credential.GetNetworkCredential().Password;
                     'client_id'  = $clientID;  'resource' = 'https://graph.microsoft.com'
             }  | Convert-AuthResponse -Save:$Script:SaveCreds -SetExpiry -Action "Logged on as "
             Write-Progress -Activity "Authenticating" -Status "Getting Token from Server" -Completed
        }
        else {
            $AuthUri  = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code'+
                          '&client_id='    + $Script:ClientID     +
                          '&scope='        + ($script:RequestedScope-join "%20") +
                          '&redirect_uri=' + $CallBackUri # for windows live: https://login.live.com/oauth20_authorize.srf?..."
            $DocComp  = { #script block for the on_document_complete event: Make URI accessible; close the form if URI has a code or an error
                $Script:uri = $web.Url.AbsoluteUri
                if ($Script:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
            }
            #Create a web browser control pointing at the Auth URI - which will contain the ClientID and be told to send back a code ...
            $web      = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=600;Height=720;Url=($AuthUri) }
            $form     = New-Object -TypeName System.Windows.Forms.Form       -Property @{Width=800;Height=820}
            $web.Add_DocumentCompleted($DocComp) #Add the event handler to the web control
            $form.Controls.Add($web)             #Add the control to the form
            $form.Add_Shown({$form.Activate()})
            $form.ShowDialog() | Out-Null

            #$URI will be set by the event handler ... so did we get a code - meaning the user logged in OK - or did we get an error ?
            if     ( $uri -match "error=([^&]*)") {Write-Warning ("Logon returned an error of " + $Matches[1]); return}
            elseif ( $uri -match "code=([^&]*)" ) {# If we got a code, request & process the token for it
                Write-Progress -Activity "Authenticating" -Status "Getting Token from Server"
                Invoke-RestMethod -Method Post -Uri $tokenUri  -Body @{
                    'grant_type'  ='authorization_code';  'code' = $Matches[1];
                    'client_id'   = $Script:ClientID;     'redirect_uri'= $CallBackUri  #some places also neet  &client_secret=xxxyyyyzzz
                } | Convert-AuthResponse -Save:$Script:SaveCreds  -SetExpiry -Action "Logged on as "
                Write-Progress -Activity "Authenticating" -Status "Getting Token from Server" -Completed
            }
        }
    }
    if     ($PassThru) {return $script:AccessToken}
    elseif (-not ($Script:AccessToken -and ((Get-Date) -lt $Script:TokenExpiry) )) {
        Write-Warning -Message "It doesn't look like there was a valid access token."
    }
}

Function Show-GraphSession {
    <#
        .Synopsis
            Returns Basic information about the current sesssion
    #>
    [CmdletBinding(DefaultParameterSetName='None')]
    Param(
        [Parameter(ParameterSetName='Who')]
        [switch]$Who,
        [Parameter(ParameterSetName='Scopes')]
        [switch]$Scopes
    )
    if (-not $script:AccessToken)   {
        Write-Warning -Message "Not Logged on"
    }
    elseif ($Scopes) {
        $Script:AuthorizedScope
    }
    elseif ($Who) {
        $Script:GraphUser
    }
    else {
        if ($Script:WorkOrSchool)  {'{0} logged on with an Azure AD Account from {1} (Tenant ID {2}).'  -f $Script:GraphUser.UserPrincipalName, $script:TenantName , $script:TenantId }
        else                       {'{0} logged on with a Windows Live account.'                        -f $Script:GraphUser.UserPrincipalName}
        'Access token has an expiry time of: {0}' -f $Script:TokenExpiry
        if ($script:RefreshToken)  {"Refresh token is present."}
        if (-not $Script:SavePath) {"Token is not being saved between sessions."}
        elseif (Test-Path -Path $Script:SavePath) {"Token has been saved."}
        'Token supports these scopes:'
        $Script:AuthorizedScope -join ", "
    }
}

Function Get-GraphOrganization  {
    <#
      .Synopsis
        Gets a summary of organization information from MSGraph
      .Description
        Can use msonline\Get-MsolCompanyInformation instead
        This needs consent to use either the User.Read or the Directory.Read.All scope
      .Example
        >(Get-GraphOrganization).verifiedDomains
        Displays a list of domains in the current subscription
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param(
    )
    Connect-MSGraph
    $webParams = @{ 'uri'     = 'https://graph.microsoft.com/v1.0/organization'
                    'Method'  = 'Get'
                    'Headers' = $Script:DefaultHeader
    }
    $result = Invoke-RestMethod @webParams
    foreach ($org in $result.value) {
        $org.pstypenames.Add('GraphOrganization')
        foreach ($d in $org.verifieddomains) {
            $d.pstypenames.add('GraphDomain')
        }
    }

    $result.value
}

Function Get-GraphDomain {
    <#
      .synopsis
        Gets domains in the current tenant
      .Description
        Requires consent to use at least the Directory.Read.All scope
    #>
    [cmdletbinding()]
    param (
    )
    Connect-MSGraph
    # if user is an admin, can add /NameReferences /serviceConfigurationRecords or /verificationDnsRecords to URI
    $result = Invoke-RestMethod  -Method get -Uri "https://graph.microsoft.com/v1.0/domains"  -headers $Script:DefaultHeader
    foreach ($r in $result.value) {
        $r.psTypeNames.add('GraphDomain')
    }
    $result.value
    #    -Uri   "https://graph.microsoft.com/v1.0/domains/{domain-name}/domainNameReferences" -Method Get -headers $Script:DefaultHeader
}

Function Get-GraphSKUList {
    <#
      .Synopsis
        Gets the SKUs organization an organization has subscribed to
      .Description
        Equivalent to  msonline\Get-MsolAccountSku
        Requires consent to the Directory.Read.All or the Directory.AccessAsUser.All scope
      .Example
        >Get-GraphSKUList | where {($_.prePaidUnits.enabled - $_.consumedunits) -lt 10}
        Lists SKUs where the number of licenses consumed is getting close to the number purchased.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param ()
    Connect-MSGraph
    $webParams = @{Method = "Get"
                   Headers = $Script:DefaultHeader
    }
    $subscribedSkus =  (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/subscribedSkus").value
    foreach ($s in $subscribedSkus) {$s.pstypenames.Add("GraphSKU")}

    $subscribedSkus
}

Function Get-GraphSKU {
    <#
      .Synopsis
        Gets details of SKUs organization an organization has subscribed to
      .Example
        Get-GraphSKUList | where skupartnumber -match "enterprise" | Get-GraphSKU -ServicePlans | sort servicePlanName | format-table
        Finds "Enterprise" SKUS and displays their service plans in alphabetical order.
    #>
    [cmdletbinding()]
    Param (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $SKU,
        #If specified just returns the Service plans for the SKU, otherwise returns the SKU with a service plans property
        [switch]$ServicePlans
    )
    Begin   {
        Connect-MSGraph
    }
    Process {
        $webParams = @{'Method'  = "Get"
                       'Headers' = $Script:DefaultHeader
        }
        foreach ($s in $sku) {
            if     ($s.id)          {$webParams["uri"] = "https://graph.microsoft.com/v1.0/subscribedSkus/$($s.id)" }
            elseif ($s -is [String]){$webParams["uri"] = "https://graph.microsoft.com/v1.0/subscribedSkus/$s" }
            else   {Write-Warning -Message 'Could not find the SKU ID from the parameter'; return}

            $result  = Invoke-RestMethod @webParams
            $result.pstypenames.Add("GraphSKU")
            foreach($s in $result.ServicePlans) {
                $s.pstypenames.Add("GraphServicePlan")
                Add-Member -InputObject $s -MemberType NoteProperty -Name "skuPartNumber" -Value $result.skuPartNumber
            }

            if ($ServicePlans) {$result.ServicePlans}
            else               {$result }
        }
    }
}

Function Get-GraphReport {
    <#
        .Synopsis
            Use BETA functionality to get reports from MS Graph
        .Example
            >Get-GraphReport -Report MailboxUsageDetail | ft "Display Name",  "Storage Used (Byte)"
            Displays mailbox storage used by users - note that
            fields have 'friendly' names which need to be wrapped in quotes
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #The report to Fetch
        [ValidateSet('TeamsUserActivityUserCounts',     'TeamsUserActivityCounts',       'TeamsUserActivityUserDetail',
                     'TeamsDeviceUsageUserDetail',      'TeamsDeviceUsageUserCounts',    'TeamsDeviceUsageDistributionUserCounts',
                    'EmailActivityCounts',              'EmailActivityUserCounts',       'EmailActivityUserDetail',
                    'EmailAppUsageAppsUserCounts',      'EmailAppUsageUserDetail',
                    'MailboxUsageMailboxCounts',        'MailboxUsageDetail',            'MailboxUsageQuotaStatusMailboxCounts', 'MailboxUsageStorage',
                    'Office365ActivationsUserDetail',   'Office365ActivationCounts',     'Office365ActivationsUserCounts',
                    'Office365ActiveUserDetail' ,       'Office365ActiveUserCounts',     'Office365ServicesUserCounts',
                    'Office365GroupsActivityDetail',    'Office365GroupsActivityCounts', 'Office365GroupsActivityGroupCounts',   'Office365GroupsActivityStorage',
                    'Office365GroupsActivityFileCounts','OneDriveActivityUserDetail',    'OneDriveActivityFileCounts',           'OneDriveActivityUserCounts',
                    'OneDriveUsageAccountDetail',       'OneDriveUsageAccountCounts',    'OneDriveUsageFileCounts' ,             'OneDriveUsageStorage',
                    'SharePointActivityUserDetail',     'SharePointActivityFileCounts',  'SharePointActivityUserCounts',         'SharePointActivityPages',
                    'SharePointSiteUsageDetail',        'SharePointSiteUsageFileCounts', 'SharePointSiteUsageSiteCounts',        'SharePointSiteUsageStorage',
                    'SharePointSiteUsagePages')]
        [parameter(Mandatory=$true)]
        $Report,
        #Date for the report - this should be a date in the past 30 days. If specified, -Period is ignored. Reports ending in Count, Storage or pages don't support date filtering
        [DateTime]$Date,
        #The range of time for the report in the form "Dn" where n is the number of days. The default is D7, except for Office365Activation activation reports
        [ValidateSet("D7", "D30", "D90", "D180")]
        $Period
    )
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    Connect-MSGraph
    $webParams = @{'Method' = 'Get'
                   'Headers' = $Script:DefaultHeader
    }
    if     ($Date)    {
        if ($report -match 'Counts$|Pages$|Storage$') {Write-Warning -Message 'Reports ending with Counts, Pages or Storage do not support date filtering' ; return }
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        if ($report -eq    'MailboxUsageDetail')      {Write-Warning -Message 'MailboxUsageDetail does not support date filtering.' ; return}
        $webParams['Uri'] = "https://graph.microsoft.com/beta/reports/get{0}(date={1:yyyy-MM-dd})" -f $Report , $Date
    }
    elseif ($Period)  {
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        $webParams['Uri'] = "https://graph.microsoft.com/beta/reports/get{0}(period='{1}')"        -f $Report , $Period
    }
    else              {
      if ($report -notmatch '^Office365Activation')  {
        $webParams['Uri'] = "https://graph.microsoft.com/beta/reports/get{0}(period='d7')"         -f $Report
      }
      else {
        $webParams['Uri'] = "https://graph.microsoft.com/beta/reports/get{0}"                      -f $Report
      }
    }
    (Invoke-RestMethod @webParams).substring(3) | ConvertFrom-Csv
}

Function Get-GraphSignInLog {
    <#
      .synopsis
        Gets the audit log -requires a priviledged account
      .Description
        This command calls https://graph.microsoft.com/beta/auditLogs/signIns
        which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.
      .Example
        >
        >Get-GraphSignInLog |
        >  select Date,UserPrincipalName,appDisplayName,ipAddress,clientAppUsed,browser,device,city,lat,long |
        >    Export-Excel -Path .\signin.xlsx -AutoSize -IncludePivotTable -PivotTableName Signins -PivotRows appdisplayName -PivotColumns browser -PivotData @{date='Count'} -show

        Gets the sign-in Log and exports it Excel, creating a PivotTable
    #>
    [cmdletbinding()]
    param (
    )
    Connect-MSGraph
    $i = 1
    Write-Progress -Activity 'Getting Sign-in Auditlog'
    try   { $result  = Invoke-RestMethod  -Method get -Uri "https://graph.microsoft.com/beta/auditLogs/signIns"  -headers $Script:DefaultHeader  }
    catch {
        if ($_.exception.response.statuscode.value__ -eq 401) {
            Write-Warning -Message "The server responded 'Unauthorized' - check that $($script:GraphUser.userPrincipalName) has rights to access the log."; return
        }
    }
    $records = $result.value
    while ($result.'@odata.nextLink') {
        $i ++
        Write-Progress -Activity 'Getting Sign-in Auditlog' -CurrentOperation "Page $i"
        $result   = Invoke-RestMethod  -Method get -Uri $result.'@odata.nextLink'  -headers $Script:DefaultHeader
        $records += $result.value
    }
    foreach ($r in $records) {
        $r.pstypenames.add('GraphSigninLog')
        Add-Member -InputObject $r -MemberType ScriptProperty -Name City    -Value {$this.location.city}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name State   -Value {$this.location.state}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Country -Value {$this.location.countryOrRegion}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Lat     -Value {$this.location.geoCoordinates.latitude}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Long    -Value {$this.location.geoCoordinates.longitude}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Browser -Value {$this.deviceDetail.browser}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Device  -Value {$this.deviceDetail.displayName;}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Date    -Value {[datetime]$this.createdDateTime}
    }
    Write-Progress -Activity 'Getting Sign-in Auditlog'-Completed

    $records
}


Function Get-GraphDirectoryLog {
    <#
      .synopsis
        Gets the Directory audit log -requires a priviledged account
      .Description
        This command calls https://graph.microsoft.com/beta/auditLogs/directoryAudits
        which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.

    #>
    [cmdletbinding()]
    param (
    )
    Connect-MSGraph
    $i = 1
    Write-Progress -Activity 'Getting Directory Audits log'
    try   { $result  = Invoke-RestMethod  -Method get -Uri "https://graph.microsoft.com/beta/auditLogs/directoryAudits"  -headers $Script:DefaultHeader  }
    catch {
        if ($_.exception.response.statuscode.value__ -eq 401) {
            Write-Warning -Message "The server responded 'Unauthorized' - check that $($script:GraphUser.userPrincipalName) has rights to access the log."; return
        }
    }
    $records = $result.value
    while ($result.'@odata.nextLink') {
        $i ++
        Write-Progress -Activity 'Getting Directory Audits log' -CurrentOperation "Page $i"
        $result   = Invoke-RestMethod  -Method get -Uri $result.'@odata.nextLink'  -headers $Script:DefaultHeader
        $records += $result.value
    }
    $defaultProperties = @('Date','User','ActivityDisplayName','result')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    foreach ($r in $records) {
        $r.pstypenames.add('GraphDirectoryLog')
        Add-Member -InputObject $r -MemberType ScriptProperty -Name User              -Value {$this.initiatedBy.user.userPrincipalName}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Date              -Value {[datetime]$this.activityDateTime}
        Add-Member -InputObject $r -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers
    }
    Write-Progress -Activity 'Getting Directory Audits log' -Completed

    $records
}

<#
 Others to explore
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directoryRoles").value | ft id,displayname,description               https://docs.microsoft.com/en-us/graph/api/directoryrole-list?view=graph-rest-1.0
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates").value | sort displayname |  ft id,displayname,description
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/groupsettingTemplates").value | ft displayname,description -wrap -aut
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices").value | ft approximateLastSignInDateTime,displayName,operatingsystem,operatingsystemversion
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices/c2221377-d362-42e7-8e16-e7d6abf80e61/registeredOwners").value
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices/c2221377-d362-42e7-8e16-e7d6abf80e61/memberof").value
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/me/owneddevices").value | ft displayname,operatingsystemversion,trusttype
#>

Get-PSCallStack | Out-File -Append ~\graph.txt


Connect-MSGraph
$Global:AccessToken = $script:AccessToken
