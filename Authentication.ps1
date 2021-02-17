#Requires -Module Microsoft.Graph.Authentication
using namespace Microsoft.Graph.PowerShell.Authentication
using namespace Microsoft.Graph.PowerShell.Models
<#
    The Connect-Graph function incorporates work to get tokens from an Azure session and
    to referesh tokens published by Justin Grote at
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Connect-MgGraphAz.ps1
    and licensed by him under the same MIT terms which apply to this module (see the LICENSE file for details)

    Portions of this file are   Copyright 2021 Justin Grote @justinwgrote
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write host used for colored information message telling user to make a change and remove the message')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Initialization clears drive cache and work or school status available outside the module')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]

#Script / global environment variables.
#Global:driveCache caches drive to name mappings, Global:__MgAzContext sets the profile for Azure logons,
#GLOBAL:__MgAzTokenExpires, records when the token will expire if we need to manage refreshing it
#global:__MgToken, allows an  global:DefaultGraphScopes
#Global:GraphUser, GLOBAL:PSDefaultParameterValues,
#$Env:GraphScopes will provide as set of scopes to request
#RefreshToken, RefreshParameters (to call connect graph with will the token expires) clent app (clientID and Client Secret) & tennant are script level vars

#Write-Host -ForegroundColor Red "Using the default / sample app ID. You should edit the .PSM1 file and either replace the ID with your own, or remove this message"
#$Script:ClientID      = "bf546ecc-067d-4030-9edd-7b0d74913411"  #You can also try  "1950a258-227b-4e31-a9cf-717495945fc2"  # Well known client ID for PowerShell
#$script:Tenant    = Guid-for-your-tennant **if the Client ID is set up as below**
$Script:ClientID      = "6413d5f3-eba6-4af4-860b-e9334ff7b762"
$Script:Client_Secret = "a1c-Pr~XaqB.Sy4.e-HN064myE_0u7USze"
$Script:Tenant        = "e6af5578-6d03-49e0-af3b-383cf5ec0b5f"

<#
    You can create an app in Azure AD or at https://apps.dev.microsoft.com/
    I created mine as a native app, with a re-direct URI of https://login.microsoftonline.com/common/oauth2/nativeclient and
    gave it a set of Microsoft graph permissions in Azure AD

    You must use this route if you want people outside your Azure AD (Microsoft accounts) to use the app
    If you ONLY Want to work with accounts in Azure AD you can set up your app with these instructions which I lifted from
    https://msunified.net/2018/12/12/post-at-microsoftteams-channel-chat-message-from-powershell-using-graph-api/
    1.  Log on to https://portal.azure.com with a GA administrator
    2.  Navigate to Azure Active Directory
    3   Go to App registrations
    4.  Click + New registration
    5.  Call it PowerShellMSGraphAPI (or another name of your choice)
    6.  Leave who can use this API on the default of single tennant and leave the Redirect URI blank
    7.  Click Register
    8.  This will bring up the details of the new APP. Under call APIS click View API permissions to grant the required group read and write permissions
    9.  Click + Add a permission
    10. Choose Microsoft Graph, then Delegated permissions and choose Group.Read.All and ReadWrite.All (remember you need to expand Group)
    12. I had to click the enterprise apps link and click "Grant admin Consent" from (this is where a GA admin is needed)
    13. You now have admin consent granted for your tenant, so users can authenticate without a consent dialog.
    14. Navigate back to Overview
    15. Copy the Application (client) ID    Paste it into this script as the value for $Script:ClientID;
    16. Also copy the tenant ID paste it into this script as the value for $script:Tenant
    17. Click Certificates and Secrets, add a secret and chose never expires (unless you want to update the script later), click add
    18. Copy the secret and paste into this script as the value for $script:clientSecret.
#>

#Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
$global:drivecache  = @{}

#The scopes requested. You can shorten this of you don't need all things provided in the module
if ($Env:GraphScopes) {$global:DefaultGraphScopes = $Env:GraphScopes -split ',\s*'}
else                  {$global:DefaultGraphScopes = @(
                'AuditLog.Read.All',
                'Directory.AccessAsUser.All', #Grant same rights to the directory as the user has
                'Calendars.ReadWrite',
                'Calendars.ReadWrite.Shared'
                'ChannelMessage.Read.All',
                'ChannelMessage.Delete',
                'ChannelMessage.Edit',
                'Contacts.ReadWrite',
                'Contacts.ReadWrite.Shared',
                'Files.ReadWrite.All',
                'Group.ReadWrite.All',# or read fails when logging on as non-admin
                'Mail.ReadWrite',
                'Mail.Send',
                'MailboxSettings.ReadWrite',
                'Notes.ReadWrite.All',
                'Notes.Create',
                'People.Read.All',
                'Presence.Read.All',
                'Reports.Read.All',
                'Sites.ReadWrite.All',
                'Sites.Manage.All',       #Needed to create lists.
                'Tasks.ReadWrite',        #Needed for Todo access
                'User.ReadWrite.all',    # Read write users and groups may not be needed if Directory is granted ?
                'openid',
                'profile'#,        'offline_access'
)}

Remove-item Alias:\Invoke-GraphRequest -ErrorAction SilentlyContinue
Function Invoke-GraphRequest {
    param(
        #Uri to call can be a segment such as /beta/me or a fully qualified https://graph.microsoft.com/beta/me
        [Parameter(Mandatory=$true, Position=1 )]
        [uri]$Uri,

        #Http Method
        [ValidateSet('GET','POST','PUT','PATCH','DELETE')]
        [Parameter(Position=2 )]
        $Method,

        #Request Body. Required when Method is Post or Patch'
        [Parameter(Position=3,ValueFromPipeline=$true)]
        $Body,

        #Optional Custom Headers
        [System.Collections.IDictionary]$Headers,

        #Output file where the response body will be saved
        [string]$OutputFilePath,

       [switch]$InferOutputFileName,

        #Input file to send in the request
        [string]$InputFilePath,

        #Indicates that the cmdlet returns the results, in addition to writing them to a file. Only valid when the OutFile parameter is also used.
        [switch]$PassThru,

        #OAuth or Bearer Token to use instead of acquired token
        [securestring]$Token,

        #Add headers to Request Header collection without validation
        [switch]$SkipHeaderValidation,

        #Body Content Type, for exmaple 'application/json'
        [string]$ContentType,

        #Graph Authentication Type - default or userProvived Token
        [Microsoft.Graph.PowerShell.Authentication.Models.GraphRequestAuthenticationType]
        $Authentication,

        #Specifies a web request session. Enter the variable name, including the dollar sign ($).You can''t use the SessionVariable and GraphRequestSession parameters in the same command.
        [Alias('SV')]
        [string]$SessionVariable,

        [Alias('RHV')]
        [string]$ResponseHeadersVariable,

        [string]$StatusCodeVariable,

        [switch]$SkipHttpErrorCheck,

        #If specified returns the .values property instead of the whole JSON object returned by the API call
        [switch]$ValueOnly,

        #If specified, loops through multi-paged results indicated by an '@odata.nextLink' property
        [switch]$AllValues,

        #If specified removes properties found in the JSON before converting to a type or returning the object
        [string[]]$ExcludeProperty,

        #If specified converts the JSON object to properties of the a new object of the requested type. Any properties which are expected in the JSON but not defined in the type should be excluded.
        [string]$AsType
    )

    begin {
        [void]$PSBoundParameters.Remove('AllValues')
        [void]$PSBoundParameters.Remove('AsType')
        [void]$PSBoundParameters.Remove('ExcludeProperty')
        [void]$PSBoundParameters.Remove('ValueOnly')
        if ($GLOBAL:__MgAzTokenExpires -and $GLOBAL:__MgAzTokenExpires -lt [DateTime]::Now) {
            if ($script:RefreshParams) {
                Write-Host -ForegroundColor DarkCyan "Token Expired! Refreshing before executing command."
                Connect-Graph @script:RefreshParams
            }
            else {Write-Warning "Token appears to have expired and no refresh information is available "}
        }
        elseif ($global:__MgToken -is [secureString]) {
            $PSBoundParameters['Token']          = $global:__MgToken
            $PSBoundParameters['Authentication'] = 'UserProvidedToken'
            Write-Debug "Using user-provided token"
        }
    }

    process {
        #my variable naming: response is an answer to a call - a partial thing
        #Result might come from processing a response or multiple responses - the end goal.
        $response = Microsoft.Graph.Authentication\Invoke-MgGraphRequest @PSBoundParameters
        if ($ValueOnly -or $AllValues) {
            $result = $response.value
            if ($AllValues) {
                while ($response.'@odata.nextLink') {
                    $PSBoundParameters['Uri'] =  $response.'@odata.nextLink'
                    $response   =   Microsoft.Graph.Authentication\Invoke-MgGraphRequest @PSBoundParameters
                    $result += $response.value
                }
            }
        }
        else  {$result = $response}
        if ($StatusCodeVariable) {Set-variable $StatusCodeVariable -Scope 1 -Value (Get-Variable $StatusCodeVariable -ValueOnly) }
        foreach ($r in $result) {
            foreach ($p in $ExcludeProperty) {$r.remove($p)}
            if ($AsType) {New-Object -TypeName $AsType -Property $r}
            else         {$r}
        }
    }
}

Remove-item Alias:\Connect-Graph -ErrorAction SilentlyContinue
Function Connect-Graph      {
    <#
        .Synopsis
            Starts a session with Microsoft Graph
        .Description
            This commands is a wrapper for Connect-MgGraph it extends the authentication methods available
            and caches information needed by other commands.
    #>
    [cmdletbinding(DefaultParameterSetName='UserParameterSet')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]
    param (
        [Parameter(ParameterSetName = 'UserParameterSet', Position = 1 )]
        #An array of delegated permissions to consent to.
        [string[]]$Scopes = $global:DefaultGraphScopes,

        #Specifies a bearer token for Microsoft Graph service. Access tokens do timeout and you'll have to handle their refresh.
        [Parameter(ParameterSetName = 'AccessTokenParameterSet', Position = 1, Mandatory = $true)]
        [string]$AccessToken,

        #A credential object to logon with an app registered in the tennant
        [Parameter(ParameterSetName = 'CredParameterSet', Position = 1, Mandatory = $true )]
        [pscredential]$Credential,

        #logon using an existing Azure session.
        [Parameter(ParameterSetName = 'AzureParameterSet', Position = 1, Mandatory = $true )]
        [switch]$FromAzureSession,

        #Refresh a the token obtained with a crednetial or Azure session logon.
        [Parameter(ParameterSetName = 'RefreshParameterSet', Position = 1, Mandatory = $true )]
        [switch]$Refresh,

        #The name of your certificate. The Certificate will be retrieved from the current user's certificate store.
        [Parameter(ParameterSetName = 'AppParameterSet', Position = 1 )]
        [Alias('CertificateSubject')]
        $CertificateName ,

        #The thumbprint of your certificate. The Certificate will be retrieved from the current user's certificate store.
        [Parameter(ParameterSetName = 'AppParameterSet', Position = 2 )]
        [String]$CertificateThumbprint ,

        #The ID for an application registered with your tennant when providing credentials, or logging on as the app.
        [Parameter(ParameterSetName = 'AppParameterSet', Position = 3  )]
        [Parameter(ParameterSetName = 'CredParameterSet', Position = 2 )]
        [alias('AppID')]
        [string]$ClientId = $Script:ClientID,

        #A secret associated with the application specified in -ClientID
        [Parameter(ParameterSetName = 'CredParameterSet', Position = 3 )]
        $ClientSecret = $Script:client_secret,

        #An x509 Certificate supplied during invocation - see https://docs.microsoft.com/en-us/graph/powershell/app-only? for configuring the host side.
        [Parameter(ParameterSetName = 'AppParameterSet' )]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        #The Az Module Context to use for the connection. You can get a list with Get-AzContext -ListAvailable. Note this parameter only accepts one context and if multiple are supplied it will only connect to the last one supplied
        [Parameter(ParameterSetName = 'AzureParameterSet', Position = 4)]
        $DefaultProfile = $GLOBAL:__MgAzContext,

        #The id of the tenant to connect to.
        [Parameter(ParameterSetName = 'AccessTokenParameterSet', Position = 4 )]
        [Parameter(ParameterSetName = 'AppParameterSet',         Position = 4 )]
        [Parameter(ParameterSetName = 'CredParameterSet',        Position = 4 )]
        [Parameter(ParameterSetName = 'UserParameterSet',        Position = 4 )]
        [string]$TenantId     = $script:Tenant,

        #Forces the command to get a new access token silently.
        [switch]$ForceRefresh ,

        #Determines the scope of authentication context. This accepts `Process` for the current process, or `CurrentUser` for all sessions started by user.
      #  [ContextScope]$ContextScope,

        #The name of the national cloud environment to connect to. By default global cloud is used.
        [Alias('EnvironmentName', 'NationalCloud')]
        [string]$Environment,

        #Dont register the global refresh handler workaround. This is required if you want to use HttpPipelinePrepend
        [Parameter(ParameterSetName = 'CredParameterSet' )]
        [Parameter(ParameterSetName = 'AzureParameterSet' )]
        [Switch]$NoRefresh,

        #Suppress the welcome messages
        [Switch]$Quiet
    )

    if ($GLOBAL:__MgAzTokenExpires) {$GLOBAL:__MgAzTokenExpires = $null}
    if ($GLOBAL:__MgToken         ) {$GLOBAL:__MgToken          = $null}

    #region to get a token for a name / password with a registerd appID and secret, or to refresh one
    if (($Refresh -or $Credential)  ){
        if (-not ($TenantId -and $ClientId -and $ClientSecret)) {
            Write-Warning "This form of logon needs a client ID and secret and a Tenant ID, and they have not been set." ; return
        }
        $tokenUri   = "https://login.microsoft.com/$TenantId/oauth2/token"
        # Send request either with grant type of password and creds, or grant type
        if ($Refresh -and -not $script:RefreshToken) {
            Write-Warning "No session to refresh" ; return
        }
        elseif ($Refresh) {
            $authresp   =   Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                'grant_type'    = 'refresh_token'
                'refresh_token' = $Script:RefreshToken
                'client_id'     = $ClientId
                'client_secret' = $ClientSecret
                'resource'      = 'https://graph.microsoft.com'
            }
        }
        else {
            $authresp   =   Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                'grant_type'    = 'password'
                'resource'      = 'https://graph.microsoft.com'
                'username'      = $Credential.UserName
                'password'      = $Credential.GetNetworkCredential().Password
                'client_id'     = $ClientId
                'client_secret' = $ClientSecret
            }
        }
        if ($authresp.access_token) {
            $null = $PSBoundParameters.Remove("ClientID")
            $null = $PSBoundParameters.Remove("Credential")
            $null = $PSBoundParameters.Remove("ClientSecret")
            $null = $PSBoundParameters.Remove("Refresh")
            $null = $PSBoundParameters.Add("AccessToken",  $authresp.access_token)
            if ($authresp.expires_in)    {$Global:__MgAzTokenExpires = (Get-Date).AddSeconds([int]$authresp.expires_in -60 )}
            if ($authresp.refresh_token) {$script:RefreshToken  = $authresp.refresh_token}
            if ($NoRefresh)              {$script:RefreshParams = $null}
            else                         {$script:RefreshParams = @{Quiet=$true; Refresh=$True}}
        }
        else {throw "No Token was returned"}
    }
    #endregion
    #region to get a token for an existing Azure Session.
    elseif ($FromAzureSession) {
        if (-not (Get-Command -Name "Get-AzAccessToken")) {
            Write-Warning "The required Azure tools are not available."
            return
        }
        if ($DefaultProfile) {
               $Global:__MgAzContext = $DefaultProfile
               $authresp =  Get-AzAccessToken -DefaultProfile $DefaultProfile -ResourceUrl 'https://graph.microsoft.com'
        }
        else { $authresp = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com'}
        if ($authresp.Token){
            $null = $PSBoundParameters.Remove("$FromAzureSession")
            $null = $PSBoundParameters.Remove("DefaultProfile")
            $null = $PSBoundParameters.Add("AccessToken",  $authresp.Token)
            $Global:__MgAzTokenExpires = $authresp.ExpiresOn
            if ($NoRefresh)              {$script:RefreshParams = $null}
            else                         {$script:RefreshParams = @{Quiet=$true; FromAzureSession=$True}}
        }
        else {throw "No Token was returned"}
    }
    #endregion

    #now connect, either using a token - passed or just fetched, or using a cert, or using the device dialog, whichever  as needed

    #region collect parameters and call Connect-MGGraph - psboundParameters won't work here because default values aren't *bound*
    $paramsinTarget       = (Get-Command Connect-MgGraph).Parameters.Keys |
                                Where-Object {$_ -notin [System.Management.Automation.Cmdlet]::CommonParameters}
    $paramsFromCurrentSet = $pscmdlet.MyInvocation.MyCommand.ParameterSets.Where({$_.name -eq $PSCmdlet.ParameterSetName})
    $paramsFromCurrentSet = $paramsFromCurrentSet.parameters.Name | Where-Object {$_ -in $paramsinTarget -and (Get-Variable $_ -ValueOnly -ErrorAction SilentlyContinue)}
    $paramsToPass         = @{}
    foreach ($p in $paramsFromCurrentSet ) {$paramsToPass[$p] = Get-Variable $P -ValueOnly}
    foreach ($p in [System.Management.Automation.Cmdlet]::CommonParameters.Where({$PSBoundParameters.ContainsKey($_)})) {
        $paramsToPass[$p] = $PSBoundParameters[$p]
    }
    $result = Connect-MgGraph @paramsToPass
    #endregion
    #region if succesful cache information about the user and session, and if necessary setup a trigger to auto-refresh tokens we fetched above
    if ($result-match "Welcome To Microsoft Graph") {
        $authcontext      = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext
        $result           = "Welcome To Microsoft Graph, $($authcontext.Account)."
        #we could call Get-Mgorganization but this way we don't depend on anything outside authentication module
        $Organization     = Invoke-GraphRequest -Method GET -Uri "$GraphURI/organization/" -ValueOnly
        if ($Organization.id) {
            Write-Verbose -Message "CONNECT: Account is from $($Organization.DisplayName)"
            Add-Member -force -InputObject $authcontext -NotePropertyName TenantName          -NotePropertyValue $Organization.DisplayName
            Add-Member -force -InputObject $authcontext -NotePropertyName WorkOrSchool        -NotePropertyValue $true
        }
        else                  {
            Write-Verbose -Message "CONNECT: Account is from Windows live"
            Add-Member -force -InputObject $authcontext -NotePropertyName TenantName          -NotePropertyValue $Organization.DisplayName
            Add-Member -force -InputObject $authcontext -NotePropertyName WorkOrSchool        -NotePropertyValue $true
        }
        $user             =   Invoke-MgGraphRequest -Method GET -Uri "$GraphURI/me/"
        $Global:GraphUser =  $user.userPrincipalName
        Add-Member -Force -InputObject $authcontext -NotePropertyName UserDisplayName        -NotePropertyValue $user.displayName
        Add-Member -Force -InputObject $authcontext -NotePropertyName UserID                 -NotePropertyValue $user.ID
        Add-Member -Force -InputObject $authcontext -NotePropertyName RefreshTokenPresent    -NotePropertyValue ($script:RefreshToken -as [bool])
        Add-Member -Force -InputObject $authcontext -NotePropertyName TokenAutoRefresh       -NotePropertyValue ($RefreshScriptBlock  -as [bool])
        if    ($Global:__MgAzTokenExpires) {
            Add-Member -Force -InputObject $authcontext -NotePropertyName TokenExpires       -NotePropertyValue ($Global:__MgAzTokenExpires)
        }
        elseif ($authcontext.TokenExpires) {$authcontext.TokenExpires = $null}
        if ($RefreshScriptBlock -and -not $GLOBAL:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend']) {
            if (-not $Quiet){
                Write-Host -Fore DarkCyan "The HttpPipelinePrepend parameter now has a default that checks for refresh tokens. Any command which uses this parameter will lose the auto-refresh"
            }
            $GLOBAL:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend'] = $RefreshScriptBlock
        }
        if ($NoRefresh -and -not $Quiet) {
            Write-Host -Fore Cyan "-NoRefresh was specified. You will need to run this command again after $($tokeninfo.ExpiresOn.LocalDateTime.ToString())"
        }
    }
    #endregion

    if (-not $Quiet) {return $result}
}

Function Show-GraphSession  {
    <#
        .Synopsis
            Returns Basic information about the current sesssion
    #>
    [CmdletBinding(DefaultParameterSetName='None')]
    [OutputType([String])]
    Param(
        [Parameter(ParameterSetName='Who')]
        [switch]$Who,
        [Parameter(ParameterSetName='Scopes')]
        [switch]$Scopes
    )
    if  ($Scopes) {[GraphSession]::Instance.AuthContext.Scopes}
    elseif ($Who) {[GraphSession]::Instance.AuthContext.Account}
    else          {Get-MgContext}
}

Function ContextHas         {
    <#
        .Syopsis
            Checks if the the current context is a work/school account and/or has access with the right scopes
    #>
    [cmdletbinding()]
    param (
        #list of scopes. will return true if at least one IS present.
        [string[]]$scopes,
        #if specifies returns ture for a work-or-school account and false for "Live" accounts
        [switch]$WorkOrSchoolAccount,
        #If specified break instead of turning false
        [switch]$BreakIfNot,
        #If specified reverses the output.
        [switch]$Not
    )
    if ($WorkOrSchoolAccount)  {
          $state =  [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.WorkOrSchool
    }
    elseif ($scopes) {
        $state =  $false
        foreach ($s in $scopes)  {
            $state = $state -or ([Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.Scopes -contains $s)
        }
    }
    if ($BreakIfNot ) {
        if ($scopes              -and -not $state) {Write-Warning "This requires the $($scopes -join ', ') scope(s)." ; break}
        if ($WorkOrSchoolAccount -and -not $state) {Write-Warning "This requires a work or school account."           ; break}
    }
    #otherwise return true or false
    else  {return ( $state -xor $not )}
}
