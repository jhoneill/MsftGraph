#Requires -Module Microsoft.Graph.Authentication
<#
    The Connect-Graph function incorporates work to get tokens from an Azure
    session and to referesh tokens which was published by Justin Grote at
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Connect-MgGraphAz.ps1
    and licensed by him under the same MIT terms which apply to this module (see the LICENSE file for details)

    Portions of this file are   Copyright 2021 Justin Grote @justinwgrote
    The remainder is Copyright 2018-2021 James O'Neill
#>
using namespace Microsoft.Graph.PowerShell.Authentication
using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Management.Automation

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write host used for colored information message telling user to make a change and remove the message')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Items needed outside the module')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]
param()

#Functions we will define are aliases in Microsoft.Graph.Authentication so remove them
Remove-item Alias:\Invoke-GraphRequest -ErrorAction SilentlyContinue
Remove-Item Alias:\Connect-Graph       -ErrorAction SilentlyContinue
Set-Alias   -Value Get-MgContext       -Name "Get-GraphContext"

function Test-GraphSession          {
<#
    .synopsis
        Connect if necessary, catch tokens needing renewal.
#>
    param ( [switch]$Quiet )

    if (-not [GraphSession]::Instance.AuthContext) {Connect-Graph -Quiet:$Quiet | Out-Host}
    elseif  ([GraphSession]::Instance.AuthContext.TokenExpires -is [datetime] -and
             [GraphSession]::Instance.AuthContext.TokenExpires -lt [datetime]::Now.AddMinutes(-1)) {
        if ($Script:RefreshParams) {
            if (-not $Quiet) { Write-Host -ForegroundColor DarkCyan "Token Expired! Refreshing before executing command."}
            Connect-Graph @script:RefreshParams
        }
        else {Write-Warning "Token appears to have expired and no refresh information is available "}
    }
}

function Invoke-GraphRequest        {
    <#
      .synopsis
        Wrappper for Invoke-MgGraphRequest.With token management and result pre-processing
      .description
        Adds -ValueOnly to return just the value part
             -AllValues to return gather multiple sets when data is paged
             -AsType to convert the retuned results to a specific type
             -ExcludeProperty  and -PropertyNotMatch for results which have properties which aren't in the specified type
    #>
    [alias('igr')]
    param   (
        #Uri to call can be a segment such as /beta/me or a fully qualified https://graph.microsoft.com/beta/me
        [Parameter(Mandatory=$true, Position=1 )]
        [uri]$Uri,

        #Http Method
        [ValidateSet('GET','POST','PUT','PATCH','DELETE')]
        [Parameter(Position=2 )]
        $Method,

        #Request body, required when Method is POST or PATCH
        [Parameter(Position=3,ValueFromPipeline=$true)]
        $Body,

        #Optional custom headers, commonly @{'ConsistencyLevel'='eventual'}
        [System.Collections.IDictionary]$Headers,

        #Output file where the response body will be saved
        [string]$OutputFilePath,

        [switch]$InferOutputFileName,

        #Input file to send in the request
        [string]$InputFilePath,

        #Indicates that the cmdlet returns the results, in addition to writing them to a file. Only valid when the OutFile parameter is also used.
        [switch]$PassThru,

        #OAuth or Bearer token to use instead of acquired token
        [securestring]$Token,

        #Add headers to request header collection without validation
        [switch]$SkipHeaderValidation,

        #Body content type, for exmaple 'application/json'
        [string]$ContentType,

        #Graph Authentication type - 'default' or 'userProvivedToken'
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
        [string[]]$ExcludeProperty =@(),

        #A regular expression for keys to be removed, for example to catch many odata properties
        [string]$PropertyNotMatch,

        #If specified converts the JSON object to properties of the a new object of the requested type. Any properties which are expected in the JSON but not defined in the type should be excluded.
        [string]$AsType
    )
    begin   {
        $null = $PSBoundParameters.Remove('AllValues') , $PSBoundParameters.Remove('AsType'), $PSBoundParameters.Remove('ExcludeProperty'), $PSBoundParameters.Remove('PropertyNotMatch') , $PSBoundParameters.Remove('ValueOnly')
        if ($ExcludeProperty -notcontains  '@odata.id') {$ExcludeProperty += '@odata.id'}
        Test-GraphSession
    }
    process {
        #I try to use "response" when it is an interim thing not the final result.
        $response = $null
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
            if ($PropertyNotMatch) {
                $keystoRemove = $r.keys -match $PropertyNotMatch
                foreach ($p in $keystoRemove) {[void]$r.remove($p)}
            }
            if ($AsType) {New-Object -TypeName $AsType -Property $r}
            else         {$r}
        }
    }
}

function Get-AccessToken            {
<#
  .Synopsis
    Requests a token for a resource path, used by connect graph but available to other tools.
  .Description
    An access token is obtained form "https://login.microsoft.com/<<tenant-ID>>/oauth2/token"
    By specifying the ID, and secret of a client app known to in that tenant,
    different modes of granting a token to access a resource (logging on) are possible:
    Extra fields passed in BodyParts    grant_type
    * Username and password            'Password'
    * Refresh_token                    'Referesh_token'
    * None (logon as the app itself)   'client_credentials'
    The Set-GraphOptions command sets the tenant ID, a client ID and a client secret
    for the session.  By default, when the module loads it looks at $env:GraphSettingsPath or for
    Microsoft.Graph.PlusPlus.settings.ps1  in the module folder, and executes it to set these values)
    Get-AccessToken relies on these if they are not set Connect-Graph removes the parameters
    which support non-intereactive logons and calling it seperately will fail

#>
    param (
        [string]$Resoure      = 'https://graph.microsoft.com',
        [Parameter(Mandatory=$true)]
        [string]$GrantType ,
        [hashtable]$BodyParts = @{}
    )
    if (-not ($Script:TenantID -and $Script:ClientID)) {
        [UnauthorizedAccessException]::new('The tenant, and ClientID need to be set with Set-GraphOptions before calling this command.')
    }
    $tokenUri  = "https://login.microsoft.com/$Script:TenantID/oauth2/token"
    $body      = $BodyParts + @{
                    'client_id'     = $Script:ClientID
                    'resource'      = $Resoure
                    'grant_type'    = $GrantType
    }
    Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body
}

function Get-AzureIdentityToken {
    [CmdletBinding(DefaultParameterSetName='scopes')]
    param (
        [Parameter(ParameterSetName = 'scopes')]
        [String[]]$Scopes = 'https://graph.microsoft.com/',
        [Parameter(ParameterSetName = 'tokenRequest', Mandatory)]
        [TokenRequestContext]$TokenRequestContext,
        [Switch]$Interactive,
        [ValidateSet('AzureCLI','Default','DeviceCode','InteractiveBrowser')]
        [string]$Type = 'Default'
    )
    end {
        if ($Type -in @('InteractiveBrowser','DeviceCode')) {
           $scopes = @('User.ReadWrite.all', 'openid', 'profile')
        }
        if (!$TokenRequestContext) {$TokenRequestContext = [Azure.Core.TokenRequestContext]::new($Scopes)}

        switch ($type) {
           "AzureCLI"           {$tokenCache = [Azure.Identity.AzureCliCredential]::new()}   #returns scopes  AuditLog.Read.All Directory.AccessAsUser.All Group.ReadWrite.All User.ReadWrite.All / client 04b07795-8ddb-461a-bbee-02f9e1bf7b46 : Microsoft Azure CLI
                                #.AuthorizationCodeCredential, would take $Script:TenantID, $Script:ClientID,  $Script:ClientSecret + AuthCode;  can't get AzurePowerShellCredential to work
                                #.ChainedTokenCredential takes array of tokencredentials, ClientCertificateCredential takes  $Script:TenantID, $Script:ClientID  cert
                                # ClientSecretCredential takes $Script:TenantID, $Script:ClientID,  $Script:ClientSecret' EnvironmentCredential uses env variables
           "Default"            {$tokenCache = [Azure.Identity.DefaultAzureCredential]::new($Interactive)}  # returns scopes  Application.ReadWrite.All email openid profile User.ReadWrite.All / client 872cd9fa-d31f-45e0-9eab-6e460a02d1f1 :    : Visual Studio
           "DeviceCode"         {$tokenCache = [Azure.Identity.DeviceCodeCredential]::new()}                # returns Scopes  AuditLog.Read.All Directory.AccessAsUser.All email Group.ReadWrite.All openid profile User.ReadWrite.All  / client 04b07795-8ddb-461a-bbee-02f9e1bf7b46 : Microsoft Azure CLI
           "InteractiveBrowser" {$tokenCache = [Azure.Identity.InteractiveBrowserCredential]::new()}        # AuditLog.Read.All Directory.AccessAsUser.All email Group.ReadWrite.All openid profile User.ReadWrite.All  / client 04b07795-8ddb-461a-bbee-02f9e1bf7b46 : Microsoft Azure CLI
                                        #ManagedIdentityCredential takes clientID ;SharedTokenCacheCredential doesn't need anything (but is empty for me!)
                                        #.UsernamePasswordCredential string username, string password, string tenantId, string clientId
        }
        $tokenCache.GetToken($TokenRequestContext,[System.Threading.CancellationToken]::new($false))
    }
}

function Connect-Graph              {
    <#
        .Synopsis
            Starts a session with Microsoft Graph
        .Description
            This commands is a wrapper for Connect-MgGraph it extends the authentication methods available
            and caches information needed by other commands.

    #>
    [cmdletbinding(DefaultParameterSetName='UserParameterSet')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]
    [Alias('New-GraphSession','GraphSession')]
    param        (
        [Parameter(ParameterSetName = 'UserParameterSet', Position = 1 )]
        #An array of delegated permissions to consent to.
        [string[]]$Scopes = $Script:DefaultGraphScopes,

        #Specifies a bearer token for Microsoft Graph service. Access tokens do timeout and you'll have to handle their refresh.
        [Parameter(ParameterSetName = 'AccessTokenParameterSet', Position = 1, Mandatory = $true)]
        [string]$AccessToken,

        #Forces the command to get a new access token silently.
        [switch]$ForceRefresh ,

        #Determines the scope of authentication context. This accepts `Process` for the current process, or `CurrentUser` for all sessions started by user.
      #  [ContextScope]$ContextScope,

        #Suppress the welcome messages
        [Switch]$Quiet
    )
    dynamicParam {
    <#
        If the Azure commands are present offer -FromAzureToken
        If ClientSecret, ClientID and TenantID have all been set, offer -Credential & -AsApp and if a refresh token was stored, -refresh
        In either of those cases offer -NoRefresh
        If client ID and TenantID have been set (with or without secret) offer the cert parameters.
    #>
        $paramDictionary     = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $NoRefreshAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        if ($Script:ClientSecret -and $Script:ClientID -and $Script:TenantID) {
            $AppParamAttribute     = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='AppSecretParameterSet';}
            #If Specified logs in as the app and gets the access granted to the app instead of logging on as a user.
            $paramDictionary.Add('AsApp',[RuntimeDefinedParameter]::new("AsApp",       [SwitchParameter],$AppParamAttribute))
        }
        if ($Script:ClientID -and $Script:TenantID) {
            $NoRefreshAttributeCollection.Add((New-Object System.Management.Automation.ParameterAttribute -Property @{       ParameterSetName='CredParameterSet'}))
            $CredParamAttribute      = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='CredParameterSet'}
            $RefreshParamAttribute   = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='RefreshParameterSet'}
            #A credential object to logon with an app registered in the tennant
            $paramDictionary.Add('Credential',[RuntimeDefinedParameter]::new("Credential", [pscredential], $CredParamAttribute))
            if ($Script:RefreshToken) {
               $paramDictionary.Add('Refresh',[RuntimeDefinedParameter]::new("Refresh", [SwitchParameter], $RefreshParamAttribute))
            }

            $CertNameAttributeSet    = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $CertNameAttributeSet.add((New-Object System.Management.Automation.Aliasattribute     -ArgumentList 'CertificateSubject'))
            $CertNameAttributeSet.add((New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AppCertNameParameterSet'; Position=1}))
            $CertThumbParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AppCertThunbParameterSet';Position=2}
            $CertParamAttribute      = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AppCertParameterSet'}
            #The name of your certificate. The Certificate will be retrieved from the current user's certificate store.
            $paramDictionary.Add('CertificateName',      [RuntimeDefinedParameter]::new('CertificateName',  [string],   $CertNameAttributeSet))
            #The thumbprint of your certificate. The Certificate will be retrieved from the current user's certificate store.
            $paramDictionary.Add('CertificateThumbprint',[RuntimeDefinedParameter]::new('CertificateThumbprint',  [string],   $CertThumbParamAttribute))
            #An x509 Certificate supplied during invocation - see https://docs.microsoft.com/en-us/graph/powershell/app-only? for configuring the host side.
            $paramDictionary.Add('Certificate',[RuntimeDefinedParameter]::new('Certificate',  [System.Security.Cryptography.X509Certificates.X509Certificate2],   $CertParamAttribute))
        }
        #the simpler  Get-command Get-AzAccessToken -ErrorAction SilentlyContinue loads az.accounts. Only offer the parameter if the module is loaded.
        if ("Get-AzAccessToken" -in (Get-Module Az.accounts).ExportedCmdlets.Keys) {
            $NoRefreshAttributeCollection.Add((New-Object System.Management.Automation.ParameterAttribute -Property @{       ParameterSetName='AzureParameterSet'}))
            $FromAzParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AzureParameterSet';Position=3}
            $DefProfParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AzureParameterSet';Position=4}
            #The Az Module Context to use for the connection. You can get a list with Get-AzContext -ListAvailable. Note this parameter only accepts one context and if multiple are supplied it will only connect to the last one supplied
            $paramDictionary.Add('FromAzureSession',[RuntimeDefinedParameter]::new('FromAzureSession', [SwitchParameter], $FromAzParamAttribute))
            $paramDictionary.Add('DefaultProfile',  [RuntimeDefinedParameter]::new('DefaultProfile',   [System.Object],   $DefProfParamAttribute))
        }
        if (Get-Command az) {
            $NoRefreshAttributeCollection.Add((New-Object System.Management.Automation.ParameterAttribute -Property @{       ParameterSetName='CLIParameterSet'}))
            $FromCLIParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='CLIParameterSet';Position=5}
            $paramDictionary.Add('FromAzCLI',[RuntimeDefinedParameter]::new('FromAzCLI', [SwitchParameter], $FromCLIParamAttribute))
        }
        if ($NoRefreshAttributeCollection.Count -ge 1) {
            #Dont register the global refresh handler workaround. This is required if you want to use HttpPipelinePrepend
            $paramDictionary.Add('NoRefresh', [RuntimeDefinedParameter]::new("NoRefresh", [SwitchParameter],$NoRefreshAttributeCollection))
        }
        return $paramDictionary
    }
    begin        {
    #Justin Grote's code makes this script block to call this function with the right parameters. If the token has timed out
    #and makes it the default httpPipelinePrepend script block for SDK commands.  We'll poke the parameters in later
    $RefreshScript = @'
    param ($req, $callback, $next)
    Write-Debug 'Checking if access token refresh required'
    #Check if global timer has expired and refresh if so
    if ([Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.TokenExpires -is [datetime] -and
        [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext.TokenExpires -lt [DateTime]::Now) {{
            Write-Host -Fore DarkCyan "Token Expired! Refreshing before executing command."
            Connect-Graph -quiet {0}
    }}
'@  }
    end          {
        $result = $null
        $bp = @{} + $PSBoundParameters #I do not know why psb doesn't work normally with dynamic params but this works round it....
        $paramsToPass         = @{}

        #Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
        if ((-not $bp.refresh) -or
            (-not $Global:DriveCache))  {$Global:DriveCache          = @{}  }

        #Justin used this variable, for checking expiry I have moved it on to an extra property of [graphSession].instance I'll remove this when I'm happy on compat.
        if ($Global:__MgAzTokenExpires) {$Global:__MgAzTokenExpires = $null}

        #credential , refresh, Azaupp, FromAzureSession are dynamic to hide them if we dont have what they need
        #If specified, get (or refresh) a token for a user name / password with a registerd appID in a known tennant,
        #or for an app-id / secret, or by calling Get-AzAccessToken in the Az.Accounts module (V2 and later)
        if ($bp.Credential -or $bp.Refresh -or $bp.AsApp -or $bp.FromAzureSession -or $bp.FromAzCLI ){
            $tokenUri   = "https://login.microsoft.com/$Script:TenantID/oauth2/token"
            if     ($bp.Refresh)              {
                Write-Verbose "CONNECT: Sending a 'Refresh_token' token request "
                $authresp      = Get-AccessToken -GrantType refresh_token -BodyParts @{'refresh_token' = $Script:RefreshToken}
            }
            elseif ($bp.Credential)           {
                Write-Verbose "CONNECT: Sending a 'Password' token request for $($bp.Credential.UserName) "
                $parts = @{ 'username' = $bp.Credential.UserName;
                            'password' = $bp.Credential.GetNetworkCredential().Password
                }
                if ($Script:ClientSecret) {
                    $parts['client_secret'] = $Script:ClientSecret
                }
                $authresp     = Get-AccessToken -GrantType password -BodyParts $parts
            }
            elseif ($bp.AsApp)                {
                Write-Verbose "CONNECT: Sending a 'client_credentials' token request for the App."
                $authresp     = Get-AccessToken  -GrantType client_credentials -BodyParts @{'client_secret' = $Script:ClientSecret}
            }
            elseif ($bp.FromAzureSession)     {
                Write-Verbose "CONNECT: getting an access token from an Azure PowerShell session."
                if ($bp.DefaultProfile) {$Global:__MgAzContext = $DefaultProfile }
                $authresp =  Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com' -DefaultProfile $Global:__MgAzContext
            }
            elseif ($bp.FromAzCLI)     {
                Write-Verbose "CONNECT: getting an access token from the Azure CLI."
                $authresp = az account get-access-token  --resource-type ms-graph  --output json | ConvertFrom-Json
            }

            #Did we get our token? If so, store it and whatever we need to refreshing
            if (-not ($authresp.access_token -or $authresp.accesstoken -or $authResp.Token))    {
                throw [System.UnauthorizedAccessException]::new("No Token was returned")
            }
            Write-Verbose ("CONNECT: Token Response= " + ($authresp | Get-Member -MemberType NoteProperty).name -join ", ")
            if       ($authresp.access_token) {
                    $null = $paramsToPass.Add("AccessToken",  $authresp.access_token)
            }
            elseif   ($authresp.accesstoken)  {
                    $null = $paramsToPass.Add("AccessToken",  $authresp.accesstoken)
            }
            elseif   ($authresp.Token)        {
                    $null = $paramsToPass.Add("AccessToken",  $authresp.Token)
            }

            if     ($authresp.scope)          {Write-Verbose "CONNECT: Scope= $($authresp.scope)" }
            if     ($authresp.refresh_token)  {$Script:RefreshToken       = $authresp.refresh_token}
            if     ($authresp.expires_in)     {$Global:__MgAzTokenExpires = (Get-Date).AddSeconds([int]$authresp.expires_in -60 )}
            elseif ($authresp.expires_on -or  $authresp.expireson )    {
                if ($authresp.expires_on) {$e = $authresp.expires_on} else {$e = $authresp.ExpiresOn}
                if     ($e -is [string] -and
                        $e -match "^\w{10}$")     {$Global:__MgAzTokenExpires = [datetime]::UnixEpoch.AddSeconds($e)}
                elseif ($e -is [string] )         {$Global:__MgAzTokenExpires = [datetime]::$e   }
                elseif ($e  -is [datetimeoffset]) {$Global:__MgAzTokenExpires = $e.LocalDateTime }
            }

            if     ($bp.NoRefresh)            {
                $Script:RefreshParams      = $null
            }
            elseif ($bp.FromAzureSession)     {
                $Script:RefreshParams      = @{'Quiet' = $true; 'FromAzureSession' = $True}
                $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' -FromAzureSession '))
            }
            elseif ($bp.FromAzCLI)           {
                $Script:RefreshParams      = @{'Quiet' = $true; 'FromAzCLI' = $true}
                $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' FromAzCLI '))
            }
            elseif ($bp.AsApp)                {
                $Script:RefreshParams      = @{'Quiet' = $true; 'AsApp' = $true}
                $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' -Refresh '))
            }

            elseif ($Script:RefreshToken)     {
                $Script:RefreshParams      = @{'Quiet' = $true; 'Refresh' = $true}
                $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' -Refresh '))
            }
        }

        #region gather params and call Connect-MGGraph with a token (passed or just fetched), or with a cert, or opening the device dialog
        #common parameters need to processed differently
        $paramsinTarget       = (Get-Command Connect-MgGraph).Parameters.Keys |
                                    Where-Object {$_ -notin [System.Management.Automation.Cmdlet]::CommonParameters}
        $paramsFromCurrentSet =  $pscmdlet.MyInvocation.MyCommand.Parameters.values.where({
                                    ($_.ParameterSets.containskey($pscmdlet.ParameterSetName) -or
                                     $_.ParameterSets.containskey('__AllParameterSets')     ) -and
                                     $_.Name -in $paramsinTarget -and
                                     (Get-Variable $_.Name -ValueOnly -ErrorAction SilentlyContinue)})

        foreach ($p in $paramsFromCurrentSet.Name ) {
            $paramsToPass[$p] = Get-Variable $P -ValueOnly   ; Write-Verbose ("{0,20} = {1}" -f $p.ToUpper(), $paramsToPass[$p])
        }
        foreach ($p in [System.Management.Automation.Cmdlet]::CommonParameters.Where({$bp.ContainsKey($_)})) {
            $paramsToPass[$p] = $bp[$p]                      ; Write-Verbose ("{0,20} = {1}" -f $p.ToUpper(), $paramsToPass[$p])
        }
        if ($Script:TenantID -and -not $paramsToPass.AccessToken) {
            $paramsToPass['TenantID'] = $Script:TenantID
        }
        if ($pscmdlet.ParameterSetName -match '^AppCert') {
            $paramsToPass['ClientId'] = $Script:ClientID
        }

        $result = Connect-MgGraph @paramsToPass
        #endregion
        #region connection succeeds, cache information about the user and session, and if necessary setup a trigger to auto-refresh tokens we fetched above
        if ($result -match "Welcome To Microsoft Graph") {
            $authcontext      = [GraphSession]::Instance.AuthContext
            #we could call Get-Mgorganization but this way we don't depend on anything outside authentication module
            $Organization     = Invoke-GraphRequest -Method GET -Uri "$GraphURI/organization/" -ValueOnly
            if ($Organization.id) {
                Write-Verbose -Message "CONNECT: Account is from $($Organization.DisplayName)"
                Add-Member -force -InputObject $authcontext -NotePropertyName TenantName          -NotePropertyValue $Organization.DisplayName
                Add-Member -force -InputObject $authcontext -NotePropertyName WorkOrSchool        -NotePropertyValue $true
            }
            else                  {
                Write-Verbose -Message "CONNECT: Account is from Windows live"
                Add-Member -force -InputObject $authcontext -NotePropertyName TenantName          -NotePropertyVa.lue $Organization.DisplayName
                Add-Member -force -InputObject $authcontext -NotePropertyName WorkOrSchool        -NotePropertyValue $true
            }
            if ($authcontext.Account) {
                $result           = 'Welcome To Microsoft Graph++, {0}.' -f $authcontext.Account
                $user             =   Invoke-MgGraphRequest -Method GET -Uri "$GraphURI/me/"
                if ($Global:GraphUser -and $Global:GraphUser -ne $user.userPrincipalName) {
                    Set-GraphOneNoteHome $null
                    Set-GraphHomeDrive   $null
                }
                $Global:GraphUser =  $user.userPrincipalName
                Add-Member -Force -InputObject $authcontext -NotePropertyName UserDisplayName     -NotePropertyValue $user.displayName
                Add-Member -Force -InputObject $authcontext -NotePropertyName UserID              -NotePropertyValue $user.ID
            }
            else {$result           = "Welcome To Microsoft Graph, connected to tenant '{1}' as the app '{0}'." -f $authcontext.AppName, $authcontext.TenantName }
            Add-Member     -Force -InputObject $authcontext -NotePropertyName RefreshTokenPresent -NotePropertyValue ($Script:RefreshToken -as [bool])
            Add-Member     -Force -InputObject $authcontext -NotePropertyName TokenAutoRefresh    -NotePropertyValue ($RefreshScriptBlock  -as [bool])
            if    ($Global:__MgAzTokenExpires) {
                Add-Member -Force -InputObject $authcontext -NotePropertyName TokenExpires         -NotePropertyValue ($Global:__MgAzTokenExpires)
            }
            elseif ($authcontext.TokenExpires) {$authcontext.TokenExpires = $null}
            if ($RefreshScriptBlock -and -not $Global:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend']) {
                if (-not $Quiet){
                    Write-Host -Fore DarkCyan "The HttpPipelinePrepend parameter now has a default that checks for refresh tokens. Any command which uses this parameter will lose the auto-refresh"
                }
                $Global:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend'] = $RefreshScriptBlock
            }
            if ($NoRefresh -and -not $Quiet) {
                Write-Host -Fore Cyan "-NoRefresh was specified. You will need to run this command again after $($tokeninfo.ExpiresOn.LocalDateTime.ToString())"
            }
        }
        #endregion

        if (-not $Quiet) {return $result}
    }
}

function Show-GraphSession          {
    <#
        .Synopsis
            Returns information about the current sesssion
    #>
    [CmdletBinding(DefaultParameterSetName='None')]
    [OutputType([String])]
    [alias('GWhoAmI')]
    param  (
        #If specified returns only the current account
        [Parameter(ParameterSetName='Who')]
        [switch]$Who,
        #If specified returns only the scopes available to the current session
        [Parameter(ParameterSetName='Scopes')]
        [switch]$Scopes,
        #If specified returns the options set using Set-GraphOption
        [Parameter(ParameterSetName='Options')]
        [switch]$Options,
        #If specified returns the current app name.
        [Parameter(ParameterSetName='AppName')]
        [switch]$AppName,
        #If specified runs Test-GraphSession to ensure a session exists.
        [switch]$Force
    )
    dynamicparam {
        $paramDictionary     = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSVersionTable.Platform -like 'win*') {
            $cachedTokenParamAttibute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='CachedToken'}
            $paramDictionary.Add('CachedToken', [RuntimeDefinedParameter]::new("CachedToken",  [SwitchParameter],$cachedTokenParamAttibute))
        }
        if ($Script:RefreshToken) {
            $refreshTokenParamAttibute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='RefreshToken'}
            $paramDictionary.Add('RefreshToken',[RuntimeDefinedParameter]::new("RefreshToken", [SwitchParameter],$refreshTokenParamAttibute))
        }
        return $paramDictionary
    }
    end {
        if     ($Force)   {Test-GraphSession -Quiet}
        if     (-not       [GraphSession]::Instance.AuthContext)  {Write-Host  "Ready for Connect-Graph."; return}
        if     ($Scopes)  {[GraphSession]::Instance.AuthContext.Scopes}
        elseif ($Who)     {[GraphSession]::Instance.AuthContext.Account}
        elseif ($AppName) {[GraphSession]::Instance.AuthContext.AppName}
        elseif ($Options) {[pscustomobject][Ordered]@{
            'TenantID'              = $Script:TenantID
            'ClientID'              = $Script:ClientID
            'ClientSecretSet'       = $Script:ClientSecret -as [bool]
            'DefaultScopes'         = $Script:DefaultGraphScopes -join ', '
            'DefaultUserProperties' = $Script:DefaultUserProperties -join ', '
            'DefaultUsageLocation'  = $Script:DefaultUsageLocation
        }}
        elseif (     $PSBoundParameters['RefreshToken']) {return $Script:RefreshToken}
        elseif (-not $PSBoundParameters['CachedToken'])  {
            if      ($Script:SkippedSubmodules) {
                Write-Host -ForegroundColor DarkGray ("Skipped " + ($Script:SkippedSubmodules -join ", ") + " because their Microsoft.Graph module(s) or private.dll file(s) were not found.")
            }
            if      ($Global:PSDefaultParameterValues['Get-GraphDrive:Drive'])           {
                Write-Host "Home OneDrive is set"
            }
            if      ($Global:PSDefaultParameterValues['*GraphOneNoteBook*:Notebook'] -and
                     $Global:PSDefaultParameterValues['Get-GraphOneNotePage:Section'])   {
                     Write-Host "Home Notebook and section are set"
            }
            elseif  ($Global:PSDefaultParameterValues['*GraphOneNoteBook*:Notebook'])    {
                     Write-Host "Home Notebook is set"
            }
            if      ($Global:PSDefaultParameterValues['Add-Graph*Tab:Team']) {
                     Write-Host "Default Group is set and team-enabled, default calendar is a team calendar."
            }
            elseif  ($Global:PSDefaultParameterValues['*-GraphEvent:Group']) {
                     Write-Host "Default Group is set and but not team-enabled, default calendar is a team calendar."
            }
            else   {Write-Host "Default Group is not set, default calendar is for the signed-in user."}
            Get-MgContext
        }
        else {
            $id               = [GraphSession]::Instance.AuthContext.ClientId
            $path             = Join-Path -ChildPath ".graph\$($id)cache.bin3" -Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))
            if (-not(Test-Path $path)) {
                    Write-Warning "Could not find a cache file for app id $id in the .graph folder."
                    return
            }
            #read and decrypted the ached file, it comes up as not very nice JSON so unpick that.
            $tokenBytes       = [System.Security.Cryptography.ProtectedData]::Unprotect( (Get-Content $path -AsByteStream) , $null, 0)
            $tokendata        = ConvertFrom-Json ([string]::new($tokenBytes)  -replace '(Token|Account|AppMetaData)":{".*?":{' ,'$1":{"X":{' -replace '"secret":".*?",','')
            $tokendata.account.x | Select-Object Name, Username, Local_account_id, Realm, Authority_type,
                                    @{n='environment';    e={$tokendata.AccessToken.x.environment}},
                                    @{n='client_id';      e={$tokendata.AccessToken.x.client_id}},
                                    @{n='credential_type';e={$tokendata.AccessToken.x.credential_type}},
                                    @{n='target';         e={$tokendata.AccessToken.x.target}},
                                    @{n='cached_at';      e={[datetime]::UnixEpoch.AddSeconds($tokendata.AccessToken.x.cached_at)}},
                                    @{n='expires_on';     e={[datetime]::UnixEpoch.AddSeconds($tokendata.AccessToken.x.expires_on)}}
        }
    }
}

function ContextHas                 {
    <#
        .Syopsis
            Checks if the the current context is a work/school account and/or has access with the right scopes
    #>
    [cmdletbinding()]
    param (
        #list of scopes. will return true if at least one IS present.
        [string[]]$Scopes,
        #if specifies returns ture for a work-or-school account and false for "Live" accounts
        [switch]$WorkOrSchoolAccount,
        #if specified returns ture if connected with a user account and false if connected as an application
        [switch]$AsUser,
        #if specified returns ture if connected as an application and false if connected with a user account
        [switch]$AsApp,
        #If specified break instead of turning false
        [switch]$BreakIfNot,
        #If specified reverses the output.
        [switch]$Not
    )
    if (-not [GraphSession]::Instance.AuthContext) { Connect-Graph | Out-Host}
    if (-not $Scopes) { $state = $true }
    else {
        $state =  $false
        foreach ($s in $Scopes)  {
            $state = $state -or ([GraphSession]::Instance.AuthContext.Scopes -contains $s)
        }
    }
    if ($WorkOrSchoolAccount)  {
        $state = $state -and [GraphSession]::Instance.AuthContext.WorkOrSchool
    }
    if ($AsUser)      {
        $state = $state -and [GraphSession]::Instance.AuthContext.Account
    }
    if ($AsApp)       {
        $state = $state -and -not [GraphSession]::Instance.AuthContext.Account
    }
    if ($BreakIfNot ) {
        if ($Scopes              -and -not $state) {Write-Warning ("This requires the {0} scope(s)." -f ($Scopes -join ', ')) ; break}
        if ($WorkOrSchoolAccount -and -not $state) {Write-Warning  "This requires a work or school account."                  ; break}
        if ($AsUser              -and -not $state) {Write-Warning  "This requires a user-logon, not an app-logon."            ; break}
        if ($AsApp               -and -not $state) {Write-Warning  "This requires an app-logon, not a user-logon."            ; break}
    }
    #otherwise return true or false
    else  {return ( $state -xor $not )}
}
