#Requires -Module Microsoft.Graph.Authentication
using namespace Microsoft.Graph.PowerShell.Authentication
using namespace Microsoft.Graph.PowerShell.Models
using namespace System.Management.Automation

<#
    The Connect-Graph function incorporates work to get tokens from an Azure session and to referesh tokens
    which was published by Justin Grote at
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Connect-MgGraphAz.ps1
    and licensed by him under the same MIT terms which apply to this module (see the LICENSE file for details)

    Portions of this file are   Copyright 2021 Justin Grote @justinwgrote
    The remainder is Copyright 2018-2021 James O'Neill
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write host used for colored information message telling user to make a change and remove the message')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Items needed outside the module')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]
param()

Remove-item Alias:\Invoke-GraphRequest -ErrorAction SilentlyContinue
function Invoke-GraphRequest {
    <#
      .synopsis
        Wrappper for Invoke-MgGraphRequest.With token management and result pre-processing
      .description
        Adds -ValueOnly to return just the value part
             -AllValues to return gather multiple sets when data is paged
             -AsType to convert the retuned results to a specific type
             -ExcludeProperty  and -PropertyNotMatch for results which have properties which aren't in the specified type
    #>
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
        [void]$PSBoundParameters.Remove('AllValues')
        [void]$PSBoundParameters.Remove('AsType')
        [void]$PSBoundParameters.Remove('ExcludeProperty')
        [void]$PSBoundParameters.Remove('PropertyNotMatch')
        [void]$PSBoundParameters.Remove('ValueOnly')
        if ([GraphSession]::Instance.AuthContext.TokenExpires -is [datetime] -and [GraphSession]::Instance.AuthContext.TokenExpires -lt [DateTime]::Now) {
            if ($script:RefreshParams) {
                Write-Host -ForegroundColor DarkCyan "Token Expired! Refreshing before executing command."
                Connect-Graph @script:RefreshParams
            }
            else {Write-Warning "Token appears to have expired and no refresh information is available "}
        }
    }
    process {
        #I try to use "response" when it is an interim thing not the final result.
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
            foreach ($p in $ExcludeProperty) {[void]$r.remove($p)}
            if ($PropertyNotMatch) {
                $keystoRemove = $r.keys -match $PropertyNotMatch
                foreach ($p in $keystoRemove) {[void]$r.remove($p)}
            }
            if ($AsType) {New-Object -TypeName $AsType -Property $r}
            else         {$r}
        }
    }
}

function Get-AccessToken     {
param (
    [string]$Resoure      = 'https://graph.microsoft.com',
    [string]$GrantType    = 'client_credentials',
    [hashtable]$BodyParts = @{}
)

$tokenUri  = "https://login.microsoft.com/$script:TenantID/oauth2/token"
$body      = $BodyParts + @{'client_id' = $script:ClientID
                        'client_secret' = $script:Client_secret
                             'resource' = $Resoure
                           'grant_type' = $GrantType
                              }

Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body

}

Remove-Item Alias:\Connect-Graph -ErrorAction SilentlyContinue
function Connect-Graph       {
    <#
        .Synopsis
            Starts a session with Microsoft Graph
        .Description
            This commands is a wrapper for Connect-MgGraph it extends the authentication methods available
            and caches information needed by other commands.
    #>
    [cmdletbinding(DefaultParameterSetName='UserParameterSet')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification='False positive for global vars.')]
    param        (
        [Parameter(ParameterSetName = 'UserParameterSet', Position = 1 )]
        #An array of delegated permissions to consent to.
        [string[]]$Scopes = $script:DefaultGraphScopes,

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
        If client_secret, ClientID and TenantID have all been set, offer -Credential & -AsApp and if a refresh token was stored, -refresh
        In either of those cases offer -NoRefresh
        If client ID and TenantID have been set (with or without secret) offer the cert parameters.
    #>
        $paramDictionary     = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $NoRefreshAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        if ($script:Client_secret -and $script:ClientID -and $script:TenantID) {
            $NoRefreshAttributeCollection.Add((New-Object System.Management.Automation.ParameterAttribute -Property @{       ParameterSetName='CredParameterSet'}))
            $CredParamAttribute    = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='CredParameterSet'}
            $AppParamAttribute     = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='AppSecretParameterSet';}
            $RefreshParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{Mandatory=$true; ParameterSetName='RefreshParameterSet'}

            #A credential object to logon with an app registered in the tennant
            $paramDictionary.Add('Credential',[RuntimeDefinedParameter]::new("Credential",  [pscredential],   $CredParamAttribute))
            #If Specified logs in as the app and gets the access granted to the app instead of logging on as a user.
            $paramDictionary.Add('AsApp',[RuntimeDefinedParameter]::new("AsApp",       [SwitchParameter],$AppParamAttribute))
            if ($script:RefreshToken) {
                $paramDictionary.Add('Refresh',[RuntimeDefinedParameter]::new("Refresh", [SwitchParameter],$RefreshParamAttribute))
            }
        }
        if ($script:ClientID -and $script:TenantID) {
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
        if (Get-command Get-AzAccessToken -ErrorAction SilentlyContinue) {
            $NoRefreshAttributeCollection.Add((New-Object System.Management.Automation.ParameterAttribute -Property @{       ParameterSetName='AzureParameterSet'}))
            $FromAzParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AzureParameterSet';Position=3}
            $DefProfParamAttribute = New-Object System.Management.Automation.ParameterAttribute -Property @{ParameterSetName='AzureParameterSet';Position=4}
            #The Az Module Context to use for the connection. You can get a list with Get-AzContext -ListAvailable. Note this parameter only accepts one context and if multiple are supplied it will only connect to the last one supplied
            $paramDictionary.Add('FromAzureSession',[RuntimeDefinedParameter]::new('FromAzureSession', [SwitchParameter], $FromAzParamAttribute))
            $paramDictionary.Add('DefaultProfile',  [RuntimeDefinedParameter]::new('DefaultProfile',   [System.Object],   $DefProfParamAttribute))
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
        $bp = @{} + $PSBoundParameters #I do not know why psb doesn't work normally with dynamic params but this works round it....
        $paramsToPass         = @{}

        #Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
        if (-not $bp.refresh)           {$global:DriveCache          = @{}  }

        #Justin used this variable, for checking expiry I have moved it on to an extra property of [graphSession].instance I'll remove this when I'm happy on compat.
        if ($global:__MgAzTokenExpires) {$global:__MgAzTokenExpires = $null}

        #region to get a token for a name / password with a registerd appID and secret, or to refresh one
        #credential , refresh, Azaupp, FromAzureSession are dynamic to hide them if we dont have what they need
        if ($bp.Credential -or $bp.Refresh -or $bp.AsApp -or $bp.FromAzureSession ){
            $tokenUri   = "https://login.microsoft.com/$script:TenantID/oauth2/token"
            # Send request with grant type of 'password' and creds, or 'refresh' or for the app use 'client_credentials'
            if     ($bp.Refresh)          {
                Write-Verbose "CONNECT: Sending a 'Refresh_token' token request "
                $authresp   =   Get-AccessToken -GrantType refresh_token -BodyParts @{'refresh_token' = $script:RefreshToken}
               <# $authresp   =   Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                    'grant_type'    = 'refresh_token'  ;
                    'refresh_token' = $script:RefreshToken
                    'client_id'     = $script:ClientID
                    'client_secret' = $script:Client_secret
                    'resource'      = 'https://graph.microsoft.com'
                } #>
            }
            elseif ($bp.Credential)       {
                Write-Verbose "CONNECT: Sending a 'Password' token request for $($bp.Credential.UserName) "
                 $authresp   =   Get-AccessToken -GrantType password -BodyParts @{ 'username' = $bp.Credential.UserName; 'password' = $bp.Credential.GetNetworkCredential().Password}
              <#
                $authresp   =   Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                    'grant_type'    = 'password'
                    'resource'      = 'https://graph.microsoft.com'
                    'username'      = $bp.Credential.UserName
                    'password'      = $bp.Credential.GetNetworkCredential().Password
                    'client_id'     = $script:ClientID
                    'client_secret' = $script:Client_secret
                }#>
            }
            elseif ($bp.AsApp)            {
                Write-Verbose "CONNECT: Sending a 'client_credentials' token request for the App."
                 $authresp   =   Get-AccessToken
                <#$authresp  = Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                    'grant_type'    = 'client_credentials';
                    'resource'      = 'https://graph.microsoft.com';
                    'client_id'     =  $script:ClientID ;
                    'client_secret' =  $script:Client_secret;
                }#>
            }
            #to leverage an existing Az Session call a command in the Az.Account module (V2 and later)
            elseif ($bp.FromAzureSession) {
                if ($bp.DefaultProfile) {$global:__MgAzContext = $DefaultProfile }
                $authresp =  Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com' DefaultProfile $global:__MgAzContext
            }
            #did it work ? If so store the token and what'll we need for refreshing it
            if ($authresp.access_token)   {
                Write-Verbose ("CONNECT: Token Response= " + ($authresp | get-member -MemberType NoteProperty).name -join ", ")
                $null = $paramsToPass.Add("AccessToken",  $authresp.access_token)
                if     ($authresp.scope)         {Write-Verbose "CONNECT: Scope= $($authresp.scope)" }
                if     ($authresp.refresh_token) {$script:RefreshToken       = $authresp.refresh_token}
                if     ($authresp.expires_in)    {$global:__MgAzTokenExpires = (Get-Date).AddSeconds([int]$authresp.expires_in -60 )}
                elseif ($authresp.expires_on)    {
                    if ($authresp.expires_on -is [string] -and
                        $authresp.expires_on -match "^\w{10}$") {
                                                  $global:__MgAzTokenExpires = [datetime]::UnixEpoch.AddSeconds($oauthAPP.expires_on)}
                elseif ($authresp.ExpiresOn  -is [datetimeoffset]){
                                                  $global:__MgAzTokenExpires = $authresp.ExpiresOn.LocalDateTime }
                }
                if     ($bp.NoRefresh)           {
                    $script:RefreshParams      = $null
                }
                elseif ($bp.FromAzureSession)    {
                    $script:RefreshParams      = @{'Quiet' = $true; 'FromAzureSession' = $True}
                    $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' -FromAzureSession '))
                }
                else                             {
                    $script:RefreshParams      = @{'Quiet' = $true; 'Refresh' = $true}
                     $RefreshScriptBlock        = [scriptblock]::Create(($RefreshScript -f ' -Refresh '))
                }
            }
            else {throw [System.UnauthorizedAccessException]::new("No Token was returned")}
        }
        #endregion

        #now connect, either using a token - passed or just fetched, or using a cert, or using the device dialog as needed

        #region collect parameters and call Connect-MGGraph
        $paramsinTarget       = (Get-Command Connect-MgGraph).Parameters.Keys |
                                    Where-Object {$_ -notin [System.Management.Automation.Cmdlet]::CommonParameters}
        $paramsFromCurrentSet = $pscmdlet.MyInvocation.MyCommand.ParameterSets.Where({$_.name -eq $PSCmdlet.ParameterSetName})
        $paramsFromCurrentSet = $paramsFromCurrentSet.parameters.Name | Where-Object {$_ -in $paramsinTarget -and (Get-Variable $_ -ValueOnly -ErrorAction SilentlyContinue)}
        foreach ($p in $paramsFromCurrentSet ) {$paramsToPass[$p] = Get-Variable $P -ValueOnly   ; Write-Verbose ("{0,20} = {1}" -f $p.ToUpper(), $paramsToPass[$p]) }
        foreach ($p in [System.Management.Automation.Cmdlet]::CommonParameters.Where({$bp.ContainsKey($_)})) {
            $paramsToPass[$p] = $bp[$p]  ; Write-Verbose ("{0,20} = {1}" -f $p.ToUpper(), $paramsToPass[$p])
        }

        $result = Connect-MgGraph @paramsToPass
        #endregion
        #region if succesful cache information about the user and session, and if necessary setup a trigger to auto-refresh tokens we fetched above
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
                $global:GraphUser =  $user.userPrincipalName
                Add-Member -Force -InputObject $authcontext -NotePropertyName UserDisplayName     -NotePropertyValue $user.displayName
                Add-Member -Force -InputObject $authcontext -NotePropertyName UserID              -NotePropertyValue $user.ID
            }
            else {$result           = "Welcome To Microsoft Graph, connected as the app '{0}'." -f $authcontext.AppName }
            Add-Member     -Force -InputObject $authcontext -NotePropertyName RefreshTokenPresent -NotePropertyValue ($script:RefreshToken -as [bool])
            Add-Member     -Force -InputObject $authcontext -NotePropertyName TokenAutoRefresh    -NotePropertyValue ($RefreshScriptBlock  -as [bool])
            if    ($global:__MgAzTokenExpires) {
                Add-Member -Force -InputObject $authcontext -NotePropertyName TokenExpires         -NotePropertyValue ($global:__MgAzTokenExpires)
            }
            elseif ($authcontext.TokenExpires) {$authcontext.TokenExpires = $null}
            if ($RefreshScriptBlock -and -not $global:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend']) {
                if (-not $Quiet){
                    Write-Host -Fore DarkCyan "The HttpPipelinePrepend parameter now has a default that checks for refresh tokens. Any command which uses this parameter will lose the auto-refresh"
                }
                $global:PSDefaultParameterValues['*-Mg*:HttpPipelinePrepend'] = $RefreshScriptBlock
            }
            if ($NoRefresh -and -not $Quiet) {
                Write-Host -Fore Cyan "-NoRefresh was specified. You will need to run this command again after $($tokeninfo.ExpiresOn.LocalDateTime.ToString())"
            }
        }
        #endregion

        if (-not $Quiet) {return $result}
    }
}

function Show-GraphSession   {
    <#
        .Synopsis
            Returns information about the current sesssion
    #>
    [CmdletBinding(DefaultParameterSetName='None')]
    [OutputType([String])]
    param  (
        [Parameter(ParameterSetName='Who')]
        [switch]$Who,
        [Parameter(ParameterSetName='Scopes')]
        [switch]$Scopes,
        [switch]$Options,
        [switch]$AppName,
        [switch]$CachedToken
    )
    if     (-not       [GraphSession]::Instance.AuthContext) {Write-Host  "Ready for Connect-Graph."; return}
    if     ($Scopes)  {[GraphSession]::Instance.AuthContext.Scopes}
    elseif ($Who)     {[GraphSession]::Instance.AuthContext.Account}
    elseif ($AppName) {[GraphSession]::Instance.AuthContext.AppName}
    elseif ($options) {[pscustomobject]@{
        'TenantID'        = $script:TenantID
        'ClientID'        = $script:ClientID
        'ClientSecretSet' = $script:Client_Secret -as [bool]
        'DefaultScopes'   = $script:DefaultGraphScopes -join ', '
    }}
    elseif (-not $CachedToken)  {Get-MgContext}
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

function ContextHas          {
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
        #if specified returns ture if connected with a user account and false if connected as an application
        [switch]$AsUser,
        #if specified returns ture if connected as an application and false if connected with a user account
        [switch]$AsApp,
        #If specified break instead of turning false
        [switch]$BreakIfNot,
        #If specified reverses the output.
        [switch]$Not
    )
    if (-not $scopes) { $state = $true }
    else {
        $state =  $false
        foreach ($s in $scopes)  {
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
        if ($scopes              -and -not $state) {Write-Warning ("This requires the {0} scope(s)." -f ($scopes -join ', ')); break}
        if ($WorkOrSchoolAccount -and -not $state) {Write-Warning  "This requires a work or school account."                  ; break}
        if ($AsUser              -and -not $state) {Write-Warning  "This requires a user0logon, not an app-logon."            ; break}
        if ($AsApp               -and -not $state) {Write-Warning  "This requires an app-logon, not a user-logon."            ; break}
    }
    #otherwise return true or false
    else  {return ( $state -xor $not )}
}

function Set-GraphConnectionOptions {
[cmdletbinding()]
param (
    #Your Tennant ID
    $TenantID,
    #Client ID if not using the SDK default of 14d82eec-204b-4c2f-b7e8-296a70dab67e. Must be known to your tennant
    $ClientID,
    #Secret set for the client ID in your $TenantID
    $Client_Secret,
    #Default Scopes to request
    $DefaultScopes
)
    if ($TenantID)        {
        if ($TenantID -notmatch $GUIDRegex) {
              Write-Warning 'TenantID should be a GUID'  ; break
        }
        else {$script:TenantID           = $TenantID}
    }
    if ($ClientID)        {
        if ($Clientid -notmatch $GUIDRegex) {
            Write-Warning 'ClientID should be a GUID'  ; break
        }
        else {$script:ClientID           = $ClientID}
    }
    if ($Client_Secret)   {
        if     ($Client_Secret -is [string]) {
               $script:Client_Secret      = $Client_Secret
        }
        elseif ($Client_Secret -is [securestring]) {
               $script:Client_Secret =  (new-object pscredential -ArgumentList "NoName", $Client_Secret).GetNetworkCredential().Password
        }
        else  {Write-Warning 'Client_secret should be a string or preferably a securestring'  ; break}
    }
    if ($script:TenantID) {Write-Verbose "TenantID: $script:TenantID , ClientID = $script:ClientID"}
    if ($DefaultScopes)   {$script:DefaultGraphScopes = $DefaultScopes}

    Write-Verbose ('Scopes: ' + ($script:DefaultGraphScopes -join ', '))
}

#calls file with calls to Set-GraphConnectionoptions
if     ($env:MGSettingsPath )                      {. $env:MGSettingsPath}
elseif(Test-Path "$PSScriptRoot\AuthSettings.ps1") {. "$PSScriptRoot\AuthSettings.ps1"}
