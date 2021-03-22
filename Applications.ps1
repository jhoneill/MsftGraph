<#
    The Get-GraphServicePrincipal function reworks work on service principals which was published by Justin Grote at
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Get-MgO365ServicePrincipal.ps1
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Get-MgManagedIdentity.ps1  and
        https://github.com/JustinGrote/JustinGrote.Microsoft.Graph.Extensions/blob/main/src/Public/Get-MgAppRole.ps1

    and licensed by him under the same MIT terms which apply to this module (see the LICENSE file for details)

    Portions of this file are Copyright 2021 Justin Grote @justinwgrote

    The remainder is Copyright 2018-2021 James O'Neill
#>
using namespace  Microsoft.Graph.PowerShell.Models
function Get-GraphServicePrincipal {
    <#
      .Synopsis
        Returns information about Service Principals
      .Description
        A replacement for the SDK's Get-MgServicePrincipal
        That has orderby which doesn't work - it's in the Docs but the API errors if you try
        It doesn't have find by name, or select Application or Managed IDs
      .Example
        PS > Get-GraphServicePrincipal "Microsoft graph*"

        Id                                   DisplayName                      AppId                                SignInAudience
        --                                   -----------                      -----                                --------------
        25b13fbf-2f44-457a-9e68-d3414fc97915 Microsoft Graph                  00000003-0000-0000-c000-000000000000 AzureADMultipleOrgs
        4e71d88a-0a46-4274-85b8-82ad86877010 Microsoft Graph Change Tracking  0bf30f3b-4a52-48df-9a82-234910c4a086 AzureADMultipleOrgs
        ...

        Run with a name the command returns service principals with matching names.
        .Example
        PS >Get-GraphServicePrincipal 25b13fbf-2f44-457a-9e68-d3414fc97915 -ExpandAppRoles

        Value                         DisplayName                Enabled Id
        -----                         -----------                ------- --
        AccessReview.Read.All         Read all access reviews    True    d07a8cc0-3d51-4b77-b3b0-32704d1f69fa
        AccessReview.ReadWrite.All    Manage all access reviews  True    ef5f7d5c-338f-44b0-86c3-351f46c8bb5f
        ...
        In this example GUID for Microsoft Graph was used from the previous example, and the command has listed the roles available to applications
    #>

    [CmdletBinding(DefaultParameterSetName='List1')]
    [OutputType([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphAppRole],ParameterSetName=('AllRoles','FilteredRoles'))]
    [OutputType([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphPermissionScope],ParameterSetName=('AllScopes','FilteredScopes'))]
    [OutputType([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphServicePrincipal],ParameterSetName=('Get2','List1','List2','List3','List4'))]
    param   (
        [Parameter(ParameterSetName='AllRoles',       Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Parameter(ParameterSetName='FilteredRoles',  Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Parameter(ParameterSetName='AllScopes',      Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Parameter(ParameterSetName='FilteredScopes', Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [Parameter(ParameterSetName='Get2',           Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        #The GUID(s) for ServicePrincipal(s). Or SP objects. If a name is given instead, the command will try to resolve matching Service principals
        $ServicePrincipalId,

        [Parameter(ParameterSetName='List5')]
        [string]$AppId,

        #Produces a list filtered to only managed identities
        [Parameter(ParameterSetName='List2')]
        [switch]$ManagedIdentity,

        #Produces a list filtered to only applications
        [Parameter(ParameterSetName='List3')]
        [switch]$Application,

        #Produces a convenience list of office 365 security principals
        [Parameter(ParameterSetName='List4')]
        [switch]$O365ServicePrincipals,

        #Select properties to be returned
        [Alias('Select')]
        [String[]]$Property,

        #Filters items by property values
        [Parameter(ParameterSetName='List1')]
        [String]$Filter,

        #Returns the list of application roles to those the role name, displayname or ID match the parameter value. Wildcards are supported
        [Parameter(ParameterSetName='AllRoles', Mandatory=$true)]
        [switch]$ExpandAppRoles,

        #Filters the list of application roles available within a SP
        [Parameter(ParameterSetName='FilteredRoles', Mandatory=$true)]
        [string]$AppRoleFilter,

        #Returns the list of (user) oauth scopes available within a SP
        [Parameter(ParameterSetName='AllScopes', Mandatory=$true)]
        [switch]$ExpandScopes,

        #Filters the list of oauth scopes to those where the scope name, displayname or ID match the parameter value. Wildcards are supported
        [Parameter(ParameterSetName='FilteredScopes', Mandatory=$true)]
        [string]$ScopeFilter
    )
    begin   {
        [String]$managedIdentityFilter = @(
            '00000001-0000-0000-c000-000000000000' #Azure ESTS Service
            '00000007-0000-0000-c000-000000000000' #Common Data Service
            '0000000c-0000-0000-c000-000000000000' #Microsoft App Access Panel'
            '00000007-0000-0ff1-ce00-000000000000' #Microsoft Exchange Online Protection
            '00000003-0000-0000-c000-000000000000' #Microsoft Graph
            '00000006-0000-0ff1-ce00-000000000000' #Microsoft Office 365 Portal
            '00000012-0000-0000-c000-000000000000' #Microsoft Rights Management Services
            '00000008-0000-0000-c000-000000000000' #Microsoft.Azure.DataMarket
            '00000002-0000-0ff1-ce00-000000000000' #Office 365 Exchange Online
            '00000003-0000-0ff1-ce00-000000000000' #Office 365 SharePoint Online
            '00000009-0000-0000-c000-000000000000' #Power BI Service
            '00000004-0000-0ff1-ce00-000000000000' #Skype for Business Online
            '00000002-0000-0000-c000-000000000000' #Windows Azure Active Directory
        ).foreach{"appId eq '$PSItem'"} -join ' or '

        if ($ExpandScopes)   {$ScopeFilter   = '*'}
        if ($ExpandAppRoles) {$AppRoleFilter = '*'}

        $webparams = @{
                    AsType          =  ([MicrosoftGraphServicePrincipal])
                    ExcludeProperty = @('resourceSpecificApplicationPermissions','@odata.context','createdDateTime','verifiedPublisher')
                    Headers         = @{'ConsistencyLevel'= 'Eventual'}
        }
    }
    process {
        if ( -not    $ServicePrincipalId)    {
            if      ($ManagedIdentity)             {$filter  = "?`$filter=servicePrincipaltype eq 'ManagedIdentity'"}
            elseif  ($Application)                 {$filter  = "?`$filter=servicePrincipaltype eq 'Application'"}
            elseif  ($AppId)                       {$filter  = "?`$filter=appid eq '$AppId'" }
            elseif  ($PSBoundParameters['Filter'] -and
                     $O365ServicePrincipal )       {$filter  = "?`$filter=( $($PSBoundParameters['Filter']) ) and $managedIdentityFilter"}
            elseif  ($PSBoundParameters['Filter']) {$filter  = "?`$filter=$($PSBoundParameters['Filter'])"}
            elseif  ($O365ServicePrincipals)       {$filter  = "?`$filter=$managedIdentityFilter"}

            if      ($Property -and $filter)       {$filter += '&$select=' +($property -join ',')}
            elseif  ($Property)                    {$filter  = '?$select=' +($property -join ',')}
            Invoke-GraphRequest "$GraphUri/servicePrincipals$filter" -ValueOnly @webparams | Sort-Object displayname
        }
        else {
            foreach ($sp in $ServicePrincipalId) {
                $result = $null
                if     ($sp -match $GUIDRegex)      {
                    $webparams['uri'] = "$GraphUri/servicePrincipals/$sp"
                    if ($Property) {$webparams['uri'] += '?$select=' +($property -join ',')}
                    try {$result = Invoke-GraphRequest @webparams}
                    catch {
                        if ($_.Exception.Response.StatusCode.value__  -eq 404) {
                            Write-Warning "$sp was not found as a service principal ID. It may be an App ID.";  continue
                        }
                        else {throw $_.Exception }
                    }
                }
                else   {
                    $webparams['uri'] = "$GraphUri/servicePrincipals?`$filter=" + (FilterString $sp)
                    if ($Property) {$webparams['uri'] += '&$select=' +($property -join ',')}
                    $result = Invoke-GraphRequest  -ValueOnly  @webparams
                }
                if     ($AppRoleFilter) {
                    $(foreach ($r in $result) {$r.approles |
                            Where-Object {$_.id -like $AppRoleFilter  -or $_.DisplayName -like $AppRoleFilter -or $_.value -like $AppRoleFilter } |
                            Add-Member -PassThru -NotePropertyName ServicePrincipal -NotePropertyValue $r.Id} ) |
                                Sort-Object -Property Value
                }
                elseif ($ScopeFilter)    {
                    $(foreach ($r in $result) {$r.Oauth2PermissionScopes |
                        Where-Object {$_.id -like $ScopeFilter  -or $_.AdminConsentDisplayName -like $ScopeFilter -or $_.value -like $ScopeFilter } |
                            Add-Member -PassThru -NotePropertyName ServicePrincipal -NotePropertyValue $r.Id} ) |
                                Sort-Object -Property Value
                }
                else {$result | Sort-Object -Property displayname}
            }
        }
    }
}

function Get-GraphApplication {
    <#
      .Synopsis
        Returns information about Applications
    #>

    [CmdletBinding(DefaultParameterSetName='List1')]
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphApplication])]
    param   (
        [Parameter(ParameterSetName='List3',Position=0)]
        [string]$Id,

        #The GUID(s) for Apps(s). Or App objects. If a name is given instead, the command will try to resolve matching App principals
        [Parameter(ParameterSetName='List2',ValueFromPipelineByPropertyName=$true)]
        [string]$AppId,

        #Select properties to be returned
        [Alias('Select')]
        [String[]]$Property,

        #Filters items by property values
        [Parameter(ParameterSetName='List1')]
        [String]$Filter
    )
    process {
        $result = @()
        $webparams = @{
                    ExcludeProperty = @('verifiedPublisher', 'applicationTemplateId','addins','@odata.context')
                    Headers         = @{'ConsistencyLevel'= 'Eventual'}
        }
        if ( -not    $Id)    {
            if      ($PSBoundParameters['Filter']) {$filter  = "?`$filter=$($PSBoundParameters['Filter'])"}
            elseif  ($AppId)                       {$filter  = "?`$filter=AppID eq '$AppId'"}
            if      ($Property -and $filter)       {$filter +=  '&$select=' +($property -join ',')}
            elseif  ($Property)                    {$filter  =  '?$select=' +($property -join ',')}

            $result += Invoke-GraphRequest "$GraphUri/Applications$filter" -ValueOnly @webparams
        }
        else {
            foreach ($app in $Id) {
                if  ($app -match $GUIDRegex)      {
                    $webparams['uri'] = "$GraphUri/applications/$app"
                    if ($Property) {$webparams['uri'] += '?$select=' +($property -join ',')}
                    try   { $result += Invoke-GraphRequest @webparams}
                    catch {
                            if ($_.Exception.Response.StatusCode.value__  -eq 404) {
                                Write-Warning "$sp was not found as an ID It may be an APPID.";  continue
                            }
                            else {throw $_.Exception }
                    }
                }
                else   {
                        $filter =  '?$filter=' + (FilterString $app)
                        if ($Property) {$filter += '&$select=' +($property -join ',')}
                        $result += Invoke-GraphRequest "$GraphUri/applications$filter" -ValueOnly  @webparams
                }
            }
        }
        foreach ($r in $result) {
            foreach ($p in $r.passwordCredentials) {
                $p.customKeyIdentifier = [byte[]][char[]]$p.customKeyIdentifier
            }
            foreach ($k in $r.keyCredentials) {
                $k.customKeyIdentifier = [byte[]][char[]]$k.customKeyIdentifier
            }
            New-Object -TypeName MicrosoftGraphApplication -Property $r
        }
    }
}