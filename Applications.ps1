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
        It doesn't have search by name, or select Application or Managed IDs
      .Example
        PS > Get-GraphServicePrincipal "Microsoft graph"

        Id                                   DisplayName                      AppId                                SignInAudience
        --                                   -----------                      -----                                --------------
        25b13fbf-2f44-457a-9e68-d3414fc97915 Microsoft Graph                  00000003-0000-0000-c000-000000000000 AzureADMultipleOrgs
        4e71d88a-0a46-4274-85b8-82ad86877010 Microsoft Graph Change Tracking  0bf30f3b-4a52-48df-9a82-234910c4a086 AzureADMultipleOrgs
        ...

        Run with just a name the command returns service principals with matching names.
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

        #Search items by search phrases
        [Parameter(ParameterSetName='List1')]
        [Parameter(ParameterSetName='List2')]
        [Parameter(ParameterSetName='List3')]
        [String]$Search,

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
        Test-GraphSession
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
    }
    process {
        if  (   -not    $ServicePrincipalId)    {
            if         ($O365ServicePrincipals  -and $PSBoundParameters['Filter']) {
                        $PSBoundParameters['Filter'] ="( $($PSBoundParameters['Filter']) ) and $managedIdentityFilter"
            }
            elseif     ($O365ServicePrincipals) {
                        $PSBoundParameters['Filter'] =  $managedIdentityFilter
            }
            elseif     ($ManagedIdentity)       {
                        $PSBoundParameters['Filter']="servicePrincipaltype eq 'ManagedIdentity'"
            }
            elseif     ($Application)           {
                        $PSBoundParameters['Filter']="servicePrincipaltype eq 'Application'"
            }
            elseif     ($AppId) {
                        $PSBoundParameters['Filter']="appid eq '$AppId'"
            }
            foreach    ($param in @('Application', 'AppId', 'ManagedIdentity', 'O365ServicePrincipals')) {
                  [void]$PSBoundParameters.Remove($param )
            }
            Microsoft.Graph.Applications.private\Get-MgServicePrincipal_List1 @PSBoundParameters -all -ConsistencyLevel Eventual | Sort-Object displayname
        }
        else {
            foreach    ($param in @('ServicePrincipalId', 'ExpandAppRoles', 'ExpandScopes', 'AppRoleFilter', 'ScopeFilter')) {
                  [void]$PSBoundParameters.Remove($param )
            }
            foreach    ($sp in $ServicePrincipalId) {
                $result = $null
                if     ($sp -match $GUIDRegex)      {
                    $uri = "$GraphUri/servicePrincipals/$sp"
                    if ($Property) {$uri += '?$select=' +($property -join ',')}
                    try {
                        $result = Invoke-GraphRequest $uri -AsType ([MicrosoftGraphServicePrincipal])
                    }
                    catch {
                        if ($_.Exception.Response.StatusCode.value__  -eq 404) {
                            Write-Warning "$sp was not found as a service principal ID. It may be an App ID.";  continue
                        }
                        else {throw $_.Exception }
                    }
                }
                else   {
                        [void]$PSBoundParameters.Remove('ServicePrincipalId')
                        $psboundParameters['Filter']="startswith(displayName,'$sp')"
                        $result = Microsoft.Graph.Applications.private\Get-MgServicePrincipal_List1 @PSBoundParameters -ConsistencyLevel Eventual | Sort-Object displayname
                }
                if     ($AppRoleFilter)  {
                        $result | Select-Object -ExpandProperty approles |
                                    Where-Object {$_.id -like $AppRoleFilter  -or $_.DisplayName -like $AppRoleFilter -or $_.value -like $AppRoleFilter } |
                                        Sort-Object -Property Value
                }
                elseif ($ExpandAppRoles) {
                        $result | Select-Object -ExpandProperty approles | Sort-Object -Property Value
                }
                elseif ($ScopeFilter)    {
                        $result | Select-Object -ExpandProperty Oauth2PermissionScopes |
                                    Where-Object {$_.id -like $ScopeFilter  -or $_.AdminConsentDisplayName -like $ScopeFilter -or $_.value -like $ScopeFilter } |
                                        Sort-Object -Property Value
                }
                elseif ($ExpandScopes)   {
                        $result | Select-Object -ExpandProperty Oauth2PermissionScopes | Sort-Object -Property Value
                }
                else   {$result}
            }
        }
    }
}