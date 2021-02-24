function Get-GraphServicePrincipal {
<#
    .Description
        A replacement for the SDK's Get-MgServicePrincipal
        That has orderby which doesn't work - the it's in the Docs but the API errors if you try
        It doesn't have search my name, or select managedIDs or Applications.
#>
[OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal])]
[CmdletBinding(DefaultParameterSetName='List1')]
    param   (
        [Parameter(ParameterSetName='Get2', Mandatory=$true, Position=0)]
        # key: id of servicePrincipal
        [String[]]$ServicePrincipalId,

        [Parameter(ParameterSetName='List2')]
        [switch]$ManagedIdentity,

        [Parameter(ParameterSetName='List3')]
        [switch]$Application,

        # Expand related entities
        [Alias('Expand')]
        [String[]]
        $ExpandProperty,

        # Select properties to be returned
        [Alias('Select')]
        [String[]]$Property,

        # Filter items by property values
        [Parameter(ParameterSetName='List1')]
        [String]$Filter,

        # Search items by search phrases
        [Parameter(ParameterSetName='List1')]
        [Parameter(ParameterSetName='List2')]
        [Parameter(ParameterSetName='List3')]
        [String]$Search,

        # Show only the first n items
        [Parameter(ParameterSetName='List1')]
        [Parameter(ParameterSetName='List2')]
        [Parameter(ParameterSetName='List3')]
        [Alias('Limit')]
        [Int32]$Top,

        [Parameter(ParameterSetName='List1')]
        [Parameter(ParameterSetName='List2')]
        [Parameter(ParameterSetName='List3')]
        [Int32]
        # Sets the page size of results.
        $PageSize,

        # List all pages.
        [Parameter(ParameterSetName='List1')]
        [Parameter(ParameterSetName='List2')]
        [Parameter(ParameterSetName='List3')]
        [switch]$All
    )

    process {
        foreach ($sp in $ServicePrincipalId) {
            if ($sp -match $GUIDRegex) {
                Microsoft.Graph.Applications.private\Get-MgServicePrincipal_Get2  @PSBoundParameters
            }
            else {
                if ($sp -and $sp -Notmatch $GUIDRegex) {
                    [void]$PSBoundParameters.Remove('ServicePrincipalId')
                    $psboundParameters['Filter']="startswith(displayName,'$sp')"
                }
                elseif ($ManagedIdentity) {
                    [void]$PSBoundParameters.Remove('ManagedIdentity')
                    $psboundParameters['Filter']="servicePrincipaltype eq 'ManagedIdentity'"
                }
                elseif ($Application) {
                    [void]$PSBoundParameters.Remove('Application')
                    $psboundParameters['Filter']="servicePrincipaltype eq 'Application'"
                }
                Microsoft.Graph.Applications.private\Get-MgServicePrincipal_List1 @PSBoundParameters -ConsistencyLevel Eventual | Sort-Object displayname
            }
        }
    }
}