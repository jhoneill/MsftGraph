<#
 for types XML
      GraphOrganization  --> Microsoft.Graph.PowerShell.Models.MicrosoftGraphOrganization
      GraphDomain        --> Microsoft.Graph.PowerShell.Models.MicrosoftGraphVerifiedDomain
      GraphSKU           --> Microsoft.Graph.PowerShell.Models.MicrosoftGraphSubscribedSku
      GraphServicePlan   --> MicrosoftGraphServicePlanInfo

#>

Function Get-GraphDomain {
    <#
      .synopsis
        Gets domains in the current tenant
      .Description
        Requires consent to use at least the Directory.Read.All scope
    #>
    [OutputType([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDomain])]
    [cmdletbinding(DefaultParameterSetName='None')]
    param (
        [parameter(Position=1, ValueFromPipeline=$true, ParameterSetName='Domain',    Mandatory=$true)]
        [parameter(Position=1, ValueFromPipeline=$true, ParameterSetName='VDRecords', Mandatory=$true)]
        [parameter(Position=1, ValueFromPipeline=$true, ParameterSetName='SCRecords', Mandatory=$true)]
        [parameter(Position=1, ValueFromPipeline=$true, ParameterSetName='NameRef',   Mandatory=$true)]
        $Domain,

        [parameter(ParameterSetName='VDRecords',Mandatory=$true)]
        [alias('VR')]
        [switch]$VerificationDNSRecords,

        [parameter(ParameterSetName='SCRecords',Mandatory=$true)]
        [switch]$ServiceConfigurationRecords,

        [parameter(ParameterSetName='NameRef',Mandatory=$true)]
        [switch]$NameReferenceList,

        [Parameter(DontShow)]
        [System.Uri]
        # The URI for the proxy server to use
        ${Proxy},

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        ${ProxyCredential},

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        ${ProxyUseDefaultCredentials}
    )

    if (-not $Domain) {Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomain_List1 @PSBoundParameters}
    else {
        $null = $PSBoundParameters.Remove("Domain")
        foreach ($d in $Domain) {
            if     ($d.id)           {$d = $d.id}
            elseif ($d -isnot [String]) {Write-Warning -Message 'Could not find the Domain ID from the parameter'}
            if ($VerificationDNSRecords) {
                $null = $PSBoundParameters.Remove("VerificationDNSRecords")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainVerificationDnsRecord_List1 -DomainId $d @PSBoundParameters
            }
            elseif ($ServiceConfigurationRecords)  {
                $null = $PSBoundParameters.Remove("ServiceConfigurationRecords")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainServiceConfigurationRecord_List1 -DomainId $d @PSBoundParameters
            }
            elseif ($NameReferenceList)  {
                $null = $PSBoundParameters.Remove("NameReferenceList")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainNameerenceByRef_List1 -DomainId $d @PSBoundParameters
            }
            else {
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomain_Get1 -DomainId $d @PSBoundParameters
            }
        }
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
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphOrganization])]
    [cmdletbinding(DefaultParameterSetName="None")]
    Param(
        $Organization,
        [Parameter(DontShow)]
        [System.Uri]
        # The URI for the proxy server to use
        ${Proxy},

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        ${ProxyCredential},

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        ${ProxyUseDefaultCredentials}
    )

    Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgOrganization_List1 @PSBoundParameters
}

Function Get-GraphSKU {
    <#
      .Synopsis
        Gets details of SKUs organization an organization has subscribed to
      .Example
        Get-GraphSKUList | where skupartnumber -match "enterprise" | Get-GraphSKU -ServicePlans | sort servicePlanName | format-table
        Finds "Enterprise" SKUS and displays their service plans in alphabetical order.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Alias('Get-GraphSubscribedSku')]
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(ParameterSetName='BySku', Mandatory=$true,ValueFromPipeline=$true)]
        $SKU,
        #If specified just returns the Service plans for the SKU, otherwise returns the SKU with a service plans property
        [parameter(ParameterSetName='BySku')]
        [switch]$ServicePlans,
        [Parameter(DontShow)]
        [System.Uri]
        # The URI for the proxy server to use
        ${Proxy},

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        ${ProxyCredential},

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        ${ProxyUseDefaultCredentials}
    )
    begin   {
        $result = @()
    }
    process {
       if (-not $sku) {$result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List1 @PSBoundParameters}
        else {
            $null = $PSBoundParameters.Remove("ServicePlans")
            $null = $PSBoundParameters.Remove("SKU")
            foreach ($s in $sku) {
                if     ($s.id)          {$result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get1 -SubscribedSkuId ($s.id) @PSBoundParameters}
                elseif ($s -is [String]){$result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get1 -SubscribedSkuId  $s     @PSBoundParameters}
                else   {Write-Warning -Message 'Could not find the SKU ID from the parameter'; continue}
            }
        }
    }
    end     {
            foreach ($r in $result) {
                foreach($plan in $r.ServicePlans) {
                    Add-Member -InputObject $plan -MemberType NoteProperty -Name "SkuPartNumber" -Value $r.SkuPartNumber
                }
            }
            if ($ServicePlans) {$result.ServicePlans}
            else               {$result }
    }
}