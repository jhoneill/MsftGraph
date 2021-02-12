using namespace System.Management.Automation

#xxxx todo: check context is a workorschool account and that it has the right scopes and warn / error / throw if not.

class SkuCompleter : IArgumentCompleter     {
    [Microsoft.Graph.PowerShell.Models.MicrosoftGraphSubscribedSku[]]$skus = @()
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    )
    {
        $result = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        if (-not $this.skus)  {$this.skus = Get-GraphSKU }
        $wildcard          = ("*" + ($wordToComplete  -replace "['""]",'' )+ "*")
        $this.skus.where({$_.skuPartNumber -like $wildcard}).skuPartNumber |
            Sort-Object | ForEach-Object {$result.Add([System.Management.Automation.CompletionResult]::new($_))}
        return $result
    }
}

class SkuPlanCompleter : IArgumentCompleter {
    [Microsoft.Graph.PowerShell.Models.MicrosoftGraphSubscribedSku[]]$skus = @()
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    )
    {
        $result = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        if (-not $this.skus)  {$this.skus = Get-GraphSKU }
        $wildcard          = ("*" + ($wordToComplete  -replace "['""]",'' )+ "*")
        if ($FakeBoundParameters['SKUID']) {
            $selectedSkus  = $this.skus.where({$_.skuID -in $FakeBoundParameters['SKUID'] -or $_.skuPartNumber -in $FakeBoundParameters['SKUID'] })
        }
        else {
            $selectedSkus = $this.skus
        }
        $selectedSkus.ServicePlans.where({$_.ServicePlanName -like $wildcard}).ServicePlanName |
            Sort-Object | ForEach-Object {$result.Add([System.Management.Automation.CompletionResult]::new($_))}
        return $result
    }
}

class DomainCompleter : IArgumentCompleter {
    [Microsoft.Graph.PowerShell.Models.MicrosoftGraphDomain[]]$Domains = @()
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    )
    {
        $result = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        if (-not $this.Domains)  {$this.Domains = Get-GraphDomain }
        $wildcard          = ("*" + ($wordToComplete  -replace "['""]",'' )+ "*")

        $this.domains.id.where({$_ -like $wildcard}) |
            Sort-Object | ForEach-Object {$result.Add([System.Management.Automation.CompletionResult]::new($_))}
        return $result
    }
}
function Get-GraphDomain          {
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
        [switch]$NameReferenceList
    )

    if (-not $Domain) {Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomain_List1 @PSBoundParameters}
    else {
        #Allow an orgnaization object to be piped in.
        if ($Domain.verifiedDomains) {$Domain = $Domain.verifiedDomains}
        $null = $PSBoundParameters.Remove("Domain")
        foreach ($d in $Domain) {
            if     ($d.id)              {$d = $d.id}
            elseif ($d.name)            {$d = $d.name}
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

function Get-GraphOrganization    {
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
        $Proxy,

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        $ProxyCredential,

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        $ProxyUseDefaultCredentials
    )

    Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgOrganization_List1 @PSBoundParameters
}

function Get-GraphSKU             {
    <#
      .Synopsis
        Gets details of SKUs organization an organization has subscribed to
      .Example
        Get-GraphSKU "enterprise*" -ServicePlans | sort servicePlanName | format-table
        Finds any SKU with a name starting "Enterprise" and displays its service plans in alphabetical order.
    #>
    [Alias('Get-GraphSubscribedSku')]
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphSubscribedSku])]
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position = 1, ValueFromPipeline=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKU = '*',
        #If specified just returns the Service plans for the SKU, otherwise returns the SKU with a service plans property
        [switch]$ServicePlans,
        [Parameter(DontShow)]
        [System.Uri]
        # The URI for the proxy server to use
        $Proxy,

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        $ProxyCredential,

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        $ProxyUseDefaultCredentials
    )
    begin   {
        $result = @()
    }
    process {
        foreach ($s in $sku) {
            $null = $PSBoundParameters.Remove("ServicePlans") ,  $PSBoundParameters.Remove("SKU")
            if ($s.skuid)          {$s = ($s.skuid) }
            if ($s -notmatch '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$') {
                $result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List @PSBoundParameters |
                            Where-Object -Property SkuPartNumber -like $s
            }
            elseif ($s -is [String]) {
                $result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get  @PSBoundParameters -SubscribedSkuId $s}
            else   {Write-Warning -Message 'Could not find the SKU ID from the parameter'; continue}
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

function Grant-GraphUserLicense   {
    <#
      .Synopsis
        Grants the licence to use a particular stock-keeping-unit (SKU) to a user
    #>
    [cmdletbinding(SupportsShouldprocess=$true)]
    param (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position=1, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,

        #ID for the user (required. "me" will select the current user)
        [parameter(Position=2, ValueFromPipeline=$true, Mandatory = $true)]
        $UserID ,

        #Disables individual parts of the the SKU
        [ArgumentCompleter([SkuPlanCompleter])]
        [string[]]$DisabledPlans,

        #Runs the command without a confirmation dialog
        [Switch]$Force
    )
    begin {
        $request        = @{'addLicenses' = @() ; 'removeLicenses' = @()}

        foreach  ($s in $SKUID) {
            if   ($s.skuid) {$s = ($s.skuid) }
            if   ($s -match '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$') {
                  $sku = Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get -SubscribedSkuId $s
            }
            else {$sku = Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List  |
                            Where-Object -Property SkuPartNumber -Like $s
            }
            if   (-not $sku -or $sku.Count -gt 1) {
                Write-Warning "$s did not match a unique SKU" ; continue
            }
            elseif ($sku.ConsumedUnits -ge $sku.PrepaidUnits.Enabled ) {
                Write-Warning "$($sku.SkuPartNumber) has used all its prepaid units" ; continue
            }
            else {
                #Be ready to convert a disabled plan from names to GUIDs
                $skuplans = @{}
                foreach ($plan in $sku.ServicePlans) {
                    $skuplans[$plan.ServicePlanName] = $plan.ServicePlanId
                }
            }
            $thisReq = @{skuId = $sku.SkuId ; disabledPlans = @()}
            #We may have been passed many skus and many disabled plans. Only apply the plans that relate to the current sku.
            foreach ($d in $DisabledPlans) {
                if     ($skuplans.ContainsValue($d)) {$thisReq.disabledPlans += $d }
                elseif ($skuplans.Containskey($d))   {$thisReq.disabledPlans += $skuplans[$d] }
            }
            $request.addLicenses += $thisReq
            $licensePartNos = @($licensePartNos , $sku.SkuPartNumber) -join ", "
        }
        $webparams      =  @{
            Contenttype =  "application/json"
            Body        =  (ConvertTo-Json $request -Depth 10)
            Method      = 'POST'
        }
        Write-Debug $webparams.body
    }
    process {
        if (-not $licensePartNos)  {
            Write-Warning "No Valid SKUs were passed"
            return
        }
        if ($UserID -is [string] -and $userid -notmatch "me|$GUIDRegex" ) {
            $userId = Get-GraphUser $UserID
        }
        foreach ($u in $UserID ) {
            #region Add the user to web parameters: allow for mulitple users - potentially with an ID or a UPN
            if ($u -eq "me") {
                    $webparams['uri']   = "$GraphUri/me/assignLicense"
                    $userDisplayName    =  $Script:GraphUser.DisplayName
            }
            elseif ($u.id)  {
                    $webparams['uri']   = "$GraphUri/users/$($u.id)/assignLicense"
                    $userDisplayName    = $u.Id  #hope to change this if we have a display name
            }
            elseif ($u.UserPrincipalName) {
                    $webparams['uri']   = "$GraphUri/users/$($u.UserPrincipalName)/assignLicense"
                    $userDisplayName    = $u.UserPrincipalName  #hope to change this if we have a display name
            }
            elseif ($u -is [string] -and $u -match $GUIDRegex) {
                    $webparams['uri']   = "$GraphUri/users/$u/assignLicense"
                    $userDisplayName    = $u
            }
            elseif ($u -is [string]) {
                $u = Get-GraphUser $u
                if ($u.count -eq 1) {
                    $webparams['uri']   = "$GraphUri/users/$($u.id)/assignLicense"
                }
                else {
                    Write-Warning "Could not resolve $u to a single user. Ignoring"
                    continue
                }
            }
            if ($u.DisplayName) {$userDisplayName = $u.DisplayName }

            if ($Force -or $Pscmdlet.Shouldprocess($userdisplayname,"Grant licence for $licensePartNos")) {
                $u = Invoke-GraphRequest  @webparams
                Write-Verbose "$licensePartNos  Grantedto $($u.userPrincipalName)"
            }
        }
    }
}

function Revoke-GraphUserLicense  {
    <#
      .Synopsis
        Revokes a user's licence to use a particular stock-keeping-unit (SKU)
    #>
    [cmdletbinding(SupportsShouldprocess=$true)]
    param (
        #The SKU to revoke either as an ID or a SKU object containing an ID
        [parameter(Position=1, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,

        #ID for the user (required. "me" will select the current user)
        [parameter(Position=2, ValueFromPipeline=$true, Mandatory = $true)]
        $UserID ,

        #Runs the command without a confirmation dialog
        [Switch]$Force,

        [Parameter(DontShow)]
        [System.Uri]
        # The URI for the proxy server to use
        $Proxy,

        [Parameter(DontShow)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        # Credentials for a proxy server to use for the remote call
        $ProxyCredential,

        [Parameter(DontShow)]
        [System.Management.Automation.SwitchParameter]
        # Use the default credentials for the proxy
        $ProxyUseDefaultCredentials
    )
    begin {
        $request        = @{'addLicenses' = @() ; 'removeLicenses' = @()}
        foreach ($s in $SKUID) {
            if  ($s.skuid) {$s = ($s.skuid) }
            if  ($s -match '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$') {
                $request.removeLicenses += $s
            }
            else {
                 $sku = Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List @invokeParams |
                            Where-Object -Property SkuPartNumber -Like $s
                if     (-not $sku -or $sku.Count -gt 1) {Write-Warning "$s did not match a unique SKU" ; continue }
                else   { $request.removeLicenses +=  $sku.SkuId}
            }
        }
        $webparams      = @{
            Contenttype =  "application/json"
            Body        =  (ConvertTo-Json $request -Depth 10)
            Method      = 'POST'
        }
        Write-Debug $webparams.body
    }
    process {
        if (-not $request.removeLicenses)  {
            Write-Warning "No Valid SKUs were passed"
            return
        }
        foreach ($u in $UserID ) {
            #region Add the user to web parameters: allow for mulitple users - potentially with an ID or a UPN
            if ($u -eq "me") {
                    $webparams['uri']   = "$GraphUri/me/assignLicense"
                    $userDisplayName    =  $Script:GraphUser.DisplayName
            }
            elseif ($u.id)  {
                    $webparams['uri']   = "$GraphUri/users/$($u.id)/assignLicense"
                    $userDisplayName    = $u.Id  #hope to change this if we have a display name
            }
            elseif ($u.UserPrincipalName) {
                    $webparams['uri']   = "$GraphUri/users/$($u.UserPrincipalName)/assignLicense"
                    $userDisplayName    = $u.UserPrincipalName  #hope to change this if we have a display name
            }
            else {  $webparams['uri']   = "$GraphUri/users/$u/assignLicense"
                    $userDisplayName    = $u
            }
            if ($u.DisplayName) {$userDisplayName = $u.DisplayName }

            if ($Force -or $Pscmdlet.Shouldprocess($userdisplayname,"Revoke licence(s) for $($request.removeLicenses.Count) SKU(s)")) {
                $u = Invoke-GraphRequest  @webparams
                Write-Verbose "Revoked licence(s) for $($request.removeLicenses.Count) SKU(s) $($u.userPrincipalName)"
            }
        }
    }
}

function Get-GraphSkuLicensedUser {
    <#
      .Synopsis
        Revokes a user's licence to use a particular stock-keeping-unit (SKU)
    #>
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position=1, ValueFromPipeline=$true, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,

        [switch]$Expand
    )
    begin   {
        $result = @()
        $idToPartNo = @{}
        $partNoToID = @{}
        Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List | ForEach-Object {
            $idToPartNo[$_.SkuId]         = $_.SkuPartNumber
            $partNoToID[$_.SkuPartNumber] = $_.SkuId
        }
    }
    process {
        foreach ($s in $SKUID) {
            if      ($s.skuid) {$s = ($s.skuid) }
            elseif  ($s -notmatch '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$' -and $partNoToID[$s]) {
                     $s =  $partNoToID[$s]
            }
            elseif  ($s -notmatch '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$') {
                Write-Warning "$s doesn't look like a valid SKU" ; continue
            }
            $uri     = '$GraphUri/users?$Select=id,displayName,userPrincipalName,assignedLicenses&$filter=assignedLicenses/any(x:x/skuId eq {0})' -f  $s
            if ($expand) {
                $result +=  Invoke-GraphRequest -Uri $uri -ValueOnly
            }
            else {
                Invoke-GraphRequest -Uri $uri -ValueOnly -AsType "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser"
            }
        }
    }
    end     {
        if ($Expand -and $result) {
            $result | ForEach-Object {
                    $Upn = $_.userPrincipalName
                    foreach ($l in $_.assignedLicenses) {
                        New-Object psobject -Property ([ordered]@{UserPrincipalName = $UPN; SkuPartNumber = $idToPartNo[$l.skuID]})
                    }
            } | Sort-Object -Property  UserPrincipalName,SkuPartNumber -Unique
        }
    }
}
