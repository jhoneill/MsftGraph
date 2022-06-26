using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
#Uses functions from  and MicrosoftGraphSubscribedSku type from  Microsoft.Graph.Identity.DirectoryManagement.private.dll

#xxxx todo: check context is a WorkOrSchool account and that it has the right scopes and warn / error / throw if not.
function Get-GraphDomain                {
    <#
      .synopsis
        Gets domains in the current tenant
      .Description
        Requires consent to use at least the Directory.Read.All scope
    #>
    [OutputType([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDomain])]
    [cmdletbinding(DefaultParameterSetName='None')]
    param (
        [parameter(Position=0, ValueFromPipeline=$true, ParameterSetName='Domain',    Mandatory=$true)]
        [parameter(Position=0, ValueFromPipeline=$true, ParameterSetName='VDRecords', Mandatory=$true)]
        [parameter(Position=0, ValueFromPipeline=$true, ParameterSetName='SCRecords', Mandatory=$true)]
        [parameter(Position=0, ValueFromPipeline=$true, ParameterSetName='NameRef',   Mandatory=$true)]
        [ArgumentCompleter([DomainCompleter])]
        $Domain,

        [parameter(ParameterSetName='VDRecords',Mandatory=$true)]
        [alias('VR')]
        [switch]$VerificationDNSRecords,

        [parameter(ParameterSetName='SCRecords',Mandatory=$true)]
        [switch]$ServiceConfigurationRecords,

        [parameter(ParameterSetName='NameRef',Mandatory=$true)]
        [switch]$NameReferenceList
    )
    Test-GraphSession
    if (-not $Domain) {Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomain_List1 @PSBoundParameters}
    else {
        #Allow an orgnaization object to be piped in.
        if ($Domain.verifiedDomains) {$Domain = $Domain.verifiedDomains}
        $null = $PSBoundParameters.Remove("Domain")
        foreach ($d in $Domain) {
            if     ($d.id)              {$d = $d.id}
            elseif ($d.name)            {$d = $d.name}
            elseif ($d -isnot [String]) {Write-Warning -Message 'Could not find the Domain ID from the parameter'}
            if     ($VerificationDNSRecords)      {
                $null = $PSBoundParameters.Remove("VerificationDNSRecords")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainVerificationDnsRecord_List1 -DomainId $d @PSBoundParameters
            }
            elseif ($ServiceConfigurationRecords) {
                $null = $PSBoundParameters.Remove("ServiceConfigurationRecords")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainServiceConfigurationRecord_List1 -DomainId $d @PSBoundParameters
            }
            elseif ($NameReferenceList)           {
                $null = $PSBoundParameters.Remove("NameReferenceList")
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomainNameReference_List1 -DomainId $d @PSBoundParameters
            }
            else   {
                Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgDomain_Get1 -DomainId $d @PSBoundParameters
            }
        }
    }
}

function Get-GraphOrganization          {
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
    param (
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
    Test-GraphSession
    Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgOrganization_List1 @PSBoundParameters
}

function Get-GraphSKU                   {
    <#
      .Synopsis
        Gets details of SKUs that an organization has subscribed to
      .Example
        Get-GraphSKU "enterprise*" -ServicePlans | sort servicePlanName | format-table
        Finds any SKU with a name starting "Enterprise" and displays its service plans in alphabetical order.
    #>
    [Alias('Get-GraphSubscribedSku')]
    [OutputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphSubscribedSku])]
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position = 0, ValueFromPipeline=$true)]
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
        Test-GraphSession
        $result = @()
    }
    process {
        foreach ($s in $sku) {
            $null = $PSBoundParameters.Remove("ServicePlans") ,  $PSBoundParameters.Remove("SKU")
            if ($s.skuid)          {$s = ($s.skuid) }
            if ($s -notmatch $GuidRegex) {
                $result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List @PSBoundParameters |
                            Where-Object -Property SkuPartNumber -like $s
            }
            elseif ($s -is [String]) {
                $result += Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get1  @PSBoundParameters -SubscribedSkuId $s}
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

function Grant-GraphLicense             {
    <#
      .Synopsis
        Grants the licence to use a particular stock-keeping-unit (SKU) to users or groups
    #>
    [cmdletbinding(SupportsShouldprocess=$true,DefaultParameterSetName='ByUserID')]
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position=0, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,

        #ID(s) for users to receive permission ("me" will select the current user), the command will accept user objects and attempt to resolve names to IDs
        [parameter(Position=1,  ParameterSetName='ByUserID', ValueFromPipeline=$true, Mandatory = $true)]
        [ArgumentCompleter([UPNCompleter])]
        $UserID ,

        #ID(s) for group(s) to receive permission, the command will accept group objects and attempt to resolve names to IDs
        [parameter(Position=2, ParameterSetName='ByGroupID', Mandatory = $true)]
        [ArgumentCompleter([GroupCompleter])]
        [Alias("Team")]
        $GroupID,

        #Disables individual parts of the the SKU
        [ArgumentCompleter([SkuPlanCompleter])]
        [string[]]$DisabledPlans,

        #A two letter country code (ISO standard 3166). Examples include: 'US', 'JP', and 'GB' Can be set/reset here
        [ValidateNotNullOrEmpty()]
        [UpperCaseTransformAttribute()]
        [ValidateCountryAttribute()]
        [string]$UsageLocation,

        #Runs the command without a confirmation dialog
        [Switch]$Force
    )
    begin   {
        Test-GraphSession
        $request        = @{'addLicenses' = @() ; 'removeLicenses' = @()}
        $SkuPartNumbers = @()
        foreach  ($s in $SKUID) {
            if   ($s.skuid) {$s = ($s.skuid) }
            if   ($s -match $GuidRegex) {
                  $sku = Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_Get1 -SubscribedSkuId $s
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
            $SkuPartNumbers += $sku.SkuPartNumber
        }
        $licensePartNos = $SkuPartNumbers -join ", "
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
        if ($UserID -is [string] -and $userid -notmatch "me|\w@\w|$GUIDRegex" ) {
            $userId = Get-GraphUser $UserID
        }
        foreach ($u in $UserID ) {
            #region Add the user to web parameters: allow for mulitple users - potentially with an ID or a UPN
            if ($u -eq "me") {
                    $baseUri  = "$GraphUri/me/"
                    $userDisplayName    =  $Global:GraphUser
            }
            elseif ($u.id)  {
                    $baseUri  = "$GraphUri/users/$($u.id)/"
                    $userDisplayName    = $u.Id  #hope to change this if we have a display name
            }
            elseif ($u.UserPrincipalName) {
                    $baseUri  = "$GraphUri/users/$($u.UserPrincipalName)/"
                    $userDisplayName    = $u.UserPrincipalName  #hope to change this if we have a display name
            }
            elseif ($u -is [string] -and $u -match "\w@\w|$GUIDRegex") {
                    $baseUri  = "$GraphUri/users/$u/"
                    $userDisplayName    = $u
            }
            elseif ($u -is [string]) {
                $u = Get-GraphUser $u
                if ($u.count -eq 1) {
                    $baseUri  = "$GraphUri/users/$($u.id)/"
                }
                else {
                    Write-Warning "Could not resolve $u to a single user. Ignoring"
                    continue
                }
            }
            $webparams['uri']  = $baseUri + "assignLicense"
            if ($u.DisplayName) {$userDisplayName = $u.DisplayName }

            if ($UsageLocation -and ($Force -or $Pscmdlet.Shouldprocess($userdisplayname,"Set usage location to '$UsageLocation'."))) {
                $null = Invoke-GraphRequest -Method PATCH -Uri $baseUri -ContentType 'application/json' -body ('{{"usageLocation": "{0}"}}' -f $UsageLocation)
            }
            if ($Force -or $Pscmdlet.Shouldprocess($userdisplayname,"License $licensePartNos to user")) {
                $u = Invoke-GraphRequest  @webparams -SkipHttpErrorCheck
                if ($u.error) {Write-Warning "Licensing $licensePartNos to user '$userDisplayName' caused error '$($u.error.message)'."  }
                else          {Write-Verbose "GRANTGRAPHLICENSE: $licensePartNos  Granted to $($u.userPrincipalName)"            }
            }
        }
        if ($Groupid -is [String]  -and  $GroupID -Notmatch $GUIDRegex)  {$groupID = Get-GraphGroup -Group $GroupID -NoTeamInfo }
        foreach ($g in $GroupID) {
            if ($g.SecurityEnabled -eq $false ) {
                Write-Warning "$($g.DisplayName) is not a security group. Only Security groups can be licensed." ; Continue
            }
            if ($g.ID) {
                    $webparams['uri']   = "$GraphUri/groups/$($g.id)/assignLicense"
                    $groupDisplayName   = $g.Id
            }
            elseif ($g -is [string] -and $g -match $GUIDRegex) {
                    $webparams['uri']   = "$GraphUri/groups/$g/assignLicense"
                    $groupDisplayName   = $g
            }
            elseif ($g -is [string]) {
                $g = Get-GraphGroup -Group $g -NoTeamInfo
                if ($g.count -eq 1 -and $g.SecurityIdentifier) {
                    $webparams['uri']   = "$GraphUri/groups/$($g.id)/assignLicense"
                }
                else {
                    Write-Warning "Could not resolve $g to a single Security group. Ignoring"
                    continue
                }
            }
            else {
                    Write-Warning "$g does not seem to be a Group. Ignoring"
                    continue
            }
            if ($g.DisplayName) {$groupDisplayName = $g.DisplayName }
            if ($Force -or $Pscmdlet.Shouldprocess($groupDisplayName,"License $licensePartNos to user")) {
                $g = Invoke-GraphRequest  @webparams -SkipHttpErrorCheck
                if ($g.error) {Write-Warning "Licensing $licensePartNos to group '$groupDisplayName' caused error '$($g.error.message)'."  }
                else          {Write-Verbose "GRANT-GRAPHLICENSE: $licensePartNos granted to group '$groupDisplayName'."            }
            }
        }
    }
}

function Revoke-GraphLicense            {
    <#
      .Synopsis
        Revokes a users or groups licences to use a particular stock-keeping-unit (SKU)
    #>
    [cmdletbinding(SupportsShouldprocess=$true,DefaultParameterSetName='ByUserID')]
    param   (
        #The SKU to revoke either as an ID or a SKU object containing an ID
        [parameter(Position=0, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,

        #ID for the user (required. "me" will select the current user)
        [parameter(Position=1, ParameterSetName='ByUserID', ValueFromPipeline=$true, Mandatory = $true)]
        [ArgumentCompleter([UPNCompleter])]
        $UserID ,

        #ID(s) for group(s) to receive permission, the command will accept group objects and attempt to resolve names to IDs
        [parameter(Position=2, ParameterSetName='ByGroupID', Mandatory = $true)]
        [ArgumentCompleter([GroupCompleter])]
        [Alias("Team")]
        $GroupID,

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
    begin   {
        Test-GraphSession
        $request        = @{'addLicenses' = @() ; 'removeLicenses' = @()}
        foreach ($s in $SKUID) {
            if  ($s.skuid) {$s = ($s.skuid) }
            if  ($s -match $GuidRegex) {
                $request.removeLicenses += $s
            }
            else {
                 $sku = Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List  |
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
                    $userDisplayName    =  $Global:GraphUser
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
                Write-Verbose "REVOKE-GRAPHUSERLICENSE - licence(s) for $($request.removeLicenses.Count) SKU(s) from $($u.userPrincipalName)"
            }
        }
        if ($Groupid -is [String]  -and  $GroupID -Notmatch $GUIDRegex)  {$groupID = Get-GraphGroup -Group $GroupID -NoTeamInfo }
        foreach ($g in $GroupID) {
            if ($g.SecurityEnabled -eq $false ) {
                Write-Warning "$($g.DisplayName) is not a security group. Only Security groups can be licensed." ; Continue
            }
            if ($g.ID) {
                    $webparams['uri']   = "$GraphUri/groups/$($g.id)/assignLicense"
                    $groupDisplayName   = $g.Id
            }
            elseif ($g -is [string] -and $g -match $GUIDRegex) {
                    $webparams['uri']   = "$GraphUri/groups/$g/assignLicense"
                    $groupDisplayName    = $g
            }
            elseif ($g -is [string]) {
                $g = Get-GraphGroup -Group $g -NoTeamInfo
                if ($g.count -eq 1 -and $g.SecurityIdentifier) {
                    $webparams['uri'] = "$GraphUri/groups/$($g.id)/assignLicense"
                }
                else {
                    Write-Warning "Could not resolve $g to a single Security group. Ignoring"
                    continue
                }
            }
            else {
                    Write-Warning "$g does not seem to be a Group. Ignoring"
                    continue
            }
            if ($g.DisplayName) {$groupDisplayName = $g.DisplayName }
            if ($Force -or $Pscmdlet.Shouldprocess($groupDisplayName,"Revoke licence(s) for $($request.removeLicenses.Count) SKU(s)")) {
                $g = Invoke-GraphRequest  @webparams -SkipHttpErrorCheck
                if ($g.error) {Write-Warning "Licensing $licensePartNos to group '$groupDisplayName' caused error '$($g.error.message)'."  }
                else          {Write-Verbose "REVOKE-GRAPHLICENSE: licence(s) for $($request.removeLicenses.Count) SKU(s) from group '$groupDisplayName'."            }
            }
        }
    }
}

function Get-GraphLicense               {
    <#
      .Synopsis
        Returns users or groups (or both) who are licensed to user a given SKU
    #>
    [cmdletbinding(DefaultParameterSetName='None')]
    param   (
        #The SKU to get either as an ID or a SKU object containing an ID
        [parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true)]
        [ArgumentCompleter([SkuCompleter])]
        $SKUID ,
        [Parameter(ParameterSetName='Users')]
        [switch]$UsersOnly,
        [Parameter(ParameterSetName='Groups')]
        [switch]$GroupsOnly
    )
    begin   {
        Test-GraphSession
        $result = @()
        $idToPartNo = @{}
        $partNoToID = @{}
        Microsoft.Graph.Identity.DirectoryManagement.private\Get-MgSubscribedSku_List1 | ForEach-Object {
            $idToPartNo[$_.SkuId]         = $_.SkuPartNumber
            $partNoToID[$_.SkuPartNumber] = $_.SkuId
        }
    }
    process {
        foreach ($s in $SKUID) {
            if      ($s.skuid) {$s = ($s.skuid) }
            elseif  ($s -notmatch $GuidRegex -and $partNoToID[$s]) {
                     $s =  $partNoToID[$s]
            }
            elseif  ($s -notmatch $GuidRegex) {
                Write-Warning "$s doesn't look like a valid SKU" ; continue
            }

            $uri     = $GraphUri + '/users?$Select=id,displayName,userPrincipalName,assignedLicenses&$filter=assignedLicenses/any(x:x/skuId eq {0})' -f  $s
            if ($UsersOnly) {Invoke-GraphRequest -Uri $uri -ValueOnly -AsType ([MicrosoftGraphUser]) }
            elseif (-not $GroupsOnly) {
                $result +=  Invoke-GraphRequest -Uri $uri -ValueOnly
            }

            $uri     = $GraphUri + '/groups?$Select=id,displayName,assignedLicenses&$filter=assignedLicenses/any(x:x/skuId eq {0})' -f  $s
            if ($GroupsOnly) {Invoke-GraphRequest -Uri $uri -ValueOnly -AsType ([MicrosoftGraphGroup]) }
            elseif (-not $UsersOnly) {
                $result +=  Invoke-GraphRequest -Uri $uri -ValueOnly
            }
         }
    }
    end     {
        if ($result -and -not ($UsersOnly -or $GroupsOnly)) {
            $result | ForEach-Object {
                    $upn = $_.userPrincipalName
                    $displayName = $_.displayName
                    foreach ($l in $_.assignedLicenses) {
                        New-Object psobject -Property ([ordered]@{
                            'DisplayName'       = $DisplayName
                            'UserPrincipalName' = $upn
                            'SkuPartNumber'     = $idToPartNo[$l.skuID]
                            'SkuID'             = $l.skuID  })
                    }
            } | Sort-Object -Property  UserPrincipalName,DisplayName,SkuPartNumber -Unique | Where-Object {$_.skupartnumber -in $SKUID -or $_.skuid -in $SKUID}
        }
    }
}

function Get-GraphDirectoryRole         {
<#
    .synopsis
        Gets an Azure AD directory role or its members
    .example
        PS C:\> Get-GraphDirectoryRole external* -Members | ft displayname,role
        Lists all members of groups whose names begin "external"
        The command adds the role name to the user object making it possible
        to show the roles and names in the output.
#>
    param   (
        #The role to get, either as a display name (wildcards allowed), an ID, or a Role object containing an ID
        [parameter(ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([RoleCompleter])]
        $Role = '*',
        #If specified returns the members of the role as user objects
        [switch]$Members
    )
    process {
        if     ($Role.count -gt 1) {
            $Role | Get-GraphDirectoryRole -Members:$Members
            return
        }
        if     ($Role -is [MicrosoftGraphDirectoryRole]) {$roles = $Role}
        elseif ($Role -is [string] -and $role -match $GUIDRegex) {
            $roles = Invoke-GraphRequest  -Uri "$GraphUri/directoryroles/$Role"        -AsType  ([MicrosoftGraphDirectoryRole] ) -ExcludeProperty '@odata.context'
        }
        else {
            $roles = Invoke-GraphRequest  -Uri "$GraphUri/directoryroles" -ValueOnly    -AsType  ([MicrosoftGraphDirectoryRole] )  |
                        Where-Object -Property displayName -like $role
        }
        #removed ?`$expand=members as it only expands 20.
        if      (-not $members) {$roles}
        else {
            foreach($r in $roles) {
                $memberlist =  igr "$GraphUri/directoryroles/$($r.id)/members" -ValueOnly
                foreach ($u in $memberlist.where({$_.'@odata.type'-match 'user$'})) {
                    $null = $u.Remove('@odata.type') ,  $u.remove('@odata.id')
                    New-object -type MicrosoftGraphUser -Property $u |
                        Add-member -NotePropertyName Role -NotePropertyValue $r.DisplayName -PassThru
                }
                foreach ($g in $memberlist.where({$_.'@odata.type'-match 'group$'})) {
                    $null = $g.Remove('@odata.type') ,  $g.remove('@odata.id'), $g.remove('@odata.context'), $g.Remove('GroupName'), $g.remove('creationOptions')
                    New-object -type MicrosoftGraphGroup -Property $g |
                        Add-member -NotePropertyName Role -NotePropertyValue $r.DisplayName -PassThru
                }
            }
        }
    }
}

function Grant-GraphDirectoryRole       {
    <#
      .synopsis
        Grants a directory role to a user or group
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
        #The role(s) to revoke, either as role names or a role objects.
        [Parameter(Position=0,Mandatory=$true)]
        [ArgumentCompleter([RoleCompleter])]
        $Role ,
        #The member to add, can be a user name, or an object representing either a group with IsAssignableToRole set or a user.
        [Parameter(ValueFromPipeline=$true,Position=1,Mandatory=$true)]
        [ArgumentCompleter([UPNCompleter])]
        $Member ,
        #Runs the command with no confirmation.
        [switch]$Force
    )
    begin   {
        $Role = $Role | Get-GraphDirectoryRole
    }
    process {
        foreach ($m in $Member) {
            if (-not $m.id) {$m = Get-GraphUserList -Name $m}
            if (-not $m -or $m.count -gt 1 -or -not $m.Id) {
                Write-Warning "Could not process the role member." ; return
            }
            foreach ($r in $role) {
                $body = ConvertTo-Json @{ '@odata.id' = "$graphUri/directoryObjects/$($m.Id)" }
                Write-Debug $body
                if ($Force -or $pscmdlet.ShouldProcess($m.displayname,"Grant access to role '$($r.displayname)'")) {
                    try   { Invoke-GraphRequest -Uri "$graphuri/directoryroles/$($r.id)/members/`$ref" -Method post -Body $body -ContentType 'application/json'}
                    catch { Write-Warning "The request failed. This may be because the member '$($m.toString())' has already been added to the '$($r.displayname)' role." }
                }
            }
        }
    }
}

function Revoke-GraphDirectoryRole      {
    <#
     .synopsis
       Removes a user or group from a an Azure AD directory role
     #>
    [cmdletbinding(SupportsShouldProcess=$True,ConfirmImpact='High')]
    param   (
        [Parameter(Position=0,Mandatory=$true)]
        #The role(s) to revoke, either as role names or a role objects.
        [ArgumentCompleter([RoleCompleter])]
        $Role ,
        [Parameter(ValueFromPipeline=$true,Position=1)]
        #The member to remove , can be a user name, or a user or group object
        $Member,
        #Runs the command without confirmation.
        [switch]$Force
    )
    begin   {
        $Role = $Role | Get-GraphDirectoryRole
    }
    process {
        if (-not $member.id) {$member = Get-GraphUserList -Name $Member}
        if ($member.count -ne 1 -or -not $member.Id) {
            Write-Warning "Could not process the role member."
        }
        foreach ($r in $role) {
            if ($Force -or $pscmdlet.ShouldProcess($Member.displayname,"Revoke access from role '$($r.displayname)'")) {
                try   {Invoke-GraphRequest -Uri "$graphuri/directoryroles/$($role.id)/members/$($member.Id)/`$ref" -Method Delete}
                catch {Write-Warning "The request failed. This may be because the member was no in thethe role"}
            }
        }
    }
}

function Get-GraphDirectoryRoleTemplate {
    <#
      .SYNOPSIS
        Gets directory role templates
    #>
    param    (
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Template = ""
    )
    process {
        $uri = "$GraphUri/identity/directoryroletemplates"
        foreach ($t in $Template) {
            if ($t -match $GUIDRegex) {
                Invoke-GraphRequest "$uri/$t" -AsType  ([MicrosoftGraphDirectoryRoleTemplate] )
            }
            elseif ($t) {
                $uri += ("?`$filter=startswith(toLower(displayName),'{0}')" -f $t.ToLower())
                Invoke-GraphRequest  -ValueOnly $uri  -AsType  ([MicrosoftGraphDirectoryRoleTemplate] )
            }
            else{
                Invoke-GraphRequest -ValueOnly $uri -AsType  ([MicrosoftGraphDirectoryRoleTemplate] )
            }
        }
    }
}

function Get-GraphDeletedObject         {
    <#
      .synopsis
        Returns deleted users or groups from the AAD recycle bin
      .description
        It can filter by name, and selects users by default or groups if -Group is selected
        The results can be piped into Restore-GraphDeletedObject
    #>
    param (
        #If specified filters the returned objects to those with a name starts with...
        $Name,
        #By default user objects are returned. This switches the choice to group objects.
        [switch]$Group
    )
    if ($name)  {$u    = '?$filter=' +(FilterString $Name)}
    else        {$u    = ''}
    if ($Group) {$type = 'Group'} else {$type='User'}
    Invoke-GraphRequest -Uri "$GraphUri/directory/deleteditems/microsoft.graph.$type$u" -AsType ([pscustomobject])  -ValueOnly
}

 function Restore-GraphDeletedObject     {
    <#
      .synopsis
        Recovers a user or group from the AAD recycle bin
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
        #A deleted object or the ID of one.
        [Parameter(ValueFromPipeline=$true,position=0)]
        $ID,
        #Specifies that the ID is associated with a group, not a user.
        [switch]$Group,
        #If specified supresses any confirmation prompt
        [switch]$Force
    )
    process {
        if ($id.displayname) {$displayname = $id.Displayname} else {$displayname = ''}
        if ($id.id) {$id = $id.id}
        if ($Force -or $PSCmdlet.ShouldProcess($displayname,'Recover directory object')) {
            Invoke-GraphRequest "$GraphUri/directory/deleteditems/$id/restore" -Method Post -body ' ' -AsType ([pscustomobject])
        }
    }
}
#  DELETE /directory/deletedItems/{id}                permanent delete
