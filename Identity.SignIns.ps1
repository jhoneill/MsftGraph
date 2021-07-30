using namespace Microsoft.Graph.PowerShell.Models
#MicrosoftGraphInvitation object is in Microsoft.Graph.Identity.SignIns.private.dll
function New-GraphInvitation   {
    <#
        .synopsis
            Invites an external user to become a guest in Azure AD
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The email address of the user being invited.
        #The characters  #~ ! $ %  ^  & * ( [ { < > } ] ) +  = \ /  | ; : " " ? , are not permitted
        #A  . or - is permitted except at the beginning or end of the name. A _  is permitted anywhere.
        [Parameter(Position=0,ValueFromPipeline=$true)]
        [string]$EmailAddress,
        #The display name of the user being invited.
        [string]$DisplayName,
        #The userType of the user being invited. By default, this is Guest. You can invite as Member if you are a company administrator.'
        [string]$UserType,
        #The URL the user should be redirected to once the invitation is redeemed. Required.
        [string]$RedirectUrl  = 'https://mysignins.microsoft.com/',
        #Indicates whether an email should be sent to the user being invited or not.
        [switch]$SendInvitationMessage
    )

    ContextHas -WorkOrSchoolAccount -BreakIfNot
    $settings = @{
        'invitedUserEmailAddress'    = $EmailAddress
        'sendInvitationMessage'      = $SendInvitationMessage -as [bool]
        'inviteRedirectUrl'          = $RedirectUrl
    }
    if ($DisplayName) {$settings['invitedUserDisplayName']  = $DisplayName}
    if ($UserType)    {$settings['invitedUserType']         = $UserType}

    $webparams = @{
        'Method'            = 'POST'
        'Uri'               = "$GraphUri/invitations"
        'Contenttype'       = 'application/json'
        'Body'              = (ConvertTo-Json $settings -Depth 5)
        'AsType'            = [MicrosoftGraphInvitation]
        'ExcludeProperty'   = '@odata.context'
    }
    Write-Debug $webparams.Body
    if ($force -or $pscmdlet.ShouldProcess($EmailAddress, 'Invite User')){
        try {
            $u = Invoke-GraphRequest @webparams
            if ($Passthru ) {return $u }
        }
        catch {
        # xxxx Todo figure out what errors need to be handled (illegal name, duplicate user)
        $_
        }
    }
}

function Get-GraphNamedLocation {
    <#
    .synopsis
        Returns named locations used in conditional access
    #>
    param   (
        #The ID or start of the display-name of a Location
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Location = ""
    )
    process {
        $uri = "$GraphUri/identity/conditionalAccess/namedlocations"
        foreach ($l in $Location) {
            if ($l -match $GUIDRegex) {
                $response = Invoke-GraphRequest "$uri/$l" -PropertyNotMatch '@odata'
            }
            elseif ($l) {
                $response = Invoke-GraphRequest  -ValueOnly  ($uri + "?`$filter=startswith(toLower(displayName),'{0}')" -f $l.ToLower())
            }
            else{
                $response = Invoke-GraphRequest -ValueOnly $uri
            }
            foreach ($r in $response) {
                if ($r.isTrusted -is [bool]) {
                    $trusted = $r.isTrusted
                    $null = $r.Remove('isTrusted')
                }
                else {$trusted = $null}
                if ($r.ipRanges) {
                    $ipranges =   $(foreach ($ipRange in $r.ipRanges) {$ipRange.cidrAddress}) -join ";"
                    $null = $r.Remove('ipranges')
                }
                else {$ipranges = $null}
                if ($r.countriesAndRegions) {
                    $countries =   $r.countriesAndRegions -join ";"
                    $null = $r.Remove('countriesAndRegions')
                }
                else {$countries = $null}
                if ($r.includeUnknownCountriesAndRegions -is [bool]) {
                    $includeUnknown =   $r.includeUnknownCountriesAndRegions
                    $null = $r.Remove('includeUnknownCountriesAndRegions')
                }
                else {$includeUnknown = $null }

                $null = $r.remove('@odata.type'),  $r.remove('@odata.id')
                New-object -TypeName  Microsoft.Graph.PowerShell.Models.MicrosoftGraphNamedLocation -Property $r  |
                        Add-Member -PassThru -NotePropertyName IsTrusted            -NotePropertyValue $trusted   |
                        Add-Member -PassThru -NotePropertyName IpRanges             -NotePropertyValue $ipranges  |
                        Add-Member -PassThru -NotePropertyName CountriesAndRegions  -NotePropertyValue $countries |
                        Add-Member -PassThru -NotePropertyName IncludeUnknownCountriesAndRegions  -NotePropertyValue $includeUnknown
            }
        }
    }
}

function Get-GraphConditionalAccessPolicy {
    param   (
        #The ID or start of the display-name of a Policy
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Policy = ""
    )
    process {
        #needs beta to work when conditions.devices is set (which is marked as preview in the GUI)
        $uri = "beta/identity/conditionalAccess/policies"
        foreach ($p in $Policy) {
            if ($p -match $GUIDRegex) {
                    Invoke-GraphRequest "$uri/$p" -AsType ([MicrosoftGraphConditionalAccessPolicy]) -PropertyNotMatch '@odata'
            }
            elseif ($P) {
                   $uri += ("?`$filter=startswith(toLower(displayName),'{0}')" -f $p.ToLower())
                  Invoke-GraphRequest -ValueOnly $uri            -AsType ([MicrosoftGraphConditionalAccessPolicy])
            }
            else{ Invoke-GraphRequest -ValueOnly $uri -AllValues -AsType ([MicrosoftGraphConditionalAccessPolicy]) }
        }
    }
}

function Expand-GraphConditionalAccessPolicy {
    param (
        #The ID or start of the display-name of a Policy
        [Parameter(ValueFromPipeline=$true,Position=0)]
        $Policy = ""
    )
    begin {
        $locations          = @{}
        Invoke-GraphRequest "$GraphUri/identity/conditionalAccess/namedlocations" -ValueOnly -AllValues |
            ForEach-Object {$locations[$_.id]         = $_.DisplayName}
        $dirRoleTemplates   = @{} ;
        Invoke-GraphRequest  -Uri "$GraphUri/directoryroletemplates/" -ValueOnly -AllValues  |
            ForEach-Object {$dirRoleTemplates[$_.id]  = $_.DisplayName}
        $servicePrincipals  = @{}
        Invoke-GraphRequest  -Uri "$GraphUri/servicePrincipals/" -ValueOnly -AllValues  |
            ForEach-Object {$servicePrincipals[$_.id] = $_.Displayname}

        $dirobjs            = @{}
        $result             = @()
        function resolveDirObj {
            param (
                [Parameter(ValueFromPipeline=$true)] $D
            )
            process {
                if ($D -notmatch $GUIDRegex ) {return $d}
                elseif (-not $dirobjs.ContainsKey($D)) {
                    $dirobjs[$d] = (Invoke-GraphRequest "$GraphUri/directoryObjects/$d").displayname
                }
                $dirobjs[$d]
            }
        }
        function tranlasteGUID {
            param (
                [Parameter(ValueFromPipeline=$true)]$G,
                [Parameter(Position=0)]$hash
            )
            process {if ($g -notmatch $GUIDRegex ) {return $g} else {$hash[$g]}}
        }
    }
    process {
        Get-GraphConditionalAccessPolicy @PSBoundParameters |
            Select-Object -Property DisplayName , Description, State,
                @{n='IncludeUsers';             e={($_.Conditions.Users.IncludeUsers                | resolveDirObj ) -join '; '}},
                @{n='ExcludeUsers';             e={($_.Conditions.Users.ExcludeUsers                | resolveDirObj ) -join '; '}},
                @{n='IncludeGroups';            e={($_.Conditions.Users.IncludeGroups               | resolveDirObj ) -join '; '}},
                @{n='ExcludeGroups';            e={($_.Conditions.Users.ExcludeGroups               | resolveDirObj ) -join '; '}},
                @{n='IncludeRoles';             e={($_.Conditions.Users.IncludeRoles                | tranlasteGUID $dirRoleTemplates)  -join '; '}},
                @{n='ExcludeRoles';             e={($_.Conditions.Users.ExcludeRoles                | tranlasteGUID $dirRoleTemplates)  -join '; '}},
                @{n='IncludeLocations';         e={($_.Conditions.Locations.IncludeLocations        | translateGuid $locations)         -join '; '}},
                @{n='ExcludeLocations';         e={($_.Conditions.Locations.ExcludeLocations        | translateGuid $locations)         -join "; "}},
                @{n='IncludeApplications';      e={($_.Conditions.Applications.IncludeApplications  | tranlasteGUID $servicePrincipals) -join '; '}},
                @{n='ExcludeApplications';      e={($_.Conditions.Applications.ExcludeApplications  | tranlasteGUID $servicePrincipals) -join '; '}},
                @{n='IncludeUserActions';       e={ $_.Conditions.Applications.IncludeUserActions     -join '; '}},
                @{n='IncludeDevices';           e={ $_.Conditions.Devices.IncludeDevices              -join '; '}},
                @{n='ExcludeDevices';           e={ $_.Conditions.Devices.ExcludeDevices              -join '; '}},
                @{n='ClientAppTypes';           e={ $_.Conditions.ClientAppTypes                      -join '; '}},
                @{n='IncludePlatforms';         e={ $_.Conditions.Platforms.IncludePlatforms          -join '; '}},
                @{n='ExcludePlatforms';         e={ $_.Conditions.Platforms.EccludePlatforms          -join '; '}},
                @{n='AppEnforcedRestrictions';  e={ $_.SessionControls.ApplicationEnforcedRestrictions.IsEnabled}},
                @{n='PersistentBrowser'; e={
                                              if   ($_.SessionControls.PersistentBrowser.isenabled) {
                                                    $_.SessionControls.PersistentBrowser.Mode}
                                              else {$null}}},
                @{n='CloudAppSecurity';  e={
                                              if   ($_.SessionControls.CloudAppSecurity.isenabled)  {
                                                    $_.SessionControls.CloudAppSecurity.CloudAppSecurityType}
                                              else {$null}}} ,
                @{n='SignInFrequency';   e={
                                              if   ($_.SessionControls.SignInFrequency.isenabled)  {
                                                    $_.SessionControls.SignInFrequency.Value + " " +
                                                    $_.SessionControls.SignInFrequency.Type }
                                            else {$null}}}
    }
}