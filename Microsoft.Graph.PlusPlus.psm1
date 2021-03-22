using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
using namespace Microsoft.Graph.PowerShell.Authentication

$Global:GraphUri                  =   'https://graph.microsoft.com/v1.0'   #Global: instead of Script: for use in cmdline invoke-graphRequest etc.
$Script:GUIDRegex                 =   '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$'
$Script:WellKnownMailFolderRegex  =   '^[/\\]?(' +  (@(
                                            'archive',       'clutter',       'conflicts', 'conversationhistory',
                                            'deleteditems',  'drafts',        'inbox',     'junkemail',
                                            'localfailures', 'msgfolderroot', 'outbox',    'recoverableitemsdeletions',
                                            'scheduled',     'searchfolders', 'sentitems', 'serverfailures',
                                            'syncissues'
                                        ) -join '|' ) + ')[/\\]?$'
$Script:DefaultUserProperties     = @(
                                        'businessPhones',   'displayName',    'givenName',  'id',  'jobTitle', 'mail',
                                        'mobilePhone',      'officeLocation', 'preferredLanguage', 'surname',  'userPrincipalName',
                                        'assignedLicenses', 'department',     'usageLocation',     'userType'
                                    )
$Script:DefaultUsageLocation      =   'GB'
$Script:SkippedSubmodules         = @(    )

#region global helper functions, completer, transformer, and validator attributes for parameters **CLASSES NEED TO BE IN PSM1
class UpperCaseTransformAttribute : ArgumentTransformationAttribute  {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$EngineIntrinsics, [object] $InputData) {
        if ($inputData -is [string]) {return $Inputdata.toUpper()}
        else                         {return ($InputData) }
    }
}

class ValidateCountryAttribute    : ValidateArgumentsAttribute {
    [void]Validate([object]$Argument, [EngineIntrinsics]$EngineIntrinsics)  {
        if ($Argument -notin [cultureInfo]::GetCultures("SpecificCultures").foreach({
                                New-Object -TypeName System.Globalization.RegionInfo -ArgumentList $_.name
                             }).TwoLetterIsoRegionName) {
            Throw [ParameterBindingException]::new("'$Argument' is not an ISO 3166 country Code")
        }
    }
}

class ChannelCompleter            : IArgumentCompleter {
    [string] $GroupID = ''
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''
        if (-not $this.GroupID) {
            $Group = $null
            if ($FakeBoundParameters['Group']) {$Group = $FakeBoundParameters['Group']}
            if ($FakeBoundParameters['Team' ]) {$Group = $FakeBoundParameters['Team']}
            #I do mean = not -eq in the elseif statements.
            elseif ($key = $Global:PSDefaultParameterValues.Keys.where({"$CommandName`:Team"  -like $_})) {
                $Group = $Global:PSDefaultParameterValues[$key]
            }
            elseif ($key = $Global:PSDefaultParameterValues.Keys.where({"$CommandName`:Group" -like $_})) {
                $Group = $Global:PSDefaultParameterValues[$key]
            }
            if     ($Group.ID)                       { $this.Groupid = $Group.id}
            elseif ($Group -is [string] -and
                    $Group -match $Script:GUIDRegex) { $this.GroupID = $Group}
            elseif ($Group -is [string])             { $this.GroupID = idfromteam $Group }
        }
        if ($this.groupID -and $this.groupID -match $Script:GUIDRegex) {
            Invoke-GraphRequest "$global:GraphUri/Teams/$($this.groupID)/Channels?$`select=id,displayname" -ValueOnly |
                ForEach-Object {
                    if ($_.displayname  -like "$wordToComplete*") {$_.displayName}
                } | Sort-Object |
                    ForEach-Object {
                        $result.Add([System.Management.Automation.CompletionResult]::new("'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
                    }
        }
        return $result
    }
}

class DomainCompleter             : IArgumentCompleter {
    [array]$Domains = @()
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    )
    {
        $result = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        if (-not $this.Domains)  {$this.Domains = Invoke-GraphRequest "$Global:GraphUri/domains?`$select=id" -ValueOnly |
                                                         ForEach-Object id | Sort-Object
        }
        $wildcard          = ('*' + ($wordToComplete  -replace '[''"]','' )+ '*')

        foreach ($d in $this.domains.where({$_ -like $wildcard}))  {$result.Add([System.Management.Automation.CompletionResult]::new($_))}
        return $result
    }
}

class GroupCompleter              : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        if ($wordToComplete) {$uri =  $Global:GraphUri +  ("/Groups/?`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $wordToComplete)}
        else                 {$uri = "$Global:GraphUri/Groups/?`$Top=20"}

        Invoke-GraphRequest -Uri $uri -ValueOnly | ForEach-Object displayname | Sort-Object | ForEach-Object {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
        }

        return $result
    }
}

class MailFolderCompleter         : IArgumentCompleter {
     [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
         [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
         [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
     ) {
         $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

         #strip quotes from word to complete - replace " or ' with nothing.
         $wordToComplete = $wordToComplete -replace '"|''', ''
         #Where interested in what's before the final / or \
         $params = @{'Select' = 'displayname'}
         $path = ''
         if ($wordToComplete -match '^[/\\]?(\w.*)[/\\].*?$') {
             $params['Name'] = $Matches[1];
             $params['ChildItems'] =$true
             $path   = $Matches[1] + '/'
         }
         if ($FakeBoundParameters['User']) {  $params['User'] = $FakeBoundParameters['User']}
         Get-GraphMailFolder @params | ForEach-Object {
             $p = $path+$_.displayname
             if ($p -like "$wordToComplete*") {
                 $result.Add([System.Management.Automation.CompletionResult]::new("'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
             }
         }
         return $result
     }
}

class OneDriveFolderCompleter     : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        If     ($wordToComplete -notmatch "/.+/" -or
                $wordToComplete -eq "/root:?/" )   {$params =@{folderPath = '/'} }
        elseif ($wordToComplete -Match '^/?root:') {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:'} } #catch after any leading / and before final /; and sandwich between / and :
        else                                       {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:'} } #catch after any leading / and before final /; and sandwich between /root/ and :

        if ($FakeBoundParameters['Drive']) {  $params['Drive'] = $FakeBoundParameters['Drive']}
        #I do mean = no -eq in the next line.
        elseif ($key = $Global:PSDefaultParameterValues.Keys.where({"$CommandName`:Drive" -like $_})) {
            $params['Drive'] = $Global:PSDefaultParameterValues[$key]
        }
        # #it would be better to order-by at the server, but consumer one drive doesn't support it.
        Get-GraphDrive @params -subFolders -quiet | Sort-Object -Property name | ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
        }
        return $result
    }
}

class OneDrivePathCompleter       : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        If     ($wordToComplete -notmatch "/.+/" -or
                $wordToComplete -eq "/root:?/" )   {$params =@{folderPath = '/'} }
        elseif ($wordToComplete -Match '^/?root:') {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:'} } #catch after any leading / and before final /; and sandwich between / and :
        else                                       {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:'} } #catch after any leading / and before final /; and sandwich between /root/ and :

        if ($FakeBoundParameters['Drive']) {  $params['Drive'] = $FakeBoundParameters['Drive']}
        #I do mean = no -eq in the next line.
        elseif ($key = $Global:PSDefaultParameterValues.Keys.where({"$CommandName`:Drive" -like $_})) {
            $params['Drive'] = $Global:PSDefaultParameterValues[$key]
        }
        # #it would be better to order-by at the server, but consumer one drive doesn't support it.
        Get-GraphDrive -quiet @params | Sort-Object -Property name | ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
        }
        return $result
    }
}

class OneNoteSectionCompleter     : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''
        $values = @()
        if ($FakeBoundParameters['Notebook'] -and $FakeBoundParameters['Notebook'].Sections  ) {$values=$FakeBoundParameters['Notebook'].Sections.DisplayName}
        #I do mean = no -eq in the next line.
        elseif ($key = $Global:PSDefaultParameterValues.Keys.where({"$CommandName`:Notebook" -like $_})) {
            $values = $Global:PSDefaultParameterValues[$key].Sections.DisplayName
        }
        foreach ($p in $values) {
            if ($P -like "$wordToComplete*" -and $p -match '^\w+$') {
                $result.Add([System.Management.Automation.CompletionResult]::new($p, $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
            elseif ($P -like "$wordToComplete*") {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
        }
        return $result
    }
}

class SkuCompleter                : IArgumentCompleter {
    [array]$skus = @()
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

class SkuPlanCompleter            : IArgumentCompleter {
    [array]$skus = @()
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

class RoleCompleter               : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        if (-not $wordToComplete) {$wordToComplete = '*'}
        else                      {$wordToComplete = "$wordToComplete*" -replace '"|''', '' }
        Invoke-GraphRequest  -Uri "$Global:GraphUri/directoryroles?`$select=displayname" -ValueOnly |
            Where-Object displayname -like $wordToComplete | Sort-Object -Property displayname | ForEach-Object {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$($_.displayname)'", $_.displayname, ([CompletionResultType]::ParameterValue) , $_.displayname) )
        }
        return $result
    }
}

class TeamCompleter               : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        if ($wordToComplete) {$uri = "$Global:GraphUri/groups?`$select=id,resourceProvisioningOptions,displayname&`$filter=startswith(displayname,'{0}')" -f $wordToComplete}
        else                 {$uri = "$Global:GraphUri/groups?`$select=id,resourceProvisioningOptions,displayname"}
        #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams so this is just completing groups for now

        Invoke-GraphRequest -Uri $uri -ValueOnly |
            ForEach-Object {if ("Team" -in $_.resourceProvisioningOptions) {$_.displayname}} |
                Sort-Object | ForEach-Object {
                $result.Add([System.Management.Automation.CompletionResult]::new("'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
        }
        return $result
    }
}

class UPNCompleter                : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result =  [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''
        if ($wordToComplete) {
            Invoke-GraphRequest -ValueOnly -headers @{'ConsistencyLevel'='eventual'} -uri "$Global:GraphUri/users?`$filter=startswith(userprincipalname,'$wordToComplete')&top=10&select=userprincipalName" |
                ForEach-Object userPrincipalName | sort-object | ForEach-Object {$result.Add([System.Management.Automation.CompletionResult]::new("'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )}
        }
        return $result
    }
}

function FilterString {
    param (
        [validatescript({
            if ($_ -is [string] -and $_ -match '\*.*\*|^\*$') {throw [ParameterBindingException]::new("Wildcard cannot be '*something*' or just '*'")} else {$true}
        })]
        [parameter(position=0,mandatory=$true)]
        [string]$SearchTerm ,
        [parameter(position=1)]
        $ExtraFields = @(),
        [switch]$ToLower
    )
    if ($toLower) {$SearchTerm = $SearchTerm.ToLower() }
    #Replace '  with '' - ensure we don't turn '' into '''' !
    $SearchTerm = $SearchTerm -replace "(?<!')'(?!')" ,"''"
    #validation blocked "* and *something*" so we have no *, * at the start, in the middle, or at the end
    if     ($SearchTerm -notmatch '\*')         {$filterStrings = ,              "displayName eq '$SearchTerm'"     }
    elseif ($SearchTerm -match   '^\*(.+)')     {$filterStrings = ,     "endswith(displayName,'$($Matches[1])')"    }
    elseif ($SearchTerm -match   '(.+)\*$')     {$filterStrings = ,   "startswith(displayName,'$($Matches[1])')"    }
    elseif ($SearchTerm -match  '^(.+)\*(.+)$') {$filterStrings = , ("(startswith(displayName,'$($Matches[1])')" +
                                                                " and endswith(displayName,'$($Matches[2])'))"  )}
    if ($ToLower) {$filterStrings[0] = $filterStrings[0] -replace 'displayName' , 'toLower(displayName)'}

    foreach ($f in $ExtraFields) {$filterStrings += $filterStrings[0]   -replace 'displayName',$f }
    $filterStrings -join ' or '
}
#endregion

#region load the bulk of the module
. "$PSScriptRoot\Authentication.ps1"

#Submodules which need the class and/or private functions from the SDK module.
$ImportCmds = [ordered]@{
  'PersonalContacts'             = @()
  'Users'                        = @('New-MgUserTodoList_CreateExpanded1','New-MgUserTodoListTask_CreateExpanded1', 'Remove-MgUserTodoList_Delete1',
                                     'Remove-MgUserTodoListTask_Delete1', 'Update-MgUserTodoListTask_UpdateExpanded1') #'Get-MgUser_List1' ,
  'Identity.DirectoryManagement' = @('Get-MgDomain_Get1', 'Get-MgDomain_List1', 'Get-MgDomainNameerenceByRef_List1',
                                     'Get-MgDomainServiceConfigurationRecord_List1' , 'Get-MgDomainVerificationDnsRecord_List1',
                                     'Get-MgOrganization_List1', 'Get-MgSubscribedSku_Get1', 'Get-MgSubscribedSku_List1')
  'Users.Functions'              = @()
  'Users.Actions'                = @()
  'Identity.SignIns'             = @()
  'Reports'                      = @()
  'Applications'                 = @()
}
foreach ($subModule in $ImportCmds.keys) {
    $result = $null
    if (Test-path     (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll")) {
        $result = Import-Module -Scope Local (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll") -Cmdlet $ImportCmds[$subModule] -PassThru
    }
    # I do mean get module and assign it to module and if it works then... not "$module -eq"
    elseif ($module = Get-Module -ListAvailable "Microsoft.Graph.$submodule" | Sort-Object -Property Version | Select-Object -Last 1) {
        $result = Import-Module -Scope Local (Join-Path (Split-Path $module.Path) -ChildPath "bin\Microsoft.Graph.$submodule.private.dll") -Cmdlet $ImportCmds[$subModule]  -PassThru
    }
    else {
        Write-Verbose "Microsoft.Graph.$subModule.private.dll  not found $subModule won't be loaded "
        $Script:SkippedSubmodules += $subModule
    }
    if ($result) {
        .  "$PSScriptRoot\$subModule.ps1"
        foreach ($cmd in $ImportCmds[$subModule]) { $c = Get-Command $cmd; $c.set_visibility('Private')  }
    }
}

if ($Script:SkippedSubmodules -contains 'Users') {
     Write-Verbose "Groups, Notes, OneDrive, Planner, and Sharepoint require the Microsoft.Graph.users module, or Microsoft.Graph.Users.private.dll in the module directory."
     $Script:SkippedSubmodules += @('Groups', 'Notes', 'OneDrive', 'Planner', 'Sharepoint')
}
else { #These submodules will work with just the users module.
    . "$PSScriptRoot\Groups.ps1"
    . "$PSScriptRoot\Notes.ps1"
    . "$PSScriptRoot\OneDrive.ps1"
    . "$PSScriptRoot\Planner.ps1"
    . "$PSScriptRoot\Sharepoint.ps1"
}
if ($Script:SkippedSubmodules) {
      Write-Host -ForegroundColor DarkGray ("Skipped " + ($Script:SkippedSubmodules -join ", ") + " because their Microsoft.Graph module(s) or private.dll file(s) were not found.")
}
#endregion

function Set-GraphOptions {
    <#
        .synopsis
            Sets defaults and the tenant client ID & Client Secret used when logging on without a web dialog
    #>
    [cmdletbinding()]
    param (
        #Your Tennant ID
        $TenantID,
        #Client ID if not using the SDK default of 14d82eec-204b-4c2f-b7e8-296a70dab67e. Must be known to your tennant
        $ClientID,
        #Secret set for the client ID in your $TenantID
        [Alias('Client_Secret,')]
        $ClientSecret,
        #Default Scopes to request
        $DefaultScopes,
        #Allows a saved Refresh Token (e.g. from Show-GraphSession) to be added to the session.
        $RefreshToken,
        #Changes the dafault properties returned by Get-GraphUser and Get-GraphUserList
        [validateSet('accountEnabled', 'ageGroup', 'assignedLicenses', 'assignedPlans', 'businessPhones', 'city',
                    'companyName', 'consentProvidedForMinor', 'country', 'createdDateTime', 'department',
                    'displayName', 'givenName', 'id', 'imAddresses', 'jobTitle', 'legalAgeGroupClassification',
                    'mail', 'mailNickname', 'mobilePhone', 'officeLocation',
                    'onPremisesDomainName', 'onPremisesExtensionAttributes', 'onPremisesImmutableId',
                    'onPremisesLastSyncDateTime', 'onPremisesProvisioningErrors', 'onPremisesSamAccountName',
                    'onPremisesSecurityIdentifier', 'onPremisesSyncEnabled', 'onPremisesUserPrincipalName',
                    'passwordPolicies', 'passwordProfile', 'postalCode', 'preferredDataLocation',
                    'preferredLanguage', 'provisionedPlans', 'proxyAddresses', 'state', 'streetAddress',
                    'surname', 'usageLocation', 'userPrincipalName', 'userType')]
        [string[]]$DefaultUserProperties,

        #Changes the default two letter (ISO  3166) country code - for new users so they can be assigned licenses.  Examples include: 'US', 'JP', and 'GB'
        [ValidateNotNullOrEmpty()]
        [string]$DefaultUsageLocation
    )

    if     ($TenantID)              {
        if ($TenantID -notmatch $GUIDRegex) {Write-Warning 'TenantID should be a GUID'  ; break }
        else {$Script:TenantID           = $TenantID}
    }
    if     ($ClientID)              {
        if ($Clientid -notmatch $GUIDRegex) {Write-Warning 'ClientID should be a GUID'  ; break}
        else {$Script:ClientID  = $ClientID}
    }
    if     ($ClientSecret)          {
        if     ($ClientSecret -is [string]) {
               $Script:ClientSecret      = $ClientSecret
        }
        elseif ($ClientSecret -is [securestring]) {
               $Script:ClientSecret =  (new-object pscredential -ArgumentList "NoName", $ClientSecret).GetNetworkCredential().Password
        }
        else  {Write-Warning 'ClientSecret should be a string or preferably a securestring'  ; break}
    }
    if     ($Script:TenantID)       {Write-Verbose "TenantID: '$Script:TenantID' , ClientID: '$Script:ClientID'"}
    if     ($DefaultScopes)         {$Script:DefaultGraphScopes     = $DefaultScopes}
    Write-Verbose ('Scopes: ' + ($Script:DefaultGraphScopes -join ', '))

    #it would be nice to the use the country validator but this goes wrong when reloading the module and calling something when everything is happening in the PSM1 file.
    if     ($DefaultUsageLocation -and -not [cultureInfo]::GetCultures("SpecificCultures").where({$_.name -match "$DefaultUsageLocation$"})) {
           Write-Warning 'DefaultUsageLocation should be an ISO 2 letter country code like GB, US or JP'  ; break
    }
    elseif ($DefaultUsageLocation ) {$Script:DefaultUsageLocation   = $DefaultUsageLocation.ToUpper() }
    if     ($DefaultUserProperties) {$Script:DefaultUserProperties  = $DefaultUserProperties}
    if     ($RefreshToken)          {$Script:RefreshToken           = $RefreshToken }
}

#call a script with calls to Set-GraphOptions
if     ($env:GraphSettingsPath )                       {. $env:GraphSettingsPath}
elseif (Test-Path "$PSScriptRoot\Microsoft.Graph.PlusPlus.settings.ps1") {. "$PSScriptRoot\Microsoft.Graph.PlusPlus.settings.ps1"}

if      ($null -eq [GraphSession]::instance.AuthContext) {Write-Host  "Ready for Connect-Graph."}
elseif ([GraphSession]::instance.AuthContext.AppName -and -not [GraphSession]::instance.AuthContext.Account) {
      Write-Host ("Already logged on as the app '$([GraphSession]::instance.AuthContext.AppName)'." )}
else {Write-Host ("Already logged on as '$([GraphSession]::instance.AuthContext.Account)'." )}
