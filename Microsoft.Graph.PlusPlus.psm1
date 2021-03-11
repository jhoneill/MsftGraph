using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models
using namespace Microsoft.Graph.PowerShell.Authentication

$global:GraphUri                  = 'https://graph.microsoft.com/v1.0'   #May want this outside the module
$script:GUIDRegex                 = '^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$'
$script:WellKnownMailFolderRegex  = '^[/\\]?(archive|clutter|conflicts|conversationhistory|deleteditems|drafts|inbox|junkemail|localfailures|msgfolderroot|outbox|recoverableitemsdeletions|scheduled|searchfolders|sentitems|serverfailures|syncissues)[/\\]?$'
$script:SkippedSubmodules         = @()

#region completer, transformer, and validator attributes for parameters **CLASSES NEED TO BE IN PSM1
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

class DomainCompleter             : IArgumentCompleter {
    [array]$Domains = @()
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    )
    {
        $result = [System.Collections.Generic.List[System.Management.Automation.CompletionResult]]::new()
        if (-not $this.Domains)  {$this.Domains = Invoke-GraphRequest "$global:GraphUri/domains?`$select=id" -ValueOnly |
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

        if ($wordToComplete) {$uri =  $global:GraphUri +  ("/Groups/?`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $wordToComplete)}
        else                 {$uri = "$global:GraphUri/Groups/?`$Top=20"}

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
        Invoke-GraphRequest  -Uri "$global:GraphUri/directoryroles?`$select=displayname" -ValueOnly |
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
        if ($wordToComplete) {$uri =  $global:GraphUri +  ("/groups?`$filter=startswith(displayname,'{0}')')" -f $wordToComplete)}
        else                 {$uri = "$global:GraphUri/Groups/?`$Top=20"}
        #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams so this is just completing groups for now
        Invoke-GraphRequest -Uri $uri -ValueOnly | ForEach-Object displayname | Sort-Object | ForEach-Object {
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
            Invoke-GraphRequest -ValueOnly -headers @{'ConsistencyLevel'='eventual'} -uri "$global:GraphUri/users?`$filter=startswith(userprincipalname,'$wordToComplete')&top=10&select=userprincipalName" |
                ForEach-Object userPrincipalName | sort-object | ForEach-Object {$result.Add([System.Management.Automation.CompletionResult]::new("'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )}
        }
        return $result
    }
}
#endregion

. "$PSScriptRoot\Authentication.ps1"

#Submodules need the class and/or private functions from the SDK module.
$ImportCmds = [ordered]@{
  'Users'                        = @('Get-MgUser_List1' , 'New-MgUserTodoList_CreateExpanded1',
                                     'New-MgUserTodoListTask_CreateExpanded1', 'Remove-MgUserTodoList_Delete1',
                                     'Remove-MgUserTodoListTask_Delete1', 'Update-MgUserTodoListTask_UpdateExpanded1')
  'Identity.DirectoryManagement' = @('Get-MgDomain_Get1', 'Get-MgDomain_List1', 'Get-MgDomainNameerenceByRef_List1',
                                     'Get-MgDomainServiceConfigurationRecord_List1' , 'Get-MgDomainVerificationDnsRecord_List1',
                                     'Get-MgOrganization_List1', 'Get-MgSubscribedSku_Get1', 'Get-MgSubscribedSku_List1')
  'Users.Functions'              = @()
  'Users.Actions'                = @()
  'Identity.SignIns'             = @()
  'Reports'                      = @()
  'Applications'                 =('Get-MgServicePrincipal_Get2','Get-MgServicePrincipal_List1')
}
foreach ($subModule in $ImportCmds.keys) {
    $result = $null
    if (Test-path     (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll")) {
        $result = Import-Module (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll") -Cmdlet $ImportCmds[$subModule] -PassThru
    }
    # I do mean get module and assign it to module and if it works then... not "$module -eq"
    elseif ($module = Get-Module -ListAvailable "Microsoft.Graph.$submodule" | Sort-Object -Property Version | Select-Object -Last 1) {
        $result = Import-Module (Join-Path (Split-Path $module.Path) -ChildPath "bin\Microsoft.Graph.$submodule.private.dll") -Cmdlet $ImportCmds[$subModule]  -PassThru
    }
    else {
        Write-Verbose "Microsoft.Graph.$subModule.private.dll  not found $subModule won't be loaded "
        $script:SkippedSubmodules += $subModule
    }
    if ($result) {
        .  "$PSScriptRoot\$subModule.ps1"
        foreach ($cmd in $ImportCmds[$subModule]) { (Get-Command $cmd).Visibility = 'Private'  }
    }
}

if ($script:SkippedSubmodules -contains 'Users') {
     Write-Verbose "Groups, Notes, OneDrive, Planner, and Sharepoint require the Microsoft.Graph.users module, or Microsoft.Graph.Users.private.dll in the module directory."
     $script:SkippedSubmodules += @('Groups', 'Notes', 'OneDrive', 'Planner', 'Sharepoint')
}
else { #These will work provided we have the users module.
    . "$PSScriptRoot\Groups.ps1"
    . "$PSScriptRoot\Notes.ps1"
    . "$PSScriptRoot\OneDrive.ps1"
    . "$PSScriptRoot\Planner.ps1"
    . "$PSScriptRoot\Sharepoint.ps1"
}
if ($script:SkippedSubmodules) {
      Write-Host ("Did not load " + ($script:SkippedSubmodules -join ", ") + " because the related Microsoft.Graph module(s) or private.dll file(s) were not be found.")
}

#call a script with calls to Set-GraphConnectionOptions
if     ($env:MGSettingsPath )                       {. $env:MGSettingsPath}
elseif (Test-Path "$PSScriptRoot\AuthSettings.ps1") {. "$PSScriptRoot\AuthSettings.ps1"}

if ($null -eq [GraphSession]::instance.AuthContext) {Write-Host  "Ready for Connect-Graph."}
elseif ([GraphSession]::instance.AuthContext.AppName -and -not [GraphSession]::instance.AuthContext.Account) {
      Write-Host ("Already logged on as the app '$([GraphSession]::instance.AuthContext.AppName)'." )}
else {Write-Host ("Already logged on as '$([GraphSession]::instance.AuthContext.Account)'." )}
