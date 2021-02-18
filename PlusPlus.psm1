using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models

#region completer, transformer, and validator attributes for parameters
class UpperCaseTransformAttribute : ArgumentTransformationAttribute  {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$EngineIntrinsics, [object] $InputData) {
        if ($inputData -is [string]) {return $Inputdata.toUpper()}
        else                         {return ($InputData) }
    }
}

class ValidateCountryAttribute    : ValidateArgumentsAttribute {
    [void]Validate([object]$Argument, [EngineIntrinsics]$EngineIntrinsics)  {
        if ($Argument -notin [cultureInfo]::GetCultures("SpecificCultures").foreach({
                                New-Object -TypeName RegionInfo -ArgumentList $_.name
                             }).TwoLetterIsoRegionName) {
            Throw [ParameterBindingException]::new("'$Argument' is not an ISO 3166 country Code")
        }
    }
}

class DomainCompleter         : IArgumentCompleter {
    [array]$Domains = @()
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

class GroupCompleter          : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result = [System.Collections.Generic.List[CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        if ($wordToComplete) {$uri =  $script:GraphUri +  ("/Groups/?`$filter=startswith(displayName,'{0}') or startswith(mail,'{0}')" -f $wordToComplete)}
        else                 {$uri = "$script:GraphUri/Groups/?`$Top=20"}

        Invoke-GraphRequest -Uri $uri -ValueOnly | ForEach-Object displayname | Sort-Object | ForEach-Object {
                $result.Add(( New-Object -TypeName CompletionResult -ArgumentList "'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
        }

        return $result
    }
}

class TeamCompleter           : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result = [System.Collections.Generic.List[CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''


        if ($wordToComplete) {$uri =  $script:GraphUri +  ("/groups?`$filter=startswith(displayname,'{0}')')" -f $wordToComplete)}
        else                 {$uri = "$script:GraphUri/Groups/?`$Top=20"}
        #had "ResourceProvisioningOptions eq 'team' and " in the filter but it removed some valid teams so this is just completing groups for now
        Invoke-GraphRequest -Uri $uri -ValueOnly | ForEach-Object displayname | Sort-Object | ForEach-Object {
                $result.Add(( New-Object -TypeName CompletionResult -ArgumentList "'$_'", $_, ([CompletionResultType]::ParameterValue) , $_) )
        }

        return $result
    }
}

class OneDrivePathCompleter   : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result = [System.Collections.Generic.List[CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        If     ($wordToComplete -notmatch "/.+/" -or
                $wordToComplete -eq "/root:?/" )   {$params =@{folderPath = '/'} }
        elseif ($wordToComplete -Match '^/?root:') {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:'} } #catch after any leading / and before final /; and sandwich between / and :
        else                                       {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:'} } #catch after any leading / and before final /; and sandwich between /root/ and :

        if ($FakeBoundParameters['Drive']) {  $params['Drive'] = $FakeBoundParameters['Drive']}
        # #it would be better to order-by at the server, but consumer one drive doesn't support it.
        Get-GraphDrive -quiet @params | Sort-Object -Property name | ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                $result.Add(( New-Object -TypeName CompletionResult -ArgumentList "'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
        }
        return $result
    }
}

class OneDriveFolderCompleter : IArgumentCompleter {
    [System.Collections.Generic.IEnumerable[CompletionResult]] CompleteArgument(
        [string]$CommandName, [string]$ParameterName, [string]$WordToComplete,
        [Language.CommandAst]$CommandAst, [System.Collections.IDictionary] $FakeBoundParameters
    ) {
        $result = [System.Collections.Generic.List[CompletionResult]]::new()

        #strip quotes from word to complete - replace " or ' with nothing
        $wordToComplete = $wordToComplete -replace '"|''', ''

        If     ($wordToComplete -notmatch "/.+/" -or
                $wordToComplete -eq "/root:?/" )   {$params =@{folderPath = '/'} }
        elseif ($wordToComplete -Match '^/?root:') {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:'} } #catch after any leading / and before final /; and sandwich between / and :
        else                                       {$params =@{folderPath = $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:'} } #catch after any leading / and before final /; and sandwich between /root/ and :

        if ($FakeBoundParameters['Drive']) {  $params['Drive'] = $FakeBoundParameters['Drive']}
        # #it would be better to order-by at the server, but consumer one drive doesn't support it.
        Get-GraphDrive @params -subFolders -quiet | Sort-Object -Property name | ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                $result.Add(( New-Object -TypeName CompletionResult -ArgumentList "'$p'", $p, ([CompletionResultType]::ParameterValue) , $p) )
            }
        }
        return $result
    }
}

class SkuCompleter            : IArgumentCompleter {
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

class SkuPlanCompleter        : IArgumentCompleter {
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
#endregion


#Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
$global:DriveCache  = @{}

$Script:GraphUri    = "https://graph.microsoft.com/v1.0"
$Script:GUIDRegex   = "^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$"


. "$PSScriptRoot\Authentication.ps1"
 $ImportCmds = [ordered]@{
  'Users'                        = @('Get-MgUser_List' , 'New-MgUserTodoList_CreateExpanded',
                                     'New-MgUserTodoListTask_CreateExpanded', 'Remove-MgUserTodoList_Delete',
                                     'Remove-MgUserTodoListTask_Delete', 'Update-MgUserTodoListTask_UpdateExpanded')
  'Identity.DirectoryManagement' = @('Get-MgDomain_Get1', 'Get-MgDomain_List1', 'Get-MgDomainNameerenceByRef_List1',
                                     'Get-MgDomainServiceConfigurationRecord_List1' , 'Get-MgDomainVerificationDnsRecord_List1',
                                     'Get-MgOrganization_List1', 'Get-MgSubscribedSku_Get', 'Get-MgSubscribedSku_List')
  'Users.Functions'              = @()
  'Users.Actions'                = @()
  'Identity.SignIns'             = @()
  'Reports'                      = @()
}
#These need the class and/or private functions from the SDK module.
foreach ($subModule in $ImportCmds.keys) {
    $result = $null
    if (Test-path     (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll")) {
         $result = Import-Module (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll") -Cmdlet $ImportCmds[$subModule] -PassThru
    }
    # I do mean get module and assign it to module and if it works then... not "$module -eq"
    elseif ($module = Get-Module -ListAvailable "Microsoft.Graph.$submodule") {
         $result = Import-Module (Join-Path (Split-Path $module.Path) -ChildPath "bin\Microsoft.Graph.$submodule.private.dll") -Cmdlet $ImportCmds[$subModule]  -PassThru
    }
    else {Write-Verbose "Microsoft.Graph.$subModule.private.dll  not found $subModule won't be loaded "}
    if ($result) {
        .  "$PSScriptRoot\$subModule.ps1"
        foreach ($cmd in $ImportCmds[$module]) { (Get-Command $cmd).Visibility = 'Private'  }
    }
}

#These will work provided we have the users module.
. "$PSScriptRoot\Groups.ps1"
. "$PSScriptRoot\Notes.ps1"
. "$PSScriptRoot\OneDrive.ps1"
. "$PSScriptRoot\Planner.ps1"

Connect-Graph | Out-Host