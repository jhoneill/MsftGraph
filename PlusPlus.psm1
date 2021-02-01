
$Script:GraphUri  = "https://graph.microsoft.com/v1.0"
$Script:GUIDRegex = "^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$"

. "$PSScriptRoot\Authentication.ps1"


foreach ($subModule in @('Users','Groups','Identity.DirectoryManagement')) {
    $result = $null
    if (Test-path     (Join-path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll")) {
        $result = Import-Module (Join-Path $PSScriptRoot -ChildPath "Microsoft.Graph.$subModule.private.dll") -PassThru
    }
    # I do mean get module and assign it to module and if it works ... not "$module -eq"
    elseif ($module = Get-Module -ListAvailable "Microsoft.Graph.$submodule") {
        $result = Import-Module (Join-Path (Split-Path $module.Path) -ChildPath "bin\Microsoft.Graph.$submodule.private.dll") -PassThru
    }
    else {Write-Verbose "Microsoft.Graph.$subModule.private.dll  not found $subModule won't be loaded "}
    if ($result) {.  "$PSScriptRoot\$subModule.ps1"}
}
