
$global:GraphUri  = "https://graph.microsoft.com/v1.0"
$global:GUIDRegex = "^\{?[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}\}?$"

#Sometimes when we want to convert an opaque drive ID (e.g. on a file or folder) to a name; save extra calls to the server by caching the id-->name
$global:drivecache  = @{}

. "$PSScriptRoot\Authentication.ps1"

foreach ($subModule in @('Users','Users.Functions','Users.Actions','Identity.DirectoryManagement','Reports')) {
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

. "$PSScriptRoot\Groups.ps1"
. "$PSScriptRoot\Notes.ps1"
. "$PSScriptRoot\OneDrive.ps1"
. "$PSScriptRoot\Planner.ps1"
#conv
Update-FormatData -AppendPath "$PSScriptRoot\Graph.format.ps1xml"

