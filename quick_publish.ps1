[cmdletbinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory=$true,Position=0)]
    $key
)

Remove-Item *.dll -ErrorAction SilentlyContinue
if (Get-Item *.dll) {throw 'Remove DLLs' ; break}
if (-not (( git status ) -match "nothing to commit, working tree clean")) {throw "Unclean !" ; break}

$files = @{}
(Import-PowerShellDataFile .\Microsoft.Graph.PlusPlus.psd1 ).filelist | ForEach-Object {
    $i = Get-Item $_ -ErrorAction stop
    $files[$i.FullName] = $true
}

#we will use git to bring these  back
Remove-Item .\.vscode\* -Force -Recurse
Remove-Item .\.gitignore
Get-ChildItem -Recurse -File -Exclude $MyInvocation.MyCommand.name |
    Where-Object {-not $files[$_.FullName] } |
     Remove-Item
if ($PSCmdlet.ShouldProcess('Publish')) {
    Publish-Module -NuGetApiKey $key -Repository PSGallery -AllowPrerelease -Name Microsoft.Graph.PlusPlus
}