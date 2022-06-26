[cmdletbinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory=$true,Position=0)]
    $key
)

Get-item *.dll |  Remove-Item -ErrorAction stop
if (Get-Item *.dll) {throw 'Remove DLLs' ; break}
if (Select-String -Path .\Microsoft.Graph.PlusPlus.settings.ps1 -Pattern 'clientsecret\s+["''](?!xxx)|TenantID\s+["''](?!xxx)') {
    throw "Settings contains secrets!"; break
}
#two chances to find things in settings
git update-index --no-skip-worktree .\Microsoft.Graph.PlusPlus.settings.ps1
if (-not (( git status ) -match "nothing to commit, working tree clean")) {throw "Unclean !" ; break}


$files = @{}
(Import-PowerShellDataFile .\Microsoft.Graph.PlusPlus.psd1 ).filelist | ForEach-Object {
    $i = Get-Item $_ -ErrorAction stop
    $files[$i.FullName] = $true
}

#we will use git to bring these  back
Remove-Item .\.vscode\* -Force -Recurse
Remove-Item .\.gitignore
Get-ChildItem -Recurse -File -Exclude $MyInvocation.MyCommand.name, .\.git |
    Where-Object {-not $files[$_.FullName] } |
     Remove-Item
if ($PSCmdlet.ShouldProcess('Publish')) {
    Publish-Module -NuGetApiKey $key -Repository PSGallery -AllowPrerelease -Name Microsoft.Graph.PlusPlus
}
 git reset HEAD --hard
 git update-index --skip-worktree .\Microsoft.Graph.PlusPlus.settings.ps1
