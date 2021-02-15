$paths = @()
$ParentHash = @{}
function hasParent {
    param($path)
    if ($path -notmatch "/") { return}
    $parent = $path -replace '^(.*)/.+?$','$1'
    if (-not $ParentHash[$parent]) {hasParent $parent }
    $ParentHash[$path] = $parent ; return $Path -replace '\(.*\)$','()' # if it ends function (params) remove params
}

Get-Item *.yml | ForEach-Object  {
    $y = convertfrom-yaml (Get-Content $_ -Raw)
    $top =  $y.servers.url -replace '/$', ''
    $ParentHash[$top] = $true
    $paths += $y.paths.Keys  |
                 Where-Object {($_ -notmatch '/\$\w+$|delta\(\)$')  -and #Filter out $ref etc and graph.delta() endings
                               ($_ -notmatch 'id1\}/|event-id\}/.') -and #Filter out id1 or event-id as these cause too much expansion
                              (($_ -split '/').count -lt 8)  } |         #Don't go too deep.
                    ForEach-Object {                                     #if we have get and post and patch and delete show them all.
                        if ($y.paths[$_].keys.count -eq 1)      { hasparent ($top + $_)}
                        else {foreach ($k in $y.paths[$_].keys) { hasParent( $top + $_ + "/" + $k )}}
                    } | Sort-Object
}
#may get duplication between yaml files
$paths = $paths | Sort-Object -unique

graph -Name g -Attributes @{rankdir="LR"; }  -ScriptBlock {
    #create
    node -name     $top -Attributes @{
          label = ($top);
          shape = 'folder'; fontname  = 'Segoe UI';
          style = 'filled'; fillcolor = 'lightyellow'; }

    foreach ($p in $paths) {
        node -Name $p -Attributes @{
           label = ($p -split "/")[-1]
           shape = 'folder'; fontname = 'Segoe UI'
           style ='filled' ; fillcolor='lightyellow' }
        edge -NodeName ($p  -replace '^(.*)/.+?$','$1') `
             -To        $p
    }
} | Export-PSGraph -ShowGraph -OutputFormat jpg
