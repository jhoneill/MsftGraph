#Parameters validated... then
$tabURI = "https://tasks.office.com/$Script:TenantId /Home/PlannerFrame?page=7&planId=$Plan"

$webparams = @{'Method'      = 'Post';
               'Uri'         = "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs" ;
               'Headers'     =  $Script:DefaultHeader;
               'ContentType' = 'application/json'
}

$json = ConvertTo-Json ([ordered]@{
            'name'          = $TabLabel
            'TeamsAppId'    = 'com.microsoft.teamspace.tab.planner'
            'configuration' = [ordered]@{
               'entityId'   = $plan
               'contentUrl' = $tabURI
               'websiteUrl' = $tabURI
               'removeUrl'  = $tabURI
            }
        })
Write-Debug $json
if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
    $result = Invoke-RestMethod @webParams -body $json
    $result.pstypeNames.add('GraphTab')
    return $result
}