#Parameters validated... then
if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
        Invoke-RestMethod -Method Post -Headers  @{Authorization = "Bearer $Script:AccessToken"} -Uri "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs" -ContentType "application/json" -body  @"
{
    "name": "$TabLabel",
    "TeamsAppId":  "com.microsoft.teamspace.tab.planner",
    "configuration":  {
                          "entityId":  "$Plan",
                          "contentUrl":  "https://tasks.office.com/$Script:TenantId/Home/PlannerFrame?page=7\u0026planId=$PlanID",
                          "websiteUrl":  "https://tasks.office.com/$Script:TenantId/Home/PlannerFrame?page=7\u0026planId=$PlanID",
                          "removeUrl":  "https://tasks.office.com/$Script:TenantId/Home/PlannerFrame?page=7\u0026planId=$PlanID"
                      }
}
"@

}