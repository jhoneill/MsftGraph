$script:PlannerTabJson =  @"
{
    "name": "{0}",
    "TeamsAppId":  "com.microsoft.teamspace.tab.planner",
    "configuration":  {
                          "entityId":  "{1}",
                          "contentUrl":  "https://tasks.office.com/{2}/Home/PlannerFrame?page=7\u0026planId={1}",
                          "websiteUrl":  "https://tasks.office.com/{2}/Home/PlannerFrame?page=7\u0026planId={1}",
                          "removeUrl":  "https://tasks.office.com/{2}/Home/PlannerFrame?page=7\u0026planId={1}
                      }
}
"@

Invoke-RestMethod -body ($script:PlannerTabJson -f $TabLabel,$Plan,$script:tennantID )

