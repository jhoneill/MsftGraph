#$azContext = Get-AzContext
#$azProfile = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile
#$profileClient = New-Object -TypeName Microsoft.Azure.Commands.ResourceManager.Common.RMProfileClient -ArgumentList ($azProfile)
#$token = $profileClient.AcquireAccessToken($azContext.Subscription.TenantId)
#$authHeader = @{'Content-Type'='application/json' ; 'Authorization'='Bearer ' + $token.AccessToken }
#$restUri = "https://management.azure.com/subscriptions/$($azContext.Subscription.id)?api-version=2020-01-01"
Invoke-RestMethod -Uri $restUri -Method Get -Headers $authHeader
<#
id                   : /subscriptions/007104fc-5081-48d6-a1ff-aa2e24a50f69
authorizationSource  : Legacy
managedByTenants     : {}
subscriptionId       : 007104fc-5081-48d6-a1ff-aa2e24a50f69
tenantId             : e6af5578-6d03-49e0-af3b-383cf5ec0b5f
displayName          : Microsoft Partner Network
state                : Enabled
subscriptionPolicies : @{locationPlacementId=Public_2014-09-01; quotaId=MPN_2014-09-01; spendingLimit=On}
#>
#Invoke-RestMethod -Method Post -Headers $authHeader -Body "{ ""subscriptions"": [""$($azContext.Subscription.Id)"" ],    ""query"": ""Resources | project name, type | limit 5""}" -Uri 'https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2019-04-01' | select -expand data
#$d =Invoke-RestMethod -Method Post -Headers $authHeader -Body "{ ""subscriptions"": [""$($azContext.Subscription.Id)"" ],    ""query"": ""Resources | project name, type | order by name asc | limit 5""}" -Uri 'https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2019-04-01' | select -expand data


$at         = Get-AccessToken -Resoure "https://management.azure.com" -GrantType "password" -BodyParts @{username=$user ;password=$passwd}
$authHeader = @{'Content-Type'='application/json' ; 'Authorization'='Bearer ' + $at.access_token }
$restUri    = "https://management.azure.com/subscriptions?api-version=2020-01-01"
$sub        = [pscustomobject]((Invoke-RestMethod -Uri $restUri -Headers $authHeader).value | Select-Object -Last 1 )

$query = "Resources | project name, type | order by name asc | limit 5"
$webparams = @{
    'body'   = (ConvertTo-Json @{'subscriptions' = @($sub.subscriptionId) ; query =$query})
    'Uri'    = 'https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2019-04-01'
    'Method' = 'Post'
    'Headers'=  $authHeader }
$d =Invoke-RestMethod @webparams  | select -expand data

<#
columns                                                rows
-------                                                ----
{@{name=name; type=string}, @{name=type; type=string}} {AzureAutomationTutorialScript microsoft.automation/automationaccounts/runbo…
#>

 foreach ($row in $d.rows) {$h = @{}; foreach ($c in  0..($d.columns.Count-1)) {$h[$d.columns[$c].name] = $row[$c] } [pscustomobject]$h}
<#
name                                  type
----                                  ----
Application Insights Smart Detection  microsoft.insights/actiongroups
AzureAutomationTutorial               microsoft.automation/automationaccounts/runbooks
AzureAutomationTutorialPython2        microsoft.automation/automationaccounts/runbooks
AzureAutomationTutorialScript         microsoft.automation/automationaccounts/runbooks
Failure Anomalies - mobulaazfunchello microsoft.alertsmanagement/smartdetectoralertrules
#>
