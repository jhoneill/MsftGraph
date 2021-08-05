Import-Module Microsoft.Graph.PlusPlus
$cred = Import-Clixml ~\Documents\PowerShell\mycred.xml # created with $cred | Export-CliXML -path mycred.xml
Connect-Graph -Credential $cred #this will fail

edit .\Microsoft.Graph.PlusPlus.settings.ps1  #Show what's in the settings file

. ..\msftgraph.safe\Microsoft.Graph.PlusPlus.settings.ps1  #use a settings file with my TenantID ClientID and ClientSecret in it.

Connect-Graph -Credential $cred   # Now it works
GWhoAmI

Get-GraphUser | Format-Table Organization

Get-GraphGroup 'Consultants' -Drive   # Example group which failed when we tried the azure logon
Get-GraphUser -Teams  |  Format-Table displayname,description  # Teams current users is in
Get-GraphTeam 'Consultants'  -site  -ov myteamsite   # OV so we can see the value and use it in the next command
Get-GraphSite $myteamsite -Lists

Get-Graphsite $myteamSite -Lists |Where-Object name -like "prob*" -ov problemslist # Example list on team site

Get-GraphList $problemslist -ColumnList
Get-GraphList $problemslist -Items -Property title,issuestatus,priority
$teamplan = Get-GraphTeam 'Consultants' -Plans | Where-Object title -like "team*"
Get-GraphPlan $teamplan -Buckets
Add-GraphPlanTask -Plan $teamplan -Bucket "To Do" -Title "Submit comments on new spec" -DueDate ([datetime]::Now).AddDays(7) -Links $teamdrive.webUrl
