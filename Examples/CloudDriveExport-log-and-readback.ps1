param (
$Outfile = "./clouddrive/auditrecords.json",  # for cloudshell
[switch]$Read,
$Infile = "s:\auditrecords.json"
)
###
if ($Outfile -and -not $Read) {
    if (-not (Get-Module -ListAvailable microsoft.graph.authentication)) {Install-Module microsoft.graph.authentication -Force }
    $logonResult =  Connect-MgGraph -AccessToken (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
    if ($logonresult -notmatch 'Welcome') {Write-Warning 'Failed to log on, bailing out.';return}
    $uri = 'v1.0/auditLogs/directoryAudits'   # $top=100'
    (Invoke-MgGraphRequest -Uri $uri).value | convertto-json -depth 100 | out-file $Outfile
}
if ($Infile -and $Read) {
    $records = (ConvertFrom-Json (Get-Content s:\auditrecords.json -Raw) -AsHashtable)
    foreach ($r in $records) {
        New-Object -TypeName  Microsoft.Graph.PowerShell.Models.MicrosoftGraphDirectoryAudit -Property $r |
        Add-Member -PassThru -MemberType ScriptProperty -Name User -Value {$this.initiatedBy.user.userPrincipalName} |
        Add-Member -PassThru -MemberType ScriptProperty -Name App  -Value {$this.initiatedBy.App.DisplayName}
    }
}