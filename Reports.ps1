using namespace Microsoft.Graph.PowerShell.Models

function Get-GraphReport       {
    <#
        .Synopsis
            Gets reports from MS Graph
        .Example
            >Get-GraphReport -Report MailboxUsageDetail | ft "Display Name",  "Storage Used (Byte)"
            Displays mailbox storage used by users - note that
            fields have 'friendly' names which need to be wrapped in quotes
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param   (
        #The report to Fetch
        [ValidateSet(
                'EmailActivityCounts', 'EmailActivityUserCounts', 'EmailActivityUserDetail',
                'EmailAppUsageAppsUserCounts', 'EmailAppUsageUserCounts', 'EmailAppUsageUserDetail', 'EmailAppUsageVersionsUserCounts',
                'MailboxUsageDetail', 'MailboxUsageMailboxCounts', 'MailboxUsageQuotaStatusMailboxCounts', 'MailboxUsageStorage',
                'Office365ActivationCounts', 'Office365ActivationsUserCounts', 'Office365ActivationsUserDetail',
                'Office365ActiveUserCounts', 'Office365ActiveUserDetail',  'Office365GroupsActivityCounts', 'Office365GroupsActivityDetail',
                'Office365GroupsActivityFileCounts', 'Office365GroupsActivityGroupCounts', 'Office365GroupsActivityStorage',
                'Office365ServicesUserCounts',
                'OneDriveActivityFileCounts', 'OneDriveActivityUserCounts', 'OneDriveActivityUserDetail', 'OneDriveUsageAccountCounts',
                'OneDriveUsageAccountDetail', 'OneDriveUsageFileCounts', 'OneDriveUsageStorage',
                'SharePointActivityFileCounts', 'SharePointActivityPages', 'SharePointActivityUserCounts',
                'SharePointActivityUserDetail', 'SharePointSiteUsageDetail', 'SharePointSiteUsageFileCounts',
                'SharePointSiteUsagePages', 'SharePointSiteUsageSiteCounts', 'SharePointSiteUsageStorage',
                'SkypeForBusinessActivityCounts', 'SkypeForBusinessActivityUserCounts', 'SkypeForBusinessActivityUserDetail',
                'SkypeForBusinessDeviceUsageDistributionUserCounts','SkypeForBusinessDeviceUsageUserCounts', 'SkypeForBusinessDeviceUsageUserDetail',
                'SkypeForBusinessOrganizerActivityCounts', 'SkypeForBusinessOrganizerActivityMinuteCounts',
                'SkypeForBusinessOrganizerActivityUserCounts', 'SkypeForBusinessParticipantActivityCounts',
                'SkypeForBusinessParticipantActivityMinuteCounts', 'SkypeForBusinessParticipantActivityUserCounts',
                'SkypeForBusinessPeerToPeerActivityCounts', 'SkypeForBusinessPeerToPeerActivityMinuteCounts', 'SkypeForBusinessPeerToPeerActivityUserCounts',
                'TeamsDeviceUsageDistributionUserCounts', 'TeamsDeviceUsageUserCounts', 'TeamsDeviceUsageUserDetail',
                'TeamsUserActivityCounts', 'TeamsUserActivityUserCounts', 'TeamsUserActivityUserDetail',
                'YammerActivityCounts', 'YammerActivityUserCounts', 'YammerActivityUserDetail', 'YammerDeviceUsageDistributionUserCounts',
                'YammerDeviceUsageUserCounts', 'YammerDeviceUsageUserDetail', 'YammerGroupsActivityCounts',
                'YammerGroupsActivityDetail', 'YammerGroupsActivityGroupCounts'
        )]
        [parameter(Mandatory=$true)]
        $Report,
        #Date for the report - this should be a date in the past 30 days. If specified, -Period is ignored. Reports ending in Count, Storage or pages don't support date filtering
        [DateTime]$Date,
        #The range of time for the report in the form "Dn" where n is the number of days. The default is D7, except for Office365Activation activation reports
        [ValidateSet("D7", "D30", "D90", "D180")]
        $Period,
        #If specified the data will be written in CSV format to the path provided, otherwise it will be output to the pipeline
        $Path
    )
    if     ($Date)    {
        if ($report -match 'Counts$|Pages$|Storage$') {Write-Warning -Message 'Reports ending with Counts, Pages or Storage do not support date filtering' ; return }
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        if ($report -eq    'MailboxUsageDetail')      {Write-Warning -Message 'MailboxUsageDetail does not support date filtering.' ; return}
        $uri = "$GraphUri/reports/microsoft.graph.Get{0}(date={1:yyyy-MM-dd})" -f $Report , $Date
    }
    elseif ($Period)  {
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        $uri = "$GraphUri/reports/microsoft.graph.Get{0}(period='{1}')"        -f $Report , $Period
    }
    else              {
      if ($report -notmatch '^Office365Activation')  {
        $uri = "$GraphUri/reports/microsoft.graph.Get{0}(period='d7')"         -f $Report
      }
      else {
        $uri = "$GraphUri/reports/microsoft.graph.Get{0}"                      -f $Report
      }
    }
    if ($Path) { Invoke-GraphRequest -Method GET -uri $uri -OutputFilePath $Path}
    else       {
        $Path = [System.IO.Path]::GetTempFileName()
        Invoke-GraphRequest -Method GET -uri $uri -OutputFilePath $Path
        Import-Csv  $Path
        Remove-Item $Path
        }
}

function Get-GraphSignInLog    {
    <#
      .synopsis
        Gets the audit log -requires a priviledged account
      .Description
        This command calls https://graph.microsoft.com/beta/auditLogs/signIns
        which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.
      .Example
        >
        >Get-GraphSignInLog |
        >  select Date,UserPrincipalName,appDisplayName,ipAddress,clientAppUsed,browser,device,city,lat,long |
        >    Export-Excel -Path .\signin.xlsx -AutoSize -IncludePivotTable -PivotTableName Signins -PivotRows appdisplayName -PivotColumns browser -PivotData @{date='Count'} -show

        Gets the sign-in Log and exports it Excel, creating a PivotTable
    #>
    [cmdletbinding()]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphSignIn])]
    param   (
            $top = 200
    )
    $i = 1
    Write-Progress -Activity 'Getting Sign-in Auditlog'

    $result  = Invoke-GraphRequest  -Method get -Uri "$GraphUri/auditLogs/signIns?`$top=$top" -SkipHttpErrorCheck -StatusCodeVariable status
    if ($result.error)              {
            Write-Progress -Activity 'Getting Sign-in Auditlog' -Completed
            Write-Warning "An error was returned: '$($result.error.message)' code: $($result.error.code) "
    }
    if ($status -notmatch "2\d\d")  {Write-Warning "Status code returned was $Status ($([System.Net.HttpStatusCode]$status)) which does not look like success."}

    $records = $result.value
    while ($result.'@odata.nextLink' -and $records.count -lt $top) {
        $i ++
        Write-Progress -Activity 'Getting Sign-in Auditlog' -CurrentOperation "Page $i"
        $result   = Invoke-GraphRequest  -Method get -Uri $result.'@odata.nextLink'
        $records += $result.value
    }

    foreach ($r in $records) {
        $r.pstypenames.add('GraphSigninLog')
        $r['RiskEventTypesV2'] = $r['RiskEventTypes_V2'] ;
        $null = $r.Remove('RiskEventTypes_V2'), $r.remove( "@odata.etag") ;
        New-Object -TypeName MicrosoftGraphSignIn -Property $r
    }
    Write-Progress -Activity 'Getting Sign-in Auditlog'-Completed


}

function Get-GraphDirectoryLog {
    <#
      .synopsis
        Gets the Directory audit log -requires a priviledged account
      .Description
        This command calls https://graph.microsoft.com/beta/auditLogs/directoryAudits
        which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.

    #>
    [cmdletbinding()]
    [outputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphDirectoryAudit])]
    param   (
    [switch]$all,
    $Top = 100
    )
    $i = 1
    Write-Progress -Activity 'Getting Directory Audits log'
    $uri = "$GraphUri/auditLogs/directoryAudits"
    if (-not $all) {$uri += "?`$Top=$Top"}
    $result  = Invoke-GraphRequest  -Method get -Uri $uri -SkipHttpErrorCheck -StatusCodeVariable status
    if ($result.error)              {Write-Warning "An error was returned: '$($result.error.message)' - code: $($result.error.code) "}

    $records = $result.value
    while ($result.'@odata.nextLink' -and  $records.Count -lt $top) {
        $i ++
        Write-Progress -Activity 'Getting Directory Audits log' -CurrentOperation "Page $i"
        $result   = Invoke-GraphRequest  -Method get -Uri $result.'@odata.nextLink'
        $records += $result.value
    }
    foreach ($r in $records) {
        New-Object -TypeName MicrosoftGraphDirectoryAudit -Property $r
    }
    Write-Progress -Activity 'Getting Directory Audits log' -Completed
}
