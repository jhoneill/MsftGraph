Function Get-GraphReport {
    <#
        .Synopsis
            Use BETA functionality to get reports from MS Graph
        .Example
            >Get-GraphReport -Report MailboxUsageDetail | ft "Display Name",  "Storage Used (Byte)"
            Displays mailbox storage used by users - note that
            fields have 'friendly' names which need to be wrapped in quotes
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
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
                'SharePointActivityUserDetail', 'SharePointSiteUsageDetail', 'SharePointSiteUsageFileCounts', '
                SharePointSiteUsagePages', 'SharePointSiteUsageSiteCounts', 'SharePointSiteUsageStorage',
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
        $Period
    )
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    if     ($Date)    {
        if ($report -match 'Counts$|Pages$|Storage$') {Write-Warning -Message 'Reports ending with Counts, Pages or Storage do not support date filtering' ; return }
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        if ($report -eq    'MailboxUsageDetail')      {Write-Warning -Message 'MailboxUsageDetail does not support date filtering.' ; return}
        $uri = "https://graph.microsoft.com/v1.0/reports/microsoft.graph.Get{0}(date={1:yyyy-MM-dd})" -f $Report , $Date
    }
    elseif ($Period)  {
        if ($report -match '^Office365Activation')    {Write-Warning -Message 'Office365Activation Reports do not support any filtering.'  ; return }
        $uri = "https://graph.microsoft.com/v1.0/reports/microsoft.graph.Get{0}(period='{1}')"        -f $Report , $Period
    }
    else              {
      if ($report -notmatch '^Office365Activation')  {
        $uri = "https://graph.microsoft.com/v1.0/reports/microsoft.graph.Get{0}(period='d7')"         -f $Report
      }
      else {
        $uri = "https://graph.microsoft.com/v1.0/reports/microsoft.graph.Get{0}"                      -f $Report
      }
    }
    Invoke-MgGraphRequest -Method GET -uri $uri | ConvertFrom-Csv
}

Function Get-GraphSignInLog {
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
    param (
    )
    Connect-MSGraph
    $i = 1
    Write-Progress -Activity 'Getting Sign-in Auditlog'
    try   { $result  = Invoke-RestMethod  -Method get -Uri "https://graph.microsoft.com/beta/auditLogs/signIns"  -headers $Script:DefaultHeader  }
    catch {
        if ($_.exception.response.statuscode.value__ -eq 401) {
            Write-Warning -Message "The server responded 'Unauthorized' - check that $($script:GraphUser.userPrincipalName) has rights to access the log."; return
        }
    }
    $records = $result.value
    while ($result.'@odata.nextLink') {
        $i ++
        Write-Progress -Activity 'Getting Sign-in Auditlog' -CurrentOperation "Page $i"
        $result   = Invoke-RestMethod  -Method get -Uri $result.'@odata.nextLink'  -headers $Script:DefaultHeader
        $records += $result.value
    }
    foreach ($r in $records) {
        $r.pstypenames.add('GraphSigninLog')
        Add-Member -InputObject $r -MemberType ScriptProperty -Name City    -Value {$this.location.city}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name State   -Value {$this.location.state}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Country -Value {$this.location.countryOrRegion}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Lat     -Value {$this.location.geoCoordinates.latitude}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Long    -Value {$this.location.geoCoordinates.longitude}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Browser -Value {$this.deviceDetail.browser}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Device  -Value {$this.deviceDetail.displayName;}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Date    -Value {[datetime]$this.createdDateTime}
    }
    Write-Progress -Activity 'Getting Sign-in Auditlog'-Completed

    $records
}

Function Get-GraphDirectoryLog {
    <#
      .synopsis
        Gets the Directory audit log -requires a priviledged account
      .Description
        This command calls https://graph.microsoft.com/beta/auditLogs/directoryAudits
        which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.

    #>
    [cmdletbinding()]
    param (
    )
    Connect-MSGraph
    $i = 1
    Write-Progress -Activity 'Getting Directory Audits log'
    try   { $result  = Invoke-RestMethod  -Method get -Uri "https://graph.microsoft.com/beta/auditLogs/directoryAudits"  -headers $Script:DefaultHeader  }
    catch {
        if ($_.exception.response.statuscode.value__ -eq 401) {
            Write-Warning -Message "The server responded 'Unauthorized' - check that $($script:GraphUser.userPrincipalName) has rights to access the log."; return
        }
    }
    $records = $result.value
    while ($result.'@odata.nextLink') {
        $i ++
        Write-Progress -Activity 'Getting Directory Audits log' -CurrentOperation "Page $i"
        $result   = Invoke-RestMethod  -Method get -Uri $result.'@odata.nextLink'  -headers $Script:DefaultHeader
        $records += $result.value
    }
    $defaultProperties = @('Date','User','ActivityDisplayName','result')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    foreach ($r in $records) {
        $r.pstypenames.add('GraphDirectoryLog')
        Add-Member -InputObject $r -MemberType ScriptProperty -Name User              -Value {$this.initiatedBy.user.userPrincipalName}
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Date              -Value {[datetime]$this.activityDateTime}
        Add-Member -InputObject $r -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers
    }
    Write-Progress -Activity 'Getting Directory Audits log' -Completed

    $records
}
