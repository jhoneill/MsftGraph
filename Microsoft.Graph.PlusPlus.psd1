@{
  Description       =   @'
  Module to work the Microsoft Graph API with both Office 365 and Microsoft accounts
  it contains over 100 functions to
  * Navigate, upload to and download from OneDrive
  * Navigate, manipulate pages and sections in OneNote Notebooks
  * Create and manage groups and teams, and Post to Teams and their channels
  * Create, read and write plans in team plans.
  * Read and write calendar events, and contacts in Outlook,
  * Read and send mail messages
  * Work with Sharepoint lists
'@
  Copyright         =   '(c) 2021 James O''Neill. All rights reserved.'
  Author            =   'James O''Neill'
  CompanyName       =   'Mobula Consulting'

  GUID              =   'f564c0f9-7d96-4452-a715-679dc47c20cc'
  ModuleVersion     =   '2.0'
  rootModule        =   '.\Microsoft.Graph.PlusPlus.psm1'
  RequiredModules   = @(@{ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '1.2.0'; })
  FormatsToProcess  =   'Microsoft.Graph.PlusPlus.format.ps1xml'
  FunctionsToExport = @('Add-FileToGraphOneNote',
                        'Add-GraphEvent',
                        'Add-GraphGroupMember',
                        'Add-GraphGroupThread',
                        'Add-GraphListItem',
                        'Add-GraphOneNotePage',
                        'Add-GraphOneNoteTab',
                        'Add-GraphPlanBucket',
                        'Add-GraphPlannerTab',
                        'Add-GraphPlanTask',
                        'Add-GraphWikiTab',
                        'Connect-Graph',
                        'ConvertTo-GraphDateTimeTimeZone',
                        'Copy-FromGraphFolder',
                        'Copy-ToGraphFolder',
                        'Export-GraphGroupMember',
                        'Export-GraphUser',
                        'Find-GraphPeople',
                        'Get-GraphBucketTaskList',
                        'Get-GraphChannel',
                        'Get-GraphChannelReply',
                        'Get-GraphContact',
                        'Get-GraphDeletedObject',
                        'Get-GraphDirectoryLog',
                        'Get-GraphDirectoryRole',
                        'Get-GraphDomain',
                        'Get-GraphDrive',
                        'Get-GraphEvent',
                        'Get-GraphGroup',
                        'Get-GraphGroupConversation',
                        'Get-GraphGroupList',
                        'Get-GraphGroupThread',
                        'Get-GraphList',
                        'Get-GraphMailFolder',
                        'Get-GraphMailItem',
                        'Get-GraphMailTips',
                        'Get-GraphOneNoteBook',
                        'Get-GraphOneNotePage',
                        'Get-GraphOneNoteSection',
                        'Get-GraphOrganization',
                        'Get-GraphPlan',
                        'Get-GraphPlanTask',
                        'Get-GraphReminderView',
                        'Get-GraphReport',
                        'Get-GraphSignInLog',
                        'Get-GraphSite',
                        'Get-GraphSiteColumn',
                        'Get-GraphSiteUserList',
                        'Get-GraphSKU',
                        'Get-GraphSkuLicensedUser',
                        'Get-GraphToDoList'
                        'Get-GraphUser',
                        'Get-GraphUserList',
                        'Grant-GraphDirectoryRole',
                        'Grant-GraphUserLicense',
                        'Import-GraphGroup',
                        'Import-GraphGroupMember',
                        'Import-GraphUser',
                        'Invoke-GraphRequest',
                        'Move-GraphMailItem',
                        'New-ContactAddress',
                        'New-EventAttendee',
                        'New-GraphAttendee,'
                        'New-GraphChannel',
                        'New-GraphChannelMessage',
                        'New-GraphChannelReply'    #All the different column types together!
                        'New-GraphColumn','New-GraphBooleanColumn', 'New-GraphCalculatedColumn', 'New-GraphChoiceColumn','New-GraphCurrencyColumn', 'New-GraphDateTimeColumn',
                                          'New-GraphLookupColumn', 'New-GraphNumberColumn','New-GraphPersonOrGroupColumn','New-GraphTextColumn',
                        # 'New-GraphContentType' ,
                        'New-GraphContact' ,
                        'New-GraphFolder',
                        'New-GraphGroup',
                        'New-GraphInvitation',
                        'New-GraphList',
                        'New-GraphMailAddress',
                        'New-GraphOneNoteSection',
                        'New-GraphPhysicalAddress',
                        'New-GraphRecipient',
                        'New-GraphRecurrence',
                        'New-GraphTeamPlan',
                        'New-GraphToDoList',
                        'New-GraphToDoTask',
                        'New-GraphUser',
                        'New-RecurrencePattern',
                        'Out-GraphOneNote',
                        'Remove-GraphChannel',
                        'Remove-GraphContact',
                        'Remove-GraphEvent',
                        'Remove-GraphGroup',
                        'Remove-GraphGroupMember',
                        'Remove-GraphGroupThread',
                        'Remove-GraphListItem',
                        'Remove-GraphOneNotePage',
                        'Remove-GraphPlan',
                        'Remove-GraphPlanbucket',
                        'Remove-GraphPlanTask',
                        'Remove-GraphToDoList',
                        'Remove-GraphToDoTask',
                        'Remove-GraphUser'
                        'Rename-GraphPlanBucket'
                        'Restore-GraphDeletedObject',
                        'Revoke-GraphDirectoryRole',
                        'Revoke-GraphUserLicense',
                        'Save-GraphMailAttachment',
                        'Send-GraphGroupReply',
                        'Send-GraphMailForward',
                        'Send-GraphMailMessage',
                        'Send-GraphMailReply',
                        'Set-GraphConnectionOptions',
                        'Set-GraphContact',
                        'Set-GraphEvent',
                        'Set-GraphGroup',
                        'Set-GraphListItem',
                        'Set-GraphPlanDetails',
                        'Set-GraphPlanTask',
                        'Set-GraphTaskDetails',
                        'Set-GraphTeam',
                        'Set-GraphUser',
                        'Show-GraphFolder',
                        'Show-GraphSession',
                        'Update-GraphOneNotePage',
                        'Update-GraphToDoTask'
  )
  AliasesToExport   = @('Add-GraphTeamChannel',
                        'Add-GraphTeamMember',
                        'Disconnect-Graph',
                        'Get-GraphConversation',
                        'Get-GraphTeam',
                        'Get-GraphTeamChannel',
                        'Get-GraphTeamConversation',
                        'Get-GraphTeamThread',
                        'New-GraphTeam',
                        'Remove-GraphTeam',
                        'Remove-GraphTeamMember',
                        'TextColumn',
                        'PersonColumn',
                        'NumberColumn',
                        'LookupColumn',
                        'DateTimeColumn',
                        'CurrencyColumn',
                        'ChoiceColumn',
                        'CalculatedColumn',
                        'BooleanColumn',
                        'ListColumn'
  )

  PrivateData = @{
       PSData    = @{
           Tags       = @('MicrosoftGraph', 'Microsoft', 'Office365', 'Graph', 'PowerShell', 'AzureAD', 'OneNote', 'OneDrive', 'Outlook', 'Sharepoint', 'Planner', 'MSGraph')
           Category   = 'Scripting office Online'
           ProjectUri = 'https://github.com/jhoneill/MsftGraph'
           LicenseUri = 'https://github.com/jhoneill/MsftGraph/blob/master/LICENSE'
        } # End of PSData hashtable
  }

  # Supported PSEditions
  # CompatiblePSEditions = @()

  # Minimum version of the Windows PowerShell engine required by this module
  # PowerShellVersion = ''

  # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
  # CLRVersion = ''

  # Type files (.ps1xml) to be loaded when importing this module
  # TypesToProcess = @()

  # HelpInfo URI of this module
  # HelpInfoURI = ''
}