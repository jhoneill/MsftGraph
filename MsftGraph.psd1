@{
  Description       =   @'
  Module to work the Microsoft Graph API with both Office 365 and Microsoft accounts
  it contains over 80 functions to
  * Navigate, upload to and download from OneDrive
  * Navigate, manipulate pages and sections in OneNote Notebooks
  * Create and manage groups and teams, and Post to Teams and their channels
  * Create, read and write plans in team plans.
  * Read and write calendar events, and contacts in Outlook,
  * Read and send mail messages
  * Work with Sharepoint lists
'@
  Copyright         =   '(c) 2019 James O''Neill. All rights reserved.'
  Author            =   'James O''Neill'
  CompanyName       =   'Mobula Consulting'

  GUID              =   'f564c0f9-7d96-4452-a715-679dc47c20cc'
  ModuleVersion     =   '1.0'

  NestedModules     = @('Graph.ps1',
                        'GroupsAndTeams.ps1',
                        'OneDrive.ps1',
                        'OneNote.ps1',
                        'Outlook-Calendar.ps1',
                        'Outlook-Contacts.ps1',
                        'Outlook-Mail.ps1',
                        'Planner.ps1',
                        'SharePoint.ps1',
                        'User.ps1'
  )
  FormatsToProcess  =   'Graph.format.ps1xml'
  FunctionsToExport = @('Add-FileToGraphOneNote',
                        'Add-GraphChannelThread',
                        'Add-GraphEvent',
                        'Add-GraphGroupMember',
                        'Add-GraphGroupThread',
                        'Add-GraphListItem',
                        'Add-GraphOneNotePage',
                        'Add-GraphOneNoteTab',
                        'Add-GraphPlannerTab',
                        'Add-GraphPlanTask',
                        'Add-GraphWikiTab',
                        'Connect-MSGraph',
                        'Copy-FromGraphFolder',
                        'Copy-ToGraphFolder',
                        'Expand-GraphTask',
                        'Find-GraphPeople',
                        'Get-GraphChannel',
                        'Get-GraphChannelReply',
                        'Get-GraphContactList',
                        'Get-GraphDomain',
                        'Get-GraphDirectoryLog',
                        'Get-GraphDrive',
                        'Get-GraphEvent',
                        'Get-GraphGroupConversation',
                        'Get-GraphGroupList',
                        'Get-GraphGroupThread',
                        'Get-GraphList',
                        'Get-GraphMailItem',
                        'Get-GraphMailFolderList',
                        'Get-GraphMailTips',
                        'Get-GraphOneNoteBook',
                        'Get-GraphOneNotePage',
                        'Get-GraphOneNoteSection',
                        'Get-GraphOrganization',
                        'Get-GraphPlan',
                        'Get-GraphPlanTask',
                        'Get-GraphReport',
                        'Get-GraphReminderView',
                        'Get-GraphSignInLog',
                        'Get-GraphSite',
                        'Get-GraphSiteColumn',
                        'Get-GraphSiteUserList',
                        'Get-GraphSKU',
                        'Get-GraphSKUList',
                        'Get-GraphTeam',
                        'Get-GraphUser',
                        'Get-GraphUserList',
                        'New-ContactAddress',   #All the different column types together!
                        'New-EventAttendee',
                        'New-GraphChannel',
                        'New-GraphColumn','New-GraphBooleanColumn', 'New-GraphCalculatedColumn', 'New-GraphChoiceColumn','New-GraphCurrencyColumn', 'New-GraphDateTimeColumn',
                                          'New-GraphLookupColumn', 'New-GraphNumberColumn','New-GraphPersonOrGroupColumn','New-GraphTextColumn',
                        'New-GraphContact' ,
                        'New-GraphFolder',
                        'New-GraphGroup',
                        'New-GraphList',
                        'New-GraphOneNoteSection',
                        'Add-GraphPlanBucket',
                        'New-GraphTeamPlan',
                        'New-Recipient',
                        'New-RecurrencePattern',
                        'Out-GraphOneNote',
                        'Remove-GraphChannel',
                        'Remove-GraphContact',
                        'Remove-GraphEvent',
                        'Remove-GraphGroup',
                        'Remove-GraphGroupMember',
                        'Remove-GraphListItem',
                        'Remove-GraphPlanbucket',
                        'Remove-GraphPlanTask',
                        'Remove-GraphGroupThread',
                        'Remove-GraphOneNotePage',
                        'Send-GraphGroupReply',
                        'Send-GraphMailForward',
                        'Send-GraphMailMessage',
                        'Send-GraphMailReply',
                        'Set-GraphContact',
                        'Set-GraphEvent',
                        'Set-GraphGroup',
                        'Set-GraphListItem',
                        'Set-GraphPlanTask',
                        'Set-GraphTeam',
                        'Set-GraphUser',
                        'Show-GraphFolder',
                        'Show-GraphSession',
                        'Update-GraphOneNotePage'
  )
  AliasesToExport   = @('Add-GraphTeamChannel',
                        'Add-GraphTeamMember',
                        'Get-GraphConversation',
                        'Get-GraphGroup',
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
           Tags       = @('Microsoft Graph', 'MSGraph', 'Office365', 'AzureAD', 'OneNote', 'OneDrive', 'Outlook', 'Sharepoint', 'Planner')
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