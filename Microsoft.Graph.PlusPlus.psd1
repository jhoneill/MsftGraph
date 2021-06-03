@{
  Copyright             =   '(c) 2021 James O''Neill. All rights reserved.'
  Author                =   'James O''Neill'
  CompanyName           =   'Mobula Consulting'
  ModuleVersion         =   '1.5.0'
  PrivateData   = @{
       PSData    = @{
           Tags         = @('MicrosoftGraph', 'Microsoft', 'Office365', 'Graph', 'PowerShell', 'AzureAD', 'OneNote', 'OneDrive', 'Outlook', 'Sharepoint', 'Planner', 'MSGraph')
           Category     = 'Functions'
           ProjectUri   = 'https://github.com/jhoneill/MsftGraph'
           LicenseUri   = 'https://github.com/jhoneill/MsftGraph/blob/master/LICENSE'
        }
  }
  Description           = @'
  Module to work the Microsoft Graph API using both AzureAD 'work or school' accounts and 'personal' Microsoft accounts
  it contains over 100 functions to
  * Navigate, upload to and download from OneDrive
  * Navigate and manipulate Pages & Sections in OneNote Notebooks
  * Create and manage Groups and Teams, and post to Teams and their Channels
  * Create, read and write Plans for Teams.
  * Read and write Calendar-events and Contacts in Outlook,
  * Read and send Mail messages in Outlook
  * Work with Sharepoint Lists
  * Access ToDo Lists.
'@
  CompatiblePSEditions  = @('Core', 'Desktop')
  PowerShellVersion     =   '5.1'
  GUID                  =   'f564c0f9-7d96-4452-a715-679dc47c20cc'
  RootModule            =   '.\Microsoft.Graph.PlusPlus.psm1'
  RequiredModules       = @(@{ModuleName = 'Microsoft.Graph.Authentication'; ModuleVersion = '1.4.0'; })
  FormatsToProcess      =   'Microsoft.Graph.PlusPlus.format.ps1xml'
  TypesToProcess        = @('Microsoft.Graph.PlusPlus.types.ps1xml')
  FunctionsToExport     = @('Add-FileToGraphOneNote',
                            'Add-GraphEvent',
                            'Add-GraphGroupMember',
                            'Add-GraphGroupThread',
                            'Add-GraphListItem',
                            'Add-GraphOneNotePage',
                            'Add-GraphOneNoteTab',
                            'Add-GraphPlanBucket',
                            'Add-GraphPlannerTab',
                            'Add-GraphPlanTask',
                            'Add-GraphSharePointTab',
                            'Add-GraphWikiTab',
                            'Connect-Graph',
                            'ConvertTo-GraphDateTimeTimeZone',
                            'Copy-FromGraphFolder',
                            'Copy-ToGraphFolder',
                            'Copy-GraphOneNotePage',
                            'Expand-GraphConditionalAccessPolicy',
                            'Export-GraphGroupMember',
                            'Export-GraphUser',
                            'Export-GraphWorkSheet',
                            'Find-GraphPeople',
                            'Get-AccessToken',
                            'Get-GraphApplication',
                            'Get-GraphBucketTaskList',
                            'Get-GraphChannel',
                            'Get-GraphChannelReply',
                            'Get-GraphConditionalAccessPolicy',
                            'Get-GraphContact',
                            'Get-GraphDeletedObject',
                            'Get-GraphDirectoryLog',
                            'Get-GraphDirectoryRole',
                            'Get-GraphDirectoryRoleTemplate',
                            'Get-GraphDomain',
                            'Get-GraphDrive',
                            'Get-GraphEvent',
                            'Get-GraphGroup',
                            'Get-GraphGroupConversation',
                            'Get-GraphGroupList',
                            'Get-GraphGroupThread',
                            'Get-GraphLicense',
                            'Get-GraphList',
                            'Get-GraphMailFolder',
                            'Get-GraphMailItem',
                            'Get-GraphMailTips',
                            'Get-GraphNamedLocation',
                            'Get-GraphOneNoteBook',
                            'Get-GraphOneNotePage',
                            'Get-GraphOneNoteSection',
                            'Get-GraphOrganization',
                            'Get-GraphPlan',
                            'Get-GraphPlanTask',
                            'Get-GraphReminderView',
                            'Get-GraphReport',
                            'Get-GraphServicePrincipal',
                            'Get-GraphSignInLog',
                            'Get-GraphSite',
                            'Get-GraphSiteColumn',
                            'Get-GraphSiteUserList',
                            'Get-GraphSKU',
                            'Get-GraphTeamsApp',
                            'Get-GraphToDoList'
                            'Get-GraphUser',
                            'Get-GraphUserList',
                            'Get-GraphWorkBook',
                            'Grant-GraphDirectoryRole',
                            'Grant-GraphLicense',
                            'Import-GraphGroup',
                            'Import-GraphGroupMember',
                            'Import-GraphUser',
                            'Import-GraphWorkSheet',
                            'Invoke-GraphRequest',
                            'Move-GraphMailItem',
                            'New-ContactAddress',
                            'New-EventAttendee',
                            'New-GraphAttendee',
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
                            'New-GraphWorkBook',
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
                            'Remove-GraphUser',
                            'Rename-GraphPlanBucket',
                            'Reset-GraphUserPassword',
                            'Restore-GraphDeletedObject',
                            'Revoke-GraphDirectoryRole',
                            'Revoke-GraphLicense',
                            'Save-GraphMailAttachment',
                            'Send-GraphGroupReply',
                            'Send-GraphMailForward',
                            'Send-GraphMailMessage',
                            'Send-GraphMailReply',
                            'Set-GraphContact',
                            'Set-GraphDefaultGroup',
                            'Set-GraphEvent',
                            'Set-GraphGroup',
                            'Set-GraphHomeDrive',
                            'Set-GraphOneNoteHome',
                            'Set-GraphOptions',
                            'Set-GraphListItem',
                            'Set-GraphPlanDetails',
                            'Set-GraphPlanTask',
                            'Set-GraphTeam',
                            'Set-GraphUser',
                            'Show-GraphFolder',
                            'Show-GraphSession',
                            'Test-GraphSession',
                            'Update-GraphOneNotePage',
                            'Update-GraphToDoTask'
  )
  AliasesToExport       = @(
                            'Add-FileToGraphNoteBook',
                            'Add-GraphNoteBookPage',
                            'Add-GraphTeamChannel',
                            'Add-GraphTeamMember',
                            'Copy-GraphNoteBookPage',
                            'Get-GraphContext',
                            'Get-GraphConversation',
                            'Get-GraphNoteBook',
                            'Get-GraphNoteBookPage',
                            'Get-GraphNoteBookSection',
                            'Get-GraphTeam',
                            'Get-GraphTeamChannel',
                            'Get-GraphTeamConversation',
                            'Get-GraphTeamThread',
                            'New-GraphGroupPlan',
                            'New-GraphNoteBookSection',
                            'New-GraphSession',
                            'New-GraphTeam',
                            'Out-GraphNoteBook',
                            'Remove-GraphNoteBookPage'
                            'Remove-GraphTeam',
                            'Remove-GraphTeamMember',
                            'Set-GraphDefaultTeam',
                            'Update-GraphNoteBookPage',
                            'BooleanColumn',
                            'CalculatedColumn',
                            'ChoiceColumn',
                            'CurrencyColumn',
                            'DateTimeColumn',
                            'ListColumn',
                            'LookupColumn',
                            'NumberColumn',
                            'PersonColumn',
                            'TextColumn',
                            'ggg',
                            'ggu',
                            'GWhoAmI',
                            'GraphSession',
                            'igr'
  )
  FileList              = @(
                            '.\README.md'
                            '.\LICENSE',
                            '.\ChangeLog.md'
                            '.\Microsoft.Graph.PlusPlus.psd1',
                            '.\Microsoft.Graph.PlusPlus.psm1',
                            '.\ActionCard.ps1',
                            '.\Applications.ps1',
                            '.\Authentication.ps1',
                            '.\Groups.ps1',
                            '.\Identity.DirectoryManagement.ps1',
                            '.\Identity.SignIns.ps1',
                            '.\Notes.ps1',
                            '.\OneDrive.ps1',
                            '.\PersonalContacts.ps1',
                            '.\Planner.ps1',
                            '.\Reports.ps1',
                            '.\Sharepoint.ps1',
                            '.\Users.Actions.ps1',
                            '.\Users.Functions.ps1',
                            '.\Users.ps1',
                            '.\Microsoft.Graph.PlusPlus.settings.ps1',
                            '.\Microsoft.Graph.PlusPlus.format.ps1xml',
                            '.\Microsoft.Graph.PlusPlus.types.ps1xml',
                            '.\Blank.xlsx',
                            '.\docs\Relationships.pdf',
                            '.\docs\Logon options.pdf',
                            '.\Examples\PlannerImportExport',
                            '.\Examples\Data-XLSx-Drive.ps1',
                            '.\Examples\Data-XLSx-Drive-dlChart.ps1',
                            '.\Examples\Demo.ps1',
                            '.\Examples\Link.png',
                            '.\Examples\New-BcAuthContext.ps1',
                            '.\Examples\OneDrive.gif',
                            '.\Examples\Team.gif',
                            '.\Examples\Template_groups.csv',
                            '.\Examples\Template_membership.csv',
                            '.\Examples\Template_users.csv'
                            '.\Examples\PlannerImportExport\Create_Planner_Template.ps1',
                            '.\Examples\PlannerImportExport\Export-planner-to-xlsx.ps1',
                            '.\Examples\PlannerImportExport\Import-Planner-From-Xlsx.ps1',
                            '.\Examples\PlannerImportExport\Planner-Export.xlsx'
  )

  # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
  # CLRVersion = ''

  # HelpInfo URI of this module
  # HelpInfoURI = ''
}