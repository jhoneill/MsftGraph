# Microsoft.Graph.PlusPlus

A group within Microsoft produces a [Graph SDK for PowerShell](https://github.com/microsoftgraph/msgraph-sdk-powershell), it is mostly auto-generated from the Open API definitions and is the starting point for other projects - like this one. 

[Multiple modules on the Powershell Gallery](https://www.powershellgallery.com/packages?q=microsoft.graph) are the raw output of that project, they cover a lot of the  API, but do not try to be good and usable PowerShell, commands for example:
* Pipeline support is missing.
* Passing parameters by position is not supported (parameters must be named). 
* Common parameters (`-Verbose` / `-ErrorAction`) are not supported, in particular `-whatif` / `-confirm` are not implemented for dangerous operations.
* In many cases the API requires a GUID as a parameter, and the auto-generated commands won't take the ID from an object (which is what would be passed via the pipeline) or translate human-readable names to IDs. This kind of intelligence whould be embedded in better PowerShell commands.
* Argument completion is weak - confined to some parameter sets. 

* The full set of modules adds over 2000 commands, making it difficult to navigate and resulting in load times of over a minute.

**This module includes 'PlusPlus' in its name because it extends and improves the Microsoft.Graph modules.**    
Some examples: 
* **User creation:** `New-GraphUser -FirstName Bob -LastName Cratchett` will use rules to generate a User Principal Name / email address (First.Last@DefaultDomain) and a display name ('First Last') and will set the user's _usage location_ so they can be licensed immediately.
* **Finding a user:** `Get-GraphUser bob` will find all users with names which start "bob", pressing \[Tab\] after `bob` will expand the matching User Principal Name(s).
* **Finding things connected with a user or group:** `Get-GraphUser bob -MemberOf` returns the groups and directory roles that "bob" has been placed in; `Get-GraphGroup 'Accounts' --Notebooks` returns the OneNote notebook for a group.
* **Easy OneDrive and OneNote access:** ` Get-GraphUser -Drive | Set-GraphHomeDrive` removes the need to specify "current user's drive" in subsequent commands `Copy-FromGraphFolder `\[tab\] then completes the names of files and folders on that drive. 

To compare just one of these 
Here's a simple task with _PlusPlus_ using the alias `ggu` for `Get-GraphUser`
```
 ggu bob | % Manager       

Display Name Job Title      Office Location Mail
------------ ---------      --------------- ----
Jacob Marley Office Manager                 Jacob@mobula_consulting.com
```
And with the SDK module
```
> $bob =  Get-MgUser -Filter "startswith(userprincipalName,'Bob')" -ExpandProperty manager
> Get-MgUser -UserId $bob.Manager.Id    

Id                                   DisplayName  Mail  
--                                   -----------  ----  
4f770fd0-6b51-4338-b66f-2b31d9048cd2 Jacob Marley Jacob@mobula_consulting.com
```

Pipeline support enables commands like 
```
Get-GraphUser Jacob -DirectReports | Set-GraphUser -Manager ebenezer@mobula_consulting.com
```    

With long UPNs, tab completion where a name or ID is expected becomes very welcome.  

## Co-existence with Other Microsoft.Graph modules
Microsoft.Graph.PlusPlus **requires** the `Microsoft.Graph.Authentication` modue and, if avaiable, will use the .private.dll files from following modules (without loading the full module). 
*  **Users**
*  Users.Functions
*  Users.Actions
*  Reports
*  Identity.SignIns
*  Identity.DirectoryManagement
*  Applications

Of these, it is **strongly recommended** that `users` is available, the others are optional.

### Private DLL or Module?
Installing the all the `Microsoft.Graph.xxx` modules creates a lot of clutter, so _PlusPlus_ just loads their private.dll files. This allows their types to be used, so for example:    
 `Get-GraphUser Bob` will return a `Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser` object.   
  _PlusPlus_ provides formatting for many of these types, and extends some of them, and only uses a small number of commands from the SDK modules.    
Instead of using `Install-Module microsoft.graph.reports` to support the _reporting_ commands, you can download the module with    
 `Save-Module Microsoft.Graph.Reports -path \temp ` and copy `Microsoft.Graph.Reports.private.dll` to the `Microsoft.Graph.PlusPlus` module Directory.    
**Whichever method you use, ensure all SDK modules are the same version**. Errors will result if one if one module tries to load a newer version of authentication than the others, for example.

### Replacement of Aliases
The `Microsoft.Graph.Authentication` module defines aliases `Invoke-GraphRequest` and `Connect-Graph`    
for `Invoke-MGGraphRequest` and `Connect-MGGraph` respectively.     
In _PlusPlus_ these names are defined as functions which call  `Invoke-MGGraphRequest` and `Connect-MGGraph`.

### Settings
When the module loads it runs `Microsoft.Graph.PlusPlus.settings.ps1` in the module directory to set 
1. The default usage location to configure for new users (so they can be granted licenses).
1. The default properties returned for users. 
1. The **scopes** which will be requested when logging on. Some commands will fail if you do not request the right scopes, but you can request fewer scopes if you do not intend to use all the commands.
1. Tennant ID, Client (App) ID and Client Secret used for additional logon methods supported by `Connect-Graph` 

If you do not want to change the file provided with the module, edit a copy of it and set the environment variable  `GraphSettingsPath` to point to the copy. These settings can also be set on demand with the `Set-GraphOptions` Command

## Expanded choice of logon options.
Depending on settings, `Connect-Graph` may offer  parameters `-Credential`  `-AsApp` `-FromAzureSession` and `-Refresh`.  
When none of these is specified it calls `Connect-MgGraph` which will see if a suitable token is cached in your .graph directory, and if not start the device logon process, with a message like this    
`To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code AB12DEFGH to authenticate.`

If a **Tennant ID, Client ID and Secret** are set (see settings above) the `-AsApp` switch parameter is enabled to allow the session to sign in as an app registered with Azure AD.    
The `-Credential` parameter will be enabled if only the Tennant ID and Client ID are set;  `Connect-Graph` treats the Client Secret as optional for allowing the login with credentials but the logon service will reject some Client IDs if no secret is provided. 
These two methods get an _Access Token_ which `Connect-MgGraph` is told to use, and an associated _Refresh Token_. These logon methods allow the module to be used non-interactively; however, they only get scopes which previously been granted to the client-app; `Show-GraphSession` will show the scopes which have been granted. `Show-GraphSession` will also allow you to get the Refresh Token to use in another session and  `Set-GraphOptions` will allow you to bring such a token into the current session.  If the session has a _Refresh_ Token, `Connect-Graph -Refresh` will update the _Access_ Token, and the module attempts to do this automatically when the _Access_ Token expires.

If the `Az.accounts` module has been loaded the `-FromAzureSession` switch parameter is enabled: this will connect as the currently signed in account, but the available scopes are fixed by the Azure accounts module.   

# Module contents

## Commands: Session management 
`Set-GraphOptions,Get-AccessToken,  Connect-Graph, Show-GraphSession ,Test-GraphSession Invoke-GraphRequest`

## Commands for working with the Directory 
### Users and service principals
`Get-GraphUserList, Get-GraphUser, New-GraphUser, Set-GraphUser, Remove-GraphUser, Import-GraphUser  Export-GraphUser, Reset-GraphUserPassword`    
`New-GraphInvitation`    
`Get-GraphServicePrincipal` 

### Groups / Teams
`Get-GraphGroupList`   
`Get-GraphGroup, New-GraphGroup, Remove-GraphGroup, Set-GraphGroup, Import-GraphGroup`   
`Add-GraphGroupMember, Remove-GraphGroupMember, Import-GraphGroupMember, Export-GraphGroupMember`   
`Set-GraphTeam`

### Licenses
`Get-GraphSKU, Get-GraphLicense, Grant-GraphLicense, Revoke-GraphLicense`

### Roles
`Get-GraphDirectoryRole, Grant-GraphDirectoryRole, Revoke-GraphDirectoryRole,  Get-GraphDirectoryRoleTemplate`

### Misc. 
`Get-GraphConditionalAccessPolicy,Expand-GraphConditionalAccessPolicy, Get-GraphNamedLocation`   
`Get-GraphDeletedObject, Restore-GraphDeletedObject`   
`Get-GraphDomain, Get-GraphOrganization, Find-GraphPeople`

### Reports and Logs.
`Get-GraphReport, Get-GraphDirectoryLog, Get-GraphSignInLog`

## Commands for working with Apps
### OneDrive & Excel files
`Get-GraphDrive, Set-GraphHomeDrive, New-GraphFolder, Show-GraphFolder, Copy-FromGraphFolder, Copy-ToGraphFolder`   
`New-GraphWorkBook, Get-GraphWorkBook, Export-GraphWorkSheet, Import-GraphWorksheet`

### OneNote
`Add-FileToGraphOneNote`
`Get-GraphOneNoteBook,  Set-GraphOneNoteHome, Get-GraphOneNoteSection, New-GraphOneNoteSection`   
`Get-GraphOneNotePage,  Copy-GraphOneNotePage, Add-GraphOneNotePage, Update-GraphOneNotePage, Remove-GraphOneNotePage, Out-GraphOneNote`

### Outlook-Messages
`Get-GraphMailTips, Get-GraphMailFolder , Save-GraphMailAttachment, Get-GraphMailItem , Move-GraphMailItem, Send-GraphMailMessage, Send-GraphMailForward , Send-GraphMailReply,`   
`New-GraphRecipient, New-GraphMailAddress`    
`Get-GraphGroupConversation, Send-GraphGroupReply, Get-GraphGroupThread, Add-GraphGroupThread, Remove-GraphGroupThread`

### Outlook-Calendar
`Get-GraphReminderView, Get-GraphEvent, Add-GraphEvent, Remove-GraphEvent, Set-GraphEvent`   
`New-GraphAttendee, New-GraphRecurrence`

### Outlook-contacts
`Get-GraphContact, Set-GraphContact,  New-GraphContact, Remove-GraphContact`   
`New-GraphPhysicalAddress`

### Planner
`Get-GraphPlan, Remove-GraphPlan, Set-GraphPlanDetails, New-GraphTeamPlan`   
`Add-GraphPlanBucket, Remove-GraphPlanBucket, Rename-GraphPlanBucket, Get-GraphBucketTaskList`   
`Get-GraphPlanTask , Add-GraphPlanTask,  Set-GraphPlanTask, Remove-GraphPlanTask, Set-GraphTaskDetails`

### Sharepoint
`Get-GraphList , New-GraphList , Add-GraphListItem, Set-GraphListItem, Remove-GraphListItem`   
`New-GraphBooleanColumn, New-GraphCalculatedColumn, New-GraphChoiceColumn, New-GraphColumn, New-GraphCurrencyColumn, New-GraphDateTimeColumn, New-GraphLookupColumn, New-GraphNumberColumn, New-GraphPersonOrGroupColumn, New-GraphTextColumn`   
`Get-GraphSite, Get-GraphSiteColumn , Get-GraphSiteUserList`

### Teams
`Get-GraphChannel , Remove-GraphChannel`   
`New-GraphChannelMessage, New-GraphChannelReply, Get-GraphChannelReply`   
`Add-GraphOneNoteTab, Add-GraphPlannerTab, Add-GraphSharePointTab, Add-GraphWikiTab`   

### ToDo
`Get-GraphToDoList, New-GraphToDoList, Remove-GraphToDoList, New-GraphToDoTask, Remove-GraphToDoTask, Update-GraphToDoTask`

## Formatting and types
The module extends 16 types from the Microsoft.Graph SDK Modules `Attachment, Calendar, ChatMessage, Contact, DirectoryAudit, Drive, DriveItem, Event, List, MailTips, Reminder, SignIn, Site, Site, TeamsTab` and  `User`

It provides formats for 47 types. `AppRole, Calendar, Channel, ChatMessage, ColumnDefinition, Contact, Conversation, ConversationThread, Device, DirectoryAudit, DirectoryRole, Domain, Drive, DriveItem, Event, GraphExtendedTask, Group, LicenseDetails, List, MailFolder, MailTips, Message, MicrosoftGraphMailboxSettings, Notebook, OnenoteOperation, Onenotepage, OnenoteSection, Organization, PermissionScope, Person, PlannerBucket, PlannerPlan, PlannerTask, Post, Presence, Reminder, ServicePlanInfo, ServicePrincipal, SignIn, Site, SubscribedSku, Team, TeamsApp, TeamsAppDefinition, TeamsTab, TodoTask, TodoTaskList, User` and` VerifiedDomain`

And it provides completers for `Domains`, `Group` names, `Mail folders`, `OneDrive folders` & `OneDrive Items`, `OneNote Section` names, `Roles` in Azure AD, `SKUs` (for licensing) & the `plans` within SKUs, and `User Principal Names`.

