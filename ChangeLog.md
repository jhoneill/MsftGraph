# Change Log

## 1.5.5

* Fixed a problem with Connect-Graph not passing dynamic parameters (like Certificate) to Connect-MgGraph
* Fixed a bug in Get-GraphGroup which prevented owners of a group being returned (small typo)
* Added control of user properties fetched for owners and members of groups in Get-GraphGroup
* Fixed a bug which meant group membership was not being expanded correctly
* Made a change to group name matching so that names which contain [ ] do not treat them as a wildcard set.

## 1.5.4

* Added Support for -UseDeviceAuthentication
* Fixed a bug logging on with personal accounts, and using them to access OneDrive.
* Fixed a issue finding types in the .PSM1 file in some configurations
* Reformatted readme to to keep markdown lint happy.

## 1.5.3

* Version issues - tried to update the gallery correctly with 1.5.2 using version number 1.5.1

## 1.5.2

* Fixed handling of two rest calls sending back extra properties (MicrosoftGraphChannel with TenantID) and Users Managers with authorizationInfo
* Fixed DomainName completer not working
* Fixed incorrect cmdlet name for getting Domain Name Reference List
* Added support for Behavior Options when creating groups
* Re-formatted this log to keep markdown lint happy.

## 1.5

* In `Grant-GraphDirectoryRole`: allow members to tab complete and fix bug granting multiple roles.
* Update photo support in `Get-GraphUser`
* Fix errors from additional properties when logged on with Azure access token
* Restore alias "New-RecurrencePattern" which had got lost.
* Set maximum column widths when formatting group objects.
* Bump version number.

## 1.4.3

### After Beta 1

* Improved auto-load behaviour so `Get-GraphUser` / `Get-GraphGroup` will autoLoad the module, login with cached credentials and execute (previously needed to load the module and run `Connect-Graph`).
* Additional type extensions (more to come).
* Warning messages in *users* now display the UPN (were showing a type name).
* Workaround `$expand=` REST parameter only expanding the first 20 directory objects.

### In Beta 1

* Added this ChangeLog.md !
* Fixed breaking typo in `Import-GraphUser`.
* Fixed a bug which prevented `Remove user` running.
* Fixed parameter sets in `Add-GraphUser` so UPN, displayName, first name and last name can all be specified together.
* Fixed bug searching for users by name. It was only searching in the `mail` field.
* In `Add-GraphGroupMember` / `Remove-GraphGroupMember` handle user members already being in the desired state.
* Cleaned up the import processes and example templates.
* Moved Personal contact functions to their own file and load the `PersonalContacts` DLL (`Contact` is incorrectly defined in `Users` so we need the right DLL).
* `Set-GraphTaskDetails` should be private - removed it from the export list in the psd1.
* Changed validation of `-DefaultUsageLocation` in `Set-GraphOptions` to work round an error when reloading the module.
* Added a helper function:  `FilterString "bob*"` returns `startsWith(displayName,'bob')`  
     and `FilterString bob  mail` returns `displayName -eq 'bob' or mail -eq 'bob'`.
* Used this helper to remove implicit wildCarding as part of a clean-up of `Get-GraphServicePrincipal`. Previously a search for 'Microsoft Graph' also returned
'Microsoft Graph Change Tracking', 'Microsoft Graph data connect Data Transfer' and 'Microsoft Graph PowerShell'.
* Added `Get-GraphApplication`.
* Removed implied wildcards generally.
* Added `ToString()` overrides to the types.ps1xml file to show UPN for users and Display name for other things. (*implicit* `.ToString()` doesn't always work!).
* Added `Set-GraphDefaultGroup`.
* Added new `ChannelCompleter` class - which will work with default Group/team. So:  
    `Set-GraphDefaultGroup Accounts`   (where accounts is a team-provisioned-group)
    `Get-GraphChannel`  \[TAB\]  
    will autocomplete the channels for the the accounts team.
* Added upn/group/team completion to parameters where the existing completers had not been added.  
* Added helpers `idFromGroup` and `idFromTeam` - previously the same code to say "is it a guid, an object-with-a-guid, or a name to resolve to get a guid" was repeated in many places. As part of this removed implicit wildcards from searching for groups/teams. The intention is that any of the group / team functions can accept objects with at least an ID property, strings containing the GUID, strings containing a name (with wild card support) either singly or as an array or via the pipeline and return the same result however the group was passed.
* Added aliases for `Connect-Graph`: "New-GraphSession" and "GraphSession".

## 1.4.1 & 1.4.2

No code changes. Fixing incorrect files bundled to the PowerShell Gallery.

## 1.4.0  (42a5e9d)

The **First release targeting the [SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell)** version numbers will now try to sync with the SDK version.
The module has changed name to Microsoft.Graph.PlusPlus and some files have changed name accordingly. Functional changes are.

* `Connect-MSGraph` is now `Connect-Graph` and is a wrapper for `Connect-MgGraph` from the Microsoft.Graph.Authentication Module (and replaces the `Connect-Graph` Alias in that module)
* Settings - especially used by the extra logon modes are in a settings file - the location of which can be set by the `GraphSettingsPath` environment variable. The default settings file will also check the `GraphScopes` environment variable for scopes to request.
* Where possible functions now return the objects defined in the SDK which are all in the `Microsoft.Graph.PowerShell.Models` name space. The models loads the Microsoft.Graph.xxx.private.DLL files needed to makes theses types available without enabling the all the functionality in the modules.  
* Functions no longer call `Invoke-RestMethod` / `Invoke-WebRequest` but call `Invoke-GraphRequest` which is a wrapper for `Invoke-MGGraphRequest` from the Microsoft.Graph.Authentication Module (and replaces the `Invoke-GraphRequest` Alias in that module). This function converts to the desired type, removing any unwanted properties and will handle getting output in multiple pages.
* Functions have been moved around .ps1 files to suit the groupings used in the SDK. Some need a private DLL loaded and will skip loading if the module isn't present.  

### Replaced functions

* `Add-GraphChannelThread` has been replaced by `New-GraphChannelMessage`
* `Connect-MSGraph`  replaced by `Connect-Graph`
* `Expand-GraphTask` an internal function which is now handled inside `Get-GraphTask`
* `Get-GraphContactList` replaced by `Get-GraphContact`
* `Get-GraphMailFolderList` replaced by `Get-GraphMailFolder`
* `Get-GraphSKUList` functionality is now in `Get-GraphSKU`
* `New-Recipient`    has been replaced by `New-GraphRecipient`

### New functions

#### Session management

* `Set-GraphHomeDrive`, `Set-GraphOneNoteHome`, `Set-GraphOptions`, Test-GraphSession

#### User / account management

* `New-GraphUser`, `Remove-GraphUser`, `Import-GraphUser`,  Export-GraphUser, Reset-GraphUserPassword
* `New-GraphInvitation` (to invite external users)
* `Get-GraphServicePrincipal` (More SP functions are planned)

#### Group / Team management

* `Import-GraphGroup`, `Import-GraphGroupMember`,  Export-GraphGroupMember
* `Set-GraphPlanDetails`, `Rename-GraphPlanBucket`, `Remove-GraphPlan`
* `New-GraphChannelMessage`, `New-GraphChannelReply`
* `Add-GraphSharePointTab`

#### Directory and access management

* `Get-GraphDeletedObject`, `Restore-GraphDeletedObject`
* `Get-GraphDirectoryRole`, `Grant-GraphDirectoryRole`, `Revoke-GraphDirectoryRole`, `Get-GraphDirectoryRoleTemplate`
* `Get-GraphLicense`, `Grant-GraphLicense`, `Revoke-GraphLicense`
* `Get-GraphConditionalAccessPolicy`, `Get-GraphNamedLocation`

#### App functions (Excel OneNote, Outlook, ToDo)

* `Get-GraphWorkBook`, `New-GraphWorkBook`, Export-GraphWorkSheet, `Import-GraphWorkSheet`
* `Copy-GraphOneNotePage`
* `Move-GraphMailItem`, `Save-GraphMailAttachment`, `New-GraphMailAddress`, `New-GraphAttendee`, `New-GraphRecurrence`, `New-GraphPhysicalAddress`
* `Get-GraphToDoList`, `New-GraphToDoList`, `Remove-GraphToDoList`, `New-GraphToDoTask`, `Remove-GraphToDoTask`, `Update-GraphToDoTask`

## 1.0 (172fbf2)

Release as the MsftGraph module.
