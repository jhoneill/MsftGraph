<#
  .description
    Executed when the module loads to set defaults for
    the scopes and information needed to support additional logon types.
    Environment Variable MGSettingsPath is set the information will be read from there.
    This file IS considered safe to edit. If it is deleted the module will
    user the default scopes provided by the PowerShell-graph-sdk.
#>
Set-GraphOptions -DefaultUsageLocation GB
Set-GraphOptions -DefaultUserProperties  @('businessPhones', 'displayName', 'givenName', 'id', 'jobTitle', 'mail', 'mobilePhone',
                                           'officeLocation', 'preferredLanguage', 'surname', 'userPrincipalName', #The preceeding fields are the APIs defaults.
                                           'assignedLicenses', 'department', 'usageLocation', 'userType')

#The scopes requested. You can shorten this of you don't need all things provided in the module
if    ($env:GraphScopes) {
        Set-GraphOptions -DefaultScopes ( $env:GraphScopes -split ',\s*')}
else  { Set-GraphOptions -DefaultScopes @(
                'AppCatalog.Read.All',
                'AuditLog.Read.All',
                'Directory.AccessAsUser.All', #Grant same rights to the directory as the user has
                'Calendars.ReadWrite',
                'Calendars.ReadWrite.Shared',
                'ChannelMessage.Read.All',
                'ChannelMessage.Delete',
                'ChannelMessage.Edit',
                'Contacts.ReadWrite',
                'Contacts.ReadWrite.Shared',
                'Files.ReadWrite.All',
                'Group.ReadWrite.All',# or read fails when logging on as non-admin
                'Mail.ReadWrite',
                'Mail.Send',
                'MailboxSettings.ReadWrite',
                'Notes.ReadWrite.All',
                'Notes.Create',
                'People.Read.All',
                'Presence.Read.All',
                'Reports.Read.All',
                'Sites.ReadWrite.All',
                'Sites.Manage.All',       #Needed to create lists.
                'Tasks.ReadWrite',        #Needed for Todo access
                'User.ReadWrite.all',    # Read write users and groups may not be needed if Directory is granted ?
                'openid',
                'profile'#,        'offline_access'
)}
<#
    If you want to logon by providing a name and password or as the app you need to provide
    1 your tenant ID (a GUID)
    2 The ID of an App (other GUID) which has been granted the right to use some
       scopes - either on its own account or delegated for a user - in your tenant
    3 A secret associated with the App.

    You can create an app in Azure AD or at https://apps.dev.microsoft.com/
    I created mine as a native app, with a re-direct URI of https://login.microsoftonline.com/common/oauth2/nativeclient and
    gave it a set of Microsoft graph permissions in Azure AD

    If you ONLY Want to work with accounts in Azure AD you can set up your app with these instructions which I lifted from
    https://msunified.net/2018/12/12/post-at-microsoftteams-channel-chat-message-from-powershell-using-graph-api/
    1.  Log on to https://portal.azure.com with a GA administrator
    2.  Navigate to Azure Active Directory
    3   Go to App registrations
    4.  Click + New registration
    5.  Call it PowerShellMSGraphAPI (or another name of your choice)
    6.  Leave who can use this API on the default of single tennant and leave the Redirect URI blank
    7.  Click Register
    8.  This will bring up the details of the new APP. Under call APIS click View API permissions to grant the required group read and write permissions
    9.  Click + Add a permission
    10. Choose Microsoft Graph, then Delegated permissions and choose Group.Read.All and ReadWrite.All (remember you need to expand Group)
    12. I had to click the enterprise apps link and click "Grant admin Consent" from (this is where a GA admin is needed)
    13. You now have admin consent granted for your tenant, so users can authenticate without a consent dialog.
    14. Navigate back to Overview
    15. Copy the Application (client) ID    Paste it into this script as the value for ClientID in Set-GraphOptions;
    16. Also copy the tenant ID paste it into this script as the value for TenantID in Set-GraphOptions
    17. Click Certificates and Secrets, add a secret and chose never expires (unless you want to update the script later), click add
    18. Copy the secret and EITHER (the dirty but portable way) paste into this script as the value for clientSecret in Set-GraphOptions
                            OR (clean but not portable) Convert it to a securestring & export: ConvertTo-SecureString -Force -AsPlainText (Get-Clipboard) | Export-Clixml myclientSecret.xml
                               Get the contents of the file as for setting clientsecret in Set-GraphOptions
#>

#YOUR tenant ID
#Set-GraphOptions -TenantID "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
#A client APP ID known to your tenant
Set-GraphOptions -ClientID "14d82eec-204b-4c2f-b7e8-296a70dab67e" #the Graph-Powershell SDK GUID. You can create your own. "1950a258-227b-4e31-a9cf-717495945fc2" is known client ID for PowerShell

#Really this should be saved somewhere else as a secure string.
#Set-GraphOptions  -ClientSecret  "xxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
#Set-GraphOptions  -Client_Secret (Import-Clixml "$PSScriptRoot\myclientSecret.xml")
