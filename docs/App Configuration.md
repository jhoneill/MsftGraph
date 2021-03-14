Microsoft.Graph.PlusPlus is a set of extensions for The Microsoft.Graph powershell modules. It makes the same calls to the Graph APIs but tries to provide more intelligent PowerShell commands. 

When anything calls the Graph APIs it needs an *access token* which says
-   This program
-   Is authorized to use some functionality
-   On behalf of some user

The program might be one the user is running on their own computer, or a service running somewhere, possibly doing something the user agreed to long ago, but doesn't need to approve every time. Although we think of authentication being a two party process - we log on to a service - actually it is a three way process, where the service wants to know if we trust a program to do things for us. 

The Microsoft.Graph.Authentication module has an _alias_ `Connect-Graph` and calling this will call the `Connect-MgGraph` cmdlet
Microsoft.Graph.PlusPlus replaces that alias with its own `Connect-Graph` _function_ with extra functionality, it still calls `Connect-MgGraph`and if no other parameters are provided `Connect-MgGraph` will log a user on using the **Device-Code method** . This method means the app tells the login service "I'm am the app with ID '14d82eec-204b-4c2f-b7e8-296a70dab67e' and **I want to use the these scopes** of functionality. If I send a user to you, can I collect a token from you after they logon?" The service responds "Tell the user to do these steps, which will identify our conversation." The app passes the instructions to the user, and polls the server until it gets a token or a failure message. 

We can **specify a set of scopes **with `Set-GraphOptions -DefaultScopes`. To avoid the need to do this every time,  Microsoft.Graph.PlusPlus runs a settings script, which contains a list of scopes. The script will check for a comma seperated list of scopes in the environment variable `GraphScopes`; and the environment variable is not set the script has its own list. The default script is named `Microsoft.Graph.PlusPlus.settings.ps1` and is in the module directory - it sets other options, you can set scopes via environment variable or via script as you choose. If you don't like changing files in the module directory you can make a copy and point to it with the environment variable `GraphSettingsPath`. If the default is not set or passed as a parameter no scopes will be passed to `Connect-MgGraph`, which will assumes a basic set.    

When the user enters the device code the login service tells them *which app* is trying to get permission to act on their behalf.    
<img Alt="The application is identified to the user before they logon" src=DeviceLogon.png width=25% />    
and before completing a logon the service can check the user is happy for application "X" to do A,B and C on for them.     
<img Alt="Interactive Consent dialog requesting access to user's OneNote" src=InteractiveConsent.png  width=25% />    
In the picture above, a different app (named "MobulaPS") has told the login service the **scope** of actions it wants to carry out. Because this user has not given this app access to that scope before, the logon service displays a **consent** dialog - consent remains unless revoked later, so the user doesn't need to re-approve apps' use of scopes. If the user agress the app receives a time-limitted access token access token and may also get a refresh token - which `ConnectMg-Graph` encrypts and saves in your `.graph` directory: it can present the *referesh* token and get a new *access* token without needing another logon.

For the first screen shot I used a few lines of powerShell to request a device logon with the App-ID of Microsoft Graph PowerShell (14d82eec-204b-4c2f-b7e8-296a70dab67e) - the dialog is not so much asking "Are you really using Microsoft Graph PowerShell" as telling me "if you logon now, whatever sent you here will be able to use scopes you have granted to this app".

*Users* can consent to apps' use of many scopes, but some scopes require *administrator* consent before use. The screen shot below shows a third app which cannot procede because it is asking to use a scope which requires admin consent.     
<img ALT="Consent dialog reporting that only an admin can grant permission" src=AAD-Admin-Consent-Needed.png Width=15%/>

To simplify the process of Admin consent, the dialog shows Admins checkbox to consent to the use of scopes and this pre-approves them for all users of the tenant.    
<img ALT="Consent dialog reporting that only an admin can grant permission" src=AAD-Admin-Consent-At-Logon.png  Width=33%/>

You can grant consents to a published app - like *Microsoft Graph PowerShell* - via Enterpise Apps in Azure Active Directory, or you can create your own apps from App registrations and use them as an alternative. 

`Connect-MgGraph` provides **"Bring your own token"** support so can we skip the Device-logon process. `Connect-Graph` uses this to provide **three additional ways to obtain a token, without a dialog between user and logon service**. You can set a custom App ID for a username and password logon by calling `Set-GraphOptions -ClientID` in the same script which sets the scopes, the default script sets the client ID to the app ID of *Microsoft Graph PowerShell*. Some app IDs also require an associated secret to the log the user on and some don't. Removing the dialog means the logon service can't ask for consent for an app to use a set of scopes, so consent must have been given to the app beforehand.

As well as users being *security principals*, who consent to some scoped permission being *delegated* to apps, apps themselves can be security principals and receive consent to  act in their own right. If you provide your Tenant ID, the ID of an app registered in the tenant, and the app's secret, you can run `Connect-Graph -AsApp` to logon as the app. This is the second of the three extra logon methods: the *Microsoft Graph PowerShell* app isn't a security principal in your tenant, and your tenant doesn't have a secret (password) for it. So it can be used for username/password logon (without a secret) - but `-AsApp` logon requires a *registered* app. 

 You can set the tenant ID in the settings script with `Set-GraphOptions -TenantID` and the secret with "`Set-GraphOptions -TenantID` If TenantID and client ID are set `Connect-Graph` enables the `Credential` parameter (which takes a credential object ) and if the client secret is set it enables the `AsApp` switch parameter
 
The tenant id appears in many places including on a registered app's *overview* page, where *Client ID* also appears:    
<img ALT="Overview page for an app in Azure AD showing the App and Tenant IDs" src=AAD-AppOverview.png width=33%/>    
If you want to logon as the app you will need a secret which you add the from _certificates and Secrets_.    
<img ALT="The app secrets page" src=AAD-AppSecrets.png width=33%/>   
You only get one chance to copy the whole secret - you can see mine begins "a1c". These 3 pieces of information go into the settings file.     
<img ALT="The auth-settings file with the secret, the client/app ID and the tenant id" SRC=AuthSettings.png width=25%/>

In *API Permissions* for a registered app you can click `+` and select the scopes that can be used by a logon  using this app ID  - often the scopes will be Microsoft Graph functions but for the example below I have used the Azure keyvault    
<img ALT="Vault Permissions showing the target URL and the scopes that may be selected" src=VaultPerms.png width=33% />    
At the top of the *permissions page* is the URL that will be called for the rest API - this forms part of the token request, and 'https://graph.microsoft.com' is hardcoded where we are using Microsoft Graph. To use this App ID to work with Key Vault, a script would specify that it wants to token for 'https://Vault.azure.com' instead. Below that the pages shows a choice: are these *delegated permissions* that the app can excercise for a user, or are they *App permissions* which it can excercise in its own right ?. In the screen shot I'm adding the only scope - to have access to the vault as the user - and this doesn't need to be authorized by administrator. Back on the permissions summary page we can see one of the graph scopes DOES need admin consent before it can be used, and the Microsoft Graph scopes are already pre-authorized. Clicking "Grant Adminsitrative consent" will pre-authorize the newly added scope for all users.   
<img ALT="Permissions showing a new scope pending admin consent and options which do and do not allow user consent" src=AppPermSummary.png width=33% /> 

**So, we can consent to scopes for both modes: logging as the app and logging on as a user.**

If `Connect-Graph` has the ID of an app with some consent Azure AD it can also be run as `Connect-Graph -Credential $someCred` it calls the login service and says "A user has given me their credentials, please can I have an access token to use my scopes as that user." Since it is handles user credentials the PowerShell code needs to be trusted. 

The final method for logging on is to piggy back on an existing Azure session current versions of the Az.Accounts module (v.2.2.6 at the time of writing, but not the V1 release), include a command `Get-AzAccessToken` - Justin Grote spotted that this could be used. It appears that Az Accounts logs in with a well-known accountID ( "1950a258-227b-4e31-a9cf-717495945fc2" - "Microsoft Azure PowerShell") Calling `Get-AzAccessToken` appears to say to the login service "I'm already logged authorised for this user and app, can I have a new token targetting a different service." `Connect-Graph -FromAzureSession` calls `Get-AzAccessToken` for you. 

All three of the Bring-your-own-token extensions try to track expiry of the access token and get a new one. There are a few cases where this not happen automatically. Running invoke-graphRequest "v1.0/me" should update the token and give you information about the logged on account. 

## Things which go wrong.

1. You don't set any information (via the settings file or otherwise). The `-Credential` and `-AsApp` paramters do not appear.
1. You use the wrong tennantID, so you try to login as you@yourdomain using my tenant. The logon process will set you straight on that.
1. You use a client (APP) ID that isn't known to your tenant, again the logon process will fail. 
1. You try to use -AsApp login with a secret which doesn't match the client ID (for example you don't change the client ID from 14d82eec-204b-4c2f-b7e8-296a70dab67e) 
1. You don't **grant the right scopes**. Either the application doesn't request them when doing an intereactive login (specify them in the settings file), or the app hasn't been assigned them. If you use the -FromAzureSession option you don't get to choose your client ID, and you can't extend the scopes.  
