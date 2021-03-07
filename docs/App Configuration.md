When anything calls the Graph APIs it needs an *access token* which says
-   This program
-   Is authorized to use some functionality
-   On behalf of some user

The program might be one the user is running on their own computer, or a service running somewhere, possibly doing something the user agreed to long ago, but doesn't need to approve every time. Although we think of authentication being a two party process - we log on to a service - actually it is a three way process, where the service wants to know if we trust a program to do things for us. 

`Connect-MgGraph` - which underpins `Connect-Graph` - logs users on with the **Device-Code method by default** . So it tells the login service "I'm am the app with ID '14d82eec-204b-4c2f-b7e8-296a70dab67e' and I want to use the these scopes of functionality. If I send a user to you to logon, can you give me a token when they do?" The service responds "Tell the user to do these steps, which will identify our conversation." The app passes the instructions to the user, and polls the server until it gets a token or a failure message. We can specify a set of scopes with `Set-GraphConnectionOptions -DefaultScopes`. When the authentication part of Microsoft.Graph.PlusPlus loads it checks the environment variable `MGSettingsPath` and if it finds it runs that script, otherwise it runs `AuthSettings.ps1`. One of the things this script does is to check another environment variable `GraphScopes` for a comma seperated list of scopes, if none is found it has a default set, and the script doesn't run `Connect-MgGraph` will assume a basic set.    

When the user enters the device code the server starts by telling them *which app* is trying to get permission to act on their behalf.    
<img Alt="The application is identified to the user before they logon" src=DeviceLogon.png width=25% />    
and before completing a logon the process can ask *"Do you want application X to be able to do A,B and C for you"*     
<img Alt="Interactive Consent dialog requesting access to user's OneNote" src=InteractiveConsent.png  width=25% />    
In the picture above, a different app (named "MobulaPS") has told the login service the *scope* of actions it wants to carry out. And the service has seen that this user has not given this app access to that scope before, so it displays a consent dialog - consent can be revoked later but it is remembered so the user doesn't need to re-approve apps use of the necessary scopes. If the user agress the app receices a time-limitted access token access token and may also get a refresh token - which `ConnectMg-Graph` encrypts and saves in your .graph directory: it can present the referesh token and get a new access token without needing another logon.

For the first screen shot I used a few lines of powerShell to request a device logon with the ID of the Microsoft Graph PowerShell app (14d82eec-204b-4c2f-b7e8-296a70dab67e) - the dialog is not so much asking "Are you really using Microsoft Graph PowerShell" as telling me "if you logon now, whatever sent you here will be able to use scopes you have granted to this app".

Users can consent to apps' use of many scopes, but some scopes must require administrator consent before use. The screen shot below shows a third app which cannot procede because it is asking to use a scope which requires admin consent.     
<img ALT="Consent dialog reporting that only an admin can grant permission" src=AAD-Admin-Consent-Needed.png Width=15%/>

To simplify the process of Admin consent, the dialog shows Admins checkbox to consent to the use of scopes and this pre-approves them for all users of the tenant.    
<img ALT="Consent dialog reporting that only an admin can grant permission" src=AAD-Admin-Consent-At-Logon.png  Width=33%/>

You can grant consents to a published app - like the Graph PowerShell SDK - via Enterpise Apps in Azure Active Directory, or you can create your own apps from App registrations and tell `Connect-MGGraph` to use their ID instead of the ID for *Microsoft Graph PowerShell* Connect-Graph makes this easier, because the same script which sets the scopes can also call `Set-GraphConnectionOptions  -ClientID` and it will be used automatically. 

`Connect-MgGraph provides` **"Bring your own token" support** so can we skip the Device logon process. `Connect-Graph` **has three ways to use this obtaining a token without a dialog between user and logon service**. This means the logon service can't ask "Do you want to grant this access to this app" this access", so consent must have been given beforehand - and if you prefer to grant permissions to your own app, Connect-Graph also streamlines the process of providing a custom app ID.

As well as users being security principals, who consent to some scoped permission being delegated to apps, apps themselves can be security principals and receive consent to  act in their own right. If you provide the right information you can run `Connect-Graph -AsApp` to logon as an app defined in your tennat. This is the first of the three extra logon methods, but you can't use it to logon as the *Microsoft Graph PowerShell* app - it isn't a security principal in your tenant, and you don't have its secret (password). 

`Connect-Graph` needs 3 peices of information to use an app you have defined to bypass the device logon. The *Tenant ID* appears in many places including on the app's *overview* page, where *Client ID* appears:    
<img ALT="Overview page for an app in Azure AD showing the App and Tenant IDs" src=AAD-AppOverview.png width=33%/>    
As well as client and tennant IDs you need a secret which you add the secret from _certificates and Secrets_.    
<img ALT="The app secrets page" src=AAD-AppSecrets.png width=33%/>
You only get one chance to copy the whole secret - you can see mine begins "a1c". These 3 pieces of information go into the settings file.     
<img ALT="The auth-settings file with the secret, the client/app ID and the tenant id" SRC=AuthSettings.png width=25%/>

Those steps provided an app to use but without any consents for it access anything. In API Permissions for a registered app you can click + and select something to grant permisions too - often this will be Microsoft Graph but for the example below I have used the Azure keyvault    
<img ALT="Vault Permissions showing the target URL and the scopes that may be selected" src=VaultPerms.png width=33% />    
At the top of the *permissions page* is the URL that will be called for the rest API - this forms part of the token request, and 'https://graph.microsoft.com' is hardcoded where we are using Microsoft Graph. To use this App ID in a script that worked with Key Vault would need a different token request. Below that the pages shows a choice: are these *delegated permissions* the app can excercise for a user, or are they *App permissions*?. In the screen shot I'm adding the only scope - to have access to the vault as the user - and this doesn't need to be authorized by administrator. Back on the permissions summary page we can see one of the graph scopes DOES need admin consent before it can be used, and the Microsoft Graph scopes are already pre-authorized. Clicking "Grant Adminsitrative consent" will pre-authorize the newly added scope for all users.

**So, we can consent to scopes for both modes of logging as the app and logging on as a user.**

If `Connect-Graph` knows the details of app an registered in Azure AD (including the secret, so PowerShell code is trusted from the App's viewpoint) it can also be run as `Connect-Graph -Credential $someCred` Since it is getting user credentials the PowerShell code needs to be trusted from the user's viewpoint as well. Now it can can call the login service, authenticate, and say "A user has given me their credentials, please can I have an access token to use my scopes as that user." 

The final method for logging on is to piggy back on an existing Azure session current versions of the Az.Accounts module (v.2.2.6) at the time of writing, but not the V1 release, include a Command `Get-AzAccessToken` - Justin Grote spotted that this could be used. It appears that Az Accounts logs in with a well-known accountID ( "1950a258-227b-4e31-a9cf-717495945fc2" - "Microsoft Azure PowerShell") Calling `Get-AzAccessToken` appears to say to the login service "I'm already logged authorised for this user and app, can I have a new token targetting a different service." `Connect-Graph -FromAzureSession` calls `Get-AzAccessToken` for you. 

All three of the Bring-your-own-token extensions try to track expiry of the access token and get a new one. There are a few cases where this not happen automatically. Running invoke-graphRequest "v1.0/me" should update the token and give you information about the logged on account. 

## Things which go wrong.

1. You don't set any information (via the settings file or otherwise)
1. You use the wrong tennantID, so you try to login as you@yourdomain using my tenant. The logon process will set you straight on that.
1. You use a client (APP) ID that isn't known to your tenant. 
1. You try to use the client ID for an app for which you don't have a corresponding secret, (like 14d82eec-204b-4c2f-b7e8-296a70dab67e Microsoft Graph PowerShell) in a mode for -credential or -asapp login
1. You don't grant the right scopes. Either the application doesn't request them when doing an intereactive login, or the app hasn't been assigned them. If you use the -FromAzureSession option you don't get to choose your client ID, and you can't extend the scopes.  
