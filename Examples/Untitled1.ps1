

$ClientID    = "bf546ecc-067d-4030-9edd-7b0d74913411" 
$scope           = @(
'People.Read.All',
'Reports.Read.All',
'User.ReadWrite.all',
'Group.ReadWrite.All',
'Files.ReadWrite.All',
'Sites.ReadWrite.All',
'Sites.Manage.All',        
'Calendars.ReadWrite',
'Calendars.ReadWrite.Shared'
'Contacts.ReadWrite',
'Contacts.ReadWrite.Shared',
'Mail.ReadWrite',
'MailboxSettings.ReadWrite',
'Notes.ReadWrite',
'Notes.Create',
'Directory.AccessAsUser.All',
'openid',
'profile',
'offline_access'
)


$CallBackUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"   
$AuthUri     = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize' + 
                '?client_id='    +  $ClientID + 
                '&scope='        + ($Scope -join '%20')  + 
                '&redirect_uri=' +  $CallBackUri+ 
                '&response_type=code' 

$DocComp  = { #script block for the on document complete event: Make URI accessible; close the form if URI has a code or an error
    $Script:uri = $web.Url.AbsoluteUri
    if ($Script:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
}
#Point a web browser control at the Auth URI 
$web      = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=2000;Height=1600;Url=$AuthUri }
$form     = New-Object -TypeName System.Windows.Forms.Form       -Property @{Width=2050;Height=1820}
$web.Add_DocumentCompleted($DocComp) #Add the event handler to the web control
$form.Controls.Add($web)             #Add the control to the form
$form.Add_Shown({$form.Activate()})
$form.ShowDialog() | Out-Null

#$URI will be set by the event handler ... so did we get a code - meaning the user logged in OK - or did we get an error ?
if     ($uri -match "error=([^&]*)") {Write-Warning ("Logon returned an error of " + $Matches[1]); return}
elseif ($Uri -match "code=([^&]*)" ) {# If we got a code, request & process the token for it
    $webparams = @{
        'Body'        = @{'grant_type'= 'authorization_code'; 'client_id' = $ClientID ; 'redirect_uri' = $CallBackUri; 'code'= $Matches[1]} 
        'uri'         = 'https://login.microsoftonline.com/common/oauth2/v2.0/token' 
        'Method'      = 'Post'
    }
    $oauthUser = Invoke-RestMethod @webParams   
}

$defaultheader = @{'Authorization' = "bearer $($oauthUser.access_token)"}
Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me" -Headers $defaultheader