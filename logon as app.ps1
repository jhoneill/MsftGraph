$tenant    = 'GUID or domain.com'
$appId     = 'GUID from Azure AD managment page'
$appSecret = 'somehwere else'

<#
Don't store app secrets in the scriptt
$Secret = "Don't_use_Pa55word"

ConvertTo-SecureString -String $Secret -AsPlainText -Force | ConvertFrom-SecureString > PasswordTest.txt

$savedPassword  = Get-Content -Path PasswordTest.txt | ConvertTo-SecureString
$Secret         = [pscredential]::new('DummyUserName',$savedPassword).GetNetworkCredential().Password

#>

$URI       =  'https://login.microsoft.com/{0}/oauth2/token' -f $tenant
$oauthAPP  = Invoke-RestMethod -Method Post -Uri $URI -Body @{
        grant_type    = 'client_credentials';
        client_id     =  $appid ;
        client_secret =  $appSecret;
        resource      = 'https://graph.microsoft.com';
}


$Defaultheader = @{Authorization="$($oauthapp.token_type) $($oauthapp.access_token)"}
