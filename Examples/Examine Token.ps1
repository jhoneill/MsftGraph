# Examine Accesstoken.

$authResponse =  Get-AccessToken -Resoure  'https://api.securitycenter.microsoft.com'
$token        = $authResponse.access_token

#$url = "https://api.security.microsoft.com/api/incidents

$token.split('.')[1] | ForEach-Object{ if ($_.Length % 4) {(( $_ + ('=' * (4 -($_.Length % 4)))))} else {$_}} |
                       ForEach-Object {[string]::new([convert]::FromBase64String($_)) |
                       ConvertFrom-Json}

# aud       = audience  <-- or resource
# iss       = issuer
# tid       = tennantId  <--
# iat       = issued at
# nbf       = not before
# exp       = expires
# Appid     = App ID
# sub       = Tennant specific AD for this app.
# appidacr  = Indicates how the client was authenticated. For a public client, the value is "0". If client ID and client secret are used, the value is "1". If a client certificate was used for authentication, the value is "2".
# idp       = ID provider same as issuer
# oid       = object ID for the account signed in .
# roles   - aka scopes

$headers = @{
    'Content-Type' = 'application/json'
    'Accept' = 'application/json'
    'Authorization' = "Bearer $token"
}