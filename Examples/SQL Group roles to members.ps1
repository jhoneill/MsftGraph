Param (
$server    = "tcp:amsqlprod.database.windows.net,1433",
$databases = @("sqldb-amDev-moj","sqldb-amSIT-moj","sqldb-amUAT-moj"),
$id        = "amsqladmin",
$pass
)
$sql  =  @"
SELECT DP1.name AS DatabaseRoleName,  DP2.name  AS DatabaseUserName ,Dp2.sid AS SID
FROM     sys.database_role_members AS DRM
    JOIN sys.database_principals   AS DP1 ON   DRM.role_principal_id   = DP1.principal_id
    JOIN sys.database_principals   AS DP2 ON   DRM.member_principal_id = DP2.principal_id
WHERE DP1.type = 'R'
  AND DP2.authentication_type_desc ='EXTERNAL'
"@
$grouplist = Get-GraphGroup | Add-Member -MemberType ScriptProperty -Name SQLID -Value { ([guid]::Parse($this.id).ToByteArray() | ForEach-Object {$_.tostring("x2")}) -join ""}  -PassThru
$roles = foreach ($db  in $databases) {
    $connectionString = "Server=$server;Initial Catalog=$db;Persist Security Info=False;User ID=$id;Password=$(Get-AzKeyVault | Get-AzKeyVaultSecret  -Name saPwd -AsPlainText);MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
     Get-SQL -ForceNew -MsSQLserver -Connection $connectionString -SQL $sql -Close | ForEach-Object {
        $sqlRole = $_.DatabaseRoleName
        $sqlId =   ($_.sid | ForEach-Object {$_.tostring('x2') }) -join ""
        $grouplist | Where-Object sqlid -eq $sqlid |
        Get-GraphGroup -Members |
            Select-Object @{n="DatabaseName";e={$db}} , @{n="DatabaseRole";e={$sqlrole}}, GroupName,Displayname,UserPrincipalName,Mail
    }
}

$roles | Sort-Object -Property DatabaseName,DatabaseRole,GroupName, DisplayName  | Format-Table  #or what you will
