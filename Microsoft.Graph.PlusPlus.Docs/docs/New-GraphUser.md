---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphUser

## SYNOPSIS
Creates a new user in Azure Active directory

## SYNTAX

### DomainFromUPNDisplay
```
New-GraphUser -UserPrincipalName <String> [-MailNickName <String>] -DisplayName <String> [-Manager <String>]
 [-UsageLocation <String>] [-Groups <Object>] [-Roles <Object>] [-Licenses <Object>]
 [-Initialpassword <String>] [-NoPasswordChange] [-ForceMFAPasswordChange] [-PasswordPolicies <String[]>]
 [-SettableProperties <Hashtable>] [-PasswordRule <ScriptBlock>] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### DomainFromUPNLast
```
New-GraphUser -UserPrincipalName <String> [-MailNickName <String>] [-DisplayName <String>] -GivenName <String>
 -Surname <String> [-Manager <String>] [-UsageLocation <String>] [-Groups <Object>] [-Roles <Object>]
 [-Licenses <Object>] [-Initialpassword <String>] [-NoPasswordChange] [-ForceMFAPasswordChange]
 [-PasswordPolicies <String[]>] [-SettableProperties <Hashtable>] [-DisplayNameRule <ScriptBlock>]
 [-NickNameRule <ScriptBlock>] [-PasswordRule <ScriptBlock>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### UPNFromDomainDisplay
```
New-GraphUser -MailNickName <String> [-Domain <String>] -DisplayName <String> [-Manager <String>]
 [-UsageLocation <String>] [-Groups <Object>] [-Roles <Object>] [-Licenses <Object>]
 [-Initialpassword <String>] [-NoPasswordChange] [-ForceMFAPasswordChange] [-PasswordPolicies <String[]>]
 [-SettableProperties <Hashtable>] [-PasswordRule <ScriptBlock>] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### UPNFromDomainLast
```
New-GraphUser [-MailNickName <String>] [-Domain <String>] -GivenName <String> -Surname <String>
 [-Manager <String>] [-UsageLocation <String>] [-Groups <Object>] [-Roles <Object>] [-Licenses <Object>]
 [-Initialpassword <String>] [-NoPasswordChange] [-ForceMFAPasswordChange] [-PasswordPolicies <String[]>]
 [-SettableProperties <Hashtable>] [-DisplayNameRule <ScriptBlock>] [-NickNameRule <ScriptBlock>]
 [-PasswordRule <ScriptBlock>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -UserPrincipalName
User principal name for the new user.
If not specified it can be built by specifying Mail nickname and domain name.

```yaml
Type: String
Parameter Sets: DomainFromUPNDisplay, DomainFromUPNLast
Aliases: UPN

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MailNickName
Mail nickname for the new user.
If not specified the part of the UPN before the @sign will be used, or using the displayname or first/last name

```yaml
Type: String
Parameter Sets: DomainFromUPNDisplay, DomainFromUPNLast, UPNFromDomainLast
Aliases: Nickname

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: String
Parameter Sets: UPNFromDomainDisplay
Aliases: Nickname

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Domain
Domain for the new user - used to create UPN name if the UPN paramater is not provided

```yaml
Type: String
Parameter Sets: UPNFromDomainDisplay, UPNFromDomainLast
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayName
The name displayed in the address book for the user.
This is usually the combination of the user''s first name, middle initial and last name.
This property is required when a user is created and it cannot be cleared during updates.

```yaml
Type: String
Parameter Sets: DomainFromUPNDisplay, UPNFromDomainDisplay
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: String
Parameter Sets: DomainFromUPNLast
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GivenName
The given name (first name) of the user.

```yaml
Type: String
Parameter Sets: DomainFromUPNLast, UPNFromDomainLast
Aliases: FirstName

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Surname
User's last / family name

```yaml
Type: String
Parameter Sets: DomainFromUPNLast, UPNFromDomainLast
Aliases: LastName

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
ID or UserPrincipalName of the user's manager

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UsageLocation
A two letter country code (ISO standard 3166).
Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries. 
Examples include: 'US', 'JP', and 'GB'

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: $Script:DefaultUsageLocation
Accept pipeline input: False
Accept wildcard characters: False
```

### -Groups
{{ Fill Groups Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Roles
{{ Fill Roles Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Licenses
{{ Fill Licenses Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Initialpassword
The initial password for the user.
If none is specified one will be generated and output by the command

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPasswordChange
If specified the user will not have to change their password on first logon

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForceMFAPasswordChange
If specified the user will need to use Multi-factor authentication when changing their password.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -PasswordPolicies
Specifies built-in password policies to apply to the user

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SettableProperties
A hash table of properties which can be passed as parameters to Set-GraphUser command after the account is created

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayNameRule
A script block specifying how the displayname should be built, by default it is {"$GivenName $Surname"};

```yaml
Type: ScriptBlock
Parameter Sets: DomainFromUPNLast, UPNFromDomainLast
Aliases:

Required: False
Position: Named
Default value: {"$GivenName $Surname"}
Accept pipeline input: False
Accept wildcard characters: False
```

### -NickNameRule
A script block specifying how the mailnickname should be built, by default it is $GivenName.$Surname with punctuation removed;

```yaml
Type: ScriptBlock
Parameter Sets: DomainFromUPNLast, UPNFromDomainLast
Aliases:

Required: False
Position: Named
Default value: {($GivenName -replace '\W','') +'.' + ($Surname -replace '\W','')}
Accept pipeline input: False
Accept wildcard characters: False
```

### -PasswordRule
A script block specifying how to create a password, by default a date between 1800 and 2199 like 10Oct2126 - easy to type and meets complexity rules.

```yaml
Type: ScriptBlock
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: {([datetime]"1/1/1800").AddDays((Get-Random 146000)).tostring("ddMMMyyyy")}
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
If specified prevents any confirmation dialog from appearing

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
