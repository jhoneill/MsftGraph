---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphContact

## SYNOPSIS
Adds an entry to the current users Outlook contacts

## SYNTAX

```
New-GraphContact [[-GivenName] <Object>] [[-MiddleName] <Object>] [[-Initials] <Object>] [[-Surname] <Object>]
 [[-NickName] <Object>] [[-FileAs] <Object>] [[-DisplayName] <Object>] [[-CompanyName] <Object>]
 [[-JobTitle] <Object>] [[-Department] <Object>] [[-Manager] <Object>] [[-Email] <Object>] [[-IM] <Object>]
 [[-MobilePhone] <Object>] [[-BusinessPhones] <Object>] [[-HomePhones] <Object>] [[-Homeaddress] <Object>]
 [[-BusinessAddress] <Object>] [[-OtherAddress] <Object>] [[-Categories] <Object>] [[-Birthday] <DateTime>]
 [[-PersonalNotes] <Object>] [[-Profession] <Object>] [[-AssistantName] <Object>] [[-Children] <Object>]
 [[-SpouseName] <Object>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Almost all the paramters can be accepted form a piped object to make import easier.

## EXAMPLES

### EXAMPLE 1
```
New-GraphContact -GivenName Pavel -Surname Bansky -Email pavelb@fabrikam.onmicrosoft.com -BusinessPhones  "+1 732 555 0102"
Creates a new contact; if no displayname is given, one will be decided using given name and suranme;
```

### EXAMPLE 2
```
>$PavelMail = New-GraphRecipient -DisplayName "Pavel Bansky [Fabikam]" -Mail  pavelb@fabrikam.onmicrosoft.com
>New-GraphContact -GivenName Pavel -Surname Bansky -Email $pavelmail  -BusinessPhones  "+1 732 555 0102"
 This creates the same contanct but sets up their email with a display name.
 New recipient creates a hash table
 @{'emailaddress' = @ {
         'name' = 'Pavel Bansky [Fabikam]'
         'address' = 'pavelb@fabrikam.onmicrosoft.com'
     }
 }
```

## PARAMETERS

### -GivenName
{{ Fill GivenName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MiddleName
{{ Fill MiddleName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Initials
{{ Fill Initials Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Surname
{{ Fill Surname Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -NickName
{{ Fill NickName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -FileAs
{{ Fill FileAs Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -DisplayName
{{ Fill DisplayName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -CompanyName
{{ Fill CompanyName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -JobTitle
{{ Fill JobTitle Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Department
{{ Fill Department Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Manager
{{ Fill Manager Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Email
One or more mail addresses, as a single string with semi colons between addresses or as an array of strings or MailAddress objects created with New-GraphMailAddress

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -IM
One or more instant messaging addresses, as an array or as a single string with semi colons between addresses

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 13
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MobilePhone
A single mobile phone number

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 14
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -BusinessPhones
One or more Business phones either as an array or as single string with semi colons between numbers

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 15
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -HomePhones
One or more home phones either as an array or as single string with semi colons between numbers

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 16
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Homeaddress
An address object created with  New-GraphPhysicalAddress

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 17
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -BusinessAddress
An address object created with  New-GraphPhysicalAddress

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 18
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -OtherAddress
An address object created with  New-GraphPhysicalAddress

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 19
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Categories
One or more categories either as an array or as single string with semi colons between them.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 20
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Birthday
The contact's Birthday as a date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: 21
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PersonalNotes
{{ Fill PersonalNotes Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 22
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Profession
{{ Fill Profession Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 23
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -AssistantName
{{ Fill AssistantName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 24
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Children
{{ Fill Children Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 25
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -SpouseName
{{ Fill SpouseName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 26
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Force
If sepcified the contact will be created without prompting for confirmation.
This is the default state but can change with the setting of confirmPreference.

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

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact
## NOTES

## RELATED LINKS
