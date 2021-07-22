---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphContact

## SYNOPSIS
Modifies or adds an entry in the current users Outlook contacts

## SYNTAX

### UpdateContact
```
Set-GraphContact [-Contact] <Object> [-GivenName <Object>] [-MiddleName <Object>] [-Initials <Object>]
 [-Surname <Object>] [-NickName <Object>] [-FileAs <Object>] [-DisplayName <Object>] [-CompanyName <Object>]
 [-JobTitle <Object>] [-Department <Object>] [-Manager <Object>] [-Email <Object>] [-IM <Object>]
 [-MobilePhone <Object>] [-BusinessPhones <Object>] [-HomePhones <Object>] [-Homeaddress <Object>]
 [-BusinessAddress <Object>] [-OtherAddress <Object>] [-Categories <Object>] [-Birthday <DateTime>]
 [-PersonalNotes <Object>] [-Profession <Object>] [-AssistantName <Object>] [-Children <Object>]
 [-SpouseName <Object>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### NewContact
```
Set-GraphContact [-IsNew] [-GivenName <Object>] [-MiddleName <Object>] [-Initials <Object>] [-Surname <Object>]
 [-NickName <Object>] [-FileAs <Object>] [-DisplayName <Object>] [-CompanyName <Object>] [-JobTitle <Object>]
 [-Department <Object>] [-Manager <Object>] [-Email <Object>] [-IM <Object>] [-MobilePhone <Object>]
 [-BusinessPhones <Object>] [-HomePhones <Object>] [-Homeaddress <Object>] [-BusinessAddress <Object>]
 [-OtherAddress <Object>] [-Categories <Object>] [-Birthday <DateTime>] [-PersonalNotes <Object>]
 [-Profession <Object>] [-AssistantName <Object>] [-Children <Object>] [-SpouseName <Object>] [-Force]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
> $pavel = Get-GraphContact -Name pavel
> Set-GraphContact $pavel -CompanyName "Fabrikam" -Birthday "1974-07-22"
The first line gets the Contact which was added in the 'New-GraphContact" example
and the second adds Birthday and Company-name attributes to the contact.
```

### EXAMPLE 2
```
> $fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
> Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
This continues from the previous example, creating an address in the first line
and adding it to the contact in the second.
```

## PARAMETERS

### -Contact
The contact to be updated either as an ID or as contact object containing an ID.

```yaml
Type: Object
Parameter Sets: UpdateContact
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IsNew
If specified, instead of providing a contact, instructs the command to create a contact instead of updating one.

```yaml
Type: SwitchParameter
Parameter Sets: NewContact
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -GivenName
{{ Fill GivenName Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -DisplayName
If not specified a display name will be generated, so updates without the display name may result in overwriting an existing one

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
Position: Named
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
