---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphContact

## SYNOPSIS
Get the user's contacts

## SYNTAX

### None (Default)
```
Get-GraphContact [-User <String>] [-Select <String[]>] [<CommonParameters>]
```

### FilterByName
```
Get-GraphContact [-User <String>] [-Select <String[]>] -Name <String> [<CommonParameters>]
```

### FilterByString
```
Get-GraphContact [-User <String>] [-Select <String[]>] -Filter <String> [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
get-graphContact -name "o'neill" | ft displayname, mobilephone
Gets contacts where the display name, given name, surname, file-as name, or email begins with
O'Neill - note the function handles apostrophe, - a single one would normal cause an error with the query.
The results are displayed as table with display name and mobile number
```

## PARAMETERS

### -User
UserID as a guid or User Principal name.
If not specified defaults to "me"

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

### -Select
A custom set of contact properties to select

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

### -Name
If specified looks for contacts where the display name, file-as Name, given name or surname beging with ...

```yaml
Type: String
Parameter Sets: FilterByName
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
A custom OData Filter String

```yaml
Type: String
Parameter Sets: FilterByString
Aliases:

Required: True
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
