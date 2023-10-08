---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphGroupList

## SYNOPSIS
Gets a list of Groups in Microsoft Graph.

## SYNTAX

### None (Default)
```
Get-GraphGroupList [-OrderBy <String>] [<CommonParameters>]
```

### FilterByName
```
Get-GraphGroupList [-Name] <String> [-OrderBy <String>] [<CommonParameters>]
```

### SelectFields
```
Get-GraphGroupList -Select <String[]> [-OrderBy <String>] [<CommonParameters>]
```

### Sort
```
Get-GraphGroupList [-OrderBy <String>] [-Descending] [<CommonParameters>]
```

### FilterByString
```
Get-GraphGroupList [-OrderBy <String>] -Filter <String> [<CommonParameters>]
```

## DESCRIPTION
The list of groups returned can be filtered by name (the beginning of the displayname and mail
address are checked) or with a custom filter string, or it can be sorted, or specific fields can
be selected.
However there is a limitation in the graph API which prevent these being combined.
Requires consent to use the Group.Read.All scope.

## EXAMPLES

### EXAMPLE 1
```
Get-GraphGroupList | Format-Table -Autosize  Displayname, SecurityEnabled, Mailenabled, Mail, ID
Displays a table of groups in the current tennant
```

### EXAMPLE 2
```
(Get-GraphGroupList -Name consult* | Get-GraphTeam -Site).weburl
Gets any group whose name begins "Consult" , finds its sharepoint site, and returns the site's URL
```

## PARAMETERS

### -Name
if specified limits the groups returned to those with names begining...

```yaml
Type: String
Parameter Sets: FilterByName
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
Field(s) to select: ID and displayname are always included;
The following are only available when getting a single group:
'allowExternalSenders','autoSubscribeNewMembers','isSubscribedByMail' 'unseenCount',

```yaml
Type: String[]
Parameter Sets: SelectFields
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OrderBy
A field to sort by - sorting is applied on the client side because filter and selct cannot be combined with server-side sort

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: DisplayName
Accept pipeline input: False
Accept wildcard characters: False
```

### -Descending
{{ Fill Descending Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Sort
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
An oData filter string; there is a graph limitation that you can't filter by description or Visibility.

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

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup
## NOTES

## RELATED LINKS
