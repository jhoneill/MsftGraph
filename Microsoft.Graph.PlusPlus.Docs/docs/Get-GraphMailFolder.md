---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphMailFolder

## SYNOPSIS
Get the user's Mailbox folders

## SYNTAX

### FilterByName (Default)
```
Get-GraphMailFolder [[-Name] <String>] [-ParentFolder <Object>] [-User <String>] [-Top <Int32>]
 [-Select <String[]>] [-ChildItems] [<CommonParameters>]
```

### Sorted
```
Get-GraphMailFolder [-ParentFolder <Object>] [-User <String>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] [-ChildItems] [<CommonParameters>]
```

### FilterByString
```
Get-GraphMailFolder [-ParentFolder <Object>] [-User <String>] [-Top <Int32>] [-Select <String[]>]
 [-Filter <String>] [-ChildItems] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphMailFolderList -Name inbox
Gets the current users inbox folder
```

## PARAMETERS

### -Name
Filter the folders returned by a name

```yaml
Type: String
Parameter Sets: FilterByName
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ParentFolder
{{ Fill ParentFolder Description }}

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

### -Top
Select the first n folders.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 100
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
fields to select in the query - will add a validate set later

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

### -OrderBy
String with orderby clause e.g.
"name", "lastmodifiedDate desc"

```yaml
Type: String
Parameter Sets: Sorted
Aliases:

Required: False
Position: Named
Default value: Displayname
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
A custom filter clause.

```yaml
Type: String
Parameter Sets: FilterByString
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChildItems
{{ Fill ChildItems Description }}

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphMailFolder
## NOTES

## RELATED LINKS
