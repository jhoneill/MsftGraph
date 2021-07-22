---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphLookupColumn

## SYNOPSIS
Creates a definition of a Sharepoint lookup column

## SYNTAX

```
New-GraphLookupColumn [-ListID] <String> [-ColumnName] <String> [-PrimaryLookupColumnID <String>]
 [-MultipleSelection] [-AllowUnlimitedLength] [<CommonParameters>]
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

### -ListID
The unique identifier of the lookup source list.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnName
The name of the lookup source column.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PrimaryLookupColumnID
If specified, this column is a secondary lookup, pulling an additional field from the list item looked up by the primary lookup.
Use the list item looked up by the primary as the source for the column named here

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

### -MultipleSelection
If Specified allows multiple/values to be specified

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

### -AllowUnlimitedLength
Specified to indicates that values in the column should be able to exceed the standard limit of 255 characters.

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

### System.Collections.Hashtable
## NOTES

## RELATED LINKS

[https://docs.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0)

