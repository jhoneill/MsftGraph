---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphSiteColumn

## SYNOPSIS
Gets a column which is defined for the whole site.

## SYNTAX

### None (Default)
```
Get-GraphSiteColumn [-AllowMultiple] [-Raw] [<CommonParameters>]
```

### Terms
```
Get-GraphSiteColumn [[-Name] <String>] [[-ColumnGroup] <String>] [[-ID] <String>] [-AllowMultiple] [-Raw]
 [<CommonParameters>]
```

### WhereClause
```
Get-GraphSiteColumn [-Where <ScriptBlock>] [-AllowMultiple] [-Raw] [<CommonParameters>]
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

### -Name
Selects column(s) by name (and possibly group)

```yaml
Type: String
Parameter Sets: Terms
Aliases:

Required: False
Position: 1
Default value: *
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ColumnGroup
Selects column(s) by group (and possibly by name)

```yaml
Type: String
Parameter Sets: Terms
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ID
Selects a column by unique ID

```yaml
Type: String
Parameter Sets: Terms
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Where
Allows a custom where clause instead of Name and/or group and/or ID

```yaml
Type: ScriptBlock
Parameter Sets: WhereClause
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowMultiple
Get-GraphSiteColumn is intended to return one column to used when creating a new list, so
    if multiple columns are returned that would be an error (i.e.
two columns have the
    same name and group wasn't given.) If -allowMultiple is specified it is *not* treated as an error

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

### -Raw
{{ Fill Raw Description }}

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

## NOTES

## RELATED LINKS
