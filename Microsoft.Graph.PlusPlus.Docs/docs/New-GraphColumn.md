---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphColumn

## SYNOPSIS
Create a new Column definition for a sharepoint list

## SYNTAX

### None (Default)
```
New-GraphColumn [-Name] <String> [-ColumnDefinition] <Hashtable> [-ColumnGroup <String>]
 [-Description <String>] [-DisplayName <String>] [-Indexed <Boolean>] [-ReadOnly <Boolean>]
 [-Required <Boolean>] [-EnforceUniqueValues <Boolean>] [-HIDden <Boolean>] [<CommonParameters>]
```

### DefaultbyFormula
```
New-GraphColumn [-Name] <String> [-ColumnDefinition] <Hashtable> [-ColumnGroup <String>]
 [-Description <String>] [-DisplayName <String>] -DefaultValueFormula <String> [-Indexed <Boolean>]
 [-ReadOnly <Boolean>] [-Required <Boolean>] [-EnforceUniqueValues <Boolean>] [-HIDden <Boolean>]
 [<CommonParameters>]
```

### DefaultbyValue
```
New-GraphColumn [-Name] <String> [-ColumnDefinition] <Hashtable> [-ColumnGroup <String>]
 [-Description <String>] [-DisplayName <String>] -DefaultValueString <String> [-Indexed <Boolean>]
 [-ReadOnly <Boolean>] [-Required <Boolean>] [-EnforceUniqueValues <Boolean>] [-HIDden <Boolean>]
 [<CommonParameters>]
```

## DESCRIPTION
New-GraphList uses column definitions to set up a new list.
Each column has a name, description, default and one of the properties from the following list
boolean, calculated, choice, currency, dateTime, lookup, number, personOrGroup or text
Flags can also be set to say if the column is indexed, Readonly and/or required.
Existing Columns defined in the site can be fetched with Get-GraphSiteColumn
New-GraphColumn defines a new column to be included in a list, and a typical list will need
multiple columns, which may be a mixture of new and existing columns.
The specifics of each of the column types is handled by a new-{typeName}Column command.
Examples appear in New-GraphList

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Name
The API-facing name of the column as it appears in the fields on a listItem.
For the user-facing name, see displayName.

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

### -ColumnDefinition
A definition created with on of the New-*Column commands for a text, currency, boolean etc

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnGroup
For site columns, the name of the group this column belongs to.
Helps organize related columns.

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

### -Description
The user-facing description of the column.

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

### -DisplayName
The user-facing name of the column.

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

### -DefaultValueFormula
Fills in the default value using a formula

```yaml
Type: String
Parameter Sets: DefaultbyFormula
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultValueString
Fills in the defaultt value using a fixed value

```yaml
Type: String
Parameter Sets: DefaultbyValue
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Indexed
If specified the column is indexed to help the perfomance of searching and grouping.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReadOnly
Specifies whether the column values can be modified.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Required
Specifies whether the column value is not optional.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -EnforceUniqueValues
If true, no two list items may have the same value for this column.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -HIDden
Specifies whether the column is displayed in the user interface.

```yaml
Type: Boolean
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

[https://docs.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0)

[New-GraphList]()

