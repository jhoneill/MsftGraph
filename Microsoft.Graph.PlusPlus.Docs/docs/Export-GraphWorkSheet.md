---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Export-GraphWorkSheet

## SYNOPSIS
{{ Fill in the Synopsis }}

## SYNTAX

### ItemNameRange
```
Export-GraphWorkSheet [-Drive <Object>] [-ItemPath] <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] -RangeName <String> [-dateFormat <String>] [-Show] [<CommonParameters>]
```

### ItemNameTable
```
Export-GraphWorkSheet [-Drive <Object>] [-ItemPath] <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] [-AsTable] [-dateFormat <String>] [-Show] [<CommonParameters>]
```

### ItemName
```
Export-GraphWorkSheet [-Drive <Object>] [-ItemPath] <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] [-dateFormat <String>] [-Show] [<CommonParameters>]
```

### ItemIDRange
```
Export-GraphWorkSheet [-Drive <Object>] -ItemID <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] -RangeName <String> [-dateFormat <String>] [-Show] [<CommonParameters>]
```

### ItemIDTable
```
Export-GraphWorkSheet [-Drive <Object>] -ItemID <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] [-AsTable] [-dateFormat <String>] [-Show] [<CommonParameters>]
```

### ItemID
```
Export-GraphWorkSheet [-Drive <Object>] -ItemID <Object> [[-SheetName] <String>] [-InputObject <Object>]
 [-NoHeader] [-dateFormat <String>] [-Show] [<CommonParameters>]
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

### -AsTable
{{ Fill AsTable Description }}

```yaml
Type: SwitchParameter
Parameter Sets: ItemNameTable, ItemIDTable
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Drive
{{ Fill Drive Description }}

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

### -InputObject
{{ Fill InputObject Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ItemID
{{ Fill ItemID Description }}

```yaml
Type: Object
Parameter Sets: ItemIDRange, ItemIDTable, ItemID
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ItemPath
{{ Fill ItemPath Description }}

```yaml
Type: Object
Parameter Sets: ItemNameRange, ItemNameTable, ItemName
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
{{ Fill NoHeader Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RangeName
{{ Fill RangeName Description }}

```yaml
Type: String
Parameter Sets: ItemNameRange, ItemIDRange
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -SheetName
{{ Fill SheetName Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases: WorkSheetName

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
{{ Fill Show Description }}

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -dateFormat
{{ Fill dateFormat Description }}

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Object
### System.Management.Automation.SwitchParameter
### System.String
## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
