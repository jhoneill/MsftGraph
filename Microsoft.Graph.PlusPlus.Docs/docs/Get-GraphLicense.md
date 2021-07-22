---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphLicense

## SYNOPSIS
Returns users or groups (or both) who are licensed to user a given SKU

## SYNTAX

### None (Default)
```
Get-GraphLicense [-SKUID] <Object> [<CommonParameters>]
```

### Users
```
Get-GraphLicense [-SKUID] <Object> [-UsersOnly] [<CommonParameters>]
```

### Groups
```
Get-GraphLicense [-SKUID] <Object> [-GroupsOnly] [<CommonParameters>]
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

### -SKUID
The SKU to get either as an ID or a SKU object containing an ID

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -UsersOnly
{{ Fill UsersOnly Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Users
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -GroupsOnly
{{ Fill GroupsOnly Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Groups
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
