---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Grant-GraphLicense

## SYNOPSIS
Grants the licence to use a particular stock-keeping-unit (SKU) to users or groups

## SYNTAX

### ByUserID (Default)
```
Grant-GraphLicense [-SKUID] <Object> [-UserID] <Object> [-DisabledPlans <String[]>] [-UsageLocation <String>]
 [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ByGroupID
```
Grant-GraphLicense [-SKUID] <Object> [-GroupID] <Object> [-DisabledPlans <String[]>] [-UsageLocation <String>]
 [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
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
Accept pipeline input: False
Accept wildcard characters: False
```

### -UserID
ID(s) for users to receive permission ("me" will select the current user), the command will accept user objects and attempt to resolve names to IDs

```yaml
Type: Object
Parameter Sets: ByUserID
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -GroupID
ID(s) for group(s) to receive permission, the command will accept group objects and attempt to resolve names to IDs

```yaml
Type: Object
Parameter Sets: ByGroupID
Aliases: Team

Required: True
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisabledPlans
Disables individual parts of the the SKU

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

### -UsageLocation
A two letter country code (ISO standard 3166).
Examples include: 'US', 'JP', and 'GB' Can be set/reset here

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

### -Force
Runs the command without a confirmation dialog

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

## NOTES

## RELATED LINKS
