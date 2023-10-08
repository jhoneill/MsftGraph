---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Show-GraphSession

## SYNOPSIS
Returns information about the current sesssion

## SYNTAX

### None (Default)
```
Show-GraphSession [-Force] [<CommonParameters>]
```

### Who
```
Show-GraphSession [-Who] [-Force] [<CommonParameters>]
```

### Scopes
```
Show-GraphSession [-Scopes] [-Force] [<CommonParameters>]
```

### Options
```
Show-GraphSession [-Options] [-Force] [<CommonParameters>]
```

### AppName
```
Show-GraphSession [-AppName] [-Force] [<CommonParameters>]
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

### -Who
If specified returns only the current account

```yaml
Type: SwitchParameter
Parameter Sets: Who
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scopes
If specified returns only the scopes available to the current session

```yaml
Type: SwitchParameter
Parameter Sets: Scopes
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
If specified returns the options set using Set-GraphOption

```yaml
Type: SwitchParameter
Parameter Sets: Options
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppName
If specified returns the current app name.

```yaml
Type: SwitchParameter
Parameter Sets: AppName
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
If specified runs Test-GraphSession to ensure a session exists.

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

### System.String
## NOTES

## RELATED LINKS
