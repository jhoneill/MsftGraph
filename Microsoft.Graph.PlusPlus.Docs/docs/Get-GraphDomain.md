---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphDomain

## SYNOPSIS
Gets domains in the current tenant

## SYNTAX

### None (Default)
```
Get-GraphDomain [<CommonParameters>]
```

### NameRef
```
Get-GraphDomain [-Domain] <Object> [-NameReferenceList] [<CommonParameters>]
```

### SCRecords
```
Get-GraphDomain [-Domain] <Object> [-ServiceConfigurationRecords] [<CommonParameters>]
```

### VDRecords
```
Get-GraphDomain [-Domain] <Object> [-VerificationDNSRecords] [<CommonParameters>]
```

### Domain
```
Get-GraphDomain [-Domain] <Object> [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use at least the Directory.Read.All scope

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Domain
{{ Fill Domain Description }}

```yaml
Type: Object
Parameter Sets: NameRef, SCRecords, VDRecords, Domain
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -VerificationDNSRecords
{{ Fill VerificationDNSRecords Description }}

```yaml
Type: SwitchParameter
Parameter Sets: VDRecords
Aliases: VR

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ServiceConfigurationRecords
{{ Fill ServiceConfigurationRecords Description }}

```yaml
Type: SwitchParameter
Parameter Sets: SCRecords
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NameReferenceList
{{ Fill NameReferenceList Description }}

```yaml
Type: SwitchParameter
Parameter Sets: NameRef
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.IMicrosoftGraphDomain
## NOTES

## RELATED LINKS
