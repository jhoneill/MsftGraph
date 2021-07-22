---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Export-GraphGroupMember

## SYNOPSIS
Exports a list of group memberships to a CSV file

## SYNTAX

```
Export-GraphGroupMember [-Group] <Object> [-Path <Object>] [-OrderByGroup] [<CommonParameters>]
```

## DESCRIPTION
Takes a list of groups (as a parameter or from the pipeline)  and creates four columns
* Action is either Add or Remove - on export it will always be add
* MemberOf the name of ONE group the user should be added to or removed from
* UserPrincipalName the name which will be used for add/remove operations.
* Displayname just to make things easier to read, especially if UPNs are opaque
If a file is specified it will be treated as CSV file for export,
otherwise the objects are output

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Group
One or more group(s) to export

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

### -Path
Destination for CSV output

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

### -OrderByGroup
If specified , output will be in Group name order (default is User name.)

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
