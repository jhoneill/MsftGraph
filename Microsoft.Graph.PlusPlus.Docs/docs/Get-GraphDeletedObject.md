---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphDeletedObject

## SYNOPSIS
Returns deleted users or groups from the AAD recycle bin

## SYNTAX

```
Get-GraphDeletedObject [[-Name] <Object>] [-Group]
```

## DESCRIPTION
It can filter by name, and selects users by default or groups if -Group is selected
The results can be piped into Restore-GraphDeletedObject

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Name
If specified filters the returned objects to those with a name starts with...

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Group
By default user objects are returned.
This switches the choice to group objects.

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

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
