---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphToDoList

## SYNOPSIS
Gets information about lists used in the To Do app.

## SYNTAX

```
Get-GraphToDoList [[-ToDoList] <Object>] [-UserId <Object>] [-Tasks] [<CommonParameters>]
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

### -ToDoList
The ID of the plan or a plan object with an ID property.
if omitted the current users planner will be assumed.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: id

Required: False
Position: 1
Default value: DefaultList
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -UserId
The User ID (GUID or UPN) of the list owner.
Defaults to the current user.

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

### -Tasks
If specified returns the tasks in the list.

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
