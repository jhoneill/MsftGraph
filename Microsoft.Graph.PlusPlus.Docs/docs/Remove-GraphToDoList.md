---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Remove-GraphToDoList

## SYNOPSIS
Removes a list from the To Do app, including any task in contains.

## SYNTAX

```
Remove-GraphToDoList [-ToDoList] <Object> [[-UserId] <String>] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
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
A To-do list object or the ID of a To-do list

```yaml
Type: Object
Parameter Sets: (All)
Aliases: TodoTaskListId, ListID

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -UserId
The User ID (GUID or UPN) of the list owner.
Defaults to the current user, and may be found on the ToDo list object

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: $Global:GraphUser
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Force
If specified, no confirmation will be displayed before deleting the list

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
