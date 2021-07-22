---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Remove-GraphEvent

## SYNOPSIS
Deletes an item from the calendar

## SYNTAX

### None (Default)
```
Remove-GraphEvent [-Event] <Object> [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### User
```
Remove-GraphEvent [-Event] <Object> -User <String> -Calendar <Object> [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

### GroupID
```
Remove-GraphEvent [-Event] <Object> -Group <Object> [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Deletes items from the calendar.
If other people have beeen invited to a meeting,
they will reveive a cancellation message.

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Event
The event to be removed either as an ID or as an event object containing an ID.

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

### -User
UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"

```yaml
Type: String
Parameter Sets: User
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Calendar
A sepecific calendar belonging to a user.

```yaml
Type: Object
Parameter Sets: User
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Group
Group ID or a Group object with an ID, whose calendar should be fetched

```yaml
Type: Object
Parameter Sets: GroupID
Aliases: Team

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
if Sepcified the event will be deleted without prompting for confirmation

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
