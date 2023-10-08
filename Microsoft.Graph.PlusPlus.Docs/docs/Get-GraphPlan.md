---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphPlan

## SYNOPSIS
Gets information about plans used in the Planner app.

## SYNTAX

### None (Default)
```
Get-GraphPlan [[-Plan] <Object>] [<CommonParameters>]
```

### Details
```
Get-GraphPlan [[-Plan] <Object>] [-Details] [<CommonParameters>]
```

### Tasks
```
Get-GraphPlan [[-Plan] <Object>] [-Tasks] [<CommonParameters>]
```

### Buckets
```
Get-GraphPlan [[-Plan] <Object>] [-Buckets] [<CommonParameters>]
```

### FullTask
```
Get-GraphPlan [[-Plan] <Object>] [-FullTasks] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphTeam -Plans | where title -eq "team planner" | get-graphplan -FullTasks
Gets the Plan(s) for the current user's team(s), and isolates those with the name "Team Planner" ;
for each of these plans gets the tasks, expanding the name, bucket name, and assignee names
```

## PARAMETERS

### -Plan
The ID of the plan or a plan object with an ID property.
if omitted the current users planner will be assumed.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Details
If Specified returns only the details of the plan

```yaml
Type: SwitchParameter
Parameter Sets: Details
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Tasks
If specified returns a list of plan tasks.

```yaml
Type: SwitchParameter
Parameter Sets: Tasks
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Buckets
If specified gets a list of plan buckets which tasks can be assigned to

```yaml
Type: SwitchParameter
Parameter Sets: Buckets
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -FullTasks
If specified fills in the plan name, Assignee Name(s) and bucket name for each task.

```yaml
Type: SwitchParameter
Parameter Sets: FullTask
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

## NOTES

## RELATED LINKS
