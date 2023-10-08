---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphOneNoteHome

## SYNOPSIS
Sets a default notebook (and optionally section).
Set to $Null to clear the setting

## SYNTAX

```
Set-GraphOneNoteHome [-Notebook] <Object> [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphGroup 'Consultants' -Notebooks | Get-GraphOneNoteBook -SectionName general*  | Set-GraphOneNoteHome -Verbose
The first command in the pipeline gets the notebook for the consultants group ,
the second finds the section in the notebook with an display name beginning "general"
and the third sets the default section for Add-FileToGraphOneNote, Add-GraphOneNotePage,
Get-GraphOneNotePage, and Out-GraphOneNote to the this section, and sets the
default Notebook for All the GraphOneNoteBook and all the GraphOneNoteSection commands
to the consultants group's notebook.
```

## PARAMETERS

### -Notebook
A note book or notebook section to set as the default location for oneNoteCommands.
Passing Null will clear the default.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
