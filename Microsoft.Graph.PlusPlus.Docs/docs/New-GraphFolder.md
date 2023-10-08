---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/datetimecolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphFolder

## SYNOPSIS
Creates a new folder on OneDrive.

## SYNTAX

```
New-GraphFolder [-Path] <String> [-Drive <Object>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
By default this will create a new folder on the user's one drive, and if the no Parent ID is specified
the folder will be created in the root of the drive.

## EXAMPLES

### EXAMPLE 1
```
New-GraphFolder -Path '/Documents/Project-x'
Creates a new folder named "Project x" in the current users Documents folder
```

### EXAMPLE 2
```
New-GraphFolder -Path 'root:/Documents/Project-Y'
Creates a new folder named "Project Y" in the current users Documents folder
Note that tab completion will change /Projects/ to root:/Projects
```

### EXAMPLE 3
```
>$drive = Get-GraphTeam -ByName Consultants -Drive
>New-GraphFolder -Drive $drive -Path 'root:/Documents/Project Firebird/Planning'
Gets the drive for the Consultants team; and adds a subfolder under documents.
As in the previous examples root:/ is how tab completion would render the path, but
'/Documents/Project Firebird/Planning' works just as well.
```

## PARAMETERS

### -Path
The name for the new folder

```yaml
Type: String
Parameter Sets: (All)
Aliases: FolderPath

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Drive
The drive holding the new folder - defaults to the user's OneDrive but can be a shared one e.g.
Drives/{ID}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Me/Drive
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
