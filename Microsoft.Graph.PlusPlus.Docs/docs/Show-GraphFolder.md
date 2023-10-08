---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Show-GraphFolder

## SYNOPSIS
Opens a OneDrive folder in a browser

## SYNTAX

### FolderName (Default)
```
Show-GraphFolder [-Path] <String> [-Drive <Object>] [<CommonParameters>]
```

### FolderID
```
Show-GraphFolder -FolderID <String> [-Drive <Object>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Show-GraphFolder -Path 'root:/Documents'
Opens the documents folder from the current user's drive in the default browser
Note that root:/documents is how tab completion will render the path, but
/documents is equally valid
```

### EXAMPLE 2
```
>$drive = Get-GraphTeam -ByName Consultants -Drive
>Show-GraphFolder -Path 'root:/Documents' -drive $drive
Finds the drive for the consultants team, and opens its
documents folder in the default browser
```

## PARAMETERS

### -Path
If Specified gets the  folder by folder ID

```yaml
Type: String
Parameter Sets: FolderName
Aliases: FolderPath

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderID
If Specified gets the  folder by folder ID

```yaml
Type: String
Parameter Sets: FolderID
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Drive
The Drive containing the path .

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
