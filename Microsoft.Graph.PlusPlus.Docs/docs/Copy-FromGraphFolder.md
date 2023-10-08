---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Copy-FromGraphFolder

## SYNOPSIS
Copies files from OneDrive to the local computer

## SYNTAX

```
Copy-FromGraphFolder [-Path] <Object> [-Drive <Object>] [[-Destination] <Object>] [-NoClobber] [-Passthru]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Copy-FromGraphFolder -Path 'root:/Scripts/Type-Info.xlsx' -Destination c:\temp
Copies a single file from a "scripts" directory on the user's drive to c:\temp.
```

### EXAMPLE 2
```
>$drive = Get-GraphTeam -ByName Consultants -Drive
>Get-GraphDrive -Drive $drive -FolderPath 'root:/Documents/Project Firebird/Planning' | Copy-FromGraphFolder -Destination c:\temp
Gets all the files in a folder on a teams drive and copies them to C:\Temp.
```

## PARAMETERS

### -Path
The path to the file on one drive

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

### -Drive
The drive, by default the current user's OneDrive.

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

### -Destination
The destination on the local computer

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: $pwd
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber
If specified prevents an existing file from being overwritten.

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

### -Passthru
If Specified the destination file will be returned (similar to Copy-Item)

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: PT

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
