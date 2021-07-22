---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphDrive

## SYNOPSIS
Gets information about a OneDrive volume

## SYNTAX

### None (Default)
```
Get-GraphDrive [-Drive <Object>] [-Subfolders] [-Quiet] [<CommonParameters>]
```

### FolderName
```
Get-GraphDrive [-Drive <Object>] [-FolderPath] <String> [-Include <String>] [-Search <String>] [-Subfolders]
 [-Quiet] [<CommonParameters>]
```

### FolderID
```
Get-GraphDrive [-Drive <Object>] -FolderID <String> [-Include <String>] [-Search <String>] [-Subfolders]
 [-Quiet] [<CommonParameters>]
```

### Special
```
Get-GraphDrive [-Drive <Object>] -SpecialFolder <String> [-Include <String>] [-Quiet] [<CommonParameters>]
```

### Recent
```
Get-GraphDrive [-Drive <Object>] [-Recent] [-Include <String>] [-Quiet] [<CommonParameters>]
```

### Shared
```
Get-GraphDrive [-Drive <Object>] [-SharedWithMe] [-Include <String>] [-Search <String>] [-Quiet]
 [<CommonParameters>]
```

### RootSearch
```
Get-GraphDrive [-Drive <Object>] [-Search <String>] [-Quiet] [<CommonParameters>]
```

### RootFolders
```
Get-GraphDrive [-Drive <Object>] [-Subfolders] [-Quiet] [<CommonParameters>]
```

### ItemName
```
Get-GraphDrive [-Drive <Object>] -ItemPath <String> [-Quiet] [<CommonParameters>]
```

### ItemID
```
Get-GraphDrive [-Drive <Object>] -ItemID <String> [-Quiet] [<CommonParameters>]
```

## DESCRIPTION
Run with no parameters this gets information about the volume for the current user.
It can get information about another volume by specifying -Drive
It can get information about the root folders, or the contents of a specific folder,
or a link to a special folder  or recent items.

## EXAMPLES

### EXAMPLE 1
```
Get-GraphDrive
Returns the drive for the current user.
```

### EXAMPLE 2
```
get-graphdrive -Subfolders
Returns the root folders in the the current user's drive. Formatting is defined
to display the items like a normal directory, but other properties are also available.
```

### EXAMPLE 3
```
>$d = get-graphteam -Drive | select -first 1
>get-graphdrive -Drive $d -SpecialFolder Documents
```

The first line gets the first team drive for a user, the second gets
the items in its Documents folder

### EXAMPLE 4
```
get-graphdrive -Drive $d -FolderPath general
This example uses the team drive found in the previous one and gets the contents of the team's "General" folder
```

### EXAMPLE 5
```
get-graphdrive -Drive $d -itemPath general
Instead of getting the the items in the General folder, this returns an object representing the folder itself
```

### EXAMPLE 6
```
Get-GraphDrive -Search preferredLanguage -FolderPath 'root:/Scripts'
This does a freetext search of "preferredLanguage" in the scripts folder; because no drive is
specified this folder is on the current user's drive.
Note that searches do not return the parent path if you need to find the folder path you can do
get-graphitem [-drive {drive}] -itemid with either the item's own ID or its parent's ID.
```

### EXAMPLE 7
```
$folder = (get-graphuser -Drive).root.children | where name -eq scripts
>get-graphdrive -Drive $folder.parentReference.driveId -FolderID $folder.id
The first command gets the users drive, and looks for a known folder as a child item in the drive-root.
This folder can't be piped into get-graphdrive, so the drive id and folder id are passed.
In this case the drive ID could be ommitted because the default is to use the user's home drive
```

## PARAMETERS

### -Drive
The drive to examine - defaults to the user's OneDrive but can be a shared one e.g.
Drives/{ID}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Me/Drive
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -FolderPath
If specified gets the items in a folder by the path from {drive}/root:

```yaml
Type: String
Parameter Sets: FolderName
Aliases: Path

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderID
If Specified gets the items in a folder by folder ID

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

### -SpecialFolder
If specified gets one of the special folders (Documents, photos etc) in the drive.
If they don't already exist the server appears to create them.

```yaml
Type: String
Parameter Sets: Special
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recent
If specified gets recent items in the drive

```yaml
Type: SwitchParameter
Parameter Sets: Recent
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SharedWithMe
If Specified gets items shared with the user

```yaml
Type: SwitchParameter
Parameter Sets: Shared
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Include
{{ Fill Include Description }}

```yaml
Type: String
Parameter Sets: FolderName, FolderID, Special, Recent, Shared
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Search
Enables a free text search of the selected content

```yaml
Type: String
Parameter Sets: FolderName, FolderID, Shared, RootSearch
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subfolders
If specified returns the subfolders - if no FolderPath or FolderID is given will return folders of the root drive

```yaml
Type: SwitchParameter
Parameter Sets: None, FolderName, FolderID, RootFolders
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ItemPath
if specified gets a file or folder by the path from {drive}/root:

```yaml
Type: String
Parameter Sets: ItemName
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ItemID
If Specified gets the a file or folder item by ID

```yaml
Type: String
Parameter Sets: ItemID
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Quiet
If specified does not display a message when a folder is empty

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
