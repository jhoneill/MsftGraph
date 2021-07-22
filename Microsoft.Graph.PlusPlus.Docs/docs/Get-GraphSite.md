---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphSite

## SYNOPSIS
Gets details of a sharepoint site, or its lists, drives or subsites

## SYNTAX

### None (Default)
```
Get-GraphSite [[-Site] <Object>] [<CommonParameters>]
```

### Lists
```
Get-GraphSite [[-Site] <Object>] [-Lists] [-HIDdenLists] [<CommonParameters>]
```

### HiddenLists
```
Get-GraphSite [[-Site] <Object>] [-HIDdenLists] [<CommonParameters>]
```

### SingleList
```
Get-GraphSite [[-Site] <Object>] -ListID <String> [<CommonParameters>]
```

### Notebooks
```
Get-GraphSite [[-Site] <Object>] [-Notebooks] [<CommonParameters>]
```

### Drives
```
Get-GraphSite [[-Site] <Object>] [-Drives] [<CommonParameters>]
```

### SubSites
```
Get-GraphSite [[-Site] <Object>] [-SubSites] [<CommonParameters>]
```

## DESCRIPTION
This interogates https://graph.microsoft.com/v1.0/sites/{id}
which requires consent to use the Sites.Read.All scope or better.
If no ID is provided it queries the Root site.
Depending on the parameters given it will return subsites, lists
detials of a single list, OneDrive Drives and on Note Notebooks.,
it

## EXAMPLES

### EXAMPLE 1
```
Get-GraphTeam -site | Get-GraphSite -Lists -Hidden
Gets the site(s) for the current user's team(s) and gets lists
from the site(s) including hidden ones.
```

## PARAMETERS

### -Site
Specifies a site, if omitted "root" will be assumed - the root site of the user's tennant.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Root
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Lists
If specified returns the lists in the site.

```yaml
Type: SwitchParameter
Parameter Sets: Lists
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -HIDdenLists
If specified returns the system lists which are hidden by default

```yaml
Type: SwitchParameter
Parameter Sets: Lists
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: SwitchParameter
Parameter Sets: HiddenLists
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListID
if Specified returns the details of one list

```yaml
Type: String
Parameter Sets: SingleList
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Notebooks
If Specified returns notebooks in the s

```yaml
Type: SwitchParameter
Parameter Sets: Notebooks
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Drives
If Specified returns the drives in the site.

```yaml
Type: SwitchParameter
Parameter Sets: Drives
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SubSites
If Specified returns the sub-sites within the site, if the user has suitable permissions.
 Needs higher permissions

```yaml
Type: SwitchParameter
Parameter Sets: SubSites
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
