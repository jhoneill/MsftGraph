---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphDirectoryRole

## SYNOPSIS
Gets an Azure AD directory role or its members

## SYNTAX

```
Get-GraphDirectoryRole [[-Role] <Object>] [-Members] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphDirectoryRole external* -Members | ft displayname,role
Lists all members of groups whose names begin "external"
The command adds the role name to the user object making it possible
to show the roles and names in the output.
```

## PARAMETERS

### -Role
The role to get, either as a display name (wildcards allowed), an ID, or a Role object containing an ID

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: *
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Members
If specified returns the members of the role as user objects

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
