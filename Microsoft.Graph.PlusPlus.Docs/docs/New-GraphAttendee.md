---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# New-GraphAttendee

## SYNOPSIS
Helper function to create a new meeting attendee, with a mail address and the type of attendance.

## SYNTAX

### Default (Default)
```
New-GraphAttendee [-Address] <String> [[-Name] <Object>] [-AttendeeType <Object>] [<CommonParameters>]
```

### PipedStrings
```
New-GraphAttendee [-AttendeeType <Object>] -InputObject <Object> [<CommonParameters>]
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

### -Address
The recipient's email address, e.g Alex@contoso.com

```yaml
Type: String
Parameter Sets: Default
Aliases: Mail

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Name
The displayname for the recipient

```yaml
Type: Object
Parameter Sets: Default
Aliases: DisplayName

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -AttendeeType
Is the attendee required or optional or a resource (such as a room).
Defaults to required

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Required
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
{{ Fill InputObject Description }}

```yaml
Type: Object
Parameter Sets: PipedStrings
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.Collections.Hashtable
## NOTES

## RELATED LINKS
