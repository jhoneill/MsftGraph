---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphReminderView

## SYNOPSIS
Returns a view of items with reminders set across all a users calendars.

## SYNTAX

```
Get-GraphReminderView [[-User] <Object>] [[-Timezone] <Object>] [[-Days] <Int32>] [[-Top] <Int32>]
 [<CommonParameters>]
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

### -User
UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: $Global:GraphUser
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timezone
Time zone to rennder event times.
By default the time zone of the local machine will me use

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: $(tzutil.exe /g)
Accept pipeline input: False
Accept wildcard characters: False
```

### -Days
Number of days of calendar to fetch from today

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: 30
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
The number of events to fetch.
Must be greater than zero, and capped at 1000

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphReminder
## NOTES

## RELATED LINKS
