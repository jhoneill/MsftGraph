---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphEvent

## SYNOPSIS
Modifies an event on a calendar

## SYNTAX

### None (Default)
```
Set-GraphEvent [-Event] <Object> [-Subject <String>] [-Start <DateTime>] [-End <DateTime>] [-Timezone <Object>]
 [-Location <Object>] [-Body <Object>] [-BodyType <Object>] [-ReminderOn] [-ReminderTime <Object>]
 [-ShowAs <String>] [-Importance <String>] [-Sensitivity <String>] [-Recurrence <Object>] [-Force] [-PassThru]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

### User
```
Set-GraphEvent [-Event] <Object> [-User <String>] [-Calendar <Object>] [-Subject <String>] [-Start <DateTime>]
 [-End <DateTime>] [-Timezone <Object>] [-Location <Object>] [-Body <Object>] [-BodyType <Object>]
 [-ReminderOn] [-ReminderTime <Object>] [-ShowAs <String>] [-Importance <String>] [-Sensitivity <String>]
 [-Recurrence <Object>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Group
```
Set-GraphEvent [-Event] <Object> -Group <Object> [-Subject <String>] [-Start <DateTime>] [-End <DateTime>]
 [-Timezone <Object>] [-Location <Object>] [-Body <Object>] [-BodyType <Object>] [-ReminderOn]
 [-ReminderTime <Object>] [-ShowAs <String>] [-Importance <String>] [-Sensitivity <String>]
 [-Recurrence <Object>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### AllDay
```
Set-GraphEvent [-Event] <Object> [-Subject <String>] -Start <DateTime> -End <DateTime> [-AllDay]
 [-Timezone <Object>] [-Location <Object>] [-Body <Object>] [-BodyType <Object>] [-ReminderOn]
 [-ReminderTime <Object>] [-ShowAs <String>] [-Importance <String>] [-Sensitivity <String>]
 [-Recurrence <Object>] [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
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

### -Event
The event to be updateds either as an ID or as an event object containing an ID.

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

### -User
UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"

```yaml
Type: String
Parameter Sets: User
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Calendar
A sepecific calendar belonging to a user.

```yaml
Type: Object
Parameter Sets: User
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Group
Group ID or a Group object with an ID whose calendar should be fetched

```yaml
Type: Object
Parameter Sets: Group
Aliases: Team

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Subject
Subject for the appointment

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Start
Start time - if -Timezone is not used this will be the in local machine's times zone

```yaml
Type: DateTime
Parameter Sets: None, User, Group
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: DateTime
Parameter Sets: AllDay
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -End
End Time - if -Timezone is not used this will be the in local machine's times zone

```yaml
Type: DateTime
Parameter Sets: None, User, Group
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: DateTime
Parameter Sets: AllDay
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllDay
Creates the event as all day - you must also set the start and end time.

```yaml
Type: SwitchParameter
Parameter Sets: AllDay
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timezone
Timezone - by default the local machine's time zone is used

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: $(tzutil.exe /g)
Accept pipeline input: False
Accept wildcard characters: False
```

### -Location
Location for the appointment

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Body
Body text - if using HTML set the body type to HTML

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyType
Type of text used for the body, Text or HTML

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Text
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReminderOn
Unless -Reminder on is specified no reminder will sound before the meeting

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

### -ReminderTime
Time in Minutes, before the start time, that the reminder should appear.
It will be set even if -ReminderOn is omitted

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowAs
Sets the task to appear as Free, Tenatative, Off-of-facility etc

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Importance
Priority setting , high , normal or low.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sensitivity
Privacy setting - normal or Private

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Recurrence
Recurrence pattern build with New-recurrencePattern

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
If specified the update will be performed without prompting for confirmation (this is the default)

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

### -PassThru
for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
Attendees is one.
link says this also sends the invite

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

[Get-GraphEvent]()

