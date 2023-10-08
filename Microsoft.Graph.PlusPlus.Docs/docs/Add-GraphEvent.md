---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphEvent

## SYNOPSIS
Adds an event to a calendar

## SYNTAX

### User
```
Add-GraphEvent [-User <String>] [-Calendar <Object>] [-Subject <String>] -Start <DateTime> [-End <DateTime>]
 [-Timezone <Object>] [-AllDay] [-Location <Object>] [-Attendees <Object>] [-ShowAs <String>] [-ReminderOn]
 [-ReminderTime <Object>] [-Body <Object>] [-BodyType <Object>] [-Importance <String>] [-Sensitivity <String>]
 [-Recurrence <Object>] [-PassThru] [<CommonParameters>]
```

### Group
```
Add-GraphEvent -Group <Object> [-Subject <String>] -Start <DateTime> [-End <DateTime>] [-Timezone <Object>]
 [-AllDay] [-Location <Object>] [-Attendees <Object>] [-ShowAs <String>] [-ReminderOn] [-ReminderTime <Object>]
 [-Body <Object>] [-BodyType <Object>] [-Importance <String>] [-Sensitivity <String>] [-Recurrence <Object>]
 [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
>$rec = New-RecurrencePattern -Weekly Friday -EndDate 2019-04-01
>Add-GraphEvent -Start "2019-01-23 15:30:00" -subject "Enter time sheet" -Recurrence $rec
Creates a recurring event. The first sets up a weekly schedule for Fridays until April 1st.
The second sets the time  (if no end is given, it is set for 30 minutes after the start),
the subject, and the recurrence pattern
```

### EXAMPLE 2
```
>$Chris = New-Attendee -Mail Chris@Contoso.com
>Add-GraphEvent -subject "Requirements for Basingstoke project" -Start "2019-02-02 10:00" -End "2019-02-02 11:00" -Attendees $chris
Creates a meeting with a second person. The first command creates an attendee - by default the attendee is 'required'
The second creates the appointment, adding the attendee and sending a meeting request.
```

### EXAMPLE 3
```
>$Chris = New-Attendee -Mail Chris@Contoso.com -display 'Chris Cross' optional
>$Phil  = New-Attendee -Mail Phil@Contoso.com
>Add-GraphEvent -subject "Phase II planning" -Start "2019-02-02 14:00" -End "2019-02-02 14:30" -Attendees $chris,$phil
Creates a meeting with a second additonal attendee. The first command creates an optional attendee with a display name
the second creates an attendee with no displayed name and the default 'required' type
Finally the meeting is created.
```

## PARAMETERS

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
Accept pipeline input: True (ByPropertyName, ByValue)
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
Parameter Sets: (All)
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
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
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

### -AllDay
Creates the event as all day.

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

### -Attendees
People or resources involved in the event.

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

### -PassThru
for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
Attendees is one.
link says this also sends the invite

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

[Get-GraphEvent]()

