---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphEvent

## SYNOPSIS
Get the  events in a calendar

## SYNTAX

### None (Default)
```
Get-GraphEvent [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>] [-OrderBy <String>]
 [<CommonParameters>]
```

### UserAndFilter
```
Get-GraphEvent -User <String> [-Calendar <Object>] [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>]
 [-Select <String[]>] [-OrderBy <String>] -Filter <String> [<CommonParameters>]
```

### UserAndSubject
```
Get-GraphEvent -User <String> [-Calendar <Object>] [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>]
 [-Select <String[]>] [-OrderBy <String>] -Subject <String> [<CommonParameters>]
```

### User
```
Get-GraphEvent -User <String> [-Calendar <Object>] [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>]
 [-Select <String[]>] [-OrderBy <String>] [<CommonParameters>]
```

### CalAndFilter
```
Get-GraphEvent -Calendar <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] -Filter <String> [<CommonParameters>]
```

### CalAndSubject
```
Get-GraphEvent -Calendar <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] -Subject <String> [<CommonParameters>]
```

### Cal
```
Get-GraphEvent -Calendar <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] [<CommonParameters>]
```

### GroupAndFilter
```
Get-GraphEvent -Group <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] -Filter <String> [<CommonParameters>]
```

### GroupAndSubject
```
Get-GraphEvent -Group <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] -Subject <String> [<CommonParameters>]
```

### GroupID
```
Get-GraphEvent -Group <Object> [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>]
 [-OrderBy <String>] [<CommonParameters>]
```

### CurrentFilter
```
Get-GraphEvent [-Timezone <Object>] [-Days <Int32>] [-Top <Int32>] [-Select <String[]>] [-OrderBy <String>]
 -Filter <String> [<CommonParameters>]
```

## DESCRIPTION
Depending on the parameters the events my come from
   * A specified calendar (retrieved by get-graphGroup or Get-GraphUser)
   * The default calendar for a group, (if only -group is provided)
   * The default calendar for a specific user, if only user is specified
   * The default calendar for the current user if no user, group, or calendar is specified.
   The request can specify the first n events in the calendar, or a number of days into
   the future, or specify the subject line or a custom filter.

## EXAMPLES

### EXAMPLE 1
```
>Get-GraphEvent -Team consultants
Finds the team (group) named "Consultants", and gets events in the team's calendar.
Note that the because "team" and "group" are used interchangably the parameter is
named "Group" with an alias of "Team"
```

### EXAMPLE 2
```
>get-graphuser -Calendars | where name -match "holidays" |
     get-graphevent -days 365 -order "start/datetime desc" -select start,end, subject |
        format-table subject, when
Gets the user's calendars and selects the national holidays one;
gets the events from this calendar for the next 365 days, sorting them to
soonest last and selecting only the dates and subject; 'when' is calculated from
start and end, so it is available to the format table command at the end of the pipeline.
```

### EXAMPLE 3
```
Get-GraphEvent -user alex@contoso.com -filter "isorganizer eq false"
Gets events from the specified user's calendar where they are not the organizer;
this requires access to have been granted access to the calendar by its owener.
```

### EXAMPLE 4
```
Get-GraphEvent  -filter "isorganizer eq false" -OrderBy start/datetime
This uses the same filter as the previous example but sorts the results at the
server before they are returned. Note that some fields like 'start' are record types,
and one of their properties may need to be specified to perform a sort, as in this case,
and the syntax is property/ChildProperty.
```

### EXAMPLE 5
```
>$userTimezone = (Get-GraphUser -MailboxSettings).timezone
>Get-GraphEvent -Days 150 -TimeZone $userTimezone -Filter "showas eq 'free'"
The first command gets the current user's preferred time zone, which may not
match the local computer, and the second requests items for the next 150 days,
where the time is shown as Free, displaying using that time zone
```

### EXAMPLE 6
```
Get-graphEvent -filter "start/dateTime ge '2019-04-01T08:00'"   | ft
Gets the events in the signed-in user's default calendar which start after April 1 2019
format-table will pick up the default display properties (Subject, When, Where and ShowAs)
```

## PARAMETERS

### -User
UserID as a guid or User Principal name, whose calendar should be fetched.
"me" can be used as a shortcut for current user

```yaml
Type: String
Parameter Sets: UserAndFilter, UserAndSubject, User
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Calendar
A sepecific calendar

```yaml
Type: Object
Parameter Sets: UserAndFilter, UserAndSubject, User
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

```yaml
Type: Object
Parameter Sets: CalAndFilter, CalAndSubject, Cal
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Group
Group ID or a Group object with an ID, whose calendar should be fetched

```yaml
Type: Object
Parameter Sets: GroupAndFilter, GroupAndSubject, GroupID
Aliases: Team

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
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
Position: Named
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
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
The neumber of events to fetch.
Must be greater than zero, and capped at 1000

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
Fields to select

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OrderBy
An order-by clause to sort the events

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

### -Subject
If specified, fetch events where the subject line begins with

```yaml
Type: String
Parameter Sets: UserAndSubject, CalAndSubject, GroupAndSubject
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Filter
A custom selection filter

```yaml
Type: String
Parameter Sets: UserAndFilter, CalAndFilter, GroupAndFilter, CurrentFilter
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
