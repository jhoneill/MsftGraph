function Get-GraphReminderView   {
    <#
      .Synopsis
        Returns a view of items with reminder sets across all a users calendars.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
         [string]$User,

        #Time zone to rennder event times. By default the time zone of the local machine will me use
        $Timezone = $(tzutil.exe /g),

        #Number of days of calendar to fetch from today
        [int]$Days =30 ,

        #The neumber of events to fetch. Must be greater than zero, and capped at 1000
        [ValidateRange(1,1000)]
        [int]$Top
    )
    Connect-MSGraph
    $webParams = @{Method = "Get"
                  Headers = $Script:DefaultHeader
    }
    if ($TimeZone) {$webParams.Headers["Prefer"]="Outlook.timezone=""$TimeZone"""}

    If ($User)   {  # get the default calendar for a specific user
            if ($User.ID) {$User=$User.ID}
            $uri = "https://graph.microsoft.com/v1.0/users/$user/reminderView(startDateTime='{0:yyyy-MM-ddTHH:mm:ss}',endDateTime='{1:yyyy-MM-ddTHH:mm:ss}')"
    }
    else {  $uri = "https://graph.microsoft.com/v1.0/me/reminderView(startDateTime='{0:yyyy-MM-ddTHH:mm:ss}',endDateTime='{1:yyyy-MM-ddTHH:mm:ss}')"  }

    $webParams['uri'] =  $uri -f [datetime]::Today, [datetime]::Today.AddDays(30)
    $result = Invoke-RestMethod @webParams
    $defaultProperties = @('Subject','When','Reminder')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
    $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $whensb = { if (  ([datetime]$this.eventStartTime.datetime).AddDays(1) -eq ([datetime]$this.eventEndTime.datetime  )) {
                      ([datetime]$this.eventStartTime.datetime).ToShortDateString()
                }
                else {([datetime]$this.eventStartTime.datetime).ToString("g") + ' to ' +  ([datetime]$this.eventEndTime.datetime  ).ToString("g") + $this.eventEndTime.timezone}
    }
    foreach ($r in $result.value) {
        Add-Member -InputObject $r -MemberType AliasProperty  -Name Subject           -Value eventSubject
        Add-Member -InputObject $r -MemberType AliasProperty  -Name Start             -Value eventStartTime
        Add-Member -InputObject $r -MemberType AliasProperty  -Name End               -Value eventEndTime
        Add-Member -InputObject $r -MemberType ScriptProperty -Name When              -Value $whenSB
        Add-Member -InputObject $r -MemberType ScriptProperty -Name Reminder          -Value {([datetime]$this.reminderFireTime.datetime).ToString("g")}
        Add-Member -InputObject $r -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers
        Add-Member -InputObject $r -MemberType AliasProperty  -Name Location          -Value eventLocation -PassThru
    }
}

function Get-GraphEvent          {
    <#
      .Synopsis
        Get the  events in a calendar
      .Description
        Depending on the parameters specified the calendar can be
           * A Specific calendar for a group, if the group and calendar are specified
             (group ID can be a calendar property)
           * The default calendar for a group, if only group is provided)
           * A specific calendar for a specific user, if user and calendar are specified
           * The default calendar for a specific user, if only user is specified
           * A specific calendar for the currrent user, if only calandar is specified
           * The default calendar for the current user if no user, group, or calendar is specified.
           The request can specify the first n events in the calendar, or a number of days into
           the future, or specify the subject line.
      .Example
        >
        >get-graphuser -Calendars | where name -match "holidays" |
             get-graphevent -days 365 -order "start/datetime desc" -select start,subject |
                ft subject, @{n="when";e={([datetime]$_.start.datetime).tostring("d")}}
        Gets the user's calendars and selects the national holidays one;
        gets the events from this calendar for the next 365 days sorts them to
        soonest last and selects only the date and subject; displays these in a table
        showing start in the local short-date format.
      .Example
        >Get-GraphEvent -user james@contoso.com -filter "isorganizer eq false"
        Gets events from the specified users calendar where they are not the organizer
      .Example
        >Get-GraphEvent  -filter "isorganizer eq false" -OrderBy start/datetime
        This uses the same filter but sorts the results at the server before they are
        returned. Note that some fields like 'start' are record types, and one of
        their propperties, as in this case, may need to be specified to perform a sort.
      .Example
        >
        >$userTimezone = (Get-GraphUser -MailboxSettings).timezone
        >Get-GraphEvent -Days 150 -TimeZone $userTimezone -Filter "showas eq 'free'"
        The first command gets the current user's time zone, and the second requests
        items for the next 150 days where the time is shown as Free, displaying using that time zone
      .Example
        >Get-graphEvent -filter "start/dateTime ge '2019-04-01T08:00'"   | ft
        Gets the events in the signed-in user's default calendar which start after April 1 2019
        format-table will pick up the default display properties. .
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( Mandatory=$true, ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="UserAndSubject",ValueFromPipelineByPropertyName=$true)]
        [Parameter( Mandatory=$true, ParameterSetName="UserAndFilter",ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [Parameter( ParameterSetName="UserAndSubject",ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Parameter( ParameterSetName="UserAndFilter",ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID, whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="GroupID",ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndSubject",ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndFilter",ValueFromPipelineByPropertyName=$true)]
        [Alias("Team")]
        $Group,

        #Time zone to rennder event times. By default the time zone of the local machine will me use
        $Timezone = $(tzutil.exe /g),

        #Number of days of calendar to fetch from today
        [int]$Days,

        #The neumber of events to fetch. Must be greater than zero, and capped at 1000
        [ValidateRange(1,1000)]
        [int]$Top,

        #Fields to select
        [ValidateSet('attendees', 'body', 'bodyPreview', 'categories', 'changeKey', 'createdDateTime', 'end', 'hasAttachments',
                     'iCalUId', 'id', 'importance', 'isAllDay', 'isCancelled', 'isOrganizer', 'isReminderOn', 'lastModifiedDateTime',
                     'location', 'locations', 'onlineMeetingUrl', 'organizer', 'originalEndTimeZone', 'originalStart',
                     'originalStartTimeZone', 'recurrence', 'reminderMinutesBeforeStart', 'responseRequested', 'responseStatus',
                    'sensitivity', 'seriesMasterId', 'showAs', 'start', 'subject', 'type', 'webLink' )]
        [string[]]$Select,

        #An order-by clause to sort the events
        [string]$OrderBy,

        #If specified, fetch events where the subject line begins with
        [Parameter(Mandatory=$true, ParameterSetName='Subject',ValueFromPipelineByPropertyName=$true)]
        [Parameter(Mandatory=$true, ParameterSetName="UserAndSubject",ValueFromPipelineByPropertyName=$true )]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndSubject",ValueFromPipelineByPropertyName=$true)]
        [string]$Subject,

        #A custom selection filter
        [Parameter(Mandatory=$true, ParameterSetName="UserAndFilter",ValueFromPipelineByPropertyName=$true )]
        [Parameter(Mandatory=$true, ParameterSetName="GroupAndFilter",ValueFromPipelineByPropertyName=$true)]
        [string]$Filter
    )
    Connect-MSGraph
    $webParams = @{Method = "Get"
                  Headers = $Script:DefaultHeader
    }
    if ($TimeZone) {$webParams.Headers["Prefer"]="Outlook.timezone=""$TimeZone"""}

    #region figure out which calendar to get. The API doesn't have v1.0/calendars/id, have to do group/calendar user/calendar or user/calendarS/id
    if     ($user -and $Calendar) { #get a specific calendar for a specific user
        if ($User.ID)     {$User     = $User.ID}
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $uri = "https://graph.microsoft.com/v1.0/users/$user/calendars/$Calendar"
    }
    elseif ($User)   {  # get the default calendar for a specific user
        if ($User.ID) {$User=$User.ID}
        $uri = "https://graph.microsoft.com/v1.0/users/$user/calendar"  #for the default calendar you can also use users/{id}/events or users//calendarView?param without "Calendar"
    }
    elseif ($Group) {  # get the [only] calendar for a group
        if ($Group.ID) {$Group=$Group.ID}
        $uri = "https://graph.microsoft.com/v1.0/groups/$Group/calendar"   #for the default calendar you can also use groups/{id}/events or groups/calendarView?param without "Calendar"
    }
    elseif ($Calendar -and $Calendar.GroupID ) { #handle piping in a group's calendar object - we have added the group ID to it, us that get the group's [only] calendar
        $uri = "https://graph.microsoft.com/v1.0/groups/$($calendar.groupID)/calendar"
    }
    elseif ($calendar) { #get a specific calendar for the current user - more normal use of the calendar parameter
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $uri = "https://graph.microsoft.com/v1.0/me/calendars/$calendar"
    }
    else  {  #no User, group or calendar specified get the current users default calendar.
        $uri = "https://graph.microsoft.com/v1.0/me/calendar"   #for the default calendar you can also use me/events or me/calendarView?params without "Calendar"
    }
    #endregion
    #region apply the selection criteria. If -days is specified use calendar view, otherwise use events and add filter, orderby, select and top as needed
    if  ($days)    {
                       $start = [datetime]::Today.ToString("yyyy-MM-dd't'HH:mm:ss")       # 'o' for ISO format time may work here.
                       $end   = [datetime]::Today.AddDays($days).tostring("yyyy-MM-dd't'HH:mm:ss")
                       $uri  += "/calendarview?`$expand=calendar&startdatetime=$start&enddatetime=$end"
    }
    else {             $uri  +=  '/events?$expand=calendar'}
    if ($Select)     { $uri  +=  '&$select=' + ($Select -join ',') }
    if ($Subject)    { $uri  += ('&$filter=startswith(subject,''{0}'')' -f $subject ) }
    elseif ($Filter) { $uri  +=  '&$Filter='  + $Filter }
    if ($OrderBy)    { $uri  +=  '&$orderby=' + $orderby }
    if ($Top)        { $uri  +=  '&$top='     + $top  }
    #endregion

    #region get the data. Cope with data being paged, add a type to help formatting and return the results.
    $eventlist      = @()
    $result         = Invoke-RestMethod @webParams -Uri $uri
    $eventlist     += $result.value
    while ($result.'@odata.nextLink') {
        $result     = Invoke-RestMethod @webParams -Uri  $result.'@odata.nextLink' ;
        $eventlist += $result.value
    }

    if ($eventlist) {
        $defaultProperties = @('Subject','When','Where','ShowAs')
        $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$defaultProperties)
        $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
        $whensb = { if (  ([datetime]$this.eventStartTime.datetime).AddDays(1) -eq ([datetime]$this.eventEndTime.datetime  )) {
                          ([datetime]$this.eventStartTime.datetime).ToShortDateString()
                    }
                    else {([datetime]$this.eventStartTime.datetime).ToString("g") + ' to ' +  ([datetime]$this.eventEndTime.datetime  ).ToString("g") + $this.eventEndTime.timezone}
        }
        foreach ($e in $eventlist) {
            $e.pstypeNames.add('GraphEvent')
            Add-Member -InputObject $e -MemberType ScriptProperty -Name When              -Value $whenSB
            Add-Member -InputObject $e -MemberType ScriptProperty -Name Where             -Value {$this.location.displayname}
            Add-Member -InputObject $e -MemberType MemberSet      -Name PSStandardMembers -Value $PSStandardMembers
        }

        return $result.value
    }
    else {Write-Warning -Message "No events were found."}
    #endregion
}

function New-RecurrencePattern   {
    <#
      .synopsis
        Creates a new recurrence pattern for an appointment
       .Description
        There are 6 patterns:
        *Daily,  Weekly, Relative & absolute Monthly and
        Relative a& absolute yearly. Each one takes an interval,
        which is 1 by default, meaning the event is scheduled for each
        matching day. 2 would be alternate matchs, 3 would be every
        third one etc. Weekly, and the two relative choice specify
        one or more week days, and realtive ones say which week of the
        month. The two absolute options specify which date in the month;
        and the two yearly options specify which month of the year.
        The pattern needs to know if it should end, so it can be given
        an end date, or a total number of occurences. If neither is
        specified the event runs until removed from the calendar.
      .Example
        >$rec = New-RecurrencePattern -Weekly Friday -EndDate 2019-04-01
        Creates a weekly meeting on every Friday until April 1st 2019
      .Example
        >$rec = New-RecurrencePattern -AbsoluteMonthly -DayOfMonth 1 -Occurrences 12
        Specifies a pattern of a meeting on the first of the month for the next
        12 months.
    #>
    [cmdletbinding()]
    param (

        # Event repeats on the every day or every [-interval] days
        [Parameter(Mandatory=$true,  ParameterSetName='daily')]
        [switch]$Daily,

        # Event repeats on the same day(s) of the week
        [Parameter(Mandatory=$true,  ParameterSetName='weekly')]
        [switch]$Weekly,

        # Event repeats on the same day of the week in the same relative position each month
        [Parameter(Mandatory=$true,  ParameterSetName='relativeMonthly')]
        [switch]$RelativeMonthly,

        # Event repeats on the same date in the month, each month
        [Parameter(Mandatory=$true,  ParameterSetName='absoluteMonthly')]
        [switch]$AbsoluteMonthly,

        #Event happens yearly on the same day week, and same relative position in the month.
        [Parameter(Mandatory=$true,  ParameterSetName='relativeYearly')]
        [switch]$RelativeYearly,

        #Event happens yearly on the same day of the month.
        [Parameter(Mandatory=$true,  ParameterSetName='AbsoluteYearly')]
        [switch]$AbsoluteYearly,

        # Which instance of the the selected day of the week will the event occur on
        [Parameter(Position=1, ParameterSetName='relativeYearly')]
        [Parameter(Position=1, ParameterSetName='relativeMonthly')]
        [ValidateSet('first','second','third','fourth','last')]
        [string]$WeekOfMonth = 'first',

        #On which day of the week will the event occur
        [Parameter(Mandatory=$true, Position=2, ParameterSetName='relativeYearly')]
        [Parameter(Mandatory=$true, Position=2, ParameterSetName='relativeMonthly')]
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='weekly')]
        [ValidateSet('Monday','Tuesday','Wednesday','Thursday','Friday', 'Saturday', 'Sunday')]
        [string[]]$Days,

        #On which date in the month will the event Occur
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='absoluteMonthly')]
        [Parameter(Mandatory=$true, Position=1, ParameterSetName='AbsoluteYearly')]
        [ValidateRange(1,31)]
        $DayOfMonth,

        #In which month does the event occur - as a number, so 9 for September.
        [Parameter(Mandatory=$true, Position=2, ParameterSetName='relativeYearly')]
        [Parameter(Mandatory=$true, Position=2, ParameterSetName='AbsoluteYearly')]
        [ValidateRange(1,12)]
        $Month,

        #How many days, weeks , months, years apart will events occur
        [Parameter(Position=3)]
        [Int]$Interval = 1,

        #Date to stop applying the pattern, the last event may not fall on this date.
        [dateTime]$EndDate,

        #The number of occurences after which the event should cease
        [Int]$Occurrences = 10

    )
    if     ($EndDate)         {$range   = @{type = 'endDate'  ; endDate= $EndDate.ToString('yyyy-MM-dd') }  }
    Elseif ($Occurrences)     {$range   = @{type = 'numbered' ; numberOfOccurrences= $Occurrences}  }
    Else                      {$range   = @{type = 'noEnd'    ; }}

    if     ($Daily)           {$pattern = @{type = 'daily'           ;interval = $Interval} }
    elseif ($Weekly)          {$pattern = @{type = 'weekly'          ;interval = $Interval; daysOfWeek= @()+$Days} }
    elseif ($RelativeMonthly) {$pattern = @{type = 'relativeMonthly' ;interval = $Interval; daysOfWeek= @()+$Days; index=$WeekOfMonth  } }
    elseif ($RelativeYearly)  {$pattern = @{type = 'relativeYearly'  ;interval = $Interval; daysOfWeek= @()+$Days; index=$WeekOfMonth ; month=$Month} }
    elseif ($AbsoluteYearly)  {$pattern = @{type = 'absoluteYearly'  ;interval = $Interval; dayOfMonth= $DayOfMonth               ; month=$Month} }
    elseif ($AbsoluteMonthly) {$pattern = @{type = 'absoluteYearly'  ;interval = $Interval; dayOfMonth= $DayOfMonth} }

    return @{'range'=$range; 'pattern'=$pattern}
}

function New-EventAttendee       {
    <#
      .Synopsis
        Creats a new meeting attendee, with a mail address and the type of attendance.
    #>
    [cmdletbinding(DefaultParameterSetName='Default')]
    param(
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true,ParameterSetName='Default',Mandatory=$true)]
        $Mail,
        #The displayname for the recipient
        [Parameter(Position=1, ValueFromPipelineByPropertyName=$true,ParameterSetName='Default')]
        $DisplayName,
        #Is the attendee required or optional or a resource (such as a room). Defaults to required
        [ValidateSet('required', 'optional', 'resource')]
        $AttendeeType = 'required',
        [Parameter(ValueFromPipeline=$true,ParameterSetName='PipedStrings',Mandatory=$true)]
        $InputObject 
    )
    @{ 'type'= $AttendeeType ; 'emailAddress' = (New-MailAddress -Mail:$mail -DisplayName:$DisplayName )}
}

function Add-GraphEvent          {
    <#
      .Synopsis
        Adds an event to a calendar
      .link
        Get-GraphEvent
      .Example
        >
        >$rec = New-RecurrencePattern -Weekly Friday -EndDate 2019-04-01
        >Add-GraphEvent -Start "2019-01-23 15:30:00" -subject "Enter time sheet" -Recurrence $rec
        Creates a recurring event. The first sets up a weekly schedule for Fridays until April 1st.
        The second sets the time  (if no end is given, it is set for 30 minutes after the start),
        the subject, and the recurrence pattern
      .Example
        >
        >$Chris = New-Attendee -Mail Chris@Contoso.com
        >Add-GraphEvent -subject "Requirements for Basingstoke project" -Start "2019-02-02 10:00" -End "2019-02-02 11:00" -Attendees $chris
        Creates a meeting with a second person. The first command creates an attendee - by default the attendee is 'required'
        The second creates the appointment, adding the attendee and sending a meeting request.
      .Example
        >
        >$Chris = New-Attendee -Mail Chris@Contoso.com -display 'Chris Cross' optional
        >$Phil  = New-Attendee -Mail Phil@Contoso.com  
        >Add-GraphEvent -subject "Phase II planning" -Start "2019-02-02 14:00" -End "2019-02-02 14:30" -Attendees $chris,$phil
        Creates a meeting with a two additonal attendee. The first command creates an optional attendee with a display name
        the second creates an attendee with no displayed name and the default 'required' type
        Finally the meeting is created.
    #>
    [cmdletbinding()]
    param (
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
        $Group,

        #Subject for the appointment
        [string]$Subject ,

        #Start time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter(Mandatory=$true)]
        [datetime]$Start,

        #End Time - if -Timezone is not used this will be the in local machine's times zone
        [datetime]$End,

        #Timezone - by default the local machine's time zone is used
        $Timezone = $(tzutil.exe /g),

        #Creates the event as all day.
        [Switch]$AllDay,

        #Location for the appointment
        $Location,

        #People or resources involved in the event.
        $Attendees,
        #Sets the task to appear as Free, Tenatative, Off-of-facility etc
        [ValidateSet('busy','free','oof','tentative','workingElsewhere')]
        [string]$ShowAs,

        #Unless -Reminder on is specified no reminder will sound before the meeting
        [switch]$ReminderOn,
        #Time in Minutes, before the start time, that the reminder should appear. It will be set even if -ReminderOn is omitted
        $ReminderTime,
        #Body text - if using HTML set the body type to HTML
        $Body  ,
        #Type of text used for the body, Text or HTML
        [ValidateSet('Text','HTML')]
        $BodyType = 'Text',
        #Priority setting , high , normal or low.
        [ValidateSet('low','normal','high')]
        [String]$Importance ,
        #Privacy setting - normal or Private
        [ValidateSet('normal','private')]
        [String]$Sensitivity,

        #Recurrence pattern build with New-recurrencePattern
        $Recurrence
        # for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
        # and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
        # Attendees is one. link says this also sends the invite

    )
    Connect-MSGraph
    $webParams = @{Method     = 'Post'
                  Contenttype = 'application/json'
                  Headers     = @{Authorization = $Script:AuthHeader ;
                                  Prefer        = "Outlook.timezone=""$TimeZone"""}
    }

    #region figure out which calendar to get. The API doesn't have v1.0/calendars/id, have to do group/calendar user/calendar or user/calendarS/id
    if     ($user -and $Calendar) { #get a specific calendar for a specific user
        if ($User.ID)     {$User=$User.ID}
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/users/$user/calendars/$Calendar/events"
    }
    elseif ($User)     {  # get the default calendar for a specific user
        if ($User.ID) {$User=$User.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/users/$user/calendar/events"  #for the default calendar you can also use users/{id}/events "Calendar"
    }
    elseif ($Group)    {  # get the [only] calendar for a group
        if ($Group.ID) {$Group=$Group.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/groups/$Group/calendar/events"   #for the default calendar you can also use groups/{id}/events
    }
    elseif ($Calendar -and $Calendar.GroupID ) { #handle piping in a group's calendar object - we have added the group ID to it, us that get the group's [only] calendar
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/groups/$($calendar.groupID)/calendar/events"
    }
    elseif ($calendar) { #get a specific calendar for the current user - more normal use of the calendar parameter
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/me/calendars/$calendar/events"
    }
    else  {  #no User, group or calendar specified get the current users default calendar.
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/me/calendar/events"   #for the default calendar you can also use me/events
    }
    #endregion

    #region assemble the body needed to create the event
    $settings = @{      'subject'=  $Subject;     'isReminderOn' = [bool]$ReminderOn}
    if ($Location)     {$settings['location']                    = @{'displayName'=$Location} }
    if ($Body)         {$settings['body']                        = @{'contentType'=$BodyType ; 'Content'=$Body}}
    if ($ShowAs)       {$settings['showAs']                      = $ShowAs}
    if ($ReminderTime) {$settings['reminderMinutesBeforeStart']  = $ReminderTime}
    if ($Importance)   {$settings['importance']                  = $Importance}
    if ($Sensitivity)  {$settings['sensitivity']                 = $Sensitivity}
    if ($AllDay)       {$settings['isAllDay']                    = $true
                        $Start = $Start.Date
                        if (-not $End)        { $End             = $Start.AddDays(1)}
                        else                  { $End             = $End.Date        }
                        if ($End -eq $Start)  { $End             = $End.AddDays(1)  }
    }
    elseif (-not $End) {$End = $Start.AddMinutes(30)    }
    $settings['start' ] = @{'timeZone'=$Timezone; 'dateTime'    = $Start.ToString("yyyy-MM-dd'T'HH:mm:ss")} ;
    $settings['end'   ] = @{'timeZone'=$Timezone; 'dateTime'    = $End.ToString(  "yyyy-MM-dd'T'HH:mm:ss")};

    if ($Recurrence)   {$settings['recurrence']                 = $Recurrence
                        $settings.recurrence.range['startDate'] = $Start.ToString('yyyy-MM-dd');
    }
    if ($Attendees)    {$settings['attendees'] = @() + $Attendees }
    $json =  (ConvertTo-Json $settings -Depth 10)
    Write-Debug $json
    #endregion

    $result = Invoke-RestMethod @webParams -Body $json
    if ($PassThru) {
        #send back the new appoinment and give it a type so it will get formatted.
        $result.pstypeNames.add('GraphEvent')
        
        $result
    }
}

function Set-GraphEvent          {
    <#
      .Synopsis
        Modifies an event on a calendar
      .link
        Get-GraphEvent
      .Example
        a
    #>
    [cmdletbinding(SupportsShouldProcess=$true,DefaultParameterSetName='None')]
    param (
        #The event to be updateds either as an ID or as an event object containing an ID.
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)]
        $Event,

        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",ValueFromPipelineByPropertyName=$true)]
        $Calendar,

        #Group ID or a Group object with an ID whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="Group", ValueFromPipelineByPropertyName=$true)]
        [Alias('Team')]
        $Group,

        #Subject for the appointment
        [string]$Subject ,

        #Start time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter( ParameterSetName="Group" )]
        [Parameter( ParameterSetName="None"  )]
        [Parameter( ParameterSetName="User"   )]
        [Parameter( ParameterSetName="AllDay", Mandatory=$true )]
        [Nullable[datetime]]$Start,

        #End Time - if -Timezone is not used this will be the in local machine's times zone
        [Parameter( ParameterSetName="Group"  )]
        [Parameter( ParameterSetName="None"   )]
        [Parameter( ParameterSetName="User"   )]
        [Parameter( ParameterSetName="AllDa#y", Mandatory=$true )]
        [Nullable[datetime]]$End,

        #Creates the event as all day - you must also set the start and end time.
        [Parameter(Mandatory=$true, ParameterSetName="AllDay")]
        [Switch]$AllDay,

        #Timezone - by default the local machine's time zone is used
        $Timezone = $(tzutil.exe /g),
        #Location for the appointment
        $Location,
        #Body text - if using HTML set the body type to HTML
        $Body  ,
        #Type of text used for the body, Text or HTML
        [ValidateSet('Text','HTML')]
        $BodyType = 'Text',
        #Unless -Reminder on is specified no reminder will sound before the meeting
        [switch]$ReminderOn,
        #Time in Minutes, before the start time, that the reminder should appear. It will be set even if -ReminderOn is omitted
        $ReminderTime,
        #Sets the task to appear as Free, Tenatative, Off-of-facility etc
        [ValidateSet('busy','free','oof','tentative','workingElsewhere')]
        [string]$ShowAs,
        #Priority setting , high , normal or low.
        [ValidateSet('low','normal','high')]
        [String]$Importance ,
        #Privacy setting - normal or Private
        [ValidateSet('normal','private')]
        [String]$Sensitivity,

        #Recurrence pattern build with New-recurrencePattern
        $Recurrence,
        #If specified the update will be performed without prompting for confirmation (this is the default)
        [switch]$Force
        # for some of things still to do see https://docs.microsoft.com/en-us/graph/api/event-update?view=graph-rest-beta
        # and https://docs.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-beta
        # Attendees is one. link says this also sends the invite

    )
    Connect-MSGraph
    $webParams = @{Method     = 'Patch'
                  Contenttype = 'application/json'
                  Headers     = @{Authorization = $Script:AuthHeader ;
                                  Prefer        = "Outlook.timezone=""$TimeZone"""}
    }

    #region figure out which calendar to get. The API doesn't have v1.0/calendars/id, have to do group/calendar user/calendar or user/calendarS/id
    if     ($user -and $Calendar) { #get a specific calendar for a specific user
        if ($User.ID)     {$User=$User.ID}
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/users/$user/calendars/$Calendar/events"
    }
    elseif ($User)   {  # get the default calendar for a specific user
        if ($User.ID) {$User=$User.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/users/$user/calendar/events"  #for the default calendar you can also use users/{id}/events "Calendar"
    }
    elseif ($Group) {  # get the [only] calendar for a group
        if ($Group.ID) {$Group=$Group.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/groups/$Group/calendar/events"   #for the default calendar you can also use groups/{id}/events
    }
    elseif ($Calendar -and $Calendar.GroupID ) { #handle piping in a group's calendar object - we have added the group ID to it, us that get the group's [only] calendar
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/groups/$($calendar.groupID)/calendar/events"
    }
    elseif ($calendar) { #get a specific calendar for the current user - more normal use of the calendar parameter
        If ($Calendar.id) {$Calendar = $Calendar.ID}
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/me/calendars/$calendar/events"
    }
    else  {  #no User, group or calendar specified get the current users default calendar.
        $webParams['uri'] = "https://graph.microsoft.com/v1.0/me/calendar/events"   #for the default calendar you can also use me/events
    }
    #endregion
    if ($Event.id) {$webParams['uri'] += "/$($Event.id)"}
    else           {$webParams['uri'] += "/$Event"}
    $settings  =   @{ }
    #region assemble the body needed to update the event
    if ($PSBoundParameters.ContainsKey('ReminderOn')) {
                           $settings['isReminderOn']    = [bool]$ReminderOn}
    if ($PSBoundParameters.ContainsKey('AllDay'))     {
                           $settings['isAllDay']        = [bool]$AllDay  }
    if ($Subject)         {$settings['subject']         = $subject };
    if ($Location)        {$settings['location']        = @{'displayName'=$Location} }
    if ($Body)            {$settings['body']            = @{'contentType'=$BodyType ; 'Content'=$Body}}
    if ($ShowAs)          {$settings['showAs']          = $ShowAs}
    if ($Importance)      {$settings['importance']      = $Importance}
    if ($Sensitivity)     {$settings['sensitivity']     = $Sensitivity}
    if ($ReminderTime)    {$settings['reminderMinutesBeforeStart'] = $ReminderTime}
    if ($Start)           {$settings['start'] = @{'timeZone'       = $Timezone; 'dateTime' = $Start.ToString("yyyy-MM-dd'T'HH:mm:ss")} }
    if ($End)             {$settings['end'  ] = @{'timeZone'       = $Timezone; 'dateTime' = $End.ToString(  "yyyy-MM-dd'T'HH:mm:ss")} }
    if ($Recurrence)      {$settings['recurrence']                 = $Recurrence
                           $settings.recurrence.range['startDate'] = $Start.ToString('yyyy-MM-dd');
    }
    $json =  (ConvertTo-Json $settings -Depth 10)
    Write-Debug $json
    #endregion

    if ($Force -or $PSCmdlet.ShouldProcess($Event.subject,'Update calendar event')) {
        $result = Invoke-RestMethod @webParams -Body $json
        $result.pstypeNames.add('GraphEvent')
        return $result
    }
}

function Remove-GraphEvent       {
    <#
      .Synopsis
        Deletes an item from the calendar
      .Description
        Deletes items from the calendar. If other people have beeen invited to a meeting,
        they will reveive a cancellation message.
    #>
    [cmdletbinding(DefaultParameterSetName="None",SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (

        #The event to be removed either as an ID or as an event object containing an ID.
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        $Event,

        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [Parameter( ParameterSetName="User",Mandatory=$true)]
        [string]$User,

        #A sepecific calendar belonging to a user.
        [Parameter( ParameterSetName="User",Mandatory=$true)]
        $Calendar,

        #Group ID or a Group object with an ID, whose calendar should be fetched
        [Parameter(Mandatory=$true, ParameterSetName="GroupID")]
        [Alias("Team")]
        $Group,

        #if Sepcified the event will be deleted without prompting for confirmation
        [switch]$Force
    )
    begin   {
        Connect-MSGraph
    }
    process {
        if     ($event.calendar.id -and -not $Calendar) {$Calendar = $event.calendar.id}

        if     ($user -and $Calendar) { #get a specific calendar for a specific user
            if ($User.ID)     {$User     = $User.ID}
            If ($Calendar.id) {$Calendar = $Calendar.ID}
            $uri = "https://graph.microsoft.com/v1.0/users/$user/calendars/$Calendar/events"
        }
        elseif ($User)   {  # get the default calendar for a specific user
            if ($User.ID) {$User=$User.ID}
            $uri = "https://graph.microsoft.com/v1.0/users/$user/calendar/events"
        }
        elseif ($Group) {  # get the [only] calendar for a group
            if ($Group.ID) {$Group=$Group.ID}
            $uri = "https://graph.microsoft.com/v1.0/groups/$Group/calendar/events"
        }
        elseif ($calendar) { #get a specific calendar for the current user - more normal use of the calendar parameter
            If ($Calendar.id) {$Calendar = $Calendar.ID}
            $uri = "https://graph.microsoft.com/v1.0/me/calendars/$calendar/events"
        }
        else  {  #no User, group or calendar specified get the current users default calendar.
            $uri = "https://graph.microsoft.com/v1.0/me/calendar/events"
        }
        if ($Force -or $PSCmdlet.ShouldProcess($Event.Subject ,'Delete from calendar')) {
            if ($event.ID) {$event = $event.id}
            Invoke-RestMethod -Method Delete -Uri "$uri/$event" -Headers $Script:DefaultHeader
        }
    }
}
