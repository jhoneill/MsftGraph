using NameSpace Microsoft.Graph.PowerShell.Models

function Get-GraphReminderView   {
    <#
      .Synopsis
        Returns a view of items with reminders set across all a users calendars.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [outputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphReminder])]
    param   (
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        $User = $Global:GraphUser,
        #Time zone to rennder event times. By default the time zone of the local machine will me use
        $Timezone = $(tzutil.exe /g),
        #Number of days of calendar to fetch from today
        [int]$Days = 30 ,
        #The number of events to fetch. Must be greater than zero, and capped at 1000
        [ValidateRange(1,1000)]
        [int]$Top
    )
    begin   {
        $webParams =  @{ Method    = 'Get'
                         ValueOnly = $true
                         AsType    = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphReminder])
        }
        if ($TimeZone) {$webParams['Headers'] = @{'Prefer' = "Outlook.timezone=""$TimeZone"""}}
        #MicrosoftGraphReminder uses strings where there should be dates.
        $whensb = {
            if (  [System.Convert]::ToDateTime($this.eventStartTime.datetime).AddDays(1) -eq
                  [System.Convert]::ToDateTime($this.eventEndTime.datetime )) {
                    $this.eventStartTime.datetime -replace '(\d{2}:\d{2}):00$','$1' -replace '00:00$','All day'
            }
            else { ($this.eventStartTime.datetime -replace '(\d{2}:\d{2}):00$','$1') + ' to ' +
                   ($this.eventEndTime.datetime   -replace '(\d{2}:\d{2}):00$','$1') + $this.eventEndTime.timezone }
        }
    }
    process {
        foreach ($u in $User) {
            if  ($u.ID) {$u=$u.ID}
            #users.functions.yml refers to  /users/$u/microsoft.graph.reminderView(StartDateTime='{0:yyyy-MM-ddTHH:mm:ss}',EndDateTime='{1:yyyy-MM-ddTHH:mm:ss}')
            #https://docs.microsoft.com/en-us/graph/api/user-reminderview?view=graph-rest-1.0&tabs=http gives the syntax here i.e. without "microsoft.graph."
            $uri = "$GraphUri/users/$u/reminderView(startDateTime='{0:yyyy-MM-ddTHH:mm:ss}',endDateTime='{1:yyyy-MM-ddTHH:mm:ss}')"

            $webParams['uri'] =  $uri -f [datetime]::Today, [datetime]::Today.AddDays($days)
            Invoke-GraphRequest @webParams  |
                Add-Member -PassThru -MemberType AliasProperty  -Name 'Subject'           -Value eventSubject  |
                Add-Member -PassThru -MemberType AliasProperty  -Name 'Location'          -Value eventLocation |
                Add-Member -PassThru -MemberType ScriptProperty -Name 'When'              -Value $whenSB       |
                Add-Member -PassThru -MemberType ScriptProperty -Name 'Start'             -Value {[System.Convert]::ToDateTime($this.eventStartTime.datetime )} |
                Add-Member -PassThru -MemberType ScriptProperty -Name 'End'               -Value {[System.Convert]::ToDateTime($this.eventEndTime.datetime )}   |
                Add-Member -PassThru -MemberType ScriptProperty -Name 'Reminder'          -Value {[System.Convert]::ToDateTime($this.reminderFireTime.datetime)}
        }
    }
}

