using NameSpace Microsoft.Graph.PowerShell.Models
# MicrosoftGraphReminder is in Microsoft.Graph.Users.Functions.private.dll
function Get-GraphReminderView   {
    <#
      .Synopsis
        Returns a view of items with reminders set across all a users calendars.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [outputType([Microsoft.Graph.PowerShell.Models.MicrosoftGraphReminder])]
    param   (
        #UserID as a guid or User Principal name, whose calendar should be fetched If not specified defaults to "me"
        [ArgumentCompleter([UPNCompleter])]
        $User = $global:GraphUser,
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
                                     #MicrosoftGraphReminder uses strings where there should be dates. Types.ps1xml adds fields for true dates

        }
        if ($TimeZone) {$webParams['Headers'] = @{'Prefer' = "Outlook.timezone=""$TimeZone"""}}

    }
    process {
        foreach ($u in $User) {
            if  ($u.ID) {$u=$u.ID}
            #users.functions.yml refers to  /users/$u/microsoft.graph.reminderView(StartDateTime='{0:yyyy-MM-ddTHH:mm:ss}',EndDateTime='{1:yyyy-MM-ddTHH:mm:ss}')
            #https://docs.microsoft.com/en-us/graph/api/user-reminderview?view=graph-rest-1.0&tabs=http gives the syntax here i.e. without "microsoft.graph."
            $uri = "$GraphUri/users/$u/reminderView(startDateTime='{0:yyyy-MM-ddTHH:mm:ss}',endDateTime='{1:yyyy-MM-ddTHH:mm:ss}')"

            $webParams['uri'] =  $uri -f [datetime]::Today, [datetime]::Today.AddDays($days)
            Invoke-GraphRequest @webParams
        }
    }
}
