
function Get-GraphGroupConversation {
    <#
      .Synopsis
        Gets details of group converstation from outlook, or its threads.
      .Description
        Requires consent to use the Group.Read.All scope
      .Example
        Get-GraphGroupList -Name consult | Get-GraphGroup -Conversations | Get-GraphGroupConversation -Threads
        Gets group(s) matching the name "consult*" , finds their conversations and for each one gets the threads in the conversation
        Note, unless you are dealing with conversations which have multiple threads, it is easier to do Get-GraphGroup -Threads
    #>
    [Cmdletbinding()]
    [Alias("Get-GraphTeamConversation","Get-GraphConversation")]  #Strictly Conversations belong to a group in Outlook, not a Team in Microsoft teams, but let either name be used.
    param(
        #The Conversation, either as an ID or an object.
        [Parameter(ValueFromPipeline=$true, Mandatory=$true, Position=0, ParameterSetName='OneConversation')]
        $Conversation,
        #The group where the conversation is found, either as an ID or as an object, if it can't be found from the conversation
        [Parameter(ParameterSetName='AllInTeam', Mandatory=$true )]
        [Parameter(ParameterSetName='OneConversation', Position=1)]
        [Alias("Team")]
        $Group,
        #If specified selects the conversation's threads, otherwise an object representing the conversation itself is returned.
        [Parameter(ParameterSetName='OneConversation', Position=1)]
        [Switch]$Threads
    )
    process {
        if ($Group  -and -not $Conversation) {
            Get-GraphGroup -Group $Group -Conversations
            return
        }
        if     (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($Conversation.Group)       {$groupID = $Conversation.Group}
        elseif ($Group.id)                 {$groupId = $Group.ID}
        elseif ($Group -is [String])       {$groupid = $Group}
        else   {Write-Warning 'Could not resolve the group - please specify the group explicitly' ; return}
        if     ($Conversation.id)          {$Conversation = $Conversation.id}

        if ($Threads) {
            $uri    = "$GraphUrl/groups/$groupID/conversations/$conversation/Threads"
            $result = Invoke-GraphRequest @webparams -Uri $uri
            foreach ($thread in $result.value) {
                $thread.pstypenames.add("GraphThread")
                Add-Member -InputObject $thread -MemberType NoteProperty -Name Group        -Value $GroupID
                Add-Member -InputObject $thread -MemberType NoteProperty -Name Conversation -Value $Conversation
            }
            return $result.value
        }
        else     {
            $c = (Invoke-GraphRequest @webparams -Uri ("$GraphUrl/groups/$groupID/conversations/$conversation"  +'?$expand=Threads'))
            $c.pstypenames.add("GraphConversation")
            Add-Member -PassThru -InputObject $c -MemberType NoteProperty -Name Group -Value $GroupID
        }
    }
}

function Get-GraphGroupThread {
    <#
      .Synopsis
        Gets a thread in a Group conversation in outlook, or its posts
      .Description
        Requires consent to use the Group.Read.All scope
      .Example
        >Get-GraphUser -Teams  | Get-GraphGroup -Threads | Get-GraphGroupThread -Posts |
             ft -a -Wrap  @{n="from";e={$_.from.emailaddress.name}},CreatedDateTime,Topic,@{n="Body";e={$_.body.content}}
        Gets a users teams, for each one gets their threads, and for each thread gets the outlook posts
        Displays the result as a table showing from, message date, thread topic and message body
        Note this uses Get-GraphGroup as an alias for Get-GraphTeams
    #>
    [Cmdletbinding()]
    [Alias("Get-GraphTeamThread")]
    param   (
        #The group thread, either as an ID or as a thread object (which may have the team/group as property)
        [Parameter(ParameterSetName='SingleThread', Position=0, ValueFromPipeline=$true, Mandatory=$true)]
        $Thread,
        #The group holding the thread, if it can't be found drm -thread
        [Alias("Team")]
        [Parameter(ParameterSetName='AllGroupThreads', Mandatory=$true)]
        [Parameter(ParameterSetName='SingleThread', Position=1)]
        $Group,
        #If specified, returns the posts in the thread
        [Parameter(ParameterSetName='SingleThread')]
        [Switch]$Posts
    )
    begin   {

        $webparams = @{Headers = @{"Prefer" ='outlook.body-content-type="text"' }}
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if     ($Group -and -not $Thread) {
            Get-GraphGroup -Group $Group -Threads
            return
        }
        if     ($Thread.Group)         {$groupid  = $Thread.group}
        elseif ($Group.id)             {$groupid  = $Group.ID}
        elseif ($Group -is [string])   {$groupid  = $Group}
        else   {Write-Warning -Message 'Could not resolve group ID'; return}

        if     ($Thread.id)            {$threadID = $Thread.id}
        elseif ($Thread -is [string])  {$threadID = $Thread}
        else   {Write-Warning -Message 'Could not resolve thread ID'; return}

        $t = Invoke-GraphRequest @webparams -Uri "$GraphUrl/groups/$Groupid/Threads/$threadID`?`$expand=Posts"
        $t.pstypenames.Add("GraphThread")
        Add-Member     -InputObject $t    -MemberType NoteProperty -Name Group -Value $Groupid
        foreach  ($post in $t.posts) {
            $Post.pstypenames.add("GraphPost")
            Add-Member -InputObject $post -MemberType NoteProperty -Name Group  -Value $Groupid
            Add-Member -InputObject $post -MemberType NoteProperty -Name Thread -Value $t.ID
            Add-Member -InputObject $post -MemberType NoteProperty -Name Topic  -Value $t.Topic
        }
        if ($Posts) {$t.posts}
        else        {$t}
    }
}

function Add-GraphGroupThread {
    <#
      .Synopsis
        Starts a new thread in a group in outlook.
      .Description
        Requires consent to use the Group.ReadWrite.All scope
      .Example
        >
        >$G = Get-GraphGroup -ByName consultants
        >Add-GraphGroupThread -Group $G -Subject "Running tests.." -Content "We will be running a full test pass this afternoon"
        Gets a group by name and creates a new thread with a message using a plain text body.
      .Example
        >$thread = Add-GraphGroupThread -passthru -Group $G -Subject "Ruuning tests.." -ContentType HTML -Content "<b><i>Drum-Roll...</i>A full test pass is running... Watch this space</i>"
        Uses the group from the previous example, and creates a thread with an HTML body, and keeps a reference to it.
      .link
        Send-GraphGroupReply
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The group where the thread will be added
        [Parameter(Mandatory=$true,Position=0)]
        [Alias("Team")]
        $Group,
        #The subject line for the thread
        [Parameter(Mandatory=$true, Position=1)]
        [Alias("Subject")]
        $ThreadTopic,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true, Position=2)]
        [String]$Content,
        #The content type, (Text by default) or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting; this is the default, unless $confirm preference has been changed
        [switch]$Force,
        #if Specified an object containing the Thread ID will be returned
        [switch]$PassThru
    )

    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

    if     ($Group.ID)           {$groupID  = $Group.ID}
    elseif ($Group -is [String]) {$groupID  = $Group   }
    else   {Write-Warning -Message 'Could not process Group parameter.'; return }

    $webparams = @{ 'uri'         = "$GraphUrl/groups/$groupID/threads/"
                    'method'      = 'Post'
                    'contentType' = 'application/json'
     }
    $Settings  = @{ 'topic'       = $ThreadTopic
                    'posts'       = @( @{body= @{'content'     = $Content
                                                 'contentType' = $ContentType}})
    }
    $json      = ConvertTo-Json $settings -Depth 5

    if ($force -or $PSCmdlet.Shouldprocess($ThreadTopic,"Create New thread")) {
        $t = Invoke-GraphRequest  @webparams -Body $json
        if ($PassThru) {
            Start-Sleep -Seconds 2
            Get-GraphGroupThread -Group $Group -Thread $t.id
        }
    }
}

function Remove-GraphGroupThread {
    <#
      .Synopsis
        Removes a thread from a group in outlook
      .Example
        Get-GraphGroup -ByName consultants -Threads | where topic -eq "Today's tests..."  | Remove-GraphGroupThread
        Finds the threads for a named group; isolates one by topic name, and removes it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param (
        #The thread to remove, either as an ID or a thread object containing an ID, and possibly a group ID
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Thread,
        #The group from which the thread is to be removed, either as an ID or a group object containing an ID
        [Alias("Team")]
        $Group,
        #if Specified the thread will be deleted without prompting.
        [switch]$Force
    )
    process {

        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

        if     ($Thread.group)        {$groupid  = $Thread.group}
        elseif ($Group.ID)            {$groupid  = $Group.ID}
        elseif ($Group -is [string])  {$groupid  = $Group}
        else   {Write-Warning 'Could not resolve the group ID' ; return}

        if     ($Thread.ID)           {$threadid = $Thread.id  }
        elseif ($Thread -is [string]) {$threadid = $Thread.id  }
        else   {Write-Warning 'Could not resolve the group ID' ; return}


        $webparams = @{
                       'uri'    =  "$GraphUrl/groups/$GroupID/threads/$threadID"
        }
        Write-Progress -Activity "Deleting thread" -Status "Checking existing thread"
        try   {$thread  = Invoke-GraphRequest -Method Get @webparams }
        catch {
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-warning 'Thread not found, it may have been deleted already'
                return
            }
            else {
                throw $_ ;
                return
            }
        }
        if (-not $thread) {throw "Could not get the thread to delete"; return}
        Write-Progress -Activity "Deleting thread" -Completed
        if ($Force -or $PSCmdlet.Shouldprocess($thread.topic,"Delete thread")) {
            Write-Progress -Activity "Deleting thread" -Status "Sending delete instruction"
            Invoke-GraphRequest -Method Delete  @webparams
            Write-Progress -Activity "Deleting thread" -Completed
        }
    }
}

function Send-GraphGroupReply {
    <#
      .Synopsis
        Replies to a group's post in outlook.
      .Example
        >$thread.posts[0] | Send-GraphGroupReply -content '<b><font color="green">Success!!</font> Go team!</b>' -ContentType HTML
        One of the examples for Add-GraphGroupThread left the result of a creating a new thread in $thread
        This takes the only post in the new thread and creates a reply to it with the content in HTML format.
      .Example
        >
        >$post = Get-GraphGroup -ByName consultants -Threads | where topic -eq "Today's tests..." | Get-GraphGroupThread -Posts | select -last 1
        >Send-GraphGroupReply $post -Content "Please join a celebration of the successful test at 4PM"
        This example finds threads for the consultants group, Isolates the one with the topic of
        "Today's Tests..." and finds the last post in the thread. It then posts are reply with the content as plain text.
      .link
        Add-GraphGroupThread
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param (
        #The Post being replied to, either as an ID or a post object containing an ID which may identify the thread and group
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        $Post,
        #The thread containing the post (if not embedded in the post itself), as an ID or object, which may identify the group
        $Thread,
        #The group containing the thread (if not embedded in the Post or thread) as an ID or object
        [Alias("Team")]
        $Group,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The type of content, text by default or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #if Specified the message will be created without prompting.
        [switch]$Force
    )

    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }

    if     ($Post.Group)          {$groupID  = $Post.group}
    elseif ($Thread.Group)        {$groupID  = $Thread.group}
    elseif ($Group.ID)            {$groupID  = $Group.ID}
    elseif ($Group -is [string])  {$groupID  = $Group}
    else   {Write-warning -Message 'Could not resolve the group ID.' ; return}

    if     ($Post.Thread)         {$threadID = $Post.Thread}
    elseif ($Thread.ID)           {$threadID = $Thread.id  }
    elseif ($Thread -is [String]) {$threadID = $Thread.id  }
    else   {Write-warning -Message 'Could not resolve the Thread ID.' ; return}

    if     ($Post.ID)             {$PostID   = $Post.ID}
    elseif ($Post -is [String])   {$PostID = $Post  }
    else   {Write-warning -Message 'Could not resolve the Post ID.' ; return}

    if (-not ($PostID -and $threadID -and $groupID)) {throw "Could not find Group, Thread and Post IDs from supplied parameters."; Return}


    $uri       =  "$GraphUrl/groups/$groupID/threads/$threadID/posts/$postid"
    Write-Progress -Activity 'Posting reply to thread' -Status 'Checking parent message'
    try   {   $p  = Invoke-GraphRequest -Method Get -uri $uri @webparams }
    catch {       throw "Could not get the post to reply to"; return}
    if (-not $p) {throw "Could not get the post to reply to"; return}

    $Settings  = @{ 'Post' = @{'body'= @{'content'=$Content; 'contentType'=$ContentType}}}
    $Json      = ConvertTo-Json $settings
    Write-Debug $Json

    if ($Force -or $PSCmdlet.Shouldprocess($thread.topic,"Reply to thread")) {
        $uri     += "/Reply"
        Write-Progress -Activity 'Posting reply to thread' -Status 'sending reply'
        Invoke-GraphRequest -Method Post -Uri $URI  @webparams -Body $Json -ContentType "application/json"
        Write-Progress -Activity 'Posting reply to thread' -Completed
    }
}

Function Get-ChannelMessagesByURI {
    <#
      .synopsis
        Helper function to add get and expand messages or replies to messages
    #>
    param (
        [parameter(Position=0,ValueFromPipeline=$true)]
        $URI
    )

    process {

        Write-progress -Activity 'Getting messages' -Status "Reading $($ch.displayname) Messages"
        $result   = Invoke-GraphRequest @webparams -Uri $uri
        $msgList  = @() + $result.value
        while ($result.'@odata.nextLink' -and $result.'@odata.count' -gt 0 ) {
            Write-Verbose  $result.'@odata.count'
            Write-progress -Activity 'Getting messages' -Status "Reading $($ch.displayname) Messages" -CurrentOperation "$($msglist.count) so far"
            $result   = Invoke-GraphRequest  @webparams -Uri $result.'@odata.nextLink'
            $msgList += $result.value
        }
        $userHash = @{}
        Write-progress -Activity 'Getting messages' -Status "Expanding User information"
        $msglist.from.user.id | Sort-Object -Unique | foreach-object {
            $userHash[$_] = ( Invoke-GraphRequest @webparams -Uri  "$GraphUrl/directoryObjects/$_").displayName
        }
        Write-progress -Activity 'Getting messages' -Completed
        foreach ($msg in $msgList) {
            $msg.pstypeNames.add("GraphTeamMsg")
            if ($msg.from.user.id) {
                Add-Member -InputObject $msg -MemberType NoteProperty -Name FromUserName -Value $userHash[$msg.from.user.id]
            }
            Add-Member     -InputObject $msg -MemberType NoteProperty -Name Created      -Value ([datetime]$msg.createdDateTime)
            Add-Member     -InputObject $msg -MemberType NoteProperty -Name team         -Value $teamID
            Add-Member     -InputObject $msg -MemberType NoteProperty -Name channel      -Value $channelID
        }

        $msgList | sort-object -Property Created
    }
}

function Get-GraphChannel {
    <#
      .Synopsis
        Gets details of a channel, or its Tabs or messages shown in Teams
      .Example
        >Get-GraphTeam -ByName consultants -ChannelName general | Get-GraphChannel -Tabs
        Gets channels for the team(s) with a name beginning 'Consultants' and selects channel(s)
        with a name beginning "general"; then gets the tabs shown in Teams for this channel
      .Example
        >Get-GraphTeam -ByName consultants -ChannelName general | Get-GraphChannel -Messages
        This follows the same method for getting the Teams but this time returns messaes in the channel
      .Example
        >Get-GraphChannel -Team $c -ByName general -Messages
        This is a variation on the previous example - here $c holds an object describing
        the consultants Team and the channel and its messages are retieved using a single command.
      .Example
        >Get-GraphChannel -Team $c -ByName -channel ""
        This previous example didn't explictly specify the channel parameter when using the
        ByName switch; this version does and specifies and empty string so it will return all
        channels (channel is a required parameter, but it can be an empty string)
    #>
    [Cmdletbinding(DefaultparameterSetName="None")]
    [Alias("Get-GraphTeamChannel")]
    param(
        #The channel either as an ID or as a channel object (which may contain the team as a property)
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)]
        $Channel,
        #If Channnel is string it is treated as an ID unless -ByName is specified
        [switch]$ByName,
        #The ID of the team if it is not in the channel object.
        $Team,
        #If specified gets the channel's Tabs
        [Parameter(parameterSetName="Tabs", Mandatory=$true)]
        [switch]$Tabs,
        #if Specified uses the beta api to get the channel's messages.
        [Parameter(parameterSetName="Msgs")]
        [Alias("Msgs")]
        [switch]$Messages,
        #If specified, returns the top n messages, otherwise the command will attempt to get all messages. The server may return more than the specified number.
        [Parameter(parameterSetName="Msgs")]
        $Top
    )

    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if ($ByName) {
            $Channel = Get-GraphTeam -Team $Team -Channels -ChannelName $channel
        }
        #ByName might return multiple channels. Support -channel being given an array of channels.
        foreach ($ch in $channel) {
            if     ($ch.Team)           {$teamID    = $ch.team }
            elseif ($Team.ID)           {$teamID    = $Team.ID }
            elseif ($Team -is [string]) {$teamID    = $Team    }
            else   {Write-Warning -Message 'Could not resolve the team for this channel'; return}
            if     ($ch.id  )           {$channelID = $ch.ID   }
            elseif ($ch -is [string])   {$channelID = $ch      }
            else   {Write-Warning -Message 'Could not resolve the channel ID'; return}
            if (-not ($teamid -and $channelID)) {Write-warning -Message "You need to provide a team ID and a Channel ID"; return}
            elseif   ($Messages -or $Top) {
                $uri      =  "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/messages"
                if ($Top) {$uri += '?$top=' + $Top }
                Get-ChannelMessagesByURI -URI $uri
                return
            }
            elseif   ($Tabs)     {
                $results = Invoke-GraphRequest @webparams -Uri  "$GraphUrl/teams/$teamID/channels/$channelID/tabs?`$expand=teamsApp"
                $t       = $results.value
                foreach ($tab in $t) {
                    $tab.pstypeNames.add('GraphTab')
                    #newly created tabs have a teamsAppId property. Existing apps have to look at the teamsApp and its ID. Make them the same!
                    Add-Member -InputObject $tab -MemberType ScriptProperty -Name teamsAppId   -Value {$this.teamsApp.ID}
                    Add-Member -InputObject $tab -MemberType ScriptProperty -Name teamsAppName -Value {$this.teamsApp.displayName}
                }
                return $t
            }
            elseif   ($ByName)   {
                #Have already fetched the channel once so don't fetch it again
                $ch
            }
            else                 {
                $result = Invoke-GraphRequest @webparams -Uri  "$GraphUrl/teams/$teamID/channels/$channelId"
                $result.pstypeNames.add("GraphChannel")
                Add-Member -InputObject $result -MemberType NoteProperty -Name Team -Value $teamID

                $result
            }
        }
    }
}

function New-GraphChannel {
    <#
      .Synopsis
        Adds a channel to a team
      .Description
        This requires the Group.ReadWrite.All scope.
      .Example
       >$newChannel  = New-GraphChannel -Team $newTeam -Name $newProjectName -Description "For anything about project $newProjectName"
       $newTeam holds the result of creating a team with New-GraphTeam...
       $newProjectName holds the name of a project the team will be working on.
       This command creates a new channel in Teams, and stores the result in a variable
       which can then be used to post messages to the channel, or add tabs to it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true)]
    [Alias("Add-GraphTeamChannel")]
    param(
        #The team where the channel will be added, either as an ID or a team object
        [Parameter( Mandatory=$true, Position=0)]
        $Team,
        #Display name for the new channel
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Alias("DisplayName")]
        [String]$Name,
        #Description for the new channel
        [String]$Description
    )
    begin  {
        if ($Team.id) {$Team = $Team.id}

    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        if (Get-GraphTeam -Channels $team | Where-Object -property displayname -eq $Name) {
            Write-warning -Message "Channel $Name already exists in this team"
        }
        $webparams = @{Method = "POST"

                    URI    = "$GraphUrl/teams/$Team/channels"
                    ContentType = "application/json"
        }
        $Settings  = @{"displayName" = $Name}
        if ($Description) {$settings["description"] = $Description}
        if ($PSCmdlet.Shouldprocess($Name,"Create channel")) {
            $channel =  Invoke-GraphRequest @webparams -body (ConvertTo-Json $settings)
            $channel.psTypenames.add('GraphChannel')
            Add-Member -InputObject $channel -MemberType NoteProperty -Name Team -Value $Team

            $channel
        }
    }
}

function Remove-GraphChannel {
    <#
      .Synopsis
        Removes a channel from a team
      .Description
        This requires the Group.ReadWrite.All scope.
      .Example
        >Get-GraphTeam -ByName Developers -ChannelName "Project Firebird" | Remove-GraphChannel
        Finds a channel by name from a named team , and removes it.
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='High')]
    param(
        #The channel to delete; either as an ID, or a channel object
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        $Team,
        #if Specified the channel will be deleted without prompting
        [switch]$Force
    )
    process {

        if ($Channel.Team) { $Team    = $Channel.team }
        elseif  ($Team.id) { $Team    = $Team.ID}
        if ($Channel.id  ) { $Channel = $Channel.ID   }

        try {
            $c = Get-GraphChannel -Channel $Channel -Team $Team
        }
        Catch {
            throw "Could not get the channel" ; return
        }
        if (-not $c)  {throw "Could not get the channel" ; return }
        if ($force -or $PSCmdlet.Shouldprocess($c.displayname, "Delete Channel")) {

            Invoke-GraphRequest -Method "Delete" -Uri "$GraphUrl/teams/$Team/channels/$Channel"
            }
        }
}

function Add-GraphChannelThread {
    <#
      .Synopsis
        Adds a new thread in a channel in Teams.
      .Description
        This uses BETA functionality.
      .Example
        >
        >$General = Get-GraphTeam $newTeam -ChannelName "General"
        >Add-GraphChannelThread -Channel $General -Content "Project Firebird now has its own channel."
        This adds a message t
    #>
    [Cmdletbinding(SupportsShouldprocess=$true, ConfirmImpact='Low')]
    param(
        #The channel to post to either as an ID or a channel object.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team object or the ID of the team, if it can't be derived from the channel.
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #Normally the message is added 'silently'. If passthru is specified, the new message will be returned.
        [Alias('PT')]
        [switch]$Passthru,
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    process {
        if     ($Channel.Team)      {$teamID    = $Channel.team }
        elseif ($Team.id)           {$teamID    = $Team.ID}
        elseif ($Team -is [string]) {$teamID    = $Team}
        else   {Write-Warning -Message 'Could not determine the team ID'; return}

        if     ($Channel.id)           {$channelID = $Channel.ID   }
        elseif ($Channel -is [string]) {$channelID = $channel  }
        else   {Write-Warning -Message 'Could not determine the channel ID'; return}

        try {$c = Get-GraphChannel -Channel $channelID -Team $teamID }
        Catch         {throw "Could not get the channel" ; return}
        if (-not $c)  {throw "Could not get the channel" ; return }
        $webparams = @{ 'Method'      = 'POST'
                         'URI'         = "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/chatThreads"
                        'ContentType' = 'application/json'
        }
        $Settings = @{ rootMessage = @{body= @{content=$Content;}}}
        if ($ContentType -eq 'HTML') {$settings.rootMessage.body['contentType'] = 1}
        else                         {$settings.rootMessage.body['contentType'] = 2}
        $json =  (ConvertTo-Json $settings)
        Write-Debug $json
        if ($force -or $PSCmdlet.Shouldprocess("Create Message")) {
            $result = Invoke-GraphRequest @webparams  -Body $json
            If ($Passthru) {
                $URI    = "https://graph.microsoft.com/beta/teams/$teamid/channels/$channelid/Messages/$($result.id)"
                $msg    = Invoke-GraphRequest -Uri $uri
                $msg.pstypenames.add('GraphTeammsg')

                $msg
            }
        }
    }
}

 # can get replies in a thread , but can't send a reply, delete or update a thread

function Get-GraphChannelReply {
    <#
      .Synopsis
        Returns replies to messages in Teams channels
      .Description
        Access to channel messages is currently in the BETA API
        It is possible to start a new thread, but not to reply to the thread.
      .Example
        >Get-GraphChannel $General -Messages | Get-GraphChannelReply -PassThru
        The GraphAPI does not return replies when requesting messages
        from a channel in Teams. By piping the messages to Get-GraphChannelReply
        it is possible to get the replies; and if -Passthru is specified
        the messages will returned, followed by their replies.
        So if $General is a channel object, the first message and the its first
        reply might be output like this.

        From          Created          Isreply Deleted Importance Content
        ----          -------          ------- ------- ---------- -------
        James O'Neill 17/02/2019 11:42 False   False   normal     Project Firebird now has its own channel.
        James O'Neill 17/02/2019 13:06 True    False   normal     And the channel has its own planner

    #>
    [Cmdletbinding()]
    param (
        #The Message to reply to as an ID or a message object containing an ID (and possibly the team and channel ID)
        [Parameter(Position=0,ValueFromPipeline)]
        $Message,
        #If the Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        $Channel,
        #If the message or channel parameters don't included the team ID, the team either as an ID or an objec containing the ID
        $Team,
        #If specified returns the message, followed by its replies. (Otherwise , only the replies are returned)
        [switch]$PassThru
    )
    #Is the team ID in the message, channel or team parameter ? Is the channel in the message parameter, and is message an object or ID?
    if     ($Message.team)         {$teamid    = $Message.team}
    elseif ($Channel.team)         {$teamid    = $Channel.team}
    elseif ($Team.id)              {$teamid    = $team.id}
    elseif ($Team -is [string])    {$teamid    = $team}
    else   {Write-Warning 'Could not determine the team ID for the message.'; return}
    if     ($Message.channel)      {$channelid = $Message.channel}
    elseif ($Channel.id)           {$channelid = $channel.id}
    elseif ($Channel -is [string]) {$channelid = $Channel}
    else   {Write-Warning 'Could not determine the Channel ID for the message.'; return}
    if     ($Message.ID)           {$msgID     = $Message.ID}
    elseif ($Message -is [string]) {$msgID     = $Message }
    else   {Write-Warning 'Could not determine the ID for the message.'; return}

    if ($PassThru) {$Message}
    Get-ChannelMessagesByURI -URI "https://graph.microsoft.com/beta/teams/$teamid/channels/$channelid/Messages/$msgID/replies"
}

#code for send reply is included here but not exported in the PSD1 file. It is described at
#https://docs.microsoft.com/en-us/graph/api/channel-post-messagereply?view=graph-rest-beta
#but calling it gives  (501) Not Implemented.
function Send-GraphChannelReply {
    <#
      .Synopsis
        Posts a reply to a message in a Teams channel
    #>
    [Cmdletbinding(SupportsShouldProcess=$true)]
    param (
        #The Message to reply to as an ID or a message object containing an ID (and possibly the team and channel ID)
        $Message,
        #If the Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID
        $Channel,
        #If the message or channel parameters don't included the team ID, the team either as an ID or an objec containing the ID
        $Team,
        #The Message body - text by default, specify -contentType if using HTML
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The format of the content, text by default , or HTML
        [ValidateSet("Text","HTML")]
        [String]$ContentType = "Text",
        #Normally the reply is added 'silently'. If passthru is specified, the new message will be returned.
        [Alias('PT')]
        [switch]$Passthru,
        #if Specified the message will be created without prompting.
        [switch]$Force
    )
    if     ($Message.team)         {$teamid    = $Message.team}
    elseif ($Channel.team)         {$teamid    = $Channel.team}
    elseif ($Team.id)              {$teamid    = $team.id}
    elseif ($Team -is [string])    {$teamid    = $team}
    else   {Write-Warning 'Could not determine the team ID for the message.'; return}
    if     ($Message.channel)      {$channelid = $Message.channel}
    elseif ($Channel.id)           {$channelid = $channel.id}
    elseif ($Channel -is [string]) {$channelid = $Channel}
    else   {Write-Warning 'Could not determine the Channel ID for the message.'; return}
    if     ($Message.ID)           {$msgID     = $Message.ID}
    elseif ($Message -is [string]) {$msgID     = $Message }
    else   {Write-Warning 'Could not determine the ID for the message.'; return}

    $uri =  "https://graph.microsoft.com/beta/teams/$teamid/channels/$channelid/Messages/$msgID"
    try   {$null = Invoke-GraphRequest -Method Get  -Uri $uri }
    catch {Write-Warning -Message 'Those parameters did correspond to a message. Cannot Continue.'; return}

    $Settings =  @{body= @{content=$Content;}}
    if ($ContentType -eq 'HTML') {$settings.body['contentType'] = 1}
    else                         {$settings.body['contentType'] = 2}
    $json =  (ConvertTo-Json $settings)
    Write-Debug $json
    if ($force -or $PSCmdlet.Shouldprocess("Post Reply")) {
        Invoke-GraphRequest -Method 'POST' -Uri "$uri/replies" -ContentType 'application/json' -Body $json
    }
}

function Add-GraphWikiTab {
    <#
      .Synopsis
        Adds a wiki tab to a channel in teams
      .Example
        >Add-GraphWikiTab -Channel $Channel -TabLabel Wiki
        Channel contains an object representing a channel in teams,
        this adds a Wiki to it. The Wiki will need to be initialized
        when the tab is first opened
    #>
    [CmdletBinding(SupportsShouldprocess=$true)]
    param(
        #An ID or Channel object which may contain the team ID
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab
        $TabLabel = "Wiki",
        #If specified the tab will be added without prompting for confirmation
        [switch]$Force,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        [Alias('PT')]
        [switch]$PassThru
    )

    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    if     ($Channel.Team)         {$teamID  = $Channel.Team }
    elseif ($Team.id)              {$teamID  = $Team.id      }
    elseif ($Team -is [String])    {$teamID  = $Team}
    else   {Write-Warning -Message 'Could not determine the team from the channel. Please Specify the team explicitly.'; return}
    if     ($Channel.id)           {$channelID = $Channel.id }
    elseif ($Channel -is [string]) {$channelID = $Channel    }
    else   {Write-Warning -Message 'Could not determine the channel ID.'; return}
    $webparams = @{'Method'      = 'Post'
                   'Uri'         = "https://graph.microsoft.com/beta/teams/$teamID/channels/$channelID/tabs"
                    'ContentType' = 'application/json'
    }
    $json = ConvertTo-Json ([ordered]@{
                    'name'       = $TabLabel
                    'TeamsAppId' = 'com.microsoft.teamspace.tab.wiki'
            })
    Write-Debug $json
    if ($Force -or $PSCmdlet.Shouldprocess($TabLabel,"Create wiki tab")) {
        $result = Invoke-GraphRequest @webparams -body $json
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'Wiki'

            $result
        }
    }
}
# Adding tab https://docs.microsoft.com/en-us/graph/api/teamstab-add?view=graph-rest-1.0
# https://products.office.com/en-us/microsoft-teams/appDefinitions.xml

function Add-GraphOneNoteTab     {
    <#
      .Synopsis
        Adds a tab in a Teams channel for a OneNote section or Notebook
      .Description
        This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
        which requires consent to use the Group.ReadWrite.All scope.
        The Notebook Parameter has an alias of 'Section' and will accept either
        a OneNote Notebook object (or its 'Self' URI - which requires the tab name to be
        set explicitly) or a Section object. If the notebook is specified it opens at the
        first section.
      .Example
        >
        > $section = Get-GraphTeam -ByName accounts -Notebooks | Select-Object -ExpandProperty sections  | where displayname -like "FY-19*"
        > $channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
        > Add-GraphOneNoteTab  $section $channel -TabLabel "FY-19 Notes"

        The first command gets the Notebook for the Accounts team and finds the "FY-19 Year End" section
        The second command gets the channels for the same team and finds the "Year end" channel
        The Third command creates a tab in the channel named 'FY-19 Notes' which opens the team notebook
        at its 'FY-19 Year End' section.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #The Notebook or Section to associate with the tab
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Section')]
        $Notebook,
        #An ID or Channel object which may contain the team ID; the tab will be created in this channel
        [Parameter(Mandatory=$true, Position=1)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab, if left blank the name of the Notebook or Section will be sued
        $TabLabel,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        [Alias('PT')]
        [switch]$PassThru,
        #If Specified the tab will be added without pausing for confirmation, this is the default unless $ConfirmPreference has been set.
        $Force
    )
    ContextHas -WorkOrSchoolAccount -BreakIfNot
    if       ($Channel.Team)           {$Team     = $Channel.Team }
    elseif   ($Team.id)                {$Team     = $Team.id}
    elseif   ($team -isnot [string])   {Write-Warning 'Unable to determine the team, please specify it explicitly'; return}
    if       ($Channel.id) {           $Channel   = $Channel.id }
    elseif   ($Channel-isnot [string]) {Write-Warning 'Unable to determine the channel'; return}
    if       (-not $TabLabel -and
                $notebook.displayName) {$TabLabel = $Notebook.displayName}
    elseif   (-not $TabLabel)          {Write-warning 'Unable to determin a name for the tab, please specify one explicitly'; return}

    $webparams = @{'Method'       = 'Post';
                   'Uri'          = "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs" ;
                   'ContentType'  = 'application/json'
    }
    #This bit had to be reverse engineered, from a beta version of the API, so if it works past next week, be happy.
    #If the "Notebook" object is actually a section, and it was fetched by one of the module commands (get-GraphTeam -notebook, or get-graphNotebook -section)
    #then $Notebook it will have a a parentNotebook ID. This IF..Else is to make sure we have the real notebook ID, and catch a sectionID if there is one.
    if   ($Notebook.parentNotebook.id) {
                    $ParamsPt2    = '&notebookSource=PickSection&sectionId='+ $Notebook.id
                    $NotebookID   = $Notebook.parentNotebook.id
          }
    else  {         $ParamsPt2    = '&notebookSource=New'
                    $NotebookID   = $Notebook.id }

    #if $Notebook is a section its url will end ?wd=(something). We need to split this off the URL and re-use it. The () need to be unescapted too,
    if ($notebook.links.oneNoteWebUrl.href -match '\?(wd=.*$)') {
                $ParamsPt2       += '&' + ( $Matches[1] -replace '%28','(' -replace '%29',')' )
                $OnenoteWebUrl    = $notebook.links.oneNoteWebUrl.href  -replace  '\?wd=.*$', ''
    }
    else      { $OnenoteWebUrl    = $notebook.links.oneNoteWebUrl.href}

    #We need the teamsite URL for the team who owns this channel, and the URL to the the Notebook. Both need to be escaped.
    $OnenoteWebUrl  = $OnenoteWebUrl                           -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'
    $siteUrl        = (Get-GraphTeam -Team $Team -Site).webUrl -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'

    #Now we need to build up the mother and father of all URIs It contains the ID and URL for the notebook (not section). The Name, the teamsite. And Section specifics if applicable.
    $URIParams      = "?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&ui={locale}&tenantId={tid}"+
                      "&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$Team%2Fnotes%2Fnotebooks%2F"+ $NotebookID   +
                      "&oneNoteWebUrl=" + $oneNoteWebUrl +
                      "&notebookName="  + [uri]::EscapeDataString( $notebook.displayName ) +
                      "&siteUrl="       + $SiteUrl +
                      $ParamsPt2

    #Now we can create the JSON. Such information as there is can be found at https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
    $json = ConvertTo-Json ([ordered]@{
                'TeamsAppId'      = '0d820ecd-def2-4297-adad-78056cde7c78'
                'name'            = $TabLabel
                'configuration'   = [ordered]@{
                    'entityId'    = ((New-Guid).tostring() + "_" +  $Notebook.ID)
                    'contentUrl'  = "https://www.onenote.com/teams/TabContent" + $URIParams
                    'removeUrl'   = "https://www.onenote.com/teams/TabRemove"  + $URIParams
                    'websiteUrl'  = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"
                }})
    $json= $json  -replace "\\u0026","&"
    Write-Debug $json
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
        $result = Invoke-GraphRequest -body $json @webparams
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'OneNote'
            return $result
        }
    }
}
