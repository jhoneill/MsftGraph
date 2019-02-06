function Get-GraphMailTips       {
    <#
      .synopsis
        Gets mail tips for one or more users (is their mailbox full, are auto-replies on etc)
    #>
    [cmdletbinding()]
    param(
        #mail addresses
        [Parameter(Mandatory=$true)]
        [string[]]$Address
    )

    $json = Convertto-Json @{EmailAddresses=$Address;
                  MailTipsOptions= "automaticReplies, mailboxFullStatus, customMailTip, deliveryRestriction, externalMemberCount, maxMessageSize, moderationStatus, recipientScope, recipientSuggestions, totalMemberCount"
     }

    Connect-MSGraph
    (Invoke-RestMethod -Method post -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/me/getMailTips" -Body $json -ContentType "application/json").value
}

function Get-GraphMailFolderList {
    <#
      .Synopsis
        Get the user's Mailbox folders
      .Example
        Get-GraphMailFolderList -Name inbox
        Gets the current users inbox folder
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$UserID,
        #Select the first n folders.
        [validaterange(1,1000)]
        [int]$Top,
        #fields to select in the query - will add a validate set later
        [string[]]$Select  ,
        #String with orderby clause e.g. "name", "lastmodifiedDate desc"
        [string]$OrderBy,
        #filter the folders returned by a name
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
        [string]$Name,
        #A custom filter clause.
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )

    Connect-MSGraph
    $webParams = @{Method = "Get"
                    Headers = $Script:DefaultHeader
    }
    #region set-up URI . If we got a user ID, use it other otherwise use the current user, add select, orderby, filter & top parameters as needed
    if ($UserID)  {$uri = "https://graph.microsoft.com/v1.0/users/$userID/mailFolders" }
    else          {$uri = "https://graph.microsoft.com/v1.0/me/mailFolders" }
    $JoinChar = "?"  #Will the next parameter be joined onto the URI with a "?"" or with "&"  ?
    if ($Select)  {$uri = $uri + '?$select=' + ($Select -join ',') ;                                 $JoinChar = "&"}
    if ($Name)    {$uri = $uri + $JoinChar + ("`$filter=startswith(displayName,'{0}') " -f $Name ) ; $JoinChar = "&"}
    if ($Filter)  {$uri = $uri + $JoinChar + '$Filter='  +$Filter                                  ; $JoinChar = "&"}
    if ($OrderBy) {$uri = $uri + $JoinChar + '$orderby='  +$Filter                                 ; $JoinChar = "&"}
    if ($Top)     {$uri = $uri + $JoinChar + '$top=' + $top }
    #endregion

    #region get the data, cope with it being paged add a type to help formatting and return the result
    $folderList    = @()
    $result       = Invoke-RestMethod @webParams -Uri $uri
    $folderList   += $result.value
    while ($result.'@odata.nextLink') {
        $result          =Invoke-RestMethod @webParams -Uri  $result.'@odata.nextLink' ;
        $folderList += $result.value
    }

    foreach ($f in $folderList) {$f.pstypenames.add("GraphMailFolder")}
    return $folderList
    #endregion
}

function Get-GraphMailItem       {
    <#
      .Synopsis
        Get items in a mail folder
      .Example
        >Get-GraphMailItem -top 5
        Gets the top 5 items in the current users Inbox
      .Example
        >Get-GraphMailItem -Mailfolder "sentitems" -top 5
        Gets the top 5 items in the current users sent items folder
      .Example
        >Get-GraphMailFolderList -Name sent | Get-GraphMailItem -top 5
        This has the same result as before but could find any folder
      .Example
        >Get-GraphMailItem -Search 'criminal'
        Searches the default folder (inbox) for 'Criminal' in any field
      .Example
        >Get-GraphMailItem -Search 'criminal' -Mailfolder ''
        Searches for 'Criminal' in any field but this time searches the whole mailbox
      .Example
        >Get-GraphMailItem -Search 'subject:criminal' -Mailfolder ''
        This time limits the search to just the subject line. from:, to: etc can be
        used in the same way as they can in a search in outlook.
      .Example
        >Get-GraphMailItem -filter "from/emailAddress/address eq 'alex@contoso.com'"
        Instead of a free text search this applies a filter on email address, looking at the inbox.
      .Example
        Get-GraphMailItem -Filter "(hasattachments eq true) and startswith(from/emailAddress/name, 'alex')"
        This shows a filter based on two conditions.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$User ,
        #The ID of a folder, or one of the well known folder names 'archive', 'clutter', 'conflicts', 'conversationhistory', 'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures', 'msgfolderroot', 'outbox', 'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'sentitems', 'serverfailures', 'syncissues'
        [Parameter(ValueFromPipeline=$true)]
        $Mailfolder = "Inbox",
        #if specified the command will return child folders instead of messages
        [switch]$ChildFolders,
        #A term to do a free text search for in the mail box (see examples)
        [string]$Search,
        #If specified returns the top X items
        [int]$Top,
        #Sorting option, defaults to sorting by SentDateTime with newest first. Searches are not sorted.
        [string]$OrderBy ='SentdateTime desc',
        #Select particular mail fields , ignored if -ChildFolders is specified; defaults to From, Subject, SentDateTime, BodyPreview, and Weblink
        [ValidateSet('bccRecipients', 'body', 'bodyPreview', 'categories', 'ccRecipients', 'changeKey', 'conversationId', 'createdDateTime',
        'flag', 'from', 'hasAttachments', 'id', 'importance', 'inferenceClassification', 'internetMessageHeaders', 'internetMessageId',
        'isDeliveryReceiptRequested', 'isDraft', 'isRead', 'isReadReceiptRequested', 'lastModifiedDateTime', 'parentFolderId',
        'receivedDateTime', 'replyTo', 'sender', 'sentDateTime', 'subject', 'toRecipients', 'uniqueBody', 'webLink' )]
        [string[]]$Select = @('From', 'Subject', 'SentDatetime', 'BodyPreview', 'weblink'),
        #A Custom filter string; for example "importance eq high" - the examples have more cases
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )
    begin   {
        Connect-MSGraph
    }
    process {
        if     ($Mailfolder.id) {$MailPath = 'mailfolders/' +  $Mailfolder.id}
        elseif ($Mailfolder)    {$MailPath = 'mailfolders/' + ($Mailfolder -replace '^/','')}
        else                    {$MailPath = ''}

        if ($User.id) {$User  = $User.id}
        if ($User)    {$uri   = "https://graph.microsoft.com/v1.0/users/$user/$MailPath" }
        else          {$uri   = "https://graph.microsoft.com/v1.0/me/$MailPath" }
        $webParams = @{Method = "Get"
                      Headers = $Script:DefaultHeader
        }

        if ($ChildFolders -and '' -ne $MailPath)    {
            $result = (Invoke-RestMethod @webParams -Uri "$uri/childfolders")
            $folderList = $result.value
            foreach ($f in $folderList) {$f.pstypenames.add("GraphMailFolder")}
            return $folderList
        }
        elseif ($ChildFolders) {
            Write-Warning -Message 'You need to specify a folder when requesting child folders.'
        }
        else {
            $webParams.Headers["Prefer"] ='outlook.body-content-type="text"'
            $uri =  $uri + '/messages?$select='  + ($Select -join ',')
            if     ($Top)    {$uri = $uri + '&$top='     + $Top              }
            if     ($Search) {$Uri = $uri + '&$search="' + $Search + '"'     }
            elseif ($Filter) {$Uri = $uri + '&$filter='  + $Filter + ''      }
            else             {$uri = $uri + '&$orderby=' + $OrderBy          }


            (Invoke-RestMethod @webParams -Uri $uri ).value |
                ForEach-Object {$_.pstypeNames.add("GraphMailMessage") ; $_ } |
                Add-Member -PassThru -MemberType ScriptProperty -Name "fromName"    -Value {$this.from.emailAddress.name} |
                Add-Member -PassThru -MemberType ScriptProperty -Name "fromAddress" -Value {$this.from.emailAddress.address} |
                Add-Member -PassThru -MemberType ScriptProperty -Name "bodyText"    -Value {$this.body.content}
        }
    }
}

function New-MailAddress         {
    param (
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        [String]$Mail,
        #The displayname for the recipient
        $DisplayName
    )
    $recip = @{address=$Mail}
    if ($DisplayName) {$recip['name'] = $DisplayName}
    
    $recip
}

function New-Recipient           {
    <#
      .Synopsis
        Creats a new meeting attendee, with a mail address and the type of attendance.
    #>
    param(
        # The recipient's email address, e.g Alex@contoso.com
        [Parameter(Mandatory=$true,Position=0, ValueFromPipeline=$true)]
        $Mail,
        #The displayname for the recipient
        $DisplayName
    )
    @{ 'emailAddress' = (New-MailAddress -Mail:$mail -DisplayName:$DisplayName )}
}

function Send-GraphMailMessage   {
    <#
      .Synopsis
        Sends Mail using the Graph API from the current user's mailbox.
      .Example
        >Send-GraphMail -To "chris@contoso.com" -subject "You left your keys behind[nt]"
        Sends a mail with a subject but no body or attachments
      .Example
        >Send-GraphMail -To "chris@contoso.com" -body "Keys are with reception" -NoSave
        Sends a mail but thi time the subject will read "No subject" and the test will be in the body.
        -NoSave means that no copy of this message will be kept in sent items
      .Example
        >Send-GraphMail -To "chris@contoso.com" -Subject "Screen shot" -body "How does this look ?" -Attachments .\Logon.png -Receipt
        #This message has an attachement and requests a read receipt.
      .Example
        >$body"<h1>New dialog</h1><br /><img src='cid:Logon.png' -alt='Look at that'><br/>what do you think"
        >$link = Send-GraphMail -To "jhoneill@waitrose.com" -Subject "Login Sreen" -body $body -BodyType HTML  -NoSave -Attachments .\Logon.png -SaveDraftOnly
        This creates an HTML body, the attached picture can be referenced in an <img> tag with cid:fileName.ext
        this time the mail is not sent but left in the user's drafts folder for review.
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param (
        #Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.
        [parameter(Mandatory=$true,Position=0)]
        $To ,
        #Recipient(s) on the "CC" line,
        $CC  ,
        #Recipient(s) on the "Bcc line" line,
        $BCC,
        #The subject of the message. A message must have a subject and/or body and/or attachments. If the subject is left blank it will be sent as "No Subject"
        [String]$Subject,
        #The content of the message; assumed to be plain text, but HTML can be specified with -BodyType
        [String]$Body    ,
        #The type of the body  content. Possible values are Text and HTML.
        [ValidateSet("Text","HTML")]
        $BodyType = "Text",
        #The importance of the message: Low, Normal or High
        [ValidateSet('Low','Normal', 'High')]
        $Importance = 'Normal' ,
        #Path to file(s) to send as attachments
        $Attachments,
        #If Specified, requests a receipt.
        [switch]$Receipt,
        #If specified leaves the message in the drafts folder without sending it and returns a link to open the message.
        [parameter(ParameterSetName='SaveDraftOnly',Mandatory=$true)]
        [switch]$SaveDraftOnly,
        #If specified specifies that a copy of the mail should not be saved
        [parameter(ParameterSetName='NoSave',Mandatory=$true)]
        [switch]$NoSave
    )

    #Do we post a message, or do we create a draft ? We need to check attacment sizes to be sure...
    $asDraft = [bool]$SaveDraftOnly
    if ($Attachments) {
        $AttachmentItems = Get-item $Attachments -ErrorAction SilentlyContinue
        if (-not $AttachmentItems) {
            Write-Warning (($Attachments -join ", ") + "Gave no items. Message sending will continue")
        }
        else {
            if ($Attachments.Where({$_.length -gt 2.85mb}))  {
                #The Maximum size for a POST is 4MB.
                #Attachments are base 64 encoded so 3MB of attachements become 4MB. Don't try closer than 95% of that
                throw ("Attachment would exceed maximum size for a POST. Maximum file size is ~ 2,900,000 bytes")
                return
            }
            elseif (-not $asDraft -and ($Attachments | Measure-Object -Sum length).sum -gt 2.7mb) {
                #If all the attaments add up to more than 90% of the possible message size, we need to
                #create a a draft and add each on its own.  BUT this method does not support "No save to sent items"
                if ($NoSave) {
                    throw ("The total size of attachments would result in an HTTP Post which greater than 4MB. Individual uploads are not possible when SaveToSentItems is disabled.")
                    return
                }
                Else {
                    Write-Verbose -Message "After BASE64 encoding attacments, message may exceed 4MB. Using Draft and sequential attachment method"
                    $asDraft= $true
                }
            }
            else { Write-Verbose -Message "$($Attachments).count attachment(s); small enough to send in a single operation"}
         }
    }
    elseif (-not $Subject -and -not $Body) {
        Write-Warning -Message "Nothing to send" ; return
    }
    elseif (-not $Subject) {$Subject = "No subject"}

    Connect-MSGraph
    $webParams = @{Headers = $Script:DefaultHeader}

    if ($asDraft) {$Uri = "https://graph.microsoft.com/v1.0/me/Messages"}
    else          {$Uri = "https://graph.microsoft.com/v1.0/me/sendmail"}

    #Build a hash table with the parts of the message, this will be coverted into JSON
    #BEWARE names are case sensitive. if you create $msgSettings.Body instead of $msgSettings.body
    #the capital B will cause a 400 bad request error.
    #My personal coding style is to use inital CAPS for parameters and inital lower case for variables (though Powershell doesn't care)
    #so the parameter is $Body and the hash table key name and JSON label is body.

    $msgSettings   =  @{   body = @{ contentType  = $BodyType;
                                         content  = $Body}
                                         subject  = $Subject
                                      importance  = $Importance
                                     toRecipients = @()
    }
    foreach ($recip in $To ) {
            if     ($recip  -is [string] ) { $msgSettings[ 'toRecipients'] += New-Recipient $recip}
            else                           { $msgSettings[ 'toRecipients'] += $recip}
    }
    if     ($CC) {
        $msgSettings['ccRecipients']      = @()
        foreach ($recip in $cc ) {
            if     ($recip  -is [string] ) { $msgSettings[ 'ccRecipients'] += New-Recipient $recip}
            else                           { $msgSettings[ 'ccRecipients'] += $recip}}
    }
    if     ($BCC) {
                $msgSettings['bccRecipients']      = @()
        foreach ($recip in $bcc ) {
            if     ($recip  -is [string] ) { $msgSettings['bccRecipients'] += New-Recipient $recip}
            else                           { $msgSettings['bccRecipients'] += $recip}}
    }
    if ($Receipt)                          { $msgSettings['isDeliveryReceiptRequested'] = $true }

    #If we are creating a draft, save it now; if sending-in-one be ready for attachments
    if ($asDraft) {
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading draft"
        $json = ConvertTo-Json $msgSettings -Depth 5 #default depth isn't enough !
        try            {$msg  = Invoke-RestMethod @webParams -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch          {throw "There was an error creating the draft message."; return }
        if (-not $msg) {throw "The draft message was not created as expected" ; return }
        else           {
            Write-Verbose -Message "Message created with id '$($msg.id)'"
            $uri = $uri + "/" + $msg.id
        }
    }
    elseif ($AttachmentItems) {
        $msgSettings["attachments"]= @()
    }

    foreach ($f in $AttachmentItems) {
        $Filesettings = @{
            '@odata.type' = '#microsoft.graph.fileAttachment';
            name          = $f.Name ;
            contentId     = $f.name ;
            contentBytes  =  [convert]::ToBase64String( [system.io.file]::readallbytes($f.FullName))

        }
        if ($asDraft) {
            Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading $($f.Name)"
            try {
                $null = Invoke-RestMethod @webParams -Method post  -uri "$uri/attachments"  -Body (ConvertTo-Json $Filesettings) -ContentType "application/json" -ErrorAction Stop
            }
            catch {
                Write-warning -Message "Error occured uploading file $($f.name) - will attempt to delete the draft message"
                Invoke-RestMethod @webParams -Method Delete  -Uri "$uri"
                throw "Failure during attachment upload"
                return
            }
        }
        else {
            $msgSettings["attachments"] += $Filesettings
        }
    }

    if ($SaveDraftOnly) {
            Write-Progress -Activity "Sending Message" -Completed
            return $msg.webLink
    }
    elseif ($asDraft) {
            Write-Progress -Activity "Sending Message" -CurrentOperation "Sending Draft"
            try {
                $msg =  Invoke-WebRequest @webParams -Method post  -uri "$uri/send"
                write-verbose -Message ($msg.StatusCode + "  " + $msg.StatusDescription)
                Write-Progress -Activity "Sending Message" -Completed
            }
            catch {throw "There was an error sending the draft message; it remains in the drafts folder"}
    }
    else {
        $mail = @{Message=$msgSettings}
        if ($NoSave) {
           $mail['saveToSentItems'] = $false
        }
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading and sending"

        $json = ConvertTo-Json $mail -Depth 10
        Write-Verbose $json
        try            {$msg  = Invoke-RestMethod @webParams -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch          {throw "There was an error sending message."; return }
        write-verbose  -Message ($msg.StatusCode + "  " + $msg.StatusDescription)
        Write-Progress -Activity "Sending Message" -Completed
    }
}

function Send-GraphMailForward   {
    <#
      .synopsis
        Forwards a mail message.
      .example
      >
      > $alex = New-Recipient Alex@contoso.com -DisplayName "Alex B."
      > Get-GraphMailItem -top 1 | Send-GraphMailForward -to $Alex -Comment "FYI :-)"
      Creates a recipient , and forwards the top mail in the users inbox to that recipent
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param (
        #Either a message ID or a Message object with an ID.
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        $Message,
        #Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.
        [parameter(Mandatory=$true,Position=1)]
        $To ,
        #Comment to attach when forwarding the message.
        $Comment
    )
    $msgSettings   =  @{     toRecipients = @() }
    foreach ($recip in $To ) {
        if     ($recip  -is [string] ) { $msgSettings[ 'toRecipients'] += New-Recipient $recip}
        else                           { $msgSettings[ 'toRecipients'] += $recip}
    }
    if ($Comment)                      { $msgSettings[ 'comment'] = $Comment}
    if ($Message.id) {$uri = "https://graph.microsoft.com/v1.0/me/Messages/$($Message.id)/forward"}
    else             {$uri = "https://graph.microsoft.com/v1.0/me/Messages/$Message/forward"}

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Verbose $Json
    Invoke-RestMethod -Method post -Uri $uri -ContentType 'application/json' -Body $json -Headers $script:DefaultHeader
}

function Send-GraphMailReply     {
    <#
      .synopsis
        Replies to a mail message.
    #>
    [Cmdletbinding(DefaultParameterSetName='None')]
    param (
        #Either a message ID or a Message object with an ID.
        [parameter(Mandatory=$true,Position=0,ValueFromPipeline)]
        $Message,
        #Comment to attach when repling to the message - blank replies aren't allowed.
        [parameter(Mandatory=$true,Position=1)]
        $Comment,
        #If specified changes reply mode from reply [to sender] to Reply-to-all
        [Alias('All')]
        [switch]$ReplyAll
    )
    $msgSettings =  @{'comment' = $Comment }
    if ($Message.id) {$uri =  "https://graph.microsoft.com/v1.0/me/Messages/$($Message.id)/"}
    else             {$uri =  "https://graph.microsoft.com/v1.0/me/Messages/$Message"}
    if ($ReplyAll)   {$uri += '/replyAll' }
    else             {$uri += '/reply' }

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Verbose $Json
    Invoke-RestMethod -Method post -Uri $uri -ContentType 'application/json' -Body $json -Headers $script:DefaultHeader
}

<#
  GET https://graph.microsoft.com/beta/me/mailFolders/inbox/messagerules
  GET https://graph.microsoft.com/beta/me/outlook/masterCategories  #Colours ...
  GET https://graph.microsoft.com/beta/me/findRooms                 #
  https://graph.microsoft.com/v1.0/me/messages('AAMkADA1M-zAAA=')/attachments('AAMkADA1M-CJKtzmnlcqVgqI=')/?$expand=microsoft.graph.itemattachment/item
  #>


<#
POST https://graph.microsoft.com/beta/me/messages/AAMkAGE1M88AADUv0uFAAA=/attachments
Content-type: application/json
Content-length: 319

{
    "@odata.type": "#microsoft.graph.referenceAttachment",
    "name": "Personal pictures",
    "sourceUrl": "https://contoso.com/personal/mario_contoso_net/Documents/Pics",
    "providerType": "oneDriveConsumer",
    "permission": "Edit",
    "isFolder": "True"
}
#>


#https://docs.microsoft.com/en-us/graph/api/resources/call?view=graph-rest-beta  calls in teams
#https://docs.microsoft.com/en-us/graph/api/resources/onlinemeeting?view=graph-rest-beta  on line meetings in teams