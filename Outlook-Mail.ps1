using namespace Microsoft.Graph.PowerShell.Models

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
                else {
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

    if ($asDraft) {$Uri = "$GraphUri/me/Messages"}
    else          {$Uri = "$GraphUri/me/sendmail"}

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
    if ($CC) {
        $msgSettings['ccRecipients']      = @()
        foreach ($recip in $cc ) {
            if     ($recip  -is [string] ) { $msgSettings[ 'ccRecipients'] += New-Recipient $recip}
            else                           { $msgSettings[ 'ccRecipients'] += $recip}}
    }
    if ($BCC) {
        $msgSettings['bccRecipients']      = @()
        foreach ($recip in $bcc ) {
            if     ($recip  -is [string] ) { $msgSettings['bccRecipients'] += New-Recipient $recip}
            else                           { $msgSettings['bccRecipients'] += $recip}}
    }
    if ($Receipt)                          { $msgSettings['isDeliveryReceiptRequested'] = $true }

    #If we are creating a draft, save it now; if sending-in-one be ready for attachments
    if     ($asDraft) {
        Write-Progress -Activity "Sending Message" -CurrentOperation "Uploading draft"
        $json = ConvertTo-Json $msgSettings -Depth 5 #default depth isn't enough !
        try            {$msg  = Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
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
                $null = Invoke-GraphRequest -Method post  -uri "$uri/attachments"  -Body (ConvertTo-Json $Filesettings) -ContentType "application/json" -ErrorAction Stop
            }
            catch {
                Write-warning -Message "Error occured uploading file $($f.name) - will attempt to delete the draft message"
                Invoke-GraphRequest -Method Delete  -Uri "$uri"
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
                Invoke-GraphRequest  -Method post  -uri "$uri/send" -Body " " # underlying stuff requires -body, but server ignores it.
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
        Write-Debug $Json
        try            {Invoke-GraphRequest -Method post  -uri $uri  -Body $json -ContentType "application/json" }
        catch          {throw "There was an error sending message."; return }
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
    if ($Message.id) {$uri = "$GraphUri/me/Messages/$($Message.id)/forward"}
    else             {$uri = "$GraphUri/me/Messages/$Message/forward"}

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Debug $Json
    Invoke-GraphRequest -Method post -Uri $uri -ContentType 'application/json' -Body $json
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
    if ($Message.id) {$uri =  "$GraphUri/me/Messages/$($Message.id)/"}
    else             {$uri =  "$GraphUri/me/Messages/$Message"}
    if ($ReplyAll)   {$uri += '/replyAll' }
    else             {$uri += '/reply' }

    $json = ConvertTo-Json $msgSettings -depth 10
    Write-Debug $Json
    Invoke-GraphMethod -Method post -Uri $uri -ContentType 'application/json' -Body $json
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