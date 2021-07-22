---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphMailItem

## SYNOPSIS
Get items in a mail folder

## SYNTAX

### None (Default)
```
Get-GraphMailItem [[-Mailfolder] <Object>] [-User <String>] [-Unread] [-Subject <String>] [-From <String>]
 [-To <String>] [-HasAttachments] [-Important] [-Today] [-Yesterday] [-Before <DateTime>] [-After <DateTime>]
 [-Search <String>] [-Top <Int32>] [-OrderBy <String>] [-Select <String[]>] [<CommonParameters>]
```

### FilterByString
```
Get-GraphMailItem [[-Mailfolder] <Object>] [-User <String>] [-Unread] [-Subject <String>] [-From <String>]
 [-To <String>] [-HasAttachments] [-Important] [-Today] [-Yesterday] [-Before <DateTime>] [-After <DateTime>]
 [-Search <String>] [-Top <Int32>] [-OrderBy <String>] [-Select <String[]>] -Filter <String>
 [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphMailItem -top 5
Gets the top 5 items in the current users Inbox
```

### EXAMPLE 2
```
Get-GraphMailItem -Mailfolder "sentitems" -top 5
Gets the top 5 items in the current users sent items folder
```

### EXAMPLE 3
```
Get-GraphMailFolderList -Name sent | Get-GraphMailItem -top 5
This has the same result as before but could find any folder
```

### EXAMPLE 4
```
Get-GraphMailItem -Search 'criminal'
Searches the default folder (inbox) for 'Criminal' in any field
```

### EXAMPLE 5
```
Get-GraphMailItem -Search 'criminal' -Mailfolder ''
Searches for 'Criminal' in any field but this time searches the whole mailbox
```

### EXAMPLE 6
```
Get-GraphMailItem -Search 'subject:criminal' -Mailfolder ''
This time limits the search to just the subject line. from:, to: etc can be
used in the same way as they can in a search in outlook.
```

### EXAMPLE 7
```
Get-GraphMailItem -filter "from/emailAddress/address eq 'alex@contoso.com'"
Instead of a free text search this applies a filter on email address, looking at the inbox.
```

### EXAMPLE 8
```
Get-GraphMailItem -Filter "(hasattachments eq true) and startswith(from/emailAddress/name, 'alex')"
This shows a filter based on two conditions.
```

## PARAMETERS

### -Mailfolder
A folder objet or the ID of a folder, or one of the well known folder names 'archive', 'clutter', 'conflicts', 'conversationhistory', 'deleteditems', 'drafts', 'inbox', 'junkemail', 'localfailures', 'msgfolderroot', 'outbox', 'recoverableitemsdeletions', 'scheduled', 'searchfolders', 'sentitems', 'serverfailures', 'syncissues'

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Inbox
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -User
UserID as a guid or User Principal name, if it can't be discovered from the mailfolder.
If not specified defaults to "me"

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

### -Unread
Selects only unread mail (equivalent to isread:no in Outlook)

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

### -Subject
Searches based on the subject field (equivalent to subject: in Outlook)

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

### -From
Searches based on the from field (equivalent to from: in Outlook)

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

### -To
Searches based on the to field (equivalent to to: in Outlook)

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

### -HasAttachments
Selects only mail with attachments (equivalent to hasAttachments:yes in Outlook).
Note this does not combine well with date based searches

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

### -Important
Selects only mail marked as important (equivalent to importance:high in Outlook).

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

### -Today
Selects only mail from today (equivalent to received:today in Outlook).

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

### -Yesterday
Selects only mail from today (equivalent to received:yesterday in Outlook).

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

### -Before
Selects only mail from before a given date

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

### -After
Selects only mail from after a given date

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

### -Search
A term to do a free text search for in the mail box (see examples)

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

### -Top
If specified returns the top X items, defaults to 100

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 100
Accept pipeline input: False
Accept wildcard characters: False
```

### -OrderBy
Sorting option, defaults to sorting by SentDateTime with newest first.
Searches are not sorted.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: SentdateTime desc
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
Select particular mail fields , ignored if -ChildFolders is specified; defaults to From, Subject, SentDateTime, BodyPreview, and Weblink

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: @('From', 'Subject', 'SentDatetime', 'hasAttachments', 'BodyPreview', 'weblink')
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
A Custom filter string; for example "importance eq high" - the examples have more cases

```yaml
Type: String
Parameter Sets: FilterByString
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphMessage
## NOTES

## RELATED LINKS
