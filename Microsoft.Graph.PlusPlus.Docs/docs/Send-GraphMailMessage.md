---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Send-GraphMailMessage

## SYNOPSIS
Sends Mail using the Graph API from the current user's mailbox.

## SYNTAX

### None (Default)
```
Send-GraphMailMessage [-To] <Object> [-CC <Object>] [-BCC <Object>] [-Subject <String>] [-Body <String>]
 [-BodyType <Object>] [-Importance <Object>] [-Attachments <Object>] [-Receipt] [<CommonParameters>]
```

### SaveDraftOnly
```
Send-GraphMailMessage [-To] <Object> [-CC <Object>] [-BCC <Object>] [-Subject <String>] [-Body <String>]
 [-BodyType <Object>] [-Importance <Object>] [-Attachments <Object>] [-Receipt] [-SaveDraftOnly]
 [<CommonParameters>]
```

### NoSave
```
Send-GraphMailMessage [-To] <Object> [-CC <Object>] [-BCC <Object>] [-Subject <String>] [-Body <String>]
 [-BodyType <Object>] [-Importance <Object>] [-Attachments <Object>] [-Receipt] [-NoSave] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Send-GraphMail -To "chris@contoso.com" -subject "You left your keys behind[nt]"
Sends a mail with a subject but no body or attachments
```

### EXAMPLE 2
```
Send-GraphMail -To "chris@contoso.com" -body "Keys are with reception" -NoSave
Sends a mail but thi time the subject will read "No subject" and the test will be in the body.
-NoSave means that no copy of this message will be kept in sent items
```

### EXAMPLE 3
```
Send-GraphMail -To "chris@contoso.com" -Subject "Screen shot" -body "How does this look ?" -Attachments .\Logon.png -Receipt
#This message has an attachement and requests a read receipt.
```

### EXAMPLE 4
```
$body"<h1>New dialog</h1><br /><img src='cid:Logon.png' -alt='Look at that'><br/>what do you think"
>$link = Send-GraphMail -To "jhoneill@waitrose.com" -Subject "Login Sreen" -body $body -BodyType HTML  -NoSave -Attachments .\Logon.png -SaveDraftOnly
This creates an HTML body, the attached picture can be referenced in an <img> tag with cid:fileName.ext
this time the mail is not sent but left in the user's drafts folder for review.
```

## PARAMETERS

### -To
Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CC
Recipient(s) on the "CC" line,

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

### -BCC
Recipient(s) on the "Bcc line" line,

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

### -Subject
The subject of the message.
A message must have a subject and/or body and/or attachments.
If the subject is left blank it will be sent as "No Subject"

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

### -Body
The content of the message; assumed to be plain text, but HTML can be specified with -BodyType

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

### -BodyType
The type of the body  content.
Possible values are Text and HTML.

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
The importance of the message: Low, Normal or High

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Normal
Accept pipeline input: False
Accept wildcard characters: False
```

### -Attachments
Path to file(s) to send as attachments

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

### -Receipt
If Specified, requests a receipt.

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

### -SaveDraftOnly
If specified leaves the message in the drafts folder without sending it and returns a link to open the message.

```yaml
Type: SwitchParameter
Parameter Sets: SaveDraftOnly
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSave
If specified specifies that a copy of the mail should not be saved

```yaml
Type: SwitchParameter
Parameter Sets: NoSave
Aliases:

Required: True
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
