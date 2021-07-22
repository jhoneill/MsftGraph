---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Send-GraphMailForward

## SYNOPSIS
Forwards a mail message.

## SYNTAX

```
Send-GraphMailForward [-Message] <Object> [-To] <Object> [-Comment <Object>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
> $alex = New-GraphRecipient Alex@contoso.com -DisplayName "Alex B."
> Get-GraphMailItem -top 1 | Send-GraphMailForward -to $Alex -Comment "FYI :-)"
Creates a recipient , and forwards the top mail in the users inbox to that recipent
```

## PARAMETERS

### -Message
Either a message ID or a Message object with an ID.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -To
Recipient(s) on the "to" line, each is either created with New-MailRecipient (a hash table), or a string holding an address.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Comment
Comment to attach when forwarding the message.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
