using namespace Microsoft.Graph.PowerShell.Models
#MicrosoftGraphMailTips object is isn Microsoft.Graph.Users.Actions.private.dll
param()

function Get-GraphMailTips       {
    <#
      .synopsis
        Gets mail tips for one or more users (is their mailbox full, are auto-replies on etc)
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="MailTip would be incorrect")]
    param   (
        #mail addresses
        [Parameter(Mandatory=$true)]
        [string[]]$Address
    )
    $webparams = @{
                    'Method'      = 'post'
                    'Uri'         = "$GraphUri/me/getMailTips"
                    'ContentType' = 'application/json'
                    'body'        = Convertto-Json @{EmailAddresses= @() + $Address;
                                                    MailTipsOptions= "automaticReplies, mailboxFullStatus, customMailTip, "+
                                                                      "deliveryRestriction, externalMemberCount, maxMessageSize, " +
                                                                      "moderationStatus, recipientScope, recipientSuggestions, totalMemberCount"
                    }
     }

    Invoke-GraphRequest @webparams -ValueOnly -AsType ([MicrosoftGraphMailTips])
}