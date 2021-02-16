using namespace Microsoft.Graph.PowerShell.Models
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='New*')]
param()
function Get-GraphMailTips       {
    <#
      .synopsis
        Gets mail tips for one or more users (is their mailbox full, are auto-replies on etc)
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification="MailTip would be incorrect")]
    param(
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

    Invoke-GraphRequest @webparams -ValueOnly -AsType ([MicrosoftGraphMailTips]) |
        Add-Member -PassThru -MemberType ScriptProperty -Name 'Address'          -Value {$this.EmailAddress.Address} |
        Add-Member -PassThru -MemberType ScriptProperty -Name 'MessageText'      -Value {$this.AutomaticReplies.Message} |
        Add-Member -PassThru -MemberType ScriptProperty -Name 'MessageStart'     -Value {$this.AutomaticReplies.scheduledStartTime.DateTime} |
        Add-Member -PassThru -MemberType ScriptProperty -Name 'MessageEnd'       -Value {$this.AutomaticReplies.scheduledEndTime.DateTime}
}