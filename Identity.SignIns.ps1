using namespace Microsoft.Graph.PowerShell.Models
#MicrosoftGraphInvitation object is in Microsoft.Graph.Identity.SignIns.private.dll
function New-GraphInvitation   {
    <#
        .synopsis
            Invites an external user to become a guest in Azure AD
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        #The email address of the user being invited.
        #The characters  #~ ! $ %  ^  & * ( [ { < > } ] ) +  = \ /  | ; : " " ? , are not permitted
        #A  . or - is permitted except at the beginning or end of the name. A _  is permitted anywhere.
        [Parameter(Position=1,ValueFromPipeline=$true)]
        [string]$EmailAddress,
        #The display name of the user being invited.
        [string]$DisplayName,
        #The userType of the user being invited. By default, this is Guest. You can invite as Member if you are a company administrator.'
        [string]$UserType,
        #The URL the user should be redirected to once the invitation is redeemed. Required.
        [string]$RedirectUrl  = 'https://mysignins.microsoft.com/',
        #Indicates whether an email should be sent to the user being invited or not.
        [switch]$SendInvitationMessage
    )

    ContextHas -WorkOrSchoolAccount -BreakIfNot
    $settings = @{
        'invitedUserEmailAddress'    = $EmailAddress
        'sendInvitationMessage'      = $SendInvitationMessage -as [bool]
        'inviteRedirectUrl'          = $RedirectUrl
    }
    if ($DisplayName) {$settings['invitedUserDisplayName']  = $DisplayName}
    if ($UserType)    {$settings['invitedUserType']         = $UserType}

    $webparams = @{
        'Method'            = 'POST'
        'Uri'               = "$GraphUri/invitations"
        'Contenttype'       = 'application/json'
        'Body'              = (ConvertTo-Json $settings -Depth 5)
        'AsType'            = [MicrosoftGraphInvitation]
        'ExcludeProperty'   = '@odata.context'
    }
    Write-Debug $webparams.Body
    if ($force -or $pscmdlet.ShouldProcess($EmailAddress, 'Invite User')){
        try {
            $u = Invoke-GraphRequest @webparams
            if ($Passthru ) {return $u }
        }
        catch {
        # xxxx Todo figure out what errors need to be handled (illegal name, duplicate user)
        $_
        }
    }
}
