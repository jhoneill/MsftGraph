[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='New*', Justification='New- Commands create definitions but do not change system state')]
Param()
Function Convert-ColorToHex     {
    <#
      .Synopsis
        Turns a [System.Drawing.Color] into a hex string e.g. Red to ff0000
    #>
    [cmdletbinding()]
    [outputType([String])]
    Param (
        [Alias("Colour")]
        [System.Drawing.Color]$Color
    )
    #return 2 hex digits for red, green and blue components
    "{0:x2}{1:x2}{2:x2}" -f $Color.r, $Color.g, $Color.b
}

Function New-CardImage          {
    [cmdletbinding()]
    [Alias('CardImage')]
    [OutPuttype([Hashtable])]
    Param (
        [parameter(Position=0, Mandatory=$true)]
        [alias('URI')]
        [string]$ImageURI,
        [string]$Title = 'Image'
    )
    @{image = 'uri' ; title='title'}
}

Function New-CardSection        {
    [CmdletBinding()]
    [alias('cardsection')]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    Param (
        [String]$Title            = '' ,
        [String]$ActivityImage    = '' ,
        [String]$ActivityTitle    = '' ,
        [String]$ActivitySubtitle = '' ,
        [String]$ActivityText     = '' ,
        [bool]$StartGroup         = $false ,
        $Images                  ,# = @( @{image = 'uri1' ; title='title1'}, @{image = 'uri' ; title='title'}) ,
        $Facts                   ,# = @( @{name1='value'}, @{name2='value2'}) ,
        $HeroimageURI             = 'HeroImageuri'  ,
        $HeroimageTitle           = 'HeroImagetitle'
    )
    $section = [ordered]@{
                    title                = $Title
                    startGroup           = $StartGroup
                    activityImage        = $ActivityImage
                    activityTitle        = $ActivityTitle
                    activitySubtitle     = $ActivitySubtitle
                    activityText         = $ActivityText
                    text                 = $Text
    }
    if ($HeroimageURI) {
                   $heroimage            = New-CardImage  -ImageURI $HeroimageURI -title $HeroimageTitle
                   $section['heroImage'] = $heroimage
    }
    if ($Facts)  { $section['facts']     = $Facts }
    if ($Images) { $section['images']    = $Images}
    return   $section
}

Function New-CardInput          {
    <#
      .synopsis
        Creates the UI imput controls for message cards
      .Example
      > New-CardInput  -InputType MultiLineText -ID 'Feedback' -Title "Let us know what you think"
      #Creates as a multi-line text box with the the ID "feedback" and a title

    #>
    [CmdletBinding()]
    [alias('cardinput')]
    Param (
        [Parameter(ParameterSetName='SingleChoice',Mandatory=$true,Position=0)]
        [ValidateSet('Text','MultiLineText','DateOnly','DateTime')]
        [Alias('Type')]
        [string]$InputType,
        [Parameter(ParameterSetName='SingleChoice')]
        [int]$MaxLength,
        [Parameter(ParameterSetName='MultiChoice',Mandatory=$true)]
        [string[]]$Choices,
        [Parameter(ParameterSetName='MultiChoice')]
        [switch]$MultiSelect,
        [Parameter(ParameterSetName='SingleChoice',Mandatory=$true,Position=1)]
        [Parameter(ParameterSetName='MultiChoice',Mandatory=$true,Position=0)]
        [string]$ID = 'feedback' ,
        [string]$Title,
        [string]$DefaultValue,
        [switch]$IsRequired
    )
    if ($InputType -in ('Text','MultilineText')) {
        $InputControl = [ordered]@{
            '@type'       = 'TextInput'
            'id'          = $ID
            'isRequired'  = [bool]$IsRequired
            'title'       = $Title
        }
        if ($InputType -eq 'MultilineText') { $InputControl['isMultiline'] = $true }
        if ($MaxLength) {                     $InputControl['maxLength']   = $MaxLength }
    }
    elseif ($InputType -in ('DateTime','Date') ) {
        $InputControl = [ordered]@{
            '@type'       = 'DateInput'
            'id'          = $ID
            'isRequired'  = [bool]$IsRequired
            'title'       = $Title
        }
        if ($InputType -eq 'DateTime') {InputControl['includeTime'] = $true}
    }
    elseif ($Choices) {
        $InputControl = [ordered]@{
            '@type'       = 'MultichoiceInput'
            'id'          = $ID
            'isRequired'  = [bool]$IsRequired
            'title'       = $Title
            'choices'     = @()
        }
        if ($MultiSelect) {$InputControl['isMultiSelect'] = $true}
        foreach ($c in $choices) {
                if ($c -is [string]) {
                    $InputControl['choices'] += @{'display' = $c ; 'value' = $c}
                }
                elseif ($c.display -and $c.value) {
                    $InputControl['choices'] += @{'display' = $c.display ; 'value' = $c.value}
                }
                else {throw 'Invalid value for a choice.'; return}
        }
    }
    if ($DefaultValue) { $InputControl['value'] = $DefaultValue }
    return $InputControl
}

Function New-CardActionHttpPost {
    <#
      .Synopsis
        Creates an Http Post action for a message card
      .Description
        See https://docs.microsoft.com/en-gb/outlook/actionable-messages/message-card-reference#httppost-action
     .Example
        New-CardActionHttpPost -Name 'Send Feedback' -Target 'http://feedback.contoso.com' -Primary
        Creates a button for a form which contains which will post to the URL.
          need more params to be useful!
    #>
    [cmdletbinding()]
    [Alias("HttpPostAction")]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        $Name ,
        [Parameter(Mandatory=$true,Position=1)]
        $Target ,
        [string[]]$Headers  ,
        $Body ,
        [ValidateSet('application/json', 'application/x-www-form-urlencoded')]
        $ContentType,
        [switch]$Primary
    )
    $action = [ordered]@{
        '@type'    = 'HttpPOST'
        'name'     = $Name
        'target'   = $target
    }
    if ($Primary)     {$action['isPrimary']       = $true}
    if ($Headers)     {$action['headers']         = $Headers}
    if ($Body)        {$action['body']            = $Body}
    if ($ContentType) {$action['bodyContentType'] = $ContentType}
    return $action
}

Function New-CardActionOpenUri  {
    <#
      .Synopsis
         Creates a button on a message card to open a link
      .Example
          New-CardActionOpenUri  -Name 'Learn more' -Targets  'https://docs.microsoft.com/outlook/actionable-messages'
          This creates an action which a appers as button on the card [Learn More] clicking it opens the link
    #>
    [cmdletbinding()]
    [Alias("OpenUriAction")]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        $Name ,
        [Parameter(Mandatory=$true,Position=1)]
        $Targets ,
        [switch]$Primary
    )
    $Action  = [ordered]@{
        '@type'    = 'OpenUri'
        'name'     = $Name
        'targets'  = @()
    }
    if ($Primary) {$Action['isPrimary']= $true}
    foreach ($t in $Targets) {
                if ($t -is [string]) {
                    $Action['targets'] += @{'os' = 'default' ; 'uri' = $t}
                }
                elseif ($t.os -and $t.uri) {
                    $Action['targets'] += @{'os' = $t.os ; 'uri' = $t.uri}
                }
                else {throw 'Invalid value for a target.'; return}

    }
    return $action
}

Function New-CardActionCard     {
    <#
      .Synopsis
        Creates an "Action card" action for a message card
      .Description
        Actions are presented on the card as buttons the user can click
        For an "action card" action this reveals a 'sub-card' with input controls
        and action buttons. The buttons must be Http posts or Open URI types, they
        can't be nest action cards.
      .Example
      >
      >$inputs     = @()
      >$inputs    += New-CardInput        ...
      >$actions    = New-CardActionHttpPost -Name 'Send Feedback' -Target 'http://....'
      >$actioncard = New-CardActionCard     -Name 'Send Feedback' -Inputs $inputs -Actions $actions
      This example starts by creating the input fields for the Action card
      It the definies a single action  - the both these steps have been truncated for brevity
      The final step creates the action card. This will then be passed in the Actions parameter for New-Message card.
      The post action will usually need to send a some of the input in the body of the post
      see under 'Input value substitution' in https://docs.microsoft.com/en-gb/outlook/actionable-messages/message-card-reference
    #>
    [cmdletbinding()]
    [Alias("CardAction")]
    param(
        [Parameter(Mandatory=$true,Position=0)]
        $Name ,
        $Inputs,
        $Actions
      )
      [ordered]@{
        '@type'    = 'ActionCard'
        'name'     = $Name
        'inputs'   = @() + $Inputs
        'actions'  = @() + $Actions
    }
}

Function New-MessageCard        {
    <#
      .Synopsis
        Creates a message card and either posts to a webhook or returns it for examination/tweaking.
      .Example
        >
        >new-messageCard  -WebHookURI $hookUri   -Title 'From powershell to teams using webhooks' -Text @'
            ![Logo](https://cdn.vsassets.io/content/notifications/teams-build-succeeded.png)**James** did a _crane job_ on the logo!
        '@
        Creates a simple card with a title and text using mark down to display a logo.
        This card is posted to the Webhook in $webhooURI.
      .Description
        See https://docs.microsoft.com/en-gb/outlook/actionable-messages/message-card-reference
    #>
    [cmdletbinding()]
    [Alias("MessageCard")]
    param(
            #The title for the card. The font for this is fixed.
            [string]$Title      = 'Visit the Outlook Dev Portal' ,
            #The body text for the card. This supports markdown, which you can use to include images.
            [String]$Text       = '',
            #Short version of the text.
            [string]$Summary    = '',
            [system.drawing.color]$Themecolor,
            $ThemeColorHex        ,#  ='0072C6',
            $Actions  ,
            [switch]$AsHashTable,
            [String]$WebHookURI

    )
    #Build the card as a HashTable. Make it ordered when we connvert to JSON the items don't look in a strange order
    $CardSettings = [ordered]@{
        '@context'       = 'https://schema.org/extensions'
        '@type'          = 'MessageCard'
        'title'          = $Title
        'text'           = $Text
        'summary'        = $Summary
    }
    #Add optional items color is a hex string but we'll take and convert system colors.
    if     ($ThemeColorHex) {$cardSettings['themeColor']      = $ThemeColorHex}
    elseif ($Themecolor)    {$cardSettings['themeColor']      = Convert-ColorToHex -Color $Themecolor}
    if     ($Actions)       {$cardSettings['potentialAction'] = @() + $Actions}

    #And either post to a webhook or return the results.
    if     ($AsHashTable) {return $cardSettings}
    elseif ($WebHookURI)  {
        Invoke-RestMethod -Method Post -Uri $WebHookURI -ContentType "application/json" -Body (ConvertTo-Json $cardSettings -Depth 99)
    }
    else {ConvertTo-Json $cardSettings -Depth 99}
}

############## Example ####################

# Your Web hook here
# $hookUri = 'https://outlook.office.com/webhook/<<teamGuid>>@<<org GUID>>/IncomingWebhook/<<id>>/<<creator guid>>'

<#
First we can use these with the commands being written out in max-verbose. and storing in variables
$inputs     = New-CardInput          -InputType MultiLineText -ID 'Feedback' -Title "Let us know what you think"
$actions    = New-CardActionHttpPost -Name 'Send Feedback' -Target 'http://....' -Primary
$actioncard = New-CardActionCard     -Name 'Send Feedback' -Inputs $inputs -Actions $actions
$link       = New-CardActionOpenUri  -Name 'Learn more'    -Targets  'https://docs.microsoft.com/outlook/actionable-messages'

$actions    = @($link,$actionCard)
$json       = New-MessageCard -Actions $actions -Title 'Visit the Outlook Dev Portal' -Text 'Click **Learn More** to learn more about Actionable Messages!'

#The same but using aliases and being terse with parameters.
$inputs     = Cardinput MultiLineText  feedback  -Title "Let us know what you think"
$actions    = HttpPostAction 'Send Feeback'  'http://....'  -Primary
$actioncard = CardAction     'Send Feedback' -Inputs $inputs -Actions $actions
$link       = OpenUriAction   'Learn more'   'https://docs.microsoft.com/outlook/actionable-messages'
$json       = MessageCard     'Visit the Outlook Dev Portal'  'Click **Learn More** to learn more about Actionable Messages!' -Actions $link,$actioncard

#And the same again but with attitude of "We don't need no stinking variables"
MessageCard -Title 'Visit the Outlook Dev Portal' -Text @'
![Logo](https://cdn.vsassets.io/content/notifications/teams-build-succeeded.png)**James** did a _crane job_ on the logo!
'@ -Actions @(
            (CardAction "Feedback" -Inputs (Cardinput MultiLineText  feedback  -Title "Let us know, what do you think?") `
                                        -Actions (HttpPostAction 'Send us feedback' 'http://....' -Primary) ) ,
            (OpenUriAction 'Learn more' 'https://docs.microsoft.com/outlook/actionable-messages')
        )      | clip # paste into https://messagecardplayground.azurewebsites.net/
#>