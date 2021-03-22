
function New-GraphPhysicalAddress {
    <#
      .synopsis
        Builds a street / postal / physical address to use in the contact commands
       .Example
        >$fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
        Creates an address - if the -Street, City,  State, Postalcode country are not explictly
        specified they will be assigned in that order. Quotes are desireable but only necessary
        when a value contains spaces.
        It can then be used like this. Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
    #>
    [cmdletbinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Does not change system state.')]
    param   (
        #Street address. This can contain carriage returns for a district, e.g. "101 London Road`r`nBotley"
        [String]$Street,
        #City, or town as people outside the US tend to call it
        [String]$City,
        #State, Province, County, the administrative level below country
        [String]$State,
        #Postal code. Even it parses as a number, as with US ZIP codes, it will be converted to a string
        [String]$PostalCode,
        #Usually a country but could be some other geographical entity
        [String]$CountryOrRegion
    )
    $Address = @{}
    foreach ($P in $PSBoundParameters.Keys.Where({$_ -notin [cmdlet]::CommonParameters})) {
        $Address[$p] + $PSBoundParameters[$p]
    }
    $Address
}

function Get-GraphContact         {
    <#
      .Synopsis
        Get the user's contacts
      .Example
        get-graphContact -name "o'neill" | ft displayname, mobilephone
        Gets contacts where the display name, given name, surname, file-as name, or email begins with
        O'Neill - note the function handles apostrophe, - a single one would normal cause an error with the query.
        The results are displayed as table with display name and mobile number
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param   (
        #UserID as a guid or User Principal name. If not specified defaults to "me"
        [string]$User,
        #A custom set of contact properties to select
        [ValidateSet('assistantName', 'birthday', 'businessAddress', 'businessHomePage', 'businessPhones',
                     'categories', 'changeKey', 'children', 'companyName', 'createdDateTime', 'department',
                     'displayName', 'emailAddresses', 'fileAs', 'generation', 'givenName', 'homeAddress',
                     'homePhones', 'id', 'imAddresses', 'initials', 'jobTitle', 'lastModifiedDateTime',
                     'manager', 'middleName', 'mobilePhone',  'nickName', 'officeLocation', 'otherAddress',
                     'parentFolderId', 'personalNotes', 'profession', 'spouseName', 'surname', 'title',
                     'yomiCompanyName', 'yomiGivenName', 'yomiSurname')]
        [string[]]$Select,

        #If specified looks for contacts where the display name, file-as Name, given name or surname beging with ...
        [Parameter(Mandatory=$true, ParameterSetName='FilterByName')]
        [string]$Name,
        #A custom OData Filter String
        [Parameter(Mandatory=$true, ParameterSetName='FilterByString')]
        [string]$Filter
    )

    #region build the URI - if we got a user ID, use it, add select, filter, orderby and/or top as needed
    if     ($User.id) {$uri = "$GraphUri/users/$($User.id)/contacts?`$top=100"}
    elseif ($User)    {$uri = "$GraphUri/users/$User/contacts?`$top=100" }
    else              {$uri = "$GraphUri/me/contacts?`$top=100" }

    if ($Select)      {$uri = $uri + '&$select=' + ($Select -join ',') }
    if ($Name)        {$uri = $uri + '&$filter=' + (FilterString $Name -ExtraFields 'companyname','givenName','surname') }
    if ($Filter)      {$uri = $uri + '&$Filter='  + $Filter   }
    #endregion

    Invoke-GraphRequest -Uri  $uri -ValueOnly -AllValues -AsType ([MicrosoftGraphContact]) -ExcludeProperty "@odata.etag"
}

function New-GraphContact         {
    <#
      .Synopsis
        Adds an entry to the current users Outlook contacts
      .Description
        Almost all the paramters can be accepted form a piped object to make import easier.
       .Example
       >New-GraphContact -GivenName Pavel -Surname Bansky -Email pavelb@fabrikam.onmicrosoft.com -BusinessPhones  "+1 732 555 0102"
       Creates a new contact; if no displayname is given, one will be decided using given name and suranme;
       .Example
       >
       >$PavelMail = New-GraphRecipient -DisplayName "Pavel Bansky [Fabikam]" -Mail  pavelb@fabrikam.onmicrosoft.com
       >New-GraphContact -GivenName Pavel -Surname Bansky -Email $pavelmail  -BusinessPhones  "+1 732 555 0102"
        This creates the same contanct but sets up their email with a display name.
        New recipient creates a hash table
        @{'emailaddress' = @ {
                'name' = 'Pavel Bansky [Fabikam]'
                'address' = 'pavelb@fabrikam.onmicrosoft.com'
            }
        }
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param   (
        [Parameter(ValueFromPipelineByPropertyName)]
        $GivenName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $MiddleName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Initials ,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Surname,
        [Parameter(ValueFromPipelineByPropertyName)]
        $NickName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $FileAs,
        [Parameter(ValueFromPipelineByPropertyName)]
        $DisplayName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $CompanyName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $JobTitle,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Department,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Manager,
        #One or more mail addresses, as a single string with semi colons between addresses or as an array of strings or MailAddress objects created with New-GraphMailAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $Email,
        #One or more instant messaging addresses, as an array or as a single string with semi colons between addresses
        [Parameter(ValueFromPipelineByPropertyName)]
        $IM,
        #A single mobile phone number
        [Parameter(ValueFromPipelineByPropertyName)]
        $MobilePhone,
        #One or more Business phones either as an array or as single string with semi colons between numbers
        [Parameter(ValueFromPipelineByPropertyName)]
        $BusinessPhones,
        #One or more home phones either as an array or as single string with semi colons between numbers
        [Parameter(ValueFromPipelineByPropertyName)]
        $HomePhones,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $Homeaddress,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $BusinessAddress,
        #An address object created with  New-GraphPhysicalAddress
        [Parameter(ValueFromPipelineByPropertyName)]
        $OtherAddress,
        #One or more categories either as an array or as single string with semi colons between them.
        [Parameter(ValueFromPipelineByPropertyName)]
        $Categories,
        #The contact's Birthday as a date
        [Parameter(ValueFromPipelineByPropertyName)]
        [dateTime]$Birthday ,
        [Parameter(ValueFromPipelineByPropertyName)]
        $PersonalNotes,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Profession,
        [Parameter(ValueFromPipelineByPropertyName)]
        $AssistantName,
        [Parameter(ValueFromPipelineByPropertyName)]
        $Children,
        [Parameter(ValueFromPipelineByPropertyName)]
        $SpouseName,
        #If sepcified the contact will be created without prompting for confirmation. This is the default state but can change with the setting of confirmPreference.
        [Switch]$Force
    )

    process {
        Set-GraphContact @PSBoundParameters -IsNew
    }
}

function Set-GraphContact         {
    <#
      .Synopsis
        Modifies or adds an entry in the current users Outlook contacts
      .Example
        >
        > $pavel = Get-GraphContact -Name pavel
        > Set-GraphContact $pavel -CompanyName "Fabrikam" -Birthday "1974-07-22"
        The first line gets the Contact which was added in the 'New-GraphContact" example
        and the second adds Birthday and Company-name attributes to the contact.
       .Example
        >
        > $fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
        > Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
        This continues from the previous example, creating an address in the first line
        and adding it to the contact in the second.

    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])]
    param   (
    #The contact to be updated either as an ID or as contact object containing an ID.
    [Parameter(ValueFromPipeline=$true,ParameterSetName='UpdateContact',Mandatory=$true, Position=0 )]
    $Contact,
    #If specified, instead of providing a contact, instructs the command to create a contact instead of updating one.
    [Parameter(ParameterSetName='NewContact',Mandatory=$true )]
    [switch]$IsNew,
    [Parameter(ValueFromPipelineByPropertyName)]
    $GivenName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $MiddleName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Initials ,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Surname,
    [Parameter(ValueFromPipelineByPropertyName)]
    $NickName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $FileAs,
    #If not specified a display name will be generated, so updates without the display name may result in overwriting an existing one
    [Parameter(ValueFromPipelineByPropertyName)]
    $DisplayName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $CompanyName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $JobTitle,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Department,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Manager,
    #One or more mail addresses, as a single string with semi colons between addresses or as an array of strings or MailAddress objects created with New-GraphMailAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $Email,
    #One or more instant messaging addresses, as an array or as a single string with semi colons between addresses
    [Parameter(ValueFromPipelineByPropertyName)]
    $IM,
    #A single mobile phone number
    [Parameter(ValueFromPipelineByPropertyName)]
    $MobilePhone,
    #One or more Business phones either as an array or as single string with semi colons between numbers
    [Parameter(ValueFromPipelineByPropertyName)]
    $BusinessPhones,
    #One or more home phones either as an array or as single string with semi colons between numbers
    [Parameter(ValueFromPipelineByPropertyName)]
    $HomePhones,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $Homeaddress,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $BusinessAddress,
    #An address object created with  New-GraphPhysicalAddress
    [Parameter(ValueFromPipelineByPropertyName)]
    $OtherAddress,
    #One or more categories either as an array or as single string with semi colons between them.
    [Parameter(ValueFromPipelineByPropertyName)]
    $Categories,
    #The contact's Birthday as a date
    [Parameter(ValueFromPipelineByPropertyName)]
    [nullable[dateTime]]$Birthday ,
    [Parameter(ValueFromPipelineByPropertyName)]
    $PersonalNotes,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Profession,
    [Parameter(ValueFromPipelineByPropertyName)]
    $AssistantName,
    [Parameter(ValueFromPipelineByPropertyName)]
    $Children,
    [Parameter(ValueFromPipelineByPropertyName)]
    $SpouseName,
    #If sepcified the contact will be created without prompting for confirmation. This is the default state but can change with the setting of confirmPreference.
    [Switch]$Force
    )
    begin   {
        $webParams = @{
            'ContentType'    = 'application/json'
            'URI'             = "$GraphUri/me/contacts"
            'AsType'          =  ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphContact])
            'ExcludeProperty' = @('@odata.etag', '@odata.context' )
        }
    }
    process {
        $contactSettings = @{  }
        if ($Email)                           {$contactSettings['emailAddresses'] = @() }
        if ($Email -is [string])              {$Email = $Email -split '\s*;\s*'}
        foreach ($e in $Email) {
            if     ($e.emailAddress)          {$contactSettings.emailAddresses    += $e.emailAddress   }
            elseif ($e -is [string])          {$contactSettings.emailAddresses    += @{'address' = $e} }
            else                              {$contactSettings.emailAddresses    += $e  }
        }
        if     ($IM             -is [string]) {$contactSettings['imAddresses']     = @() + $IM             -split '\s*;\s*'}
        elseif ($IM                         ) {$contactSettings['imAddresses']     =       $IM}
        if     ($Categories     -is [string]) {$contactSettings['categories']      = @() + $Categories     -split '\s*;\s*'}
        elseif ($Categories                 ) {$contactSettings['categories']      =       $Categories}
        if     ($Children       -is [string]) {$contactSettings['children']        = @() + $Children       -split '\s*;\s*'}
        elseif ($Children                   ) {$contactSettings['children']        =       $Children}
        if     ($BusinessPhones -is [string]) {$contactSettings['businessPhones']  = @() + $BusinessPhones -split '\s*;\s*'}
        elseif ($BusinessPhones             ) {$contactSettings['businessPhones']  =       $BusinessPhones}
        if     ($HomePhones     -is [string]) {$contactSettings['homePhones']      = @() + $HomePhones     -split '\s*;\s*'}
        elseif ($HomePhones                 ) {$contactSettings['homePhones']      =       $HomePhones  }
        if     ($MobilePhone                ) {$contactSettings['mobilePhone']     =       $MobilePhone}
        if     ($GivenName                  ) {$contactSettings['givenName']       =       $GivenName}
        if     ($MiddleName                 ) {$contactSettings['middleName']      =       $MiddleName}
        if     ($Initials                   ) {$contactSettings['initials']        =       $Initials}
        if     ($Surname                    ) {$contactSettings['surname']         =       $Surname}
        if     ($NickName                   ) {$contactSettings['nickName']        =       $NickName}
        if     ($FileAs                     ) {$contactSettings['fileAs']          =       $FileAs}
        if     ($DisplayName                ) {$contactSettings['displayName']     =       $DisplayName}
        if     ($Manager                    ) {$contactSettings['manager']         =       $Manager}
        if     ($JobTitle                   ) {$contactSettings['jobTitle']        =       $JobTitle}
        if     ($Department                 ) {$contactSettings['department']      =       $Department}
        if     ($CompanyName                ) {$contactSettings['companyName']      =      $CompanyName}
        if     ($PersonalNotes              ) {$contactSettings['personalNotes']   =       $PersonalNotes}
        if     ($Profession                 ) {$contactSettings['profession']      =       $Profession}
        if     ($AssistantName              ) {$contactSettings['assistantName']   =       $AssistantName}
        if     ($Children                   ) {$contactSettings['children']        =       $Children}
        if     ($SpouseName                 ) {$contactSettings['spouseName']      =       $spouseName}
        if     ($Homeaddress                ) {$contactSettings['homeaddress']     =       $Homeaddress}
        if     ($BusinessAddress            ) {$contactSettings['businessAddress'] =       $BusinessAddress}
        if     ($OtherAddress               ) {$contactSettings['otherAddress']    =       $OtherAddress}
        if     ($Birthday                   ) {$contactSettings['birthday']        =       $Birthday.tostring('yyyy-MM-dd')} #note this is a different date format to most things !

        $webParams['body'] = ConvertTo-Json $contactSettings
        Write-Debug $webParams.body
        if ($IsNew) {
            if ($force -or $PSCmdlet.ShouldProcess($DisplayName,'Create Contact')) {
                Invoke-GraphRequest @webParams -method Post
            }
        }
        else {#if Contact Passed
            if ($force -or $PSCmdlet.ShouldProcess($Contact.DisplayName,'Update Contact')) {
                if ($Contact.id)   {$webParams.uri += '/' + $Contact.ID}
                else               {$webParams.uri += '/' + $Contact }
                Invoke-GraphRequest @webParams -method Patch
            }
        }
    }
}

function Remove-GraphContact      {
    <#
      .synopsis
         Deletes a contact from the default user's contacts
      .Example
        > Get-GraphContact -Name pavel | Remove-GraphContact
        Finds and removes any user whose given name, surname, email or display name
        matches Pavel*. This might return unexpected users, fortunately there is a prompt
        before deleting - the prompt it can be supressed by using the -Force switch if you
        are confident you have the right contact selected.
    #>
    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param   (
        #The contact to remove, as an ID or as a contact object containing an ID
        [parameter(Position=0,ValueFromPipeline=$true,Mandatory=$true )]
        $Contact,
        #If specified the contact will be removed without prompting for confirmation
        $Force
    )
    process {
        if ($force -or $pscmdlet.ShouldProcess($Contact.DisplayName, 'Delete contact')) {
            if ($Contact.id) {$Contact = $Contact.id}
            Invoke-GraphRequest -Method Delete -uri "$GraphUri/me/contacts/$Contact"
        }
    }
}
