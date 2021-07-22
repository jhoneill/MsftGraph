---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphUser

## SYNOPSIS
Sets properties of  a user (the current user by default)

## SYNTAX

```
Set-GraphUser [[-UserID] <Object>] [-AboutMe <String>] [-Mail <String>] [-OtherMails <String[]>]
 [-MobilePhone <String>] [-BusinessPhones <String[]>] [-MySite <String>] [-UsageLocation <String>]
 [-DisplayName <String>] [-GivenName <String>] [-Surname <String>] [-JobTitle <String>] [-Department <String>]
 [-OfficeLocation <String>] [-CompanyName <Object>] [-Manager <String>] [-EmployeeID <String>]
 [-EmployeeType <String>] [-EmployeeHireDate <DateTime>] [-ExternalUserState <Object>]
 [-StreetAddress <Object>] [-City <Object>] [-State <Object>] [-Country <Object>] [-PostalCode <Object>]
 [-Birthday <DateTime>] [-Interests <String[]>] [-PastProjects <String[]>] [-Photo <String>]
 [-Responsibilities <String[]>] [-Schools <String[]>] [-Skills <String[]>] [-AccountDisabled] [-PassThru]
 [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Needs consent to use the User.ReadWrite, User.ReadWrite.All, Directory.ReadWrite.All,
or Directory.AccessAsUser.All scope.

## EXAMPLES

### EXAMPLE 1
```
Set-GraphUser -Birthday "31 march 1965"  -Aboutme "Lots to say" -PastProjects "Phoenix","Excalibur" -interests "Photography","F1" -Skills "PowerShell","Active Directory","Networking","Clustering","Excel","SQL","Devops","Server builds","Windows Server","Office 365" -Responsibilities "Design","Implementation","Audit"
Sets the current user, giving lists for projects, interests and skills
```

## PARAMETERS

### -UserID
ID for the user if not the current user

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Me
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -AboutMe
A freeform text entry field for the user to describe themselves.

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

### -Mail
The SMTP address for the user, for example, 'Alex@contoso.onmicrosoft.com'

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

### -OtherMails
A list of additional email addresses for the user; for example: \['bob@contoso.com', 'Robert@fabrikam.com'\].

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MobilePhone
User's mobile phone number

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

### -BusinessPhones
The telephone numbers for the user.
NOTE: Although this is a string collection, only one number can be set for this property

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MySite
Url for user's personal site.

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

### -UsageLocation
A two letter country code (ISO standard 3166).
Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries. 
Examples include: 'US', 'JP', and 'GB'

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

### -DisplayName
The name displayed in the address book for the user.
This is usually the combination of the user''s first name, middle initial and last name.
This property is required when a user is created and it cannot be cleared during updates.

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

### -GivenName
The given name (first name) of the user.

```yaml
Type: String
Parameter Sets: (All)
Aliases: FirstName

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Surname
User's last / family name

```yaml
Type: String
Parameter Sets: (All)
Aliases: LastName

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -JobTitle
The user's job title

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

### -Department
The name for the department in which the user works.

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

### -OfficeLocation
The office location in the user's place of business.

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

### -CompanyName
The company name which the user is associated.
This property can be useful for describing the company that an external user comes from.
The maximum length of the company name is 64 chararcters.

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

### -Manager
ID or UserPrincipalName of the user's manager

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

### -EmployeeID
The employee identifier assigned to the user by the organization

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

### -EmployeeType
Captures enterprise worker type: Employee, Contractor, Consultant, Vendor, etc.

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

### -EmployeeHireDate
The date and time when the user was hired or will start work in case of a future hire

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

### -ExternalUserState
For an external user invited to the tenant using the invitation API, this property represents the invited user's invitation status.
For invited users, the state can be PendingAcceptance or Accepted, or null for all other users.

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

### -StreetAddress
The street address of the user's place of business.

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

### -City
The city in which the user is located.

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

### -State
The state, province or county in the user's address.

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

### -Country
The country/region in which the user is located; for example, 'US' or 'UK'

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

### -PostalCode
The postal code for the user's postal address, specific to the user's country/region.
In the United States of America, this attribute contains the ZIP code.

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

### -Birthday
User's birthday as a date.
If passing a string it can be "March 31 1965", "31 March 1965", "1965/03/31" or  "3/31/1965" - this layout will always be read as US format.

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

### -Interests
List of user's interests

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PastProjects
List of user's past projects

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Photo
Path to a .jpg file holding the users photos

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

### -Responsibilities
List of user's responsibilities

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Schools
List of user's Schools

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Skills
List of user's skills

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AccountDisabled
Set to disable the user account, to re-enable an account use $AccountDisabled:$false

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

### -PassThru
If specified the modified user will be returned

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

### -Force
Supresses any confirmation prompt

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

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

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

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser
## NOTES

## RELATED LINKS
