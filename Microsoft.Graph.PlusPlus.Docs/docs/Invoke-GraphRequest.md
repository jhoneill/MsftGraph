---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Invoke-GraphRequest

## SYNOPSIS
Wrappper for Invoke-MgGraphRequest.With token management and result pre-processing

## SYNTAX

```
Invoke-GraphRequest [-Uri] <Uri> [[-Method] <Object>] [[-Body] <Object>] [-Headers <IDictionary>]
 [-OutputFilePath <String>] [-InferOutputFileName] [-InputFilePath <String>] [-PassThru]
 [-Token <SecureString>] [-SkipHeaderValidation] [-ContentType <String>]
 [-Authentication <GraphRequestAuthenticationType>] [-SessionVariable <String>]
 [-ResponseHeadersVariable <String>] [-StatusCodeVariable <String>] [-SkipHttpErrorCheck] [-ValueOnly]
 [-AllValues] [-ExcludeProperty <String[]>] [-PropertyNotMatch <String>] [-AsType <String>]
 [<CommonParameters>]
```

## DESCRIPTION
Adds -ValueOnly to return just the value part
     -AllValues to return gather multiple sets when data is paged
     -AsType to convert the retuned results to a specific type
     -ExcludeProperty  and -PropertyNotMatch for results which have properties which aren't in the specified type

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Uri
Uri to call can be a segment such as /beta/me or a fully qualified https://graph.microsoft.com/beta/me

```yaml
Type: Uri
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Method
Http Method

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Body
Request body, required when Method is POST or PATCH

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Headers
Optional custom headers, commonly @{'ConsistencyLevel'='eventual'}

```yaml
Type: IDictionary
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFilePath
Output file where the response body will be saved

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

### -InferOutputFileName
{{ Fill InferOutputFileName Description }}

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

### -InputFilePath
Input file to send in the request

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

### -PassThru
Indicates that the cmdlet returns the results, in addition to writing them to a file.
Only valid when the OutFile parameter is also used.

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

### -Token
OAuth or Bearer token to use instead of acquired token

```yaml
Type: SecureString
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipHeaderValidation
Add headers to request header collection without validation

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

### -ContentType
Body content type, for exmaple 'application/json'

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

### -Authentication
Graph Authentication type - 'default' or 'userProvivedToken'

```yaml
Type: GraphRequestAuthenticationType
Parameter Sets: (All)
Aliases:
Accepted values: Default, UserProvidedToken

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SessionVariable
Specifies a web request session.
Enter the variable name, including the dollar sign ($).You can''t use the SessionVariable and GraphRequestSession parameters in the same command.

```yaml
Type: String
Parameter Sets: (All)
Aliases: SV

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResponseHeadersVariable
{{ Fill ResponseHeadersVariable Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases: RHV

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StatusCodeVariable
{{ Fill StatusCodeVariable Description }}

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

### -SkipHttpErrorCheck
{{ Fill SkipHttpErrorCheck Description }}

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

### -ValueOnly
If specified returns the .values property instead of the whole JSON object returned by the API call

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

### -AllValues
If specified, loops through multi-paged results indicated by an '@odata.nextLink' property

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

### -ExcludeProperty
If specified removes properties found in the JSON before converting to a type or returning the object

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: @()
Accept pipeline input: False
Accept wildcard characters: False
```

### -PropertyNotMatch
A regular expression for keys to be removed, for example to catch many odata properties

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

### -AsType
If specified converts the JSON object to properties of the a new object of the requested type.
Any properties which are expected in the JSON but not defined in the type should be excluded.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
