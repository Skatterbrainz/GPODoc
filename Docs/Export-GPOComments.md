---
external help file: GPODoc-help.xml
Module Name: GPODoc
online version: https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Export-GPOComments.md
schema: 2.0.0
---

# Export-GPOComments

## SYNOPSIS
Compile HTML Report of GPO comments

## SYNTAX

```
Export-GPOComments [-GPOName] <String[]> [-ReportFile] <String> [[-StyleSheet] <String>] [<CommonParameters>]
```

## DESCRIPTION
Compile an HTML report of comments embedded within GPOs, GPO Settings
and GP Preferences settings

## EXAMPLES

### EXAMPLE 1
```
Export-GPOComments -GPOName '*' -ReportFile ".\gpo.htm"
```

### EXAMPLE 2
```
$GpoNames | Export-GPOComments -ReportFile ".\gpo.htm"
```

### EXAMPLE 3
```
Export-GPOComments -ReportFile ".\gpo.htm" -StyleSheet ".\mystyles.css"
```

## PARAMETERS

### -GPOName
Name(s) of Group Policy Objects or '*' for all GPOs

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ReportFile
Path and name of new HTML report file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StyleSheet
Path and name of custom CSS template file (default is /GPODoc/assets/default.css)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
1.2.0 - 7/6/2022

## RELATED LINKS

[https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Export-GPOComments.md](https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Export-GPOComments.md)

