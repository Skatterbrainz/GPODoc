---
external help file: GPODoc-help.xml
Module Name: GPOdoc
online version: 
schema: 2.0.0
---

# Export-GPOComments

## SYNOPSIS
Compile HTML Report of GPO comments

## SYNTAX

```
Export-GPOComments [-GPOName] <String[]> [-ReportFile] <String> [[-StyleSheet] <String>]
```

## DESCRIPTION
Compile an HTML report of comments embedded within GPOs, GPO Settings
and GP Preferences settings

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Export-GPOComments -GPOName '*' -ReportFile ".\gpo.htm"
```

### -------------------------- EXAMPLE 2 --------------------------
```
$GpoNames | Export-GPOComments -ReportFile ".\gpo.htm"
```

### -------------------------- EXAMPLE 3 --------------------------
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
Path and name of CSS template file

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

## INPUTS

## OUTPUTS

## NOTES
1.1.0 - 11/14/2017 - David Stein

## RELATED LINKS

