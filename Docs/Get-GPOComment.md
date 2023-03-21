---
external help file: GPODoc-help.xml
Module Name: GPODoc
online version: https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOComment.md
schema: 2.0.0
---

# Get-GPOComment

## SYNOPSIS
Retrieve comments/descriptions embedded in GPOs

## SYNTAX

```
Get-GPOComment [-GPOName] <String[]> [-PolicyGroup] <String> [-ShowInfo] [<CommonParameters>]
```

## DESCRIPTION
Retrieve comments and descriptions embedded in GPOs, GPO Settings and GP Preferences Settings

## EXAMPLES

### EXAMPLE 1
```
Get-GPOComment -GPOName '*' -PolicyGroup 'Policy'
```

### EXAMPLE 2
```
$GpoNames | Get-GPOComment -PolicyGroup 'Policy'
```

### EXAMPLE 3
```
Get-GPOComment -GPOName '*' -PolicyGroup 'Settings'
```

### EXAMPLE 4
```
Get-GPOComment -GPOName '*' -PolicyGroup 'Preferences'
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

### -PolicyGroup
What aspects of each GPO is to be queried
List = Policy, Settings, Preferences

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

### -ShowInfo
{{ Fill ShowInfo Description }}

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
version 1.2.0 - 7/6/2022

## RELATED LINKS

[https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOComment.md](https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOComment.md)

