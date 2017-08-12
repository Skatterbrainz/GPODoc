# GPPDoc
Version 1.0.2 - 2017.08.11.  Group Policy and Group Policy Preferences Documentation Tools.
Available from the PowerShell Gallery. To import this module, use _Install-Module GPODoc_ and _Import-Module GPODoc_

## Get-GPOComment

* Retrieves Descriptions for GPOs, Comments for embedded GPO settings, and Comments for embedded Group Policy Preference settings.
* Examples:
  * Get-GPOComment -GPOName "*" -PolicyGroup Policy
  * Get-GPOComment -GPOName "GPO1","GPO2" -PolicyGroup Preferences

## Set-GPOComment
Coming soon

## Export-GPOCommentReport

* Generates an HTML report of comments embedded within GPOs, GPO Settings and GPO Preferences.
* Example: Export-GPOCommentReport -GPOName "*" -ReportFile ".\gpo.htm"
* Example: Get-GPO -All | ?{$_.DisplayName -like "User *"} | Export-GPOCommentReport -ReportFile ".\userGPOs.htm"
