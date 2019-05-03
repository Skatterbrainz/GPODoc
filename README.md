# GPPDoc

# Introduction

GPODoc is a Group Policy Object documenter and document query tool.  Basically, it can query for comments attached to GPO's directly, as well as comments attached to settings and preferences.  Results are returned via pipeline, so they can be filtered, sorted and manipulated as needed.

# History

Version 1.1.2 - 2018.01.03.  Group Policy and Group Policy Preferences Documentation Tools.
Version 1.2.0 - 2018.08.01.  Updated to support explicit AD domain references (cross-forest reporting)
Available from the PowerShell Gallery. To import this module, use _Install-Module GPODoc_ and _Import-Module GPODoc_

# Functions

## Get-GPOComment

* Retrieves Descriptions for GPOs, Comments for embedded GPO settings, and Comments for embedded Group Policy Preference settings.

## Export-GPOCommentReport

* Generates an HTML report of comments embedded within GPOs, GPO Settings and GPO Preferences.

Review markdown help documents under the /Docs folder for more details and examples.
