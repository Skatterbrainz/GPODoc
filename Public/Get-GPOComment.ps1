#requires -Modules GroupPolicy

function Get-GPOComment {
	<#
	.SYNOPSIS
		Retrieve comments/descriptions embedded in GPOs
	.DESCRIPTION
		Retrieve comments and descriptions embedded in GPOs, GPO Settings and GP Preferences Settings
	.PARAMETER GPOName
		Name(s) of Group Policy Objects or '*' for all GPOs
	.PARAMETER PolicyGroup
		What aspects of each GPO is to be queried
		List = Policy, Settings, Preferences
	.EXAMPLE
		Get-GPOComment -GPOName '*' -PolicyGroup 'Policy'
	.EXAMPLE
		$GpoNames | Get-GPOComment -PolicyGroup 'Policy'
	.EXAMPLE
		Get-GPOComment -GPOName '*' -PolicyGroup 'Settings'
	.EXAMPLE
		Get-GPOComment -GPOName '*' -PolicyGroup 'Preferences'
	.NOTES
		version 1.2.0 - 7/6/2022
	.LINK
		https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOComment.md
	#>
	[CmdletBinding()]
	param (
		[parameter(Mandatory = $True, ValueFromPipeline = $True, HelpMessage = 'Name of Policy or Policies')]
			[ValidateNotNullOrEmpty()]
			[string[]] $GPOName,
		[parameter(Mandatory = $True, HelpMessage = 'Policy group to query')]
			[ValidateSet('Policy', 'Settings', 'Preferences')]
			[string] $PolicyGroup,
		[parameter(Mandatory = $False)]
			[switch] $ShowInfo
	)
	if ($ShowInfo) {
		$ModuleData = Get-Module GPODoc
		$ModuleVer  = $ModuleData.Version -join '.'
		Write-Host "GPODoc $ModuleVer - https://github.com/Skatterbrainz/GPODoc" -ForegroundColor Cyan
	}
	switch ($PolicyGroup) {
		'Policy' {
			if ($GPOName -eq '*') {
				Write-Verbose "loading all policy objects"
				$gpos = Get-GPO -All | Sort-Object -Property DisplayName
			}
			else {
				Write-Verbose "loading specific policy objects"
				try {
					$gpos = $GPOName | Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select-Object Id, DisplayName, DomainName, Description}
				}
				catch {
					Write-Warning "Unable to load GPO"
					break
				}
			}
			foreach ($gpo in $gpos) {
				Write-Verbose "policy: $($gpo.DisplayName)"
				$data = [ordered]@{
					PolicyID    = $gpo.Id
					Name        = $gpo.DisplayName
					Description = $gpo.Description
				}
				New-Object -TypeName PSObject -Property $data
			}
			break
		}
		'Settings' {
			if ($GPOName -eq '*') {
				Write-Verbose "loading all policy objects"
				$gpos = Get-GPO -All | Sort-Object -Property DisplayName
			}
			else {
				Write-Verbose "loading specific policy objects"
				$gpos = $GPOName | 
					Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select-Object Id, DisplayName, DomainName, Description}
			}
			foreach ($gpo in $gpos) {
				$configs = @('Machine', 'User')
				Write-Verbose "policy: $($gpo.DisplayName)"
				foreach ($config in $configs) {
					$commFile = Get-GpoCommentFile -Gpo $gpo -ConfigType $config
					if ($commFile -ne "") {
						Write-Verbose "loading... $commFile"
						[xml]$cmtx = Get-Content -Path $commFile
						$clinks = $cmtx.policyComments.resources.stringTable.string
						foreach ($clink in $clinks) {
							$data = [ordered]@{
								PolicyID      = $gpo.Id
								Name          = $gpo.DisplayName
								Configuration = $config
								Setting       = $clink.Id
								Comment       = $clink.InnerText.Trim()
							}
							New-Object -TypeName PSObject -Property $data
						}
					}
					else {
						Write-Verbose "there are no $config comments in $($gpo.DisplayName)"
						$data = [ordered]@{
							PolicyID      = $gpo.ID
							Name          = $gpo.DisplayName
							Configuration = $config
							Setting       = $null
							Comment       = $null
						}
						New-Object -TypeName PSObject -Property $data
					}
				}
			}
			break
		}
		'Preferences' {
			if ($GPOName -eq '*') {
				Write-Verbose "loading all policy objects: preferences"
				$gpos = Get-GPO -All | Sort-Object -Property DisplayName
			}
			else {
				Write-Verbose "loading specific policy objects"
				$gpos = $GPOName | Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select-Object Id, DisplayName, DomainName, Description}
			}
			foreach ($gpo in $gpos) {
				$configs = @('Machine', 'User')
				Write-Verbose "policy: $($gpo.DisplayName)"
				foreach ($config in $configs) {
					$gppPath = Get-GppCommentPath -Gpo $gpo -ConfigType $config
					if ($gppPath -ne "") {
						Write-Verbose "there are preferences for this policy object"
						foreach ($section in Get-ChildItem -Path $gppPath -Directory | Select-Object -ExpandProperty Name) {
							$gppFile = "$gppPath\$section\$section.xml"
							if (Test-Path $gppFile) {
								Write-Verbose "gppref file: $gppFile"
								[xml]$gppXML = Get-Content $gppFile
								$Element = $section.substring(0, $section.Length - 1)
								foreach ($Element in $gppXML."$section"."$Element") {
									$data = [ordered]@{
										PolicyID      = $gpo.ID
										Name          = $gpo.DisplayName
										Configuration = $config
										Section       = $section
										Element       = $Element.name
										Comment       = $Element.desc
									}
									New-Object -TypeName PSObject -Property $data
								}
							}
							else {
								Write-Verbose "no preference xml data found for $section"
							}
						}
					}
					else {
						Write-Verbose "there are no preferences for this policy object"
					}
				}
			}
			break
		}
	}
}
