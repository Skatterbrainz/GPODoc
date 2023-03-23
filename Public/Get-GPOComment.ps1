function Get-GPOComment {
	<#
	.SYNOPSIS
		Return comments and description embedded in GPOs
	.DESCRIPTION
		Return comments and description embedded in GPOs, individual
		GPO settings, and Group Policy Preferences settings.
	.PARAMETER Name
		Optional. Name of GPO to query. If blank, will return all GPO's in the domain.
	.PARAMETER PolicyGroup
		Optional. Which aspect of GPO information to query comments from

		* Policy (default) - returns GPO comments
		* Settings - returns comments entered for individual settings within the GPOs
		* Preferences - returns comments entered for individual preferences within the GPOs

	.EXAMPLE
		Get-GPOComment -Name "Workstation Power Settings"

		Return GPO comments for named GPO
	.EXAMPLE
		Get-GPOComment -Name "Workstation Power Settings" -PolicyGroup Settings

		Return GPO comments on individual settings within GPO
	.EXAMPLE
		Get-GPOComment

		Return GPO comments for all GPO's in the current AD domain
	.LINK
		https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOComment.md
	#>
	[CmdletBinding()]
	param (
		[parameter()][string]$Name = "",
		[parameter()][ValidateSet('Policy','Settings','Preferences')][string] $PolicyGroup = "Policy"
	)
	if ([string]::IsNullOrWhiteSpace($Name)) {
		$GpoSet = @(Get-GPO -All | Sort-Object DisplayName)
	} else {
		$GpoSet = @(Get-GPO -Name $Name | Select-Object Id, DisplayName, DomainName, Description)
	}
	foreach ($gpo in $GpoSet) {
		Write-Verbose "policy: $($gpo.DisplayName)"
		switch ($PolicyGroup) {
			'Policy' {
				[pscustomobject]@{
					PolicyID    = $gpo.Id
					Name        = $gpo.DisplayName
					Description = $gpo.Description
				}
			}
			'Settings' {
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
				} # foreach config
			}
			'Preferences' {
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
									[pscustomobject]@{
										PolicyID      = $gpo.ID
										Name          = $gpo.DisplayName
										Configuration = $config
										Section       = $section
										Element       = $Element.name
										Comment       = $Element.desc
									}
								}
							} else {
								Write-Verbose "no preference xml data found for $section"
							}
						}
					} else {
						Write-Verbose "there are no preferences for this policy object"
					}
				} # foreach config
			}
		} # switch
	} # foreach gpo
}
