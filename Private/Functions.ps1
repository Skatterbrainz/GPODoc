function Get-GpoCommentFile {
	<#
	.SYNOPSIS
	Get Gpo Comment File
	.DESCRIPTION
	Get GPO comment file path
	.PARAMETER Gpo
	Group Policy Object reference
	.PARAMETER ConfigType
	Configuration Type: Machine or User
	.EXAMPLE
	$filepath = Get-GpoCommentFile -Gpo $GPO
	.NOTES
	.INPUTS
		Group Policy Object reference
	.OUTPUTS
		Path and Name of Comment CMTX file
	#>
	param (
		[parameter(Mandatory=$True, HelpMessage="Group Policy ID")]
			[ValidateNotNullOrEmpty()]
			$Gpo,
		[parameter(Mandatory=$True, HelpMessage="Configuration Type: Machine or User")]
			[ValidateSet('Machine','User')]
			[string] $ConfigType
	)
	$result = "\\$($gpo.DomainName)\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.ID)}\$ConfigType\comment.cmtx"
	if (Test-Path $result) {
		Write-Output $result
	}
	else {
		Write-Output ""
	}
}

function Get-GppCommentPath {
	<#
	.SYNOPSIS
	Get Group Policy Preferences Comment file path
	.DESCRIPTION
	Get Group Policy Preferences Comment file path
	.PARAMETER Gpo
	Group Policy Object reference
	.PARAMETER ConfigType
	Configuration Type: Machine or User
	.EXAMPLE
	$gppPath = Get-GppCommentPath -Gpo $gpo -ConfigType 'Machine'
	.NOTES
	General notes
	#>
	param (
		[parameter(Mandatory=$True, HelpMessage="Group Policy Object")]
			[ValidateNotNullOrEmpty()]
			$Gpo,
		[parameter(Mandatory=$True, HelpMessage="Machine or User")]
			[ValidateSet('Machine','User')]
			[string] $ConfigType
	)
	$result = "\\$($gpo.DomainName)\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.Id)}\$ConfigType\Preferences"
	if (Test-Path $result) {
		Write-Output $result
	}
	else {
		Write-Output ""
	}
}

function Get-GpoLinkID {
	<#
	.SYNOPSIS
	Returns GPO link GUID from OU links
	.DESCRIPTION
	Same as the Synopsis
	.PARAMETER OuLinks
	Array of OU links
	#>
	param (
		[parameter(Mandatory)]$OuLinks
	)
	$guidpattern = '[A-Z0-9]{8}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{12}'
	foreach ($item in $OuLinks) {
		($item | Select-String $guidpattern).Matches.Value
	}
}

function Get-GPOByID {
	<#
	.SYNOPSIS
	Get GPO by ID (GUID)
	.DESCRIPTION
	OMG do I need to say more?
	.PARAMETER ID
	GUID of the GPO
	.PARAMETER GpoArray
	Array of GPOs
	#>
	[CmdletBinding()]
	param (
		[parameter(Mandatory)]
		[ValidateScript({
			try {
				[System.Guid]::Parse($_) | Out-Null
				$true
			} catch {
				$false
			}
		})]
		[string]$ID,
		[parameter(Mandatory)]$GpoArray
	)
	$GpoArray.Where({$_.Id -eq $ID})
}
