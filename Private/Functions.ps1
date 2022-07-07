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
