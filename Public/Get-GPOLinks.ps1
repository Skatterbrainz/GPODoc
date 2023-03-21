function Get-GPOLinks {
	<#
	.SYNOPSIS
		Return OUs and associated (linked) GPO's
	.DESCRIPTION
		Return OU's and associated GPO's. All OU's or starting with a SearchBase location in AD.
	.NOTES
		Heavily and deeply, almost spiritually and metaphysically, inspired by https://github.com/BananaJama/PowerShell/
	.EXAMPLE
		Get-GPOLinks

		Returns all OUs and GPO links for each
	.EXAMPLE
		Get-GPOLinks -SearchBase "OU=Servers,OU=CORP,DC=contoso,DC=local"

		Returns OUs starting at and below the path specified by SearchBase, and their GPO links

	.LINK
		https://github.com/Skatterbrainz/GPODoc/blob/master/Docs/Get-GPOlinks.md
	#>
	[CmdletBinding()]
	param (
		[parameter()][string]$SearchBase = ""
	)
	if (![string]::IsNullOrWhiteSpace($SearchBase)) {
		$ous = @(Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -SearchScope Subtree)
	} else {
		$ous = @(Get-ADOrganizationalUnit -Filter *)
	}
	$allGPOs = @(Get-GPO -All)
	foreach ($ou in $ous) {
		$links = $ou.LinkedGroupPolicyObjects
		$ids = Get-GpoLinkID -OuLinks $links
		$gpos = $ids | Foreach-Object { Get-GPOByID -ID $_ -GpoArray $allGPOs }
		foreach ($gpo in $gpos) {
			[pscustomobject]@{
				OuName  = $ou.Name
				OuPath  = $ou.DistinguishedName
				GpoName = $gpo.DisplayName
				GpoID   = $gpo.Id
			}
		}
	}
}