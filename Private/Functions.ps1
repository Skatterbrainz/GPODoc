function Get-GpoCommentFile {
    <#
    .SYNOPSIS
    Get Gpo Comment File
    
    .DESCRIPTION
    Get GPO comment file path
    
    .PARAMETER Gpo
    Group Policy Object reference
    
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
        $Gpo
    )
    $result = "\\$($gpo.DomainName)\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.ID)}\$config\comment.cmtx"
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
    
    .EXAMPLE
    $gppPath = Get-GppCommentPath -Gpo $gpo
    
    .NOTES
    General notes
    #>
    param (
        [parameter(Mandatory=$True, HelpMessage="Group Policy Object")]
        [ValidateNotNullOrEmpty()]
        $Gpo,
        [parameter(Mandatory=$True, HelpMessage="Computer or User")]
        [ValidateNotNullOrEmpty()]
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