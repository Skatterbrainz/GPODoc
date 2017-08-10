#requires -modules GroupPolicy

function Get-GPOComment {
    param (
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string] $GPOName
    )
    try {
        $gpo = Get-GPO -Name $GPOName
        $result = $($gpo.Description).Trim()
    }
    catch {
        Write-Warning "Failed to read GPO: $GPOName"
    }
    Write-Output $result
}
