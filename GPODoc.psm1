#Module vars
$ModulePath = $PSScriptRoot

$Public  = Get-ChildItem $PSScriptRoot\Public\*.ps1 -ErrorAction Stop
$Private = Get-ChildItem $PSScriptRoot\Private\*.ps1 -ErrorAction Stop
[string[]]$PrivateModules = Get-ChildItem $PSScriptRoot\Private -ErrorAction Stop |
    Where-Object {$_.PSIsContainer} |
        Select-Object -ExpandProperty FullName

Write-Output "$($PrivateModules.count) modules"

# dot source the files
if ($Private) {
    foreach ($import in $Private) {
        try {
            . $import.FullName
        }
        catch {
            Write-Error "Failed to import function $($import.FullName): $_"
        }
    }
}
if ($Public) {
    foreach ($import in $Public) {
        try {
            . $import.FullName
        }
        catch {
            Write-Error "Failed to import function $($import.FullName): $_"
        }
    }
}

# load dependency modules
foreach ($Module in $PrivateModules) {
    try {
        Import-Module $Module -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to import module $Module`: $_"
    }
}
