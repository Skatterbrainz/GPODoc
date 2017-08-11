#requires -modules GroupPolicy
<#
.DESCRIPTION
    Returns comments embedded in GPO policy settings

.PARAMETER GPOName
    [string] (required) Name of GPO

.PARAMETER Settings
    [switch] (optional) Retrieve comments associated with embedded settings

.PARAMETER Configuration
    [string] (optional) list: Machine or User

.PARAMETER Preferences
    [switch] (optional) Retrieve comments associated with embedded preferences

.EXAMPLE
    Get-GPOComment -GPOName "C - Kiosk Settings"

.EXAMPLE
    Get-GPOComment -GPOName "C - Kiosk Settings" -Settings -Configuration Machine

.EXAMPLE
    Get-GPOComment -GPOName "C - Kiosk Settings" -MachinePreferences -ComputerSection Folders

.EXAMPLE
    Get-GPOComment -GPOName "C - Kiosk Settings" -UserPreferences -UserSection InternetSettings

.EXAMPLE
    Get-GPOComment -GPOName "C - Kiosk Settings" -UserPreferences -UserSection Folders -UserItem SharedDocs

.NOTES
    2017.08.10.01 - DS
#>

function Get-GPOComment {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='GPO Name')]
            [ValidateNotNullOrEmpty()]
            [string[]] $GPOName,

        [parameter(ParameterSetName='Settings', Mandatory=$False, HelpMessage='Query GPO Settings')]
            [switch] $Settings,

        [parameter(ParameterSetName='Settings', Mandatory=$False, HelpMessage='Settings Configuration Group Name')]
            [ValidateSet('Machine','User')]
            [string] $Configuration,

        [parameter(ParameterSetName='MachinePreferences', Mandatory=$False, HelpMessage='Query Machine Preferences')]
            [switch] $MachinePreferences,
        [parameter(ParameterSetName='MachinePreferences', Mandatory=$False, HelpMessage="Preferences Section")]
            [ValidateSet('DataSources','Devices','EnvironmentVariables','Files','Folders','Groups','IniFiles','NetworkShares','PowerOptions','Printers','Registry','ScheduledTasks','Shortcuts')]
            [string] $ComputerSection,
        [parameter(ParameterSetName='MachinePreferences', Mandatory=$False, HelpMessage="Name of Individual Setting")]
            [string] $MachineItem = "",

        [parameter(ParameterSetName='UserPreferences', Mandatory=$False, HelpMessage='Query User Preferences')]
            [switch] $UserPreferences,
        [parameter(ParameterSetName='UserPreferences', Mandatory=$False, HelpMessage="Preferences Section")]
            [ValidateSet('Drives','EnvironmentVariables','Files','Folders','InternetSettings','RegionalOptions','StartMenuTaskbar')]
            [string] $UserSection,
        [parameter(ParameterSetName='UserPreferences', Mandatory=$False, HelpMessage="Name of Individual Setting")]
            [string] $UserItem = ""

    )
    
    Write-Verbose "$((Get-Module GPODoc | Select -ExpandProperty Version) -join '.')"
    if ($Settings) {
        if (-not ($Configuration)) {
            Write-Warning "-Settings requires -Configuration to specify Machine or User"
            break
        }
        Write-Verbose "querying GPO settings"
        foreach ($GN in $GPOName) {
            Write-Verbose "querying GPO: $GN"
            try {
                $gpo = Get-GPO -Name $GN
            }
            catch {
                Write-Warning "could not open policy object: $GN"
                break
            }
            $policyID     = $gpo.ID
            $policyDomain = $gpo.DomainName
            $policyName   = $gpo.DisplayName
            $policyVpath  = "\\$($policyDomain)\SYSVOL\$($policyDomain)\Policies\{$($policyID)}\$Configuration\comment.cmtx"
            Write-Verbose "filepath... $policyVpath"
            if (Test-Path $policyVpath) {
                Write-Verbose "loading... $policyVpath"
                [xml]$cmtx = Get-Content -Path $policyVpath
                $clinks = $cmtx.policyComments.resources.stringTable.string
                foreach ($clink in $clinks) {
                    $result = New-Object -TypeName PSObject
                    $result | Add-Member -MemberType NoteProperty -Name PolicyRef -Value $clink.id
                    $result | Add-Member -MemberType NoteProperty -Name Comment -Value $clink.InnerText.Trim()
                    Write-Output $result
                }
            }
            else {
                Write-Warning "$GN has no embedded preferences comments"
            }
        } # end foreach
    }
    elseif ($MachinePreferences -or $UserPreferences) {
        Write-Verbose "querying GP preferences"
        foreach ($GN in $GPOName) {
            try {
                $gpo = Get-GPO -Name $GN -ErrorAction Stop
            }
            catch {
                Write-Error "Group Policy $GN not found"
                break
            }

            $policyID     = $gpo.ID
            $policyDomain = $gpo.DomainName
            $policyName   = $gpo.DisplayName

            if ($MachinePreferences) { 
                $Config   = 'Machine'
                $Section  = $ComputerSection
                $ItemName = $MachineItem
            } 
            else { 
                $Config   = 'User'
                $Section  = $UserSection
                $ItemName = $UserItem
            }
            $ElementName = $Section.Substring(0,$Section.Length-1)
            $policyVpath = "\\$($policyDomain)\SYSVOL\$($policyDomain)\Policies\{$($policyID)}\$Config\Preferences\$Section\$Section.xml"
    
            Write-Verbose "loading...... $policyVpath"
            Write-Verbose "section...... $Section"
            Write-Verbose "elementname.. $ElementName"

            if (Test-Path $policyVpath) {
                [xml]$SettingXML = Get-Content $policyVpath
                Write-Verbose "$($SettingXML.ChildNodes.Count) nodes found"
                if ($ItemName -ne "") {
                    $result = $SettingXML."$Section"."$ElementName" | 
                        Where-Object {$_.name -eq $ItemName} |
                            Select-Object -ExpandProperty desc
                    Write-Output $result
                }
                else {
                    foreach ($Element in $SettingXML."$Section"."$ElementName") {
                        $result = New-Object -TypeName PSObject
                        $result | Add-Member -MemberType NoteProperty -Name ItemName -Value $Element.name
                        $result | Add-Member -MemberType NoteProperty -Name Description -Value $Element.desc
                        Write-Output $result
                    }
                }
            }
            else {
                Write-Error "unable to load xml data for $GN"
            }
        } # end foreach
    }
    else {
        Write-Verbose "querying GPO descriptions"
        foreach ($GN in $GPOName) {
            Write-Verbose "querying GPO $GN"
            try {
                $gpo = Get-GPO -Name $GN
                $comment = $($gpo.Description).Trim()
            }
            catch {
                Write-Warning "$GN has no comment"
                break
            }
            $result = New-Object -TypeName PSObject
            $result | Add-Member -MemberType NoteProperty -Name Name -Value $GN
            $result | Add-Member -MemberType NoteProperty -Name Comment -Value $comment
            Write-Output $result
        } # end foreach
    }
}
