#requires -modules GroupPolicy
<#
.DESCRIPTION
  returns comment values associated with GPO Preferences items
.EXAMPLE
  Get-GPPrefComment -GPOName "GPO1" -Machine -ComputerSection Folders
.EXAMPLE
  Get-GPPrefComment -GPOName "GPO1" -Machine -ComputerSection Folders -ItemName 'MyFolder'
.EXAMPLE
  Get-GPPrefComment -GPOName "GPO1" -User -UserSection Files
.EXAMPLE
  Get-GPPrefComment -GPOName "GPO1" -User -UserSection Files -ItemName 'TestFile.txt'
.NOTES
  If ItemName is not specified, the results are returned as a hash
  If ItemName is specified, the result is returned as text
#>

function Get-GPPrefComment {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$True, HelpMessage="Name of Group Policy Object")]
            [ValidateNotNullOrEmpty()]
            [string] $GPOName,
        [parameter(ParameterSetName='Machine')]
            [switch] $Machine,
        [parameter(ParameterSetName='Machine', HelpMessage="Preferences Section")]
            [ValidateSet('DataSources','Devices','EnvironmentVariables','Files','Folders','Groups','IniFiles','NetworkShares','PowerOptions','Printers','Registry','ScheduledTasks','Shortcuts')]
            [string] $ComputerSection,
        [parameter(ParameterSetName='User')]
            [switch] $User,
        [parameter(ParameterSetName='User', HelpMessage="Preferences Section")]
            [ValidateSet('Drives','EnvironmentVariables','Files','Folders','InternetSettings','RegionalOptions','StartMenuTaskbar')]
            [string] $UserSection,
        [parameter(Mandatory=$False, HelpMessage="Name of Individual Setting")]
            [string] $ItemName = ""
    )
    try {
        $gpo = Get-GPO -Name $GPOName -ErrorAction Stop
    }
    catch {
        Write-Error "Group Policy $GPOName not found"
        break
    }

    $policyID = $gpo.ID
    $policyDomain = $gpo.DomainName
    $policyName   = $gpo.DisplayName

    if ($Machine) { 
        $Config  = 'Machine'
        $Section = $ComputerSection
    } 
    else { 
        $Config  = 'User'
        $Section = $UserSection
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
        Write-Error "unable to load xml data for $GPOName"
    }
}
