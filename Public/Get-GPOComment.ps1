#requires -Modules GroupPolicy
<#
.SYNOPSIS
    Retrieve comments/descriptions embedded in GPOs
.DESCRIPTION
    Retrieve comments and descriptions embedded in GPOs, GPO Settings and GP Preferences Settings
.PARAMETER GPOName
    [string[]] (required) Name(s) of Group Policy Objects or '*' for all GPOs
.PARAMETER PolicyGroup
    [string] (required) What aspects of each GPO is to be queried
    List = Policy, Settings, Preferences
.NOTES
    version 1.0.3 - DS - 2017.08.11
.EXAMPLE
    Get-GPOComment -GPOName '*' -PolicyGroup 'Policy'
.EXAMPLE
    $GpoNames | Get-GPOComment -PolicyGroup 'Policy'
.EXAMPLE
    Get-GPOComment -GPOName '*' -PolicyGroup 'Settings'
.EXAMPLE
    Get-GPOComment -GPOName '*' -PolicyGroup 'Preferences'
#>

function Get-GPOComment {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='Name of Policy or Policies')]
            [ValidateNotNullOrEmpty()]
            [string[]] $GPOName,
        [parameter(Mandatory=$True, HelpMessage='Policy group to query')]
            [ValidateSet('Policy','Settings','Preferences')]
            [string] $PolicyGroup
    )
    if ($PolicyGroup -eq 'Policy') {
        if ($GPOName -eq '*') {
            Write-Verbose "loading all policy objects"
            $gpos = Get-GPO -All
        }
        else {
            Write-Verbose "loading specific policy objects"
            try {
                $gpos = $GPOName | Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select Id,DisplayName,DomainName,Description}
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
    }
    elseif ($PolicyGroup -eq 'Settings') {
        if ($GPOName -eq '*') {
            Write-Verbose "loading all policy objects"
            $gpos = Get-GPO -All
        }
        else {
            Write-Verbose "loading specific policy objects"
            $gpos = $GPOName | Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select Id,DisplayName,DomainName,Description}
        }
        foreach ($gpo in $gpos) {
            $configs   = @('Machine','User')
            Write-Verbose "policy: $($gpo.DisplayName)"
            foreach ($config in $configs) {
                $commFile = "\\$($gpo.DomainName)\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.ID)}\$config\comment.cmtx"
                if (Test-Path $commFile) {
                    Write-Verbose "loading... $commFile"
                    [xml]$cmtx = Get-Content -Path $commFile
                    $clinks = $cmtx.policyComments.resources.stringTable.string
                    foreach ($clink in $clinks) {
                        $data = [ordered]@{
                            PolicyID = $gpo.Id
                            Name     = $gpo.DisplayName
                            Configuration = $config
                            Setting  = $clink.Id
                            Comment  = $clink.InnerText.Trim()
                        }
                        New-Object -TypeName PSObject -Property $data
                    }
                }
                else {
                    Write-Verbose "there are no $config comments in $($gpo.DisplayName)"
                    $data = [ordered]@{
                        PolicyID = $gpo.ID
                        Name     = $gpo.DisplayName
                        Configuration = $config
                        Setting  = $null
                        Comment  = $null
                    }
                    New-Object -TypeName PSObject -Property $data
                }
            }
        }
    }
    elseif ($PolicyGroup -eq 'Preferences') {
        if ($GPOName -eq '*') {
            Write-Verbose "loading all policy objects: preferences"
            $gpos = Get-GPO -All
        }
        else {
            Write-Verbose "loading specific policy objects"
            $gpos = $GPOName | Foreach-Object {Get-GPO -Name $_ -ErrorAction Stop | Select Id,DisplayName,DomainName,Description}
        }
        foreach ($gpo in $gpos) {
            $configs   = @('Machine','User')
            Write-Verbose "policy: $($gpo.DisplayName)"
            foreach ($config in $configs) {
                $gppPath = "\\$($gpo.DomainName)\SYSVOL\$($gpo.DomainName)\Policies\{$($gpo.Id)}\$config\Preferences"
                if (Test-Path $gppPath) {
                    Write-Verbose "there are preferences for this policy object"
                    foreach ($section in Get-ChildItem -Path $gppPath -Directory | Select-Object -ExpandProperty Name) {
                        $gppFile = "$gppPath\$section\$section.xml"
                        Write-Verbose "gppref file: $gppFile"
                        [xml]$gppXML = Get-Content $gppFile
                        $Element = $section.substring(0, $section.Length-1)
                        foreach ($Element in $gppXML."$section"."$Element") {
                            $data = [ordered]@{
                                PolicyID = $gpo.ID
                                Name     = $gpo.DisplayName
                                Configuration = $config
                                Section  = $section
                                Element  = $Element.name
                                Comment  = $Element.desc
                            }
                            New-Object -TypeName PSObject -Property $data
                        }
                    }
                }
                else {
                    Write-Verbose "there are no preferences for this policy object"
                }
            }
        }
    }
}
