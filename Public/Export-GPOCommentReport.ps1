#requires -Modules GroupPolicy

function Export-GPOCommentReport {
    param (
        [parameter(Mandatory=$True, ValueFromPipeline=$True, HelpMessage='Name of Policy or Policies')]
            [ValidateNotNullOrEmpty()]
            [string[]] $GPOName,
        [parameter(Mandatory=$True, HelpMessage='Path to report file')]
            [ValidateNotNullOrEmpty()]
            [string] $ReportFile
    )
    if ($GPOName -eq '*') {
        Write-Verbose "loading all policy objects: preferences"
        $gpos = Get-GPO -All
    }
    else {
        Write-Verbose "loading specific policy objects"
        $gpos = $GPOName | Foreach-Object {Get-GPO -Name $_}
    }

    $fragments = @()
    $fragments += "<h1>Group Policy Report</h1>"

    foreach ($gpo in $gpos) {
        $gpoName = $gpo.DisplayName
        Write-Output "GPO: $gpoName"
        $desc = Get-GpoComment -GPOName $gpoName -PolicyGroup Policy
        $sett = Get-GpoComment -GPOName $gpoName -PolicyGroup Settings
        $pref = Get-GpoComment -GPOName $gpoName -PolicyGroup Preferences
        
        Write-Verbose $desc

        $fragments += "<h2>$gpoName</h2><br/><p>$desc</p>" | ConvertTo-Html -As List

        $fragments += "<h3>Policy Settings</h3>"
        $fragments += $sett | ConvertTo-Html -Fragment

        $fragments += "<h3>Preferences</h3>"
        $fragments += $pref | ConvertTo-Html -Fragment
    }
    $fragments += "<p class='footer'>$(Get-Date)</p>"

    $convertParams = @{ 
  head = @"
 <Title>Group Policy Comments - $($env:COMPUTERNAME)</Title>
<style>
body { background-color:#E5E4E2;
       font-family:Calibri;
       font-size:10pt; }
td, th { border:0px solid black; 
         border-collapse:collapse;
         white-space:pre; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px ;white-space:pre; }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
h2 {
    font-family:Tahoma;
    color:#000;
}
h2 {
    font-family:Tahoma;
    color:#6D7B8D;
}
h3 {
    font-family:Tahoma;
    color:#6D7B8D;
}
.alert {
 color: red; 
 }
.footer 
{ color:green; 
  margin-left:10px; 
  font-family:Tahoma;
  font-size:8pt;
  font-style:italic;
}
</style>
"@
 body = $fragments
}
 
    ConvertTo-Html @convertParams | Out-File $ReportFile

}
