<# 
.SYNOPSIS
  WSUS weekly compliance report -> HTML + Markdown -> optional Teams post

.PREREQS
  - Account with WSUS read.
  - WSUS admin console installed (Microsoft.UpdateServices.Administration).
  - Teams incoming webhook if you want to post.

.PARAMS
  -ServerName   WSUS server DNS
  -Port         8530 default (8531 if SSL)
  -UseSsl       switch
  -TargetGroup  optional WSUS computer target group
  -OutDir       output directory
  -TeamsWebhook optional Teams webhook URL

.NOTES
  Save as: wsus-weekly-report.ps1
#>

param(
  [Parameter(Mandatory=$true)][string]$ServerName,
  [int]$Port = 8530,
  [switch]$UseSsl,
  [string]$TargetGroup,
  [Parameter(Mandatory=$true)][string]$OutDir,
  [string]$TeamsWebhook,
  [int]$ActiveDays = 14 # ignore clients older than this LastSync
)

$ErrorActionPreference = 'Stop'
if (!(Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

function Connect-WSUS {
  [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.UpdateServices.Administration')
  [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ServerName, [bool]$UseSsl, $Port)
}

function Get-ComputerScope {
  param($Server, $TargetGroup)
  if ([string]::IsNullOrWhiteSpace($TargetGroup)) {
    $Server.GetComputerTargets()
  } else {
    $group = $Server.GetComputerTargetGroups() | Where-Object { $_.Name -eq $TargetGroup }
    if (-not $group) { throw "Target group '$TargetGroup' not found." }
    $group.GetComputerTargets()
  }
}

function Get-ComputerCompliance {
  <#
    Returns per-computer counts:
    Needed, Failed, Installed, NotApplicable, PendingReboot
  #>
  param($Server, $Computers)

  $results = @()

  # Scope: only approved updates. Include main install states.
  $scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
  $scope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
  $scope.IncludedInstallationStates = (
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Installed,
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::InstalledPendingReboot,
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotApplicable,
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Needed,
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed,
    [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
  )

  foreach ($c in $Computers) {

    $infos = $Server.GetUpdateInstallationInfoPerComputerTarget($c.Id, $scope)

    $counts = @{
      Installed        = 0
      NotApplicable    = 0
      Needed           = 0
      Failed           = 0
      PendingReboot    = 0
    }

    foreach ($info in $infos) {
      switch ($info.UpdateInstallationState) {
        'Installed'               { $counts.Installed++ }
        'NotApplicable'           { $counts.NotApplicable++ }
        'NotInstalled'            { $counts.Needed++ }
        'Needed'                  { $counts.Needed++ }
        'Failed'                  { $counts.Failed++ }
        'InstalledPendingReboot'  { $counts.PendingReboot++ }
      }
    }

    $results += [pscustomobject]@{
      ComputerName   = $c.FullDomainName
      GroupNames     = ($c.GetComputerTargetGroups() | ForEach-Object Name) -join ', '
      OSDescription  = $c.ClientVersion + ' | ' + $c.OSDescription
      LastSync       = $c.LastSyncTime
      Installed      = $counts.Installed
      NotApplicable  = $counts.NotApplicable
      Needed         = $counts.Needed
      Failed         = $counts.Failed
      PendingReboot  = $counts.PendingReboot
    }
  }
  $results
}

function Summarize-Fleet {
  param($Rows)
  [pscustomobject]@{
    TotalMachines    = $Rows.Count
    ActiveMachines   = ($Rows | Where-Object { $_.LastSync -gt (Get-Date).AddDays(-$ActiveDays) }).Count
    AnyNeeded        = ($Rows | Where-Object { $_.Needed -gt 0 }).Count
    AnyFailed        = ($Rows | Where-Object { $_.Failed -gt 0 }).Count
    AnyPendingReboot = ($Rows | Where-Object { $_.PendingReboot -gt 0 }).Count
    NeededUpdates    = ($Rows.Needed | Measure-Object -Sum).Sum
    FailedUpdates    = ($Rows.Failed | Measure-Object -Sum).Sum
  }
}

function New-MarkdownReport {
  param($ScopeName, $Summary, $Rows, $ActiveDays)

  $date = Get-Date -Format 'yyyy-MM-dd'
  $md = @()
  $md += "# WSUS Weekly Patch Report — $date"
  $md += ""
  $md += "**Scope:** $ScopeName"
  $md += ""
  $md += "_Note: ignores clients with LastSync older than $ActiveDays days in the per-machine table._"
  $md += ""
  $md += "## Fleet Summary"
  $md += ""
  $md += "| Metric | Count |"
  $md += "|---|---:|"
  $md += "| Total machines (all) | {0} |" -f $Summary.TotalMachines
  $md += "| Active machines (≤{0}d) | {1} |" -f $ActiveDays, $Summary.ActiveMachines
  $md += "| Machines needing updates | {0} |" -f $Summary.AnyNeeded
  $md += "| Machines with failed updates | {0} |" -f $Summary.AnyFailed
  $md += "| Machines pending reboot | {0} |" -f $Summary.AnyPendingReboot
  $md += "| Total needed updates | {0} |" -f $Summary.NeededUpdates
  $md += "| Total failed updates | {0} |" -f $Summary.FailedUpdates
  $md += ""
  $md += "## Per-Machine Detail (top 50 by Needed, active only)"
  $md += ""
  $md += "| Computer | Needed | Failed | PendingReboot | Installed | NotApplicable | LastSync | Groups |"
  $md += "|---|---:|---:|---:|---:|---:|---|---|"

  $Rows |
    Where-Object { $_.LastSync -gt (Get-Date).AddDays(-$ActiveDays) } |
    Sort-Object Needed -Descending |
    Select-Object -First 50 |
    ForEach-Object {
      $md += "| {0} | {1} | {2} | {3} | {4} | {5} | {6:yyyy-MM-dd HH:mm} | {7} |" -f `
        $_.ComputerName, $_.Needed, $_.Failed, $_.PendingReboot, $_.Installed, $_.NotApplicable, $_.LastSync, ($_.GroupNames -replace '\|','-')
    }

  $md -join "`r`n"
}

function New-HtmlReport {
  param($Markdown)

  # Minimal markdown to HTML
  $lines = $Markdown -split "`r`n"
  $out = New-Object System.Collections.Generic.List[string]

  foreach ($line in $lines) {
    if ($line -match '^\#\# (.+)$') { $out.Add("<h2>$($Matches[1])</h2>"); continue }
    if ($line -match '^\# (.+)$')   { $out.Add("<h1>$($Matches[1])</h1>"); continue }
    if ($line -match '^\|') {
      # table mode
      $cells = $line.Trim('|').Split('|').ForEach({ $_.Trim() })
      if ($cells[0] -eq '---') { continue } # skip table separators
      $row = '<tr>' + ($cells | ForEach-Object { "<td>$([System.Web.HttpUtility]::HtmlEncode($_))</td>" }) -join '' + '</tr>'
      if (-not $out[-1] -or -not $out[-1].EndsWith('</table>')) {
        if (-not ($out | Select-String '<table>$' -SimpleMatch)) { $out.Add('<table>') }
      }
      $out.Add($row)
      # next non-table line will close table
    } else {
      if ($out.Count -gt 0 -and $out[-1].StartsWith('<tr>')) { $out.Add('</table>') }
      $line = $line -replace '\*\*(.+?)\*\*','<strong>$1</strong>'
      if ($line.Trim().Length -eq 0) { $out.Add('<br/>') } else { $out.Add("<p>$line</p>") }
    }
  }
  if ($out.Count -gt 0 -and $out[-1].StartsWith('<tr>')) { $out.Add('</table>') }

@"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>WSUS Weekly Patch Report</title>
  <style>
    body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
    h1 { margin-bottom: .2rem; }
    table { border-collapse: collapse; width: 100%; margin: 1rem 0; }
    td, th { border: 1px solid #ddd; padding: 6px 8px; }
    tr:nth-child(even) { background: #fafafa; }
  </style>
</head>
<body>
$($out -join "`r`n")
</body>
</html>
"@
}

function Publish-Teams {
  param($WebhookUrl, $Title, $Markdown)
  if (-not $WebhookUrl) { return }
  $payload = @{
    title = $Title
    text  = $Markdown
  } | ConvertTo-Json -Depth 4
  Invoke-RestMethod -Method Post -Uri $WebhookUrl -Body $payload -ContentType 'application/json'
}

# -------- Main --------
$wsus = Connect-WSUS
$computers = Get-ComputerScope -Server $wsus -TargetGroup $TargetGroup
$rows = Get-ComputerCompliance -Server $wsus -Computers $computers
$summary = Summarize-Fleet -Rows $rows

$scopeName = if ($TargetGroup) { "Group: $TargetGroup" } else { "All Computers" }
$md = New-MarkdownReport -ScopeName $scopeName -Summary $summary -Rows $rows -ActiveDays $ActiveDays
$mdPath = Join-Path $OutDir ("WSUS-Weekly-{0}.md" -f (Get-Date -Format 'yyyyMMdd'))
$md | Out-File -FilePath $mdPath -Encoding utf8

$html = New-HtmlReport -Markdown $md
$htmlPath = Join-Path $OutDir ("WSUS-Weekly-{0}.html" -f (Get-Date -Format 'yyyyMMdd'))
$html | Out-File -FilePath $htmlPath -Encoding utf8

Publish-Teams -WebhookUrl $TeamsWebhook -Title "WSUS Weekly Patch Report" -Markdown $md

Write-Host "Markdown: $mdPath"
Write-Host "HTML    : $htmlPath"