<# 
.SYNOPSIS
  WSUS weekly compliance report -> HTML + Markdown -> publish (Teams or email)

.PREREQS
  - Run with an account that can read WSUS.
  - WSUS admin console installed (for Microsoft.UpdateServices.Administration).
  - For Teams posting: an incoming webhook URL.
  - For email: SMTP relay reachable.

.PARAMS
  -ServerName   WSUS server DNS name
  -Port         WSUS port (default 8530; 8531 if SSL)
  -UseSsl       Switch for SSL
  -TargetGroup  Optional WSUS computer target group to scope
  -OutDir       Output directory for report files
  -TeamsWebhook Optional Teams incoming webhook URL
  -Email*       Optional SMTP settings for mail

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
  [string]$EmailFrom,
  [string]$EmailTo,
  [string]$EmailSubject = "WSUS Weekly Patch Report",
  [string]$SmtpServer,
  [int]$SmtpPort = 25
)

$ErrorActionPreference = 'Stop'
if (!(Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

function Connect-WSUS {
  [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.UpdateServices.Administration')
  $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ServerName, [bool]$UseSsl, $Port)
  return $wsus
}

function Get-ComputerScope {
  param($Server, $TargetGroup)
  if ([string]::IsNullOrWhiteSpace($TargetGroup)) {
    return $Server.GetComputerTargets()
  } else {
    $group = $Server.GetComputerTargetGroups() | Where-Object { $_.Name -eq $TargetGroup }
    if (-not $group) { throw "Target group '$TargetGroup' not found." }
    return $group.GetComputerTargets()
  }
}

function Get-ComputerCompliance {
  <#
    Output per-computer object with counts:
    Needed, Failed, Installed, NotApplicable, PendingReboot
    Note: WSUS API surfaces per-computer install states via UpdateInstallationInfo.
  #>
  param($Server, $Computers)

  $results = @()
  foreach ($c in $Computers) {
    # Some environments have thousands of updates; limit to “approved” to keep it fast.
    $scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $scope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
    $scope.IncludedInstallationStates = (
      [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Installed,
      [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::InstalledPendingReboot,
      [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotApplicable,
      [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Needed,
      [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed
    )

    # Pull per-computer installation info
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
        'NotInstalled'            { $counts.Needed++ }            # WSUS sometimes reports as NotInstalled
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
  return $results
}

function Summarize-Fleet {
  param($Rows)
  $total = $Rows.Count
  [pscustomobject]@{
    TotalMachines   = $total
    AnyNeeded       = ($Rows | Where-Object { $_.Needed -gt 0 }).Count
    AnyFailed       = ($Rows | Where-Object { $_.Failed -gt 0 }).Count
    AnyPendingReboot= ($Rows | Where-Object { $_.PendingReboot -gt 0 }).Count
    NeededUpdates   = ($Rows.Needed | Measure-Object -Sum).Sum
    FailedUpdates   = ($Rows.Failed | Measure-Object -Sum).Sum
  }
}

function New-MarkdownReport {
  param($ScopeName, $Summary, $Rows)

  $date = Get-Date -Format 'yyyy-MM-dd'
  $md = @()
  $md += "# WSUS Weekly Patch Report — $date"
  $md += ""
  $md += "**Scope:** $ScopeName"
  $md += ""
  $md += "## Fleet Summary"
  $md += ""
  $md += "| Metric | Count |"
  $md += "|---|---:|"
  $md += "| Total machines | {0} |" -f $Summary.TotalMachines
  $md += "| Machines needing updates | {0} |" -f $Summary.AnyNeeded
  $md += "| Machines with failed updates | {0} |" -f $Summary.AnyFailed
  $md += "| Machines pending reboot | {0} |" -f $Summary.AnyPendingReboot
  $md += "| Total needed updates | {0} |" -f $Summary.NeededUpdates
  $md += "| Total failed updates | {0} |" -f $Summary.FailedUpdates
  $md += ""
  $md += "## Per-Machine Detail (top 50 by Needed)"
  $md += ""
  $md += "| Computer | Needed | Failed | PendingReboot | Installed | NotApplicable | LastSync | Groups |"
  $md += "|---|---:|---:|---:|---:|---:|---|---|"

  $Rows | Sort-Object Needed -Descending | Select-Object -First 50 | ForEach-Object {
    $md += "| {0} | {1} | {2} | {3} | {4} | {5} | {6:yyyy-MM-dd HH:mm} | {7} |" -f `
      $_.ComputerName, $_.Needed, $_.Failed, $_.PendingReboot, $_.Installed, $_.NotApplicable, $_.LastSync, ($_.GroupNames -replace '\|','-')
  }

  return ($md -join "`r`n")
}

function New-HtmlReport {
  param($Markdown)

  # Simple Markdown -> HTML for headings, tables, bold. Not a full parser.
  $html = $Markdown `
    -replace '^\# (.+)$','<h1>$1</h1>' `
    -replace '^\#\# (.+)$','<h2>$1</h2>' `
    -replace '\*\*(.+?)\*\*','<strong>$1</strong>' `
    -replace '^\|(.*)\|$','<tr><td>$1</td></tr>'

  # Fix table row cells into <td>
  $html = ($html -split "`r`n") | ForEach-Object {
    if ($_ -match '^<tr><td>') {
      $cells = ($_ -replace '^<tr><td>','' -replace '</td></tr>$','').Split('|').Trim()
      '<tr>' + ($cells | ForEach-Object { "<td>$($_)</td>" }) -join '' + '</tr>'
    } else { $_ }
  } | Out-String

  @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>WSUS Weekly Patch Report</title>
  <style>
    body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
    h1 { margin-bottom: 0.2rem; }
    table { border-collapse: collapse; width: 100%; margin: 1rem 0; }
    td, th { border: 1px solid #ddd; padding: 6px 8px; }
    th { background: #f2f2f2; text-align: left; }
    tr:nth-child(even) { background: #fafafa; }
    code { background:#f3f3f3; padding:2px 4px; }
  </style>
</head>
<body>
$(
# convert header rows of tables
  ($html -split "`r`n") | ForEach-Object {
    if ($_ -match '^\|') { $_ } else { $_ }
  } | Out-String
)
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

function Send-Email {
  param($From,$To,$Subject,$Html,$SmtpServer,$SmtpPort)
  if (-not ($From -and $To -and $SmtpServer)) { return }
  $msg = New-Object System.Net.Mail.MailMessage
  $msg.From = $From
  $To.Split(',') | ForEach-Object { [void]$msg.To.Add($_.Trim()) }
  $msg.Subject = $Subject
  $msg.IsBodyHtml = $true
  $msg.Body = $Html
  $client = New-Object System.Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
  $client.Send($msg)
  $msg.Dispose()
  $client.Dispose()
}

# -------- Main --------
$wsus = Connect-WSUS
$computers = Get-ComputerScope -Server $wsus -TargetGroup $TargetGroup
$rows = Get-ComputerCompliance -Server $wsus -Computers $computers
$summary = Summarize-Fleet -Rows $rows

$scopeName = if ($TargetGroup) { "Group: $TargetGroup" } else { "All Computers" }
$md = New-MarkdownReport -ScopeName $scopeName -Summary $summary -Rows $rows
$mdPath = Join-Path $OutDir ("WSUS-Weekly-{0}.md" -f (Get-Date -Format 'yyyyMMdd'))
$md | Out-File -FilePath $mdPath -Encoding utf8

$html = New-HtmlReport -Markdown $md
$htmlPath = Join-Path $OutDir ("WSUS-Weekly-{0}.html" -f (Get-Date -Format 'yyyyMMdd'))
$html | Out-File -FilePath $htmlPath -Encoding utf8

Publish-Teams -WebhookUrl $TeamsWebhook -Title "WSUS Weekly Patch Report" -Markdown $md
Send-Email -From $EmailFrom -To $EmailTo -Subject $EmailSubject -Html $html -SmtpServer $SmtpServer -SmtpPort $SmtpPort

Write-Host "Markdown: $mdPath"
Write-Host "HTML    : $htmlPath"
