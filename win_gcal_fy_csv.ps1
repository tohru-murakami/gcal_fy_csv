# ===============================
# Google Calendar ICS → CSV (FY)
# Sorted by start datetime
# ===============================

# ===== ICS URL（ここを書き換える）=====
$IcsUrl = "https://calendar.google.com/calendar/ical/XXXXXXXX/basic.ics"
# ======================================

# 出力先（このスクリプトと同じフォルダ）
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir
$OutCsv = Join-Path $ScriptDir "gunma_FY.csv"

# 今年度（4/1〜翌年3/31）
$now = Get-Date
if ($now.Month -lt 4) {
  $fyStartYear = $now.Year - 1
  $fyEndYear   = $now.Year
} else {
  $fyStartYear = $now.Year
  $fyEndYear   = $now.Year + 1
}
$fyStart = [int]("{0}0401" -f $fyStartYear)  # inclusive
$fyEndEx = [int]("{0}0401" -f $fyEndYear)    # exclusive

Write-Host ""
Write-Host "Google Calendar → CSV Export"
Write-Host ("Fiscal Year: {0}-04-01 → {1}-03-31" -f $fyStartYear, $fyEndYear)
Write-Host ""

# ICSをダウンロード
$tmp = Join-Path $env:TEMP ("gcal_" + [guid]::NewGuid().ToString() + ".ics")

Write-Host "Downloading ICS..."
Invoke-WebRequest -Uri $IcsUrl -OutFile $tmp
Write-Host "Downloaded."
Write-Host ""

# RFC5545 折り返し行を展開（行頭スペースは前行に連結）
$rawLines = Get-Content -LiteralPath $tmp -Encoding UTF8
$lines = New-Object System.Collections.Generic.List[string]
foreach ($line in $rawLines) {
  $l = $line.TrimEnd("`r")
  if ($l.StartsWith(" ") -and $lines.Count -gt 0) {
    $lines[$lines.Count-1] = $lines[$lines.Count-1] + $l.Substring(1)
  } else {
    $lines.Add($l)
  }
}

function Unescape-Ics([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace '\\n', ' / '
  $s = $s -replace '\\,', ','
  $s = $s -replace '\\;', ';'
  $s = $s -replace '\\\\', '\'
  return $s
}

function Get-ValueAfterColon([string]$line) {
  $i = $line.IndexOf(":")
  if ($i -lt 0) { return "" }
  return $line.Substring($i + 1)
}

function Format-Dt([string]$dt) {
  if ([string]::IsNullOrEmpty($dt)) { return "" }
  # 20260226T090000Z / 20260226T090000 / 20260226
  if ($dt -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})') {
    return "{0}-{1}-{2} {3}:{4}" -f $matches[1],$matches[2],$matches[3],$matches[4],$matches[5]
  }
  if ($dt -match '^(\d{4})(\d{2})(\d{2})$') {
    return "{0}-{1}-{2}" -f $matches[1],$matches[2],$matches[3]
  }
  return $dt
}

function YmdInt([string]$dt) {
  if ([string]::IsNullOrEmpty($dt)) { return 0 }
  if ($dt.Length -ge 8 -and $dt.Substring(0,8) -match '^\d{8}$') {
    return [int]$dt.Substring(0,8)
  }
  return 0
}

# VEVENT解析（RRULE展開なし）
$events = @()
$inEvent = $false
$cur = @{ DTSTART=""; DTEND=""; SUMMARY=""; LOCATION=""; DESCRIPTION="" }

foreach ($line in $lines) {
  if ($line -eq "BEGIN:VEVENT") {
    $inEvent = $true
    $cur = @{ DTSTART=""; DTEND=""; SUMMARY=""; LOCATION=""; DESCRIPTION="" }
    continue
  }
  if ($line -eq "END:VEVENT") {
    if ($inEvent) {
      $d = YmdInt $cur.DTSTART
      if ($d -ge $fyStart -and $d -lt $fyEndEx) {
        $events += [pscustomobject]@{
          start       = (Format-Dt $cur.DTSTART)
          end         = (Format-Dt $cur.DTEND)
          summary     = (Unescape-Ics $cur.SUMMARY)
          location    = (Unescape-Ics $cur.LOCATION)
          description = (Unescape-Ics $cur.DESCRIPTION)
        }
      }
    }
    $inEvent = $false
    continue
  }

  if (-not $inEvent) { continue }

  if ($line -match '^DTSTART')      { $cur.DTSTART      = Get-ValueAfterColon $line; continue }
  if ($line -match '^DTEND')        { $cur.DTEND        = Get-ValueAfterColon $line; continue }
  if ($line -match '^SUMMARY:')     { $cur.SUMMARY      = Get-ValueAfterColon $line; continue }
  if ($line -match '^LOCATION:')    { $cur.LOCATION     = Get-ValueAfterColon $line; continue }
  if ($line -match '^DESCRIPTION:') { $cur.DESCRIPTION  = Get-ValueAfterColon $line; continue }
}

# ===== 日時順ソート（startで昇順）=====
$eventsSorted = $events | Sort-Object -Property start

# CSV出力（UTF-8 BOM付き：Excel対策）
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvLines = $eventsSorted | ConvertTo-Csv -NoTypeInformation
[System.IO.File]::WriteAllLines($OutCsv, $csvLines, $utf8Bom)

Write-Host "Created:"
Write-Host $OutCsv
Write-Host ("Rows output: {0}" -f $eventsSorted.Count)

Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue