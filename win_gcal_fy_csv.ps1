# =======================================
# Google Calendar ICS → CSV (Fiscal Year)
# URLハードコード版（ここを書き換える）
# =======================================

$IcsUrl = "https://calendar.google.com/calendar/ical/XXXXXXXX/basic.ics"

$OutCsv = "gunma_FY.csv"

# ---- FY range (Apr 1 to Mar 31) ----
$now = Get-Date
if ($now.Month -lt 4) {
  $fyStartYear = $now.Year - 1
  $fyEndYear   = $now.Year
} else {
  $fyStartYear = $now.Year
  $fyEndYear   = $now.Year + 1
}
$fyStart = [int]("{0}0401" -f $fyStartYear)  # inclusive
$fyEndEx = [int]("{0}0401" -