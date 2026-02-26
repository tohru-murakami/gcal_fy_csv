$IcsUrl = "https://calendar.google.com/calendar/ical/XXXX/basic.ics"

$OutCsv = "gunma_FY.csv"

Write-Host "Downloading calendar..."

$tmp = Join-Path $env:TEMP "gcal.ics"

Invoke-WebRequest -Uri $IcsUrl -OutFile $tmp

$now = Get-Date

if ($now.Month -lt 4) {
 $fyStartYear = $now.Year - 1
 $fyEndYear = $now.Year
}
else {
 $fyStartYear = $now.Year
 $fyEndYear = $now.Year + 1
}

$fyStart = [int]("{0}0401" -f $fyStartYear)
$fyEnd = [int]("{0}0401" -f $fyEndYear)

Write-Host "Creating CSV..."

$lines = Get-Content $tmp

$events = @()

$event = $false

$start=""
$end=""
$sum=""
$loc=""
$desc=""

foreach($line in $lines){

 if($line -eq "BEGIN:VEVENT"){
  $event=$true
  $start=""
  $end=""
  $sum=""
  $loc=""
  $desc=""
 }

 elseif($line -eq "END:VEVENT"){

  if($start.Length -ge 8){

   $d=[int]$start.Substring(0,8)

   if($d -ge $fyStart -and $d -lt $fyEnd){

    $events+=New-Object PSObject -Property @{

     start = $start.Substring(0,4)+"-"+$start.Substring(4,2)+"-"+$start.Substring(6,2)

     end = $end.Substring(0,4)+"-"+$end.Substring(4,2)+"-"+$end.Substring(6,2)

     summary=$sum

     location=$loc

     description=$desc

    }

   }

  }

  $event=$false

 }

 elseif($event){

  if($line.StartsWith("DTSTART")){
   $start=$line.Split(":")[1]
  }

  elseif($line.StartsWith("DTEND")){
   $end=$line.Split(":")[1]
  }

  elseif($line.StartsWith("SUMMARY")){
   $sum=$line.Split(":")[1]
  }

  elseif($line.StartsWith("LOCATION")){
   $loc=$line.Split(":")[1]
  }

  elseif($line.StartsWith("DESCRIPTION")){
   $desc=$line.Split(":")[1]
  }

 }

}

$events | Export-Csv $OutCsv -Encoding UTF8 -NoTypeInformation

Write-Host ""
Write-Host "Created:"
Write-Host $OutCsv