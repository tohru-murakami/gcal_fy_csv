#!/bin/bash
set -euo pipefail

# Finderダブルクリック対策（必須）
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# ===== ICS URL（ここを書き換える） =====

ICS_URL="https://calendar.google.com/calendar/ical/XXXXXXXX/basic.ics"

# =======================================

OUT="$SCRIPT_DIR/gunma_FY.csv"
TMP="$(mktemp)"

trap 'rm -f "$TMP"' EXIT

echo
echo "Google Calendar → CSV Export"
echo

echo "Downloading calendar..."

curl -fsSL "$ICS_URL" -o "$TMP"

echo "Downloaded."

echo

# ===== 年度計算（4月〜翌年3月） =====

Y=$(date +%Y)
M=$(date +%m)
M=$((10#$M))

if (( M < 4 )); then
 FY_START_YEAR=$((Y-1))
 FY_END_YEAR=$Y
else
 FY_START_YEAR=$Y
 FY_END_YEAR=$((Y+1))
fi

FY_START="${FY_START_YEAR}0401"
FY_END="${FY_END_YEAR}0401"

echo "Fiscal Year:"
echo "${FY_START_YEAR}-04-01 → ${FY_END_YEAR}-03-31"

echo

echo "Creating CSV..."

echo "start,end,summary,location,description" > "$OUT"

awk -v FY_START="$FY_START" -v FY_END="$FY_END" '

function ymd(x){
 return substr(x,1,8)
}

function fmt(x){

 if(length(x)>=13)
  return substr(x,1,4) "-" substr(x,5,2) "-" substr(x,7,2) " " substr(x,10,2) ":" substr(x,12,2)

 else
  return substr(x,1,4) "-" substr(x,5,2) "-" substr(x,7,2)
}

function unesc(x){

 gsub(/\\n/," / ",x)
 gsub(/\\,/ ,",",x)
 gsub(/\\;/ ,";",x)

 return x
}

function esc(x){

 x=unesc(x)

 gsub(/"/,"\"\"",x)

 return "\"" x "\""
}

BEGIN{
event=0
}

/BEGIN:VEVENT/{
event=1
start=end=sum=loc=desc=""
}

/END:VEVENT/{

d=ymd(start)

if(d>=FY_START && d<FY_END){

printf "%s,%s,%s,%s,%s\n",

fmt(start),
fmt(end),
esc(sum),
esc(loc),
esc(desc)

}

event=0
}

/^DTSTART/ && event{
sub(/.*:/,"")
start=$0
}

/^DTEND/ && event{
sub(/.*:/,"")
end=$0
}

/^SUMMARY:/ && event{
sub(/^SUMMARY:/,"")
sum=$0
}

/^LOCATION:/ && event{
sub(/^LOCATION:/,"")
loc=$0
}

/^DESCRIPTION:/ && event{
sub(/^DESCRIPTION:/,"")
desc=$0
}

' "$TMP" >> "$OUT"

echo
echo "Created:"
echo "$OUT"
echo

echo "Rows:"
wc -l "$OUT"

echo

read -p "Press Enter to close"