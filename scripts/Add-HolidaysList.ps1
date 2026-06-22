<#
.SYNOPSIS
  Creates the PM_Holidays list the task generator reads to skip company
  holidays, and seeds a starter set for the rest of 2026. Idempotent: safe to
  re-run; it won't duplicate the list, columns, or already-present dates.

.DESCRIPTION
  Columns:
    Title       (built-in)  - the holiday name, e.g. "Independence Day (observed)"
    HolidayDate (Date only)  - the date generation should be skipped
    IsActive    (Yes/No)     - uncheck to disable a holiday without deleting it

  To add/remove holidays later you DON'T need this script — just edit the
  PM_Holidays list directly in SharePoint (Site contents -> PM_Holidays).
  The generator skips any active row whose HolidayDate matches the run day
  (Eastern time). Weekends are skipped automatically and don't need entries.

.NOTES
  Requires PnP.PowerShell + interactive sign-in.
#>

$ErrorActionPreference = "Stop"

$SiteUrl  = "https://vanrockre.sharepoint.com/sites/NewshirePM"
$ClientId = "63567714-59eb-4d4f-b3f0-f827e58d9a59"  # CAHP Provisioning Shell — interactive PnP sign-in
$ListName = "PM_Holidays"

try {
  Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
} catch {
  Write-Host "Sign-in failed or timed out: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}

# --- list ---
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $list) {
  $list = New-PnPList -Title $ListName -Template GenericList
  Write-Host "Created list '$ListName'." -ForegroundColor Green
} else {
  Write-Host "List '$ListName' already exists." -ForegroundColor Yellow
}

# --- fields ---
$existing = Get-PnPField -List $ListName | Select-Object -ExpandProperty InternalName
if ($existing -notcontains "HolidayDate") {
  Add-PnPField -List $ListName -DisplayName "Holiday Date" -InternalName "HolidayDate" -Type DateTime -AddToDefaultView | Out-Null
  Set-PnPField -List $ListName -Identity "HolidayDate" -Values @{ DisplayFormat = 0 } | Out-Null  # 0 = Date only
  Write-Host "Added 'HolidayDate' (Date only)." -ForegroundColor Green
}
if ($existing -notcontains "IsActive") {
  Add-PnPField -List $ListName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView | Out-Null
  Write-Host "Added 'IsActive' (Yes/No)." -ForegroundColor Green
}

# --- seed: NewShire's standard 10 holidays, 2026-2028 ---
# Floating holidays (Good Friday, Memorial/Labor Day, Thanksgiving) are computed
# per year. Lines marked (WEEKEND) fall on a Sat/Sun in that year — the generator
# already skips weekends, so those rows are harmless but do nothing on their own.
# If NewShire OBSERVES those on an adjacent weekday, change the date to that
# weekday (or tell me your rule and I'll bake it in).
$seed = @(
  # ---- rest of 2026 ----
  @{ Name = "Day Before Independence Day"; Date = "2026-07-03" }
  @{ Name = "Independence Day";            Date = "2026-07-04" }  # (WEEKEND - Sat)
  @{ Name = "Labor Day";                   Date = "2026-09-07" }
  @{ Name = "Thanksgiving Day";            Date = "2026-11-26" }
  @{ Name = "Day After Thanksgiving";      Date = "2026-11-27" }
  @{ Name = "Christmas Day";               Date = "2026-12-25" }
  @{ Name = "Day After Christmas";         Date = "2026-12-26" }  # (WEEKEND - Sat)
  # ---- 2027 ----
  @{ Name = "New Year's Day";              Date = "2027-01-01" }
  @{ Name = "Good Friday";                 Date = "2027-03-26" }
  @{ Name = "Memorial Day";                Date = "2027-05-31" }
  @{ Name = "Day Before Independence Day"; Date = "2027-07-03" }  # (WEEKEND - Sat)
  @{ Name = "Independence Day";            Date = "2027-07-04" }  # (WEEKEND - Sun)
  @{ Name = "Labor Day";                   Date = "2027-09-06" }
  @{ Name = "Thanksgiving Day";            Date = "2027-11-25" }
  @{ Name = "Day After Thanksgiving";      Date = "2027-11-26" }
  @{ Name = "Christmas Day";               Date = "2027-12-25" }  # (WEEKEND - Sat)
  @{ Name = "Day After Christmas";         Date = "2027-12-26" }  # (WEEKEND - Sun)
  # ---- 2028 ----
  @{ Name = "New Year's Day";              Date = "2028-01-01" }  # (WEEKEND - Sat)
  @{ Name = "Good Friday";                 Date = "2028-04-14" }
  @{ Name = "Memorial Day";                Date = "2028-05-29" }
  @{ Name = "Day Before Independence Day"; Date = "2028-07-03" }
  @{ Name = "Independence Day";            Date = "2028-07-04" }
  @{ Name = "Labor Day";                   Date = "2028-09-04" }
  @{ Name = "Thanksgiving Day";            Date = "2028-11-23" }
  @{ Name = "Day After Thanksgiving";      Date = "2028-11-24" }
  @{ Name = "Christmas Day";               Date = "2028-12-25" }
  @{ Name = "Day After Christmas";         Date = "2028-12-26" }
)

# Existing dates (so re-runs don't duplicate)
$have = @{}
Get-PnPListItem -List $ListName -PageSize 500 -Fields "HolidayDate" | ForEach-Object {
  $d = $_.FieldValues.HolidayDate
  if ($d) { $have[([datetime]$d).ToString("yyyy-MM-dd")] = $true }
}

$added = 0
foreach ($h in $seed) {
  if ($have.ContainsKey($h.Date)) { continue }
  Add-PnPListItem -List $ListName -Values @{
    Title       = $h.Name
    HolidayDate = (Get-Date $h.Date)
    IsActive    = $true
  } | Out-Null
  Write-Host "  seeded $($h.Date)  $($h.Name)" -ForegroundColor Gray
  $added++
}

Write-Host "`nDone. Seeded $added new holiday row(s). Edit the PM_Holidays list anytime to add/remove dates." -ForegroundColor Green
