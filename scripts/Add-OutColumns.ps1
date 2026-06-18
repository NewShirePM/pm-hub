<#
.SYNOPSIS
  Adds the OutSince and OutUntil date columns to the Employees list so PM Hub
  can stamp when someone is marked Out / back In (absence tracking).

.NOTES
  - Both are Date-only columns (no time-of-day).
  - Idempotent: skips any field that already exists.
  - Stops on the first real error so it never prints a false success.
  - Requires PnP.PowerShell + interactive sign-in.
#>

$ErrorActionPreference = "Stop"

$SiteUrl  = "https://vanrockre.sharepoint.com/sites/NewshirePM"
$ClientId = "63567714-59eb-4d4f-b3f0-f827e58d9a59"  # CAHP Provisioning Shell — set up for interactive PnP sign-in
$ListName = "Employees"
$Fields   = @("OutSince", "OutUntil")

try {
  Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
} catch {
  Write-Host "Sign-in failed or timed out: $($_.Exception.Message)" -ForegroundColor Red
  Write-Host "Re-run the script and complete the browser sign-in prompt." -ForegroundColor Red
  exit 1
}

$existing = Get-PnPField -List $ListName | Select-Object -ExpandProperty InternalName

foreach ($name in $Fields) {
  if ($existing -contains $name) {
    Write-Host "Field '$name' already exists on '$ListName' — skipping." -ForegroundColor Yellow
    continue
  }

  Add-PnPField -List $ListName `
    -DisplayName $name `
    -InternalName $name `
    -Type DateTime `
    -AddToDefaultView | Out-Null

  # Make it Date-only (DisplayFormat 0 = DateOnly, 1 = DateTime).
  Set-PnPField -List $ListName -Identity $name -Values @{ DisplayFormat = 0 } | Out-Null

  Write-Host "Added '$name' (Date only) to '$ListName'." -ForegroundColor Green
}

Write-Host "Done." -ForegroundColor Green
