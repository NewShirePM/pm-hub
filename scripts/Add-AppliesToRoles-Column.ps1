<#
.SYNOPSIS
  Adds the AppliesToRoles column to the PM_RecurringTasks list so recurring
  tasks can be assigned to a PMHubRole (e.g. "pm,maintenance") instead of
  hand-picking individuals.

.NOTES
  - Single line of text; stores a comma-joined list of lowercased PMHubRole
    values (matches what the PM Hub admin UI writes).
  - Idempotent: skips creation if the field already exists.
  - Requires the PnP.PowerShell module and interactive sign-in.
#>

$SiteUrl  = "https://vanrockre.sharepoint.com/sites/NewshirePM"
$ClientId = "32e75ffa-747a-4cf0-8209-6a19150c4547"  # same app reg the PM Hub uses
$ListName = "PM_RecurringTasks"
$Internal = "AppliesToRoles"

Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

$existing = Get-PnPField -List $ListName | Where-Object { $_.InternalName -eq $Internal }
if ($existing) {
  Write-Host "Field '$Internal' already exists on '$ListName' — nothing to do." -ForegroundColor Yellow
  return
}

Add-PnPField -List $ListName `
  -DisplayName "Applies To Roles" `
  -InternalName $Internal `
  -Type Text `
  -AddToDefaultView | Out-Null

Write-Host "Added '$Internal' (Single line of text) to '$ListName'." -ForegroundColor Green
