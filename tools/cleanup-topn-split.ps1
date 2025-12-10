param(
  [string]$Root = "C:\Users\14ugy\Projects\topN_distributor_v1\data\output\split",
  [switch]$DryRun = $true
)

# use default console encoding; avoid emoji/Japanese to prevent mojibake
$ErrorActionPreference = "Stop"

Write-Host "Scanning: $Root ..."

$targets = Get-ChildItem -Path $Root -Recurse -File |
  Where-Object { $_.Name -match '温惣菜|冷総菜' }

if ($targets.Count -eq 0) {
  Write-Host "No matching files."
  exit
}

Write-Host "Targets:"
$targets | Select-Object FullName, LastWriteTime, Length | Format-Table -AutoSize

if ($DryRun) {
  Write-Host ""
  Write-Host "DryRun: no files were deleted. To delete, run with -DryRun:`$false"
} else {
  $targets | Remove-Item -Force
  Write-Host ""
  Write-Host ("Deleted: {0} file(s)" -f $targets.Count)
}
