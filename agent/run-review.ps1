param(
  [string]$Month,
  [string]$Year
)

$agentRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $agentRoot
$inputFolder = Join-Path $projectRoot "monthly-input"
$processor = Join-Path $agentRoot "run-monthly-agent.ps1"

New-Item -ItemType Directory -Path $inputFolder -Force | Out-Null
$files = @(Get-ChildItem -LiteralPath $inputFolder -File | Where-Object {
  $_.Extension -in @(".pptx", ".ppt")
})

if ($files.Count -eq 0) {
  Write-Error "No PowerPoint files were found in: $inputFolder"
  Write-Output "Copy a monthly .pptx file into that folder and run this command again."
  exit 1
}

Write-Output ""
Write-Output "Available PowerPoint files:"
for ($index = 0; $index -lt $files.Count; $index++) {
  Write-Output ("[{0}] {1}" -f ($index + 1), $files[$index].Name)
}

$selection = Read-Host "Enter the number of the presentation to process"
$selectedNumber = 0
if (-not [int]::TryParse($selection, [ref]$selectedNumber) -or $selectedNumber -lt 1 -or $selectedNumber -gt $files.Count) {
  Write-Error "Please enter one of the numbers shown in the list."
  exit 1
}

$selectedFile = $files[$selectedNumber - 1]
$months = @("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
if (-not $Month) { $Month = Read-Host "Enter the month, for example July" }
if ($months -notcontains $Month) {
  Write-Error "Month must be a full month name, such as January or July."
  exit 1
}

if (-not $Year) { $Year = Read-Host "Enter the year, for example 2026" }
$yearNumber = 0
if (-not [int]::TryParse($Year, [ref]$yearNumber) -or $yearNumber -lt 2000 -or $yearNumber -gt 2100) {
  Write-Error "Year must be a four-digit number between 2000 and 2100."
  exit 1
}

Write-Output ""
Write-Output ("Processing: {0}" -f $selectedFile.Name)
Write-Output ("Period:     {0} {1}" -f $Month, $Year)
& $processor -PowerPointPath $selectedFile.FullName -Month $Month -Year $Year -OutputRoot (Join-Path $projectRoot "monthly-output")
if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}

$generatedPage = Join-Path $projectRoot ("monthly-output\{0}-{1}\index.html" -f $Month, $Year)
Write-Output ""
Write-Output "Completed successfully."
Write-Output ("Webpage: {0}" -f $generatedPage)
Write-Output "Open it with Microsoft Edge using the command below:"
Write-Output ('& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" "{0}"' -f $generatedPage)
