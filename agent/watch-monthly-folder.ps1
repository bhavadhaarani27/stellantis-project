param(
  [Parameter(Mandatory = $false)]
  [string]$InputFolder,
  [Parameter(Mandatory = $false)]
  [string]$OutputRoot
)

$agentRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $agentRoot
$config = Get-Content (Join-Path $agentRoot "config.json") -Raw | ConvertFrom-Json
if (-not $InputFolder) { $InputFolder = Join-Path $projectRoot ($config.inputFolder -replace '^\.\\', '') }
if (-not $OutputRoot) { $OutputRoot = Join-Path $projectRoot ($config.outputFolder -replace '^\.\\', '') }
New-Item -ItemType Directory -Path $InputFolder -Force | Out-Null

$processed = @{}
Write-Output "Watching $InputFolder for monthly PowerPoint files..."

while ($true) {
  Get-ChildItem -LiteralPath $InputFolder -File -Include *.pptx,*.ppt | ForEach-Object {
    $file = $_
    $stamp = "$($file.FullName)|$($file.LastWriteTimeUtc.Ticks)"
    if ($processed.ContainsKey($file.FullName) -and $processed[$file.FullName] -eq $stamp) { return }

    Start-Sleep -Seconds 3
    try {
      & (Join-Path $agentRoot "run-monthly-agent.ps1") -PowerPointPath $file.FullName -OutputRoot $OutputRoot
      if ($LASTEXITCODE -eq 0) {
        $processed[$file.FullName] = $stamp
        Write-Output "Processed $($file.Name)"
      }
    } catch {
      Write-Warning ("Could not process {0}: {1}" -f $file.Name, $_.Exception.Message)
    }
  }
  Start-Sleep -Seconds ([int]$config.pollSeconds)
}
