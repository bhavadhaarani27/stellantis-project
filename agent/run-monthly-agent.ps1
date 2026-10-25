param(
  [Parameter(Mandatory = $true)]
  [string]$PowerPointPath,
  [Parameter(Mandatory = $false)]
  [string]$Month,
  [Parameter(Mandatory = $false)]
  [string]$Year,
  [Parameter(Mandatory = $false)]
  [string]$OutputRoot
)

$agentRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $agentRoot
$config = Get-Content (Join-Path $agentRoot "config.json") -Raw | ConvertFrom-Json

if (-not $Month) { $Month = (Get-Date).ToString("MMMM") }
if (-not $Year) { $Year = (Get-Date).Year.ToString() }
if (-not $OutputRoot) { $OutputRoot = Join-Path $projectRoot "monthly-output" }

$templatePath = Join-Path $projectRoot ($config.template -replace '^\.\\', '')
$resolvedPowerPoint = (Resolve-Path $PowerPointPath).Path
if (-not (Test-Path -LiteralPath $templatePath)) {
  throw "Reference template not found: $templatePath"
}
$reviewFolder = Join-Path $OutputRoot ("{0}-{1}" -f $Month, $Year)
$slidesFolder = Join-Path $reviewFolder "slides"
New-Item -ItemType Directory -Path $slidesFolder -Force | Out-Null

function ConvertTo-HtmlText([string]$Value) {
  if ($null -eq $Value) { return "" }
  return [System.Net.WebUtility]::HtmlEncode($Value.Trim())
}

function Get-SectionName([int]$Index, [int]$Total) {
  if ($Index -eq 1) { return "home" }
  if ($Index -eq 2) { return "agenda" }
  if ($Index -le 7) { return "eeq" }
  if ($Index -le 18) { return "fqm" }
  if ($Index -le 29) { return "hwe" }
  if ($Index -le 40) { return "eeav" }
  if ($Index -le 42) { return "fea" }
  if ($Index -le 45) { return "nea" }
  if ($Index -eq 46) { return "closing" }
  if ($Index -eq 47) { return "target" }
  return "additional"
}

function Get-SectionLabel([string]$Id) {
  switch ($Id) {
    "home" { return "Home" }
    "agenda" { return "Agenda" }
    "eeq" { return "EEQ Dashboard" }
    "fqm" { return "FQM / EEHD" }
    "hwe" { return "HWE" }
    "eeav" { return "EEAV" }
    "fea" { return "FEA Quality" }
    "nea" { return "NEA 1.2 J4U/D85" }
    "closing" { return "Closing" }
    "target" { return "Quality Target 2027" }
    default { return "Additional Slides" }
  }
}

$powerPoint = New-Object -ComObject PowerPoint.Application
$presentation = $null
$slideManifest = @()
try {
  $presentation = $powerPoint.Presentations.Open($resolvedPowerPoint, 0, 1, 0)
  $slideCount = $presentation.Slides.Count

  for ($index = 1; $index -le $slideCount; $index++) {
    $slide = $presentation.Slides.Item($index)
    $imageName = "slide-{0:D3}.png" -f $index
    $imagePath = Join-Path $slidesFolder $imageName
    $slide.Export($imagePath, "PNG", [int]$config.exportWidth, [int]$config.exportHeight)

    $texts = @()
    foreach ($shape in $slide.Shapes) {
      try {
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
          $text = ($shape.TextFrame.TextRange.Text -replace '\s+', ' ').Trim()
          if ($text) { $texts += $text }
        }
      } catch {
        # Some PowerPoint shape types do not expose text properties.
      }
    }
    $slideManifest += [pscustomobject]@{
      slide = $index
      image = "slides/$imageName"
      section = Get-SectionName $index $slideCount
      title = if ($texts.Count) { $texts[0] } else { "Slide $index" }
      text = $texts
    }
  }
} finally {
  if ($presentation) { $presentation.Close() }
  $powerPoint.Quit()
}

$template = Get-Content $templatePath -Raw
$style = ([regex]::Match($template, '(?s)<style>(.*?)</style>')).Groups[1].Value
$scriptMatches = [regex]::Matches($template, '(?s)<script>(.*?)</script>')
$navigationScript = $scriptMatches[$scriptMatches.Count - 1].Groups[1].Value

$sectionOrder = @("home", "agenda", "eeq", "fqm", "hwe", "eeav", "fea", "nea", "closing", "target", "additional")
$usedSections = @($slideManifest | ForEach-Object { $_.section } | Select-Object -Unique)
$navLinks = New-Object System.Collections.Generic.List[string]
$sectionMarkup = New-Object System.Collections.Generic.List[string]

foreach ($sectionId in $sectionOrder) {
  if ($usedSections -notcontains $sectionId) { continue }
  $sectionSlides = @($slideManifest | Where-Object { $_.section -eq $sectionId })
  $firstSlide = $sectionSlides[0].slide
  $lastSlide = $sectionSlides[$sectionSlides.Count - 1].slide
  $range = if ($firstSlide -eq $lastSlide) { "$firstSlide" } else { "$firstSlide-$lastSlide" }
  $navLinks.Add("<a class=""nav-link"" href=""#$sectionId"" data-target=""$sectionId""><span>$(Get-SectionLabel $sectionId)</span><small>$range</small></a>")
  $cards = foreach ($slideInfo in $sectionSlides) {
    "<article class=""slide-card"" data-slide=""$($slideInfo.slide)""><div class=""slide-number"">Slide $($slideInfo.slide)</div><img src=""$($slideInfo.image)"" alt=""$(ConvertTo-HtmlText $slideInfo.title)""></article>"
  }
  $sectionMarkup.Add("<section id=""$sectionId"" class=""deck-section""><div class=""section-heading""><h2>$(Get-SectionLabel $sectionId)</h2><span>Slides $range</span></div>$($cards -join '')</section>")
}

$html = @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>EHE Quality Review | $Month $Year</title><style>$style</style></head><body><aside class="sidebar"><div class="brand"><b>EHE QUALITY REVIEW</b><span>Interactive presentation webpage</span></div>$($navLinks -join '')</aside><div class="slide-controls"><button id="prevSlide" aria-label="Previous slide">&#8249;</button><div class="slide-counter" id="slideCounter">1 / $slideCount</div><button id="nextSlide" aria-label="Next slide">&#8250;</button></div><main>$($sectionMarkup -join '')</main><script>$navigationScript</script></body></html>
"@

$outputHtml = Join-Path $reviewFolder "index.html"
Set-Content -LiteralPath $outputHtml -Value $html -Encoding UTF8
$manifestPath = Join-Path $reviewFolder "slide-manifest.json"
[pscustomobject]@{
  sourcePowerPoint = $resolvedPowerPoint
  month = $Month
  year = $Year
  generatedAt = (Get-Date).ToString("o")
  slideCount = $slideCount
  slides = $slideManifest
} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Output "Generated: $outputHtml"
Write-Output "Manifest:  $manifestPath"
Write-Output "Slides:    $slideCount"
