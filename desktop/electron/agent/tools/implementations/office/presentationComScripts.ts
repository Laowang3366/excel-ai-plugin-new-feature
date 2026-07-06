export function slideTextShapesScript(): string {
  return `
function Get-ShapeTextInfo($shape) {
  $text = ''; $hasText = $false
  try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = [string]$shape.TextFrame.TextRange.Text; $hasText = $true } } catch {}
  [pscustomobject]@{ name = $shape.Name; index = $shape.Id; type = $shape.Type; hasText = $hasText; text = $text }
}`;
}

export function slideLayoutResolverScript(): string {
  return `
function Normalize-LayoutName($value) {
  if ($null -eq $value) { return '' }
  return ([string]$value).ToLowerInvariant() -replace '[\\s_\\-]+', ''
}
function Get-LayoutAliases($layoutKey) {
  switch (Normalize-LayoutName $layoutKey) {
    'title' { return @('title','titleslide','onlytitle','标题','标题幻灯片') }
    'blank' { return @('blank','空白','空白幻灯片') }
    default { return @('titlebody','titleandcontent','titlecontent','titleandtext','标题和内容','标题与内容','标题和正文','标题与正文') }
  }
}
function Resolve-CustomSlideLayout($presentation, $layoutKey) {
  $aliases = Get-LayoutAliases $layoutKey | ForEach-Object { Normalize-LayoutName $_ }
  foreach ($customLayout in $presentation.SlideMaster.CustomLayouts) {
    if ($aliases -contains (Normalize-LayoutName $customLayout.Name)) { return $customLayout }
  }
  return $null
}`;
}
