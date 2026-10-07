# A stand-in for PowerPoint, used only by the automated checks.
#
# The helper loads this file when SLIDESIZE_HELPER_TEST_ENGINE points at it.
# It lets the queue, the time limits, cancelling and recovery be exercised on
# a machine with no PowerPoint. It converts nothing. The "presentations" it
# opens are small JSON files such as {"slides":6,"hidden":[3]} and the PDFs
# it writes are blank pages.
#
# It reports its own name, and the page refuses any result whose engine is
# not Microsoft PowerPoint, so a stand-in can never pass for the real thing.

function Engine-Name { return 'SlideSize test double (not PowerPoint)' }
function Engine-Installed { return ($env:SLIDESIZE_FAKE_NOT_INSTALLED -ne '1') }
function Engine-RunningBefore { return $false }
function Engine-InstalledFonts { return @('Arial', 'Carlito', 'Test Sans') }
function Engine-Start([string]$root) {
  if ($env:SLIDESIZE_FAKE_START_FAILS -eq '1') { throw 'The stand-in was told to fail at start.' }
  return @{ app = @{ fake = $true }; info = @{ version = '0.0'; build = 'test'; userPresentations = 0 } }
}
function Engine-Stop($app, [bool]$owned) { }
function Engine-Recover([bool]$owned) { return '' }

function New-TestPdf([string]$path, [int]$pages, [double]$w, [double]$h) {
  $sb = New-Object System.Text.StringBuilder
  $offsets = New-Object System.Collections.Generic.List[int]
  [void]$sb.Append("%PDF-1.4`n")
  $kids = ''
  for ($i = 0; $i -lt $pages; $i++) { $kids += (3 + $i).ToString() + ' 0 R ' }
  $objs = @()
  $objs += '<< /Type /Catalog /Pages 2 0 R >>'
  $objs += ('<< /Type /Pages /Count ' + $pages + ' /Kids [ ' + $kids + '] >>')
  for ($i = 0; $i -lt $pages; $i++) {
    $objs += ('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + $w.ToString([System.Globalization.CultureInfo]::InvariantCulture) + ' ' + $h.ToString([System.Globalization.CultureInfo]::InvariantCulture) + '] /Resources << >> >>')
  }
  for ($i = 0; $i -lt $objs.Count; $i++) {
    $offsets.Add($sb.Length)
    [void]$sb.Append(($i + 1).ToString() + " 0 obj`n" + $objs[$i] + "`nendobj`n")
  }
  $xref = $sb.Length
  [void]$sb.Append("xref`n0 " + ($objs.Count + 1) + "`n0000000000 65535 f `n")
  foreach ($o in $offsets) { [void]$sb.Append($o.ToString('D10') + " 00000 n `n") }
  [void]$sb.Append("trailer`n<< /Size " + ($objs.Count + 1) + " /Root 1 0 R /ID [<00112233445566778899aabbccddeeff> <00112233445566778899aabbccddeeff>] >>`nstartxref`n" + $xref + "`n%%EOF`n")
  [System.IO.File]::WriteAllText($path, $sb.ToString(), (New-Object System.Text.ASCIIEncoding))
}

function Engine-Convert($app, $t, [string]$inPath, [string]$pdfPath, [string]$tmpPdf, [string]$refDir) {
  $o = $t.options
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  Set-Phase 'opening'
  $deck = [System.IO.File]::ReadAllText($inPath) | ConvertFrom-Json
  if ($deck.stall) { Start-Sleep -Seconds 3600 }
  if ($deck.password) { throw (New-Failure 'password' 'The stand-in deck is password protected.') }
  if ($deck.fail) { throw (New-Failure 'open-failed' ([string]$deck.fail)) }
  if ($deck.slow) { Start-Sleep -Milliseconds ([int]$deck.slow) }
  $n = [int]$deck.slides
  $hidden = @()
  if ($deck.hidden) { $hidden = @($deck.hidden | ForEach-Object { [int]$_ }) }
  $w = 960.0; $h = 540.0
  if ($deck.w) { $w = [double]$deck.w; $h = [double]$deck.h }
  $from = 1; $to = $n
  if ($o.range) { $from = [System.Math]::Max(1, [int]$o.range[0]); $to = [System.Math]::Min($n, [int]$o.range[1]) }
  $map = @()
  for ($i = $from; $i -le $to; $i++) { if ($o.includeHidden -or ($hidden -notcontains $i)) { $map += $i } }
  if ($map.Count -eq 0) { throw (New-Failure 'no-slides' 'No slides are left to export with this slide range and the hidden slides left out.') }
  Set-Phase 'exporting'
  New-TestPdf $tmpPdf $map.Count $w $h
  if (Test-Path -LiteralPath $pdfPath) {
    if (-not $t.overwrite) { [System.IO.File]::Delete($tmpPdf); throw (New-Failure 'exists' ('The output file already exists: ' + [System.IO.Path]::GetFileName($pdfPath))) }
    [System.IO.File]::Delete($pdfPath)
  }
  [System.IO.File]::Move($tmpPdf, $pdfPath)
  $r = @{
    facts = @{ slides = $n; hidden = @($hidden); slideWidthPt = $w; slideHeightPt = $h; fonts = @(@{ name = 'Carlito'; embedded = $false; embeddable = $true }); readOnly = $true }
    export = @{ method = 'ExportAsFixedFormat2'; notApplied = @(); placeholders = @(); pageMap = @($map); removedSlides = @() }
    timing = @{ openMs = 1; exportMs = 1 }
    pdf = @{ bytes = [long](Get-Item -LiteralPath $pdfPath).Length }
    reference = $null
  }
  foreach ($p in @($o.placeholders)) { if ($p) { $r.export.placeholders += @{ slide = [int]$p.slide; ok = $true } } }
  if ($o.reference) {
    Set-Phase 'reference'
    Ensure-Dir $refDir
    $want = @($map)
    if ($o.reference.mode -eq 'sample') { $want = @(Select-Sample $want) }
    foreach ($sn in $want) { [System.IO.File]::WriteAllText((Join-Path $refDir ('slide-' + ('{0:D4}' -f [int]$sn) + '.png')), 'not a real picture') }
    $r.reference = @{ dir = ('ref/' + $t.id); slides = @($want); width = 1280; height = 720; pattern = 'slide-%04d.png' }
  }
  $r.timing.totalMs = [int]$sw.ElapsedMilliseconds
  return $r
}
