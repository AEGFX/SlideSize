# SlideSize PowerPoint to PDF helper for Windows.
#
# Drives the copy of Microsoft PowerPoint installed on this computer so that
# slidesize.com can batch convert presentations to PDF. Nothing is installed
# and nothing is sent anywhere. The page and this script exchange small JSON
# files inside the _slidesize folder of the output folder you chose:
#
#   queue\   job tickets written by the page
#   active\  the ticket being worked on
#   done\    one result per ticket, written by this script
#   in\      a temporary copy of each presentation, deleted after use
#   out\     PDFs the page still has to edit
#   ref\     PowerPoint's own pictures of slides, for the fidelity check
#
# The script runs as two processes. This one, the watchdog, hands out
# tickets and keeps time. A second hidden one, the worker, talks to
# PowerPoint. If PowerPoint stalls on a file the watchdog stops the worker,
# records the failure and carries on with the next file.
#
# Presentations you already have open are never closed or changed. Each file
# is converted from a temporary copy that is opened read only, without a
# window, and never saved.
#
# Written for Windows PowerShell 5.1, which every Windows 10 and 11 PC has.
# Keep this file plain ASCII. PowerShell 5.1 misreads anything else.

param(
  [switch]$Worker,
  [string]$Root = '',
  [int]$ParentPid = 0,
  [int]$Owns = 0
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$HelperVersion = '1.0.0'
$Protocol = 1
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$IsWin = ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)

# ---------------------------------------------------------------- files

function Read-Json([string]$path) {
  try {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $text = [System.IO.File]::ReadAllText($path, $Utf8)
    if (-not $text) { return $null }
    return ($text | ConvertFrom-Json)
  } catch { return $null }
}

function Write-JsonAtomic([string]$path, $obj) {
  $json = ConvertTo-Json -InputObject $obj -Depth 12 -Compress
  $tmp = $path + '.' + $PID + '.tmp'
  [System.IO.File]::WriteAllText($tmp, $json, $Utf8)
  for ($i = 0; $i -lt 20; $i++) {
    try {
      if (Test-Path -LiteralPath $path) { [System.IO.File]::Delete($path) }
      [System.IO.File]::Move($tmp, $path)
      return
    } catch { Start-Sleep -Milliseconds 50 }
  }
  try { [System.IO.File]::Delete($tmp) } catch {}
}

function Ensure-Dir([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { [void](New-Item -ItemType Directory -Path $path -Force) }
}

# The invariant culture keeps the colons. Some regional settings would write a full stop in their place.
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
function Format-Iso([System.DateTime]$t) { return $t.ToString('yyyy-MM-ddTHH:mm:ss.fffZ', $Invariant) }
function Now-Iso { return (Format-Iso ([System.DateTime]::UtcNow)) }

function Write-Log([string]$text) {
  try {
    $line = (Now-Iso) + ' ' + $(if ($Worker) { 'worker   ' } else { 'watchdog ' }) + $text + [System.Environment]::NewLine
    [System.IO.File]::AppendAllText((Join-Path $script:LogDir 'helper.log'), $line, $Utf8)
  } catch {}
}

function New-Failure([string]$code, [string]$message) {
  $e = New-Object System.Exception($message)
  $e.Data['code'] = $code
  return $e
}

function Get-FailureCode($err) {
  try {
    $ex = $err.Exception
    while ($ex) {
      if ($ex.Data -and $ex.Data.Contains('code')) { return [string]$ex.Data['code'] }
      $ex = $ex.InnerException
    }
  } catch {}
  return ''
}

function Get-FailureMessage($err) {
  try {
    $ex = $err.Exception
    while ($ex.InnerException) { $ex = $ex.InnerException }
    return ([string]$ex.Message).Trim()
  } catch { return [string]$err }
}

# A path inside a ticket must stay inside the folders this helper owns.
function Resolve-Inside([string]$base, [string]$relative, [string]$mustBeUnder) {
  $full = [System.IO.Path]::GetFullPath((Join-Path $base $relative))
  $under = [System.IO.Path]::GetFullPath($mustBeUnder)
  $sep = [System.IO.Path]::DirectorySeparatorChar
  if (-not $under.EndsWith([string]$sep)) { $under = $under + $sep }
  $cmp = [System.StringComparison]::OrdinalIgnoreCase
  if (-not $full.StartsWith($under, $cmp)) { throw (New-Failure 'bad-ticket' ('The ticket points outside the output folder: ' + $relative)) }
  return $full
}

# ---------------------------------------------------------------- PowerPoint
#
# Everything that touches PowerPoint is in the Engine-* functions below.
# The automated checks replace them with stand-ins so that the queue, the
# time limits and the recovery can be exercised on a machine with no
# PowerPoint. The stand-in reports its own name, and the page refuses any
# result whose engine is not Microsoft PowerPoint.

function Engine-Name { return 'Microsoft PowerPoint' }

function Engine-Installed {
  if (-not $IsWin) { return $false }
  try { return (Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application') } catch { return $false }
}

function Engine-RunningBefore {
  try { return [bool](Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) } catch { return $false }
}

function Engine-InstalledFonts {
  $names = New-Object 'System.Collections.Generic.HashSet[string]'
  try {
    Add-Type -AssemblyName System.Drawing
    $c = New-Object System.Drawing.Text.InstalledFontCollection
    foreach ($f in $c.Families) { [void]$names.Add($f.Name) }
  } catch {}
  try {
    Add-Type -AssemblyName PresentationCore
    foreach ($fam in [System.Windows.Media.Fonts]::SystemFontFamilies) {
      foreach ($n in $fam.FamilyNames.Values) { [void]$names.Add([string]$n) }
      foreach ($face in $fam.FamilyTypefaces) {
        foreach ($n in $fam.FamilyNames.Values) {
          foreach ($fn in $face.AdjustedFaceNames.Values) { [void]$names.Add(([string]$n + ' ' + [string]$fn)) }
        }
      }
    }
  } catch {}
  foreach ($key in @('HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts', 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')) {
    try {
      $props = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
      foreach ($p in $props.PSObject.Properties) {
        if ($p.Name -like 'PS*') { continue }
        $n = ($p.Name -replace '\s*\((TrueType|OpenType|All res|VGA res)\)\s*$', '')
        foreach ($part in ($n -split '\s+&\s+')) { if ($part) { [void]$names.Add($part.Trim()) } }
      }
    } catch {}
  }
  return @($names | Sort-Object)
}

function Engine-Start([string]$root) {
  $app = New-Object -ComObject PowerPoint.Application
  $inDir = (Join-Path $root 'in')
  # a copy left open by a worker that was stopped mid file belongs to us. Close it. Nothing else is touched.
  try {
    for ($i = $app.Presentations.Count; $i -ge 1; $i--) {
      $p = $app.Presentations.Item($i)
      if ([string]$p.FullName -like ($inDir + '\*')) { try { $p.Saved = -1; $p.Close() } catch {} }
    }
  } catch {}
  $info = @{ version = [string]$app.Version; build = ''; userPresentations = 0 }
  try { $info.build = [string]$app.Build } catch {}
  try { $info.userPresentations = [int]$app.Presentations.Count } catch {}
  return @{ app = $app; info = $info }
}

function Engine-Stop($app, [bool]$owned) {
  if (-not $app) { return }
  try {
    # only quit a PowerPoint this helper started, and only when nothing of yours is open in it
    if ($owned -and ($app.Presentations.Count -eq 0)) { $app.Quit() }
  } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
}

# Called by the watchdog after it has stopped a stalled worker.
# Returns '' when work can continue, or a sentence saying why it cannot.
function Engine-Recover([bool]$owned) {
  $procs = @(Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
  if ($procs.Count -eq 0) { return '' }
  $visible = @($procs | Where-Object { $_.MainWindowHandle -ne 0 })
  if ($owned -and $visible.Count -eq 0) {
    foreach ($p in $procs) { try { $p.Kill() } catch {} }
    Start-Sleep -Milliseconds 800
    return ''
  }
  $stuck = @($procs | Where-Object { -not $_.Responding })
  if ($stuck.Count -gt 0) {
    return 'PowerPoint is not responding and it has windows open, so the helper will not close it. Answer any question PowerPoint is showing, or save your work and close PowerPoint, then press Retry.'
  }
  return ''
}

function Open-Deck($app, [string]$path, [bool]$mayPrompt) {
  # A wrong password on the end of the path makes a protected file fail at once
  # where it would otherwise stop and ask. An unprotected file ignores it.
  try {
    return $app.Presentations.Open($path + '::slidesize-no-password::', -1, 0, 0)
  } catch {
    $m = Get-FailureMessage $_
    if ($m -match 'password') { throw (New-Failure 'password' $m) }
  }
  try {
    return $app.Presentations.Open($path, -1, 0, 0)
  } catch {
    $m = Get-FailureMessage $_
    if ($m -match 'password') { throw (New-Failure 'password' $m) }
    throw (New-Failure 'open-failed' $m)
  }
}

function Select-Sample($slides) {
  $list = @($slides)
  if ($list.Count -le 5) { return $list }
  $out = @()
  for ($i = 0; $i -lt 5; $i++) {
    $v = $list[[int][System.Math]::Floor($i * ($list.Count - 1) / 4.0 + 0.5)]
    if ($out -notcontains $v) { $out += $v }
  }
  return $out
}

function Engine-Convert($app, $t, [string]$inPath, [string]$pdfPath, [string]$tmpPdf, [string]$refDir) {
  $o = $t.options
  $r = @{ facts = @{}; export = @{ method = ''; notApplied = @(); placeholders = @(); pageMap = @(); removedSlides = @() }; timing = @{}; reference = $null }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $pres = $null
  $oldAlerts = $null
  $oldSecurity = $null
  try {
    if ($IsWin) { try { Unblock-File -LiteralPath $inPath -ErrorAction SilentlyContinue } catch {} }
    try { $oldAlerts = $app.DisplayAlerts; $app.DisplayAlerts = 1 } catch {}
    try { $oldSecurity = $app.AutomationSecurity; $app.AutomationSecurity = 3 } catch {}

    Set-Phase 'opening'
    $pres = Open-Deck $app $inPath ([bool]$t.mayPrompt)
    $r.timing.openMs = [int]$sw.ElapsedMilliseconds

    $n = [int]$pres.Slides.Count
    $hidden = @()
    for ($i = 1; $i -le $n; $i++) {
      if ($pres.Slides.Item($i).SlideShowTransition.Hidden -ne 0) { $hidden += $i }
    }
    $fonts = @()
    try {
      for ($i = 1; $i -le $pres.Fonts.Count; $i++) {
        $f = $pres.Fonts.Item($i)
        $fonts += @{ name = [string]$f.Name; embedded = ($f.Embedded -ne 0); embeddable = ($f.Embeddable -ne 0) }
      }
    } catch {}
    $w = [double]$pres.PageSetup.SlideWidth
    $h = [double]$pres.PageSetup.SlideHeight
    $r.facts = @{ slides = $n; hidden = @($hidden); slideWidthPt = $w; slideHeightPt = $h; fonts = @($fonts); readOnly = ($pres.ReadOnly -ne 0) }

    if ($t.collect -and $t.collect.media) {
      $media = @()
      for ($i = 1; $i -le $n; $i++) {
        try {
          foreach ($sh in $pres.Slides.Item($i).Shapes) {
            if ($sh.Type -eq 16) {
              $kind = 'video'
              try { if ($sh.MediaType -eq 2) { $kind = 'audio' } } catch {}
              $media += @{ slide = $i; type = $kind; name = [string]$sh.Name }
            }
          }
        } catch {}
      }
      $r.facts.media = @($media)
    }
    if ($n -eq 0) { throw (New-Failure 'no-slides' 'The presentation has no slides.') }

    $from = 1
    $to = $n
    $rangeType = 1
    if ($o.range) {
      $from = [System.Math]::Max(1, [int]$o.range[0])
      $to = [System.Math]::Min($n, [int]$o.range[1])
      $rangeType = 4
    }
    $map = @()
    for ($i = $from; $i -le $to; $i++) {
      if ($o.includeHidden -or ($hidden -notcontains $i)) { $map += $i }
    }
    if ($map.Count -eq 0) { throw (New-Failure 'no-slides' 'No slides are left to export with this slide range and the hidden slides left out.') }
    $r.export.pageMap = @($map)

    foreach ($p in @($o.placeholders)) {
      if (-not $p) { continue }
      try {
        $sl = $pres.Slides.Item([int]$p.slide)
        $sh = $sl.Shapes.AddShape(1, [single]$p.left, [single]$p.top, [single]$p.width, [single]$p.height)
        $sh.Fill.Visible = -1
        $sh.Fill.Solid()
        $sh.Fill.ForeColor.RGB = 0x2A2A2A
        $sh.Line.Visible = -1
        $sh.Line.ForeColor.RGB = 0x2A30E7
        $sh.Line.Weight = 3
        $sh.Line.DashStyle = 4
        $tr = $sh.TextFrame.TextRange
        $tr.Text = [string]$p.label
        $tr.Font.Name = 'Arial'
        $tr.Font.Size = 14
        $tr.Font.Bold = -1
        $tr.Font.Color.RGB = 0xFFFFFF
        $tr.ParagraphFormat.Alignment = 2
        $r.export.placeholders += @{ slide = [int]$p.slide; ok = $true }
      } catch {
        $r.export.placeholders += @{ slide = [int]$p.slide; ok = $false; error = (Get-FailureMessage $_) }
      }
    }

    Set-Phase 'exporting'
    $t0 = $sw.ElapsedMilliseconds
    $outputType = 1
    if ($o.output -eq 'notes') { $outputType = 5 }
    $intent = 2
    if ($o.intent -eq 'screen') { $intent = 1 }
    $hid = 0
    if ($o.includeHidden) { $hid = -1 }
    # PowerPoint rejects the call unless it is handed a real print range, even for all slides
    $ranges = $pres.PrintOptions.Ranges
    $ranges.ClearAll()
    $pr = $ranges.Add($from, $to)
    if (Test-Path -LiteralPath $tmpPdf) { [System.IO.File]::Delete($tmpPdf) }
    $err2 = ''
    $err1 = ''
    try {
      $pres.ExportAsFixedFormat2($tmpPdf, 2, $intent, 0, 1, $outputType, $hid, $pr, $rangeType, '', [bool]$o.docProps, $true, [bool]$o.tags, [bool]$o.bitmapText, [bool]$o.pdfa, [bool]$o.markup)
      $r.export.method = 'ExportAsFixedFormat2'
    } catch {
      $err2 = Get-FailureMessage $_
      try {
        $pres.ExportAsFixedFormat($tmpPdf, 2, $intent, 0, 1, $outputType, $hid, $pr, $rangeType, '', [bool]$o.docProps, $true, [bool]$o.tags, [bool]$o.bitmapText, [bool]$o.pdfa)
        $r.export.method = 'ExportAsFixedFormat'
        if ($o.markup) { $r.export.notApplied += @{ option = 'markup'; reason = 'This PowerPoint has no ExportAsFixedFormat2, which is the call that can include comments and ink.' } }
      } catch {
        $err1 = Get-FailureMessage $_
        # Last resort. Plain Save As has no options, so only use it when none were asked for.
        $plain = (-not $o.pdfa) -and (-not $o.range) -and ($o.output -ne 'notes') -and (-not $o.includeHidden) -and ($o.intent -ne 'screen')
        if (-not $plain) { throw (New-Failure 'export-failed' ($err2 + ' / ' + $err1)) }
        try { $pres.SaveAs($tmpPdf, 32, 0) } catch { throw (New-Failure 'export-failed' ($err2 + ' / ' + $err1 + ' / ' + (Get-FailureMessage $_))) }
        $r.export.method = 'Save As PDF'
        $r.export.fallbackReason = $err2
        $r.export.pageMap = $null
        foreach ($k in @('bitmapText', 'tags', 'docProps', 'markup')) {
          $r.export.notApplied += @{ option = $k; reason = 'Save As PDF uses PowerPoint defaults.' }
        }
      }
    }
    if (-not (Test-Path -LiteralPath $tmpPdf)) { throw (New-Failure 'export-failed' 'PowerPoint reported success but wrote no file.') }
    if ((Get-Item -LiteralPath $tmpPdf).Length -le 0) { throw (New-Failure 'export-failed' 'PowerPoint wrote an empty file.') }
    $r.timing.exportMs = [int]($sw.ElapsedMilliseconds - $t0)

    if (Test-Path -LiteralPath $pdfPath) {
      if (-not $t.overwrite) { throw (New-Failure 'exists' ('The output file already exists: ' + [System.IO.Path]::GetFileName($pdfPath))) }
      [System.IO.File]::Delete($pdfPath)
    }
    [System.IO.File]::Move($tmpPdf, $pdfPath)
    $r.pdf = @{ bytes = [long](Get-Item -LiteralPath $pdfPath).Length }

    if ($o.reference -and $o.output -ne 'notes' -and $r.export.pageMap) {
      Set-Phase 'reference'
      $t1 = $sw.ElapsedMilliseconds
      Ensure-Dir $refDir
      $edge = [int]$o.reference.longEdge
      if ($edge -lt 320) { $edge = 320 }
      if ($w -ge $h) { $pw = $edge; $ph = [int][System.Math]::Round($edge * $h / $w) } else { $ph = $edge; $pw = [int][System.Math]::Round($edge * $w / $h) }
      $want = @($r.export.pageMap)
      if ($o.reference.mode -eq 'sample') { $want = @(Select-Sample $want) }
      $made = @()
      foreach ($sn in $want) {
        try {
          $file = Join-Path $refDir ('slide-' + ('{0:D4}' -f [int]$sn) + '.png')
          $pres.Slides.Item([int]$sn).Export($file, 'PNG', $pw, $ph)
          if (Test-Path -LiteralPath $file) { $made += [int]$sn }
        } catch {}
      }
      $r.reference = @{ dir = ('ref/' + $t.id); slides = @($made); width = $pw; height = $ph; pattern = 'slide-%04d.png' }
      $r.timing.referenceMs = [int]($sw.ElapsedMilliseconds - $t1)
    }
  } finally {
    Set-Phase 'closing'
    if ($pres) {
      try { $pres.Saved = -1 } catch {}
      try { $pres.Close() } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {}
    }
    if ($oldAlerts -ne $null) { try { $app.DisplayAlerts = $oldAlerts } catch {} }
    if ($oldSecurity -ne $null) { try { $app.AutomationSecurity = $oldSecurity } catch {} }
    if (Test-Path -LiteralPath $tmpPdf) { try { [System.IO.File]::Delete($tmpPdf) } catch {} }
  }
  $r.timing.totalMs = [int]$sw.ElapsedMilliseconds
  return $r
}

# ---------------------------------------------------------------- veraPDF

function Find-VeraPdf {
  $names = @('verapdf.bat', 'verapdf')
  foreach ($n in $names) {
    try { $c = Get-Command $n -ErrorAction SilentlyContinue; if ($c) { return [string]$c.Source } } catch {}
  }
  $homeDir = [System.Environment]::GetFolderPath('UserProfile')
  foreach ($p in @((Join-Path $homeDir 'verapdf\verapdf.bat'), (Join-Path $homeDir 'verapdf/verapdf'), 'C:\Program Files\veraPDF\verapdf.bat')) {
    try { if ($p -and (Test-Path -LiteralPath $p)) { return $p } } catch {}
  }
  return ''
}

function Invoke-Validate([string]$pdfPath) {
  $vera = Find-VeraPdf
  if (-not $vera) { return $null }
  try {
    $version = ''
    try { $version = [string]((& $vera --version 2>$null) | Select-Object -First 1) } catch {}
    $outText = [string]((& $vera --format mrr $pdfPath 2>$null) -join "`n")
    $m = [regex]::Match($outText, 'isCompliant="(true|false)"')
    if (-not $m.Success) { return @{ validator = ('veraPDF ' + $version).Trim(); result = 'error'; detail = 'veraPDF ran but gave no verdict.' } }
    $profileName = ''
    $pm = [regex]::Match($outText, 'profileName="([^"]*)"')
    if ($pm.Success) { $profileName = $pm.Groups[1].Value }
    $failed = ''
    $fm = [regex]::Match($outText, 'failedRules="(\d+)"')
    if ($fm.Success) { $failed = $fm.Groups[1].Value + ' failed rules.' }
    $res = 'failed'
    if ($m.Groups[1].Value -eq 'true') { $res = 'passed' }
    return @{ validator = ('veraPDF ' + ($version -replace '^veraPDF\s*', '')).Trim(); result = $res; profile = $profileName; detail = $failed }
  } catch {
    return @{ validator = 'veraPDF'; result = 'error'; detail = (Get-FailureMessage $_) }
  }
}

# ---------------------------------------------------------------- worker

function Set-Phase([string]$phase) {
  $script:Phase = $phase
  Write-WorkerState 'busy'
}

function Write-WorkerState([string]$state, [string]$message = '') {
  Write-JsonAtomic (Join-Path $Root 'worker.json') @{
    pid = $PID; state = $state; phase = $script:Phase; id = $script:CurrentId; heartbeat = (Now-Iso)
    message = $message; powerpoint = $script:PptInfo; verapdf = $script:Vera
  }
}

function Test-ParentAlive {
  if ($ParentPid -le 0) { return $true }
  try { return [bool](Get-Process -Id $ParentPid -ErrorAction SilentlyContinue) } catch { return $false }
}

function Invoke-Ticket($app, [string]$ticketPath) {
  $t = Read-Json $ticketPath
  $id = [System.IO.Path]::GetFileNameWithoutExtension($ticketPath)
  $script:CurrentId = $id
  $started = Now-Iso
  $res = @{ protocol = $Protocol; id = $id; ok = $false; status = 'failed'; startedAt = $started
            engine = @{ name = (Engine-Name); version = $script:PptInfo.version; build = $script:PptInfo.build; platform = $(if ($IsWin) { 'windows' } else { 'other' }) } }
  $inPath = ''
  try {
    if (-not $t -or $t.protocol -ne $Protocol) { throw (New-Failure 'bad-ticket' 'The ticket could not be read or comes from a different version of the page.') }
    $outRoot = [System.IO.Path]::GetDirectoryName($Root)
    $task = 'convert'
    if ($t.task) { $task = [string]$t.task }
    $pdfPath = Resolve-Inside $Root ([string]$t.pdf) $outRoot
    if (-not $pdfPath.ToLower().EndsWith('.pdf')) { throw (New-Failure 'bad-ticket' 'The output name must end in .pdf.') }

    if ($task -eq 'validate') {
      Set-Phase 'validating'
      $res.pdfa = Invoke-Validate $pdfPath
      $res.ok = $true
      $res.status = 'done'
    } else {
      $inPath = Resolve-Inside $Root ([string]$t.input) (Join-Path $Root 'in')
      if (-not (Test-Path -LiteralPath $inPath)) { throw (New-Failure 'not-found' 'The temporary copy of the presentation is missing.') }
      $tmpPdf = Join-Path (Join-Path $Root 'out') ($id + '.export.pdf')
      $refDir = Join-Path (Join-Path $Root 'ref') $id
      if (Test-Path -LiteralPath $refDir) { try { Remove-Item -LiteralPath $refDir -Recurse -Force } catch {} }
      Write-Log ('convert ' + $id + ' ' + [string]$t.sourceName)
      $r = Engine-Convert $app $t $inPath $pdfPath $tmpPdf $refDir
      foreach ($k in $r.Keys) { $res[$k] = $r[$k] }
      if ($t.options.validatePdfa -and $t.options.pdfa -and $t.validateNow) {
        Set-Phase 'validating'
        $res.pdfa = Invoke-Validate $pdfPath
      }
      $res.ok = $true
      $res.status = 'done'
    }
  } catch {
    $code = Get-FailureCode $_
    if (-not $code) { $code = 'convert-failed' }
    $res.error = @{ code = $code; message = (Get-FailureMessage $_) }
    Write-Log ('failed ' + $id + ' ' + $code + ' ' + $res.error.message)
  } finally {
    if ($inPath -and (Test-Path -LiteralPath $inPath)) { try { [System.IO.File]::Delete($inPath) } catch {} }
  }
  $res.finishedAt = Now-Iso
  Write-JsonAtomic (Join-Path (Join-Path $Root 'done') ($id + '.json')) $res
  try { [System.IO.File]::Delete($ticketPath) } catch {}
  $script:CurrentId = ''
  $script:Phase = ''
  Write-WorkerState 'ready'
}

function Start-WorkerLoop {
  $script:Phase = ''
  $script:CurrentId = ''
  $script:PptInfo = @{ version = ''; build = ''; userPresentations = 0 }
  $script:Vera = $false
  Write-WorkerState 'starting'
  $app = $null
  try {
    $started = Engine-Start $Root
    $app = $started.app
    $script:PptInfo = $started.info
  } catch {
    Write-Log ('could not start PowerPoint: ' + (Get-FailureMessage $_))
    Write-WorkerState 'error' ('PowerPoint could not be started. ' + (Get-FailureMessage $_))
    exit 2
  }
  try { $script:Vera = [bool](Find-VeraPdf) } catch {}
  try {
    if (-not (Test-Path -LiteralPath (Join-Path $Root 'fonts.json'))) {
      Write-JsonAtomic (Join-Path $Root 'fonts.json') @{ platform = $(if ($IsWin) { 'windows' } else { 'other' }); families = @(Engine-InstalledFonts) }
    }
  } catch {}
  Write-WorkerState 'ready'
  $active = Join-Path $Root 'active'
  $last = [System.DateTime]::UtcNow
  try {
    while ($true) {
      if (-not (Test-ParentAlive)) { break }
      if (Test-Path -LiteralPath (Join-Path $active '_stop')) { break }
      $next = @(Get-ChildItem -LiteralPath $active -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -First 1)
      if ($next.Count -gt 0) {
        Invoke-Ticket $app $next[0].FullName
        $last = [System.DateTime]::UtcNow
      } else {
        if (([System.DateTime]::UtcNow - $last).TotalSeconds -ge 2) { Write-WorkerState 'ready'; $last = [System.DateTime]::UtcNow }
        Start-Sleep -Milliseconds 200
      }
    }
  } finally {
    Engine-Stop $app ([bool]$Owns)
    Write-WorkerState 'stopped'
  }
}

# ---------------------------------------------------------------- watchdog

function Find-Root {
  $here = $PSScriptRoot
  if (-not $here) { $here = (Get-Location).Path }
  if ($Root) { return [System.IO.Path]::GetFullPath($Root) }
  if (Test-Path -LiteralPath (Join-Path $here 'marker.json')) { return $here }
  if (Test-Path -LiteralPath (Join-Path (Join-Path $here '_slidesize') 'marker.json')) { return (Join-Path $here '_slidesize') }
  Write-Host ''
  Write-Host 'This helper was started from outside the output folder.'
  Write-Host 'Choose the output folder you picked on the SlideSize page.'
  try {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Choose the output folder you picked on the SlideSize page'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $picked = Join-Path $dlg.SelectedPath '_slidesize'
      if (Test-Path -LiteralPath (Join-Path $picked 'marker.json')) { return $picked }
      Write-Host 'That folder has not been set up by the SlideSize page. Choose it on the page first.'
    }
  } catch {}
  return ''
}

function Start-Worker([bool]$owns) {
  $exe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $exe
  $psi.Arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -Worker -Root "' + $Root + '" -ParentPid ' + $PID + ' -Owns ' + $(if ($owns) { 1 } else { 0 })
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  return [System.Diagnostics.Process]::Start($psi)
}

function Stop-Worker($proc) {
  if (-not $proc) { return }
  try { if (-not $proc.HasExited) { $proc.Kill(); [void]$proc.WaitForExit(5000) } } catch {}
}

function Write-TimeoutResult([string]$id, [string]$status, [int]$timeoutSec, [string]$message) {
  $res = @{ protocol = $Protocol; id = $id; ok = $false; status = $status; timeoutSec = $timeoutSec
            error = @{ code = $status; message = $message }; finishedAt = (Now-Iso)
            engine = @{ name = (Engine-Name); version = ''; build = ''; platform = $(if ($IsWin) { 'windows' } else { 'other' }) } }
  Write-JsonAtomic (Join-Path (Join-Path $Root 'done') ($id + '.json')) $res
}

function Start-Watchdog {
  $Root = Find-Root
  if (-not $Root) {
    Write-Host ''
    Write-Host 'No SlideSize output folder was found. Go back to the page, choose an output folder, then start this helper again.'
    exit 1
  }
  $script:Root = $Root
  foreach ($d in @('queue', 'active', 'done', 'in', 'out', 'ref', 'log')) { Ensure-Dir (Join-Path $Root $d) }
  $script:LogDir = Join-Path $Root 'log'

  $lock = $null
  try {
    $lock = [System.IO.File]::Open((Join-Path $Root 'helper.lock'), [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  } catch {
    Write-Host ''
    Write-Host 'A helper is already running for this folder. Use that one, or close its window first.'
    exit 1
  }

  # anything half done by a previous run goes back in the queue
  foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $Root 'active') -ErrorAction SilentlyContinue)) {
    if ($f.Name -eq '_stop') { Remove-Item -LiteralPath $f.FullName -Force; continue }
    try { Move-Item -LiteralPath $f.FullName -Destination (Join-Path (Join-Path $Root 'queue') $f.Name) -Force } catch {}
  }
  try { Remove-Item -LiteralPath (Join-Path $Root 'worker.json') -Force -ErrorAction SilentlyContinue } catch {}
  try { Remove-Item -LiteralPath (Join-Path $Root 'fonts.json') -Force -ErrorAction SilentlyContinue } catch {}

  $installed = Engine-Installed
  $runningBefore = Engine-RunningBefore
  $owns = (-not $runningBefore)
  $control = Read-Json (Join-Path $Root 'control.json')
  $lastSeq = 0
  if ($control -and $control.seq) { $lastSeq = [long]$control.seq }

  Write-Host ''
  Write-Host 'SlideSize PowerPoint to PDF helper' $HelperVersion
  Write-Host 'Folder  ' ([System.IO.Path]::GetDirectoryName($Root))
  if (-not $installed) {
    Write-Host ''
    Write-Host 'Microsoft PowerPoint was not found on this computer.'
    Write-Host 'Microsoft PowerPoint must be installed for PowerPoint-based conversion.'
  } else {
    if ($runningBefore) { Write-Host 'PowerPoint is already open. Your presentations will be left alone.' }
    Write-Host 'Waiting for the SlideSize page. Leave this window open. Close it to stop.'
  }
  Write-Log ('start ' + $HelperVersion + ' installed=' + $installed + ' runningBefore=' + $runningBefore)

  $proc = $null
  $state = 'starting'
  $message = ''
  $activeId = ''
  $activeSince = $null
  $activeTimeout = 600
  $workerStarted = $null
  $lastBeat = [System.DateTime]::MinValue
  $converted = 0
  $failed = 0
  $stop = $false
  $caps = @{ pdfa = $true; notes = $true; quality = $true; bitmapText = $true; tags = $true; docProps = $true; markup = $true
             hidden = $true; range = $true; reference = $true; placeholders = $true }

  try {
    while (-not $stop) {
      $now = [System.DateTime]::UtcNow

      # the page can ask to stop, or to give up on the file in hand
      $control = Read-Json (Join-Path $Root 'control.json')
      $abort = ''
      if ($control -and $control.seq -and ([long]$control.seq -gt $lastSeq)) {
        $lastSeq = [long]$control.seq
        if ($control.stop) { $stop = $true }
        if ($control.abort) { $abort = [string]$control.abort }
        if ($control.retry -and $state -eq 'blocked') { $state = 'starting'; $message = '' }
      }

      if ($installed -and $state -ne 'blocked' -and -not $stop) {
        if (-not $proc -or $proc.HasExited) {
          if ($proc -and $activeId) {
            # the worker died with a file in hand
            Write-TimeoutResult $activeId 'failed' $activeTimeout 'The helper process that talks to PowerPoint stopped unexpectedly while converting this file.'
            try { Remove-Item -LiteralPath (Join-Path (Join-Path $Root 'active') ($activeId + '.json')) -Force -ErrorAction SilentlyContinue } catch {}
            $failed++
            $activeId = ''
            $activeSince = $null
          }
          $w = Read-Json (Join-Path $Root 'worker.json')
          if ($proc -and $w -and $w.state -eq 'error') {
            $state = 'blocked'
            $message = [string]$w.message
            $proc = $null
          } else {
            $proc = Start-Worker $owns
            $workerStarted = $now
            $state = 'starting'
          }
        }
      }

      $w = Read-Json (Join-Path $Root 'worker.json')
      if ($proc -and -not $proc.HasExited -and $w -and $w.pid -eq $proc.Id) {
        if ($w.state -eq 'ready' -or $w.state -eq 'busy') { if ($state -eq 'starting') { $state = 'idle' } }
        if ($w.state -eq 'starting' -and $workerStarted -and (($now - $workerStarted).TotalSeconds -gt 90)) {
          Stop-Worker $proc
          $proc = $null
          $why = Engine-Recover $owns
          if (-not $why) { $why = 'PowerPoint did not start within 90 seconds. Open PowerPoint by hand to see whether it is asking a question, then press Retry.' }
          $state = 'blocked'
          $message = $why
        }
      }

      # a ticket that has finished
      if ($activeId -and (Test-Path -LiteralPath (Join-Path (Join-Path $Root 'done') ($activeId + '.json')))) {
        $d = Read-Json (Join-Path (Join-Path $Root 'done') ($activeId + '.json'))
        if ($d) {
          if ($d.ok) { $converted++; Write-Host ('  done    ' + $activeId + '  ' + [int](($now - $activeSince).TotalSeconds) + ' s') }
          else { $failed++; Write-Host ('  FAILED  ' + $activeId + '  ' + [string]$d.error.message) }
          $activeId = ''
          $activeSince = $null
          if ($state -ne 'blocked') { $state = 'idle' }
        }
      }

      # a ticket that has run out of time, or that the page gave up on
      if ($activeId -and $activeSince) {
        $late = (($now - $activeSince).TotalSeconds -gt $activeTimeout)
        if ($late -or ($abort -eq $activeId)) {
          Stop-Worker $proc
          $proc = $null
          $status = 'timeout'
          $text = 'PowerPoint did not finish within ' + $activeTimeout + ' seconds.'
          if (-not $late) { $status = 'cancelled'; $text = 'Skipped from the page while it was converting.' }
          Write-TimeoutResult $activeId $status $activeTimeout $text
          try { Remove-Item -LiteralPath (Join-Path (Join-Path $Root 'active') ($activeId + '.json')) -Force -ErrorAction SilentlyContinue } catch {}
          Write-Host ('  ' + $status.ToUpper() + ' ' + $activeId)
          Write-Log ($status + ' ' + $activeId)
          $failed++
          $activeId = ''
          $activeSince = $null
          $why = Engine-Recover $owns
          if ($why) { $state = 'blocked'; $message = $why } else { $state = 'starting' }
        }
      }
      if ($abort -and -not $activeId) {
        # the page gave up on a ticket that had not started
        try { Remove-Item -LiteralPath (Join-Path (Join-Path $Root 'queue') ($abort + '.json')) -Force -ErrorAction SilentlyContinue } catch {}
      }

      # hand out the next ticket
      if (-not $activeId -and $state -eq 'idle' -and $proc -and -not $proc.HasExited) {
        $next = @(Get-ChildItem -LiteralPath (Join-Path $Root 'queue') -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -First 1)
        if ($next.Count -gt 0) {
          $t = Read-Json $next[0].FullName
          if ($t) {
            $id = [System.IO.Path]::GetFileNameWithoutExtension($next[0].Name)
            $activeTimeout = 600
            if ($t.timeoutSec) { $activeTimeout = [int]$t.timeoutSec }
            try {
              Move-Item -LiteralPath $next[0].FullName -Destination (Join-Path (Join-Path $Root 'active') $next[0].Name) -Force
              $activeId = $id
              $activeSince = $now
              $state = 'working'
              $label = $(if ($t.task -eq 'validate') { 'validate' } else { 'convert ' })
              Write-Host ('  ' + $label + ' ' + $id + '  ' + [string]$t.sourceName)
            } catch {}
          }
        }
      }
      if ($activeId -and $state -eq 'idle') { $state = 'working' }

      if (($now - $lastBeat).TotalMilliseconds -ge 1500) {
        $lastBeat = $now
        $ppt = @{ installed = [bool]$installed; version = ''; build = ''; runningBefore = [bool]$runningBefore; userPresentations = 0 }
        $vera = $false
        if ($w) {
          if ($w.powerpoint) { $ppt.version = [string]$w.powerpoint.version; $ppt.build = [string]$w.powerpoint.build; $ppt.userPresentations = [int]$w.powerpoint.userPresentations }
          if ($w.verapdf) { $vera = $true }
        }
        $current = $null
        if ($activeId) {
          $phase = ''
          if ($w -and $w.id -eq $activeId) { $phase = [string]$w.phase }
          $current = @{ id = $activeId; phase = $phase; since = (Format-Iso $activeSince) }
        }
        $shown = $state
        if (-not $installed) { $shown = 'idle' }
        Write-JsonAtomic (Join-Path $Root 'helper.json') @{
          protocol = $Protocol; helper = $HelperVersion; platform = $(if ($IsWin) { 'windows' } else { 'other' })
          os = [string][System.Environment]::OSVersion.VersionString; pid = $PID; heartbeat = (Now-Iso)
          state = $shown; message = $message; powerpoint = $ppt; caps = $caps; verapdf = $vera
          engine = (Engine-Name); current = $current; converted = $converted; failed = $failed
        }
      }
      Start-Sleep -Milliseconds 250
    }
  } finally {
    try { [System.IO.File]::WriteAllText((Join-Path (Join-Path $Root 'active') '_stop'), '', $Utf8) } catch {}
    if ($proc) { try { [void]$proc.WaitForExit(6000) } catch {}; Stop-Worker $proc }
    try { Remove-Item -LiteralPath (Join-Path (Join-Path $Root 'active') '_stop') -Force -ErrorAction SilentlyContinue } catch {}
    try {
      $h = Read-Json (Join-Path $Root 'helper.json')
      if ($h) { $h.state = 'stopped'; $h.heartbeat = (Now-Iso); Write-JsonAtomic (Join-Path $Root 'helper.json') $h }
    } catch {}
    if ($lock) { try { $lock.Close() } catch {} }
    try { Remove-Item -LiteralPath (Join-Path $Root 'helper.lock') -Force -ErrorAction SilentlyContinue } catch {}
    Write-Log 'stop'
    Write-Host ''
    Write-Host 'Helper stopped.'
  }
}

# ---------------------------------------------------------------- start

# Test stand-ins for PowerPoint. Never set outside the automated checks.
if ($env:SLIDESIZE_HELPER_TEST_ENGINE -and (Test-Path -LiteralPath $env:SLIDESIZE_HELPER_TEST_ENGINE)) { . $env:SLIDESIZE_HELPER_TEST_ENGINE }

if ($Worker) {
  $Root = [System.IO.Path]::GetFullPath($Root)
  $script:LogDir = Join-Path $Root 'log'
  Start-WorkerLoop
} else {
  Start-Watchdog
}
