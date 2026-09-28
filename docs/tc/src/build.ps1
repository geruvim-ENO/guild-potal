$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
function LoadJson($p) { [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }
$cfg = LoadJson (Join-Path $dir 'cfg.json')
$basic = LoadJson (Join-Path $dir 'tc_basic.json')
$others = LoadJson (Join-Path $dir 'tc_others.json')

$data = [ordered]@{}
$data[$cfg.basicSheet] = @($basic)
foreach ($n in $cfg.order) { if ($n -ne $cfg.basicSheet) { $data[$n] = @($others.$n) } }

function Expand([string]$s) {
  if ($null -eq $s) { return '' }
  foreach ($t in $cfg.tokens) { $s = $s.Replace($t[0], $t[1]) }
  return $s
}
function Numbered($arr) {
  $out = @(); $i = 1
  foreach ($a in $arr) { $out += ("{0}. {1}" -f $i, (Expand $a)); $i++ }
  return $out
}
# estimate wrapped line count for text in a column of given width (Excel width units, 9pt font)
function Lines([string]$text, [double]$width) {
  if ([string]::IsNullOrEmpty($text)) { return 1 }
  $cap = [Math]::Max(4.0, $width - 1.2)
  $total = 0
  foreach ($para in $text.Split("`n")) {
    $w = 0.0
    foreach ($ch in $para.ToCharArray()) { if ([int]$ch -lt 128) { $w += 0.86 } else { $w += 1.62 } }
    $total += [Math]::Max(1, [Math]::Ceiling($w / $cap))
  }
  return $total
}

Copy-Item -LiteralPath $cfg.src -Destination $cfg.dst -Force
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false; $x.ScreenUpdating = $false
$wb = $x.Workbooks.Open($cfg.dst)
$counts = @{}
try {
  $widths = @{ 7 = 27; 8 = 58; 9 = 46; 12 = 30; 15 = 24 }
  foreach ($name in $data.Keys) {
    $ws = $wb.Worksheets.Item($name)
    $tcs = $data[$name]
    # title merge B2:O3
    $ws.Range('B2:N3').UnMerge()
    $ws.Range('B2:O3').Merge()
    $ws.Range('B2:O3').Interior.Color = $ws.Range('B2').Interior.Color
    # header O12
    $ws.Range('N12').Copy($ws.Range('O12')) | Out-Null
    $ws.Range('O12').Value2 = $cfg.hdrO
    foreach ($k in $widths.Keys) { $ws.Columns.Item($k).ColumnWidth = $widths[$k] }
    # clear old rows
    $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    if ($last -ge 13) { $ws.Range("13:$last").Delete() | Out-Null }

    $totalRows = 0; foreach ($tc in $tcs) { $totalRows += @($tc.ex).Count }
    $arr = New-Object 'object[,]' $totalRows, 14
    $groups = @(); $r = 0
    foreach ($tc in $tcs) {
      $ex = @(Numbered @($tc.ex)); $n = $ex.Count
      $steps = @(Numbered @($tc.st)) -join "`n"
      $pre = Expand $tc.pre
      for ($i = 0; $i -lt $n; $i++) {
        $row = $r + $i
        if ($i -eq 0) {
          $arr[$row,0] = $tc.id; $arr[$row,1] = $tc.c; $arr[$row,2] = $tc.d; $arr[$row,3] = $tc.e; $arr[$row,4] = $tc.p
          $arr[$row,5] = $pre; $arr[$row,6] = $steps; $arr[$row,10] = $tc.ref; $arr[$row,13] = $tc.tq
        }
        $arr[$row,7] = $ex[$i]
      }
      $groups += ,@($r, $n, $pre, $steps, $tc.ref, $tc.tq, $tc.e, $ex)
      $r += $n
    }
    $counts[$name] = @($tcs.Count, $totalRows)
    $endRow = 12 + $totalRows
    $rng = $ws.Range("B13:O$endRow")
    for ($ri = 0; $ri -lt $totalRows; $ri++) {
      for ($ci = 0; $ci -lt 14; $ci++) {
        $val = $arr[$ri,$ci]
        if ($null -ne $val -and "$val" -ne '') {
          $cell = $ws.Cells.Item(13 + $ri, 2 + $ci)
          try { $cell.Value2 = [string]$val }
          catch { try { $cell.Value2 = "'" + [string]$val } catch { Write-Host ("FAIL {0}!R{1}C{2}: {3}" -f $name, (13+$ri), (2+$ci), ([string]$val).Substring(0, [Math]::Min(60, ([string]$val).Length))) } }
        }
      }
    }

    # formatting
    $rng.Font.Name = $cfg.font; $rng.Font.Size = 9; $rng.Font.Bold = $false; $rng.Font.Color = 0
    $rng.Interior.Pattern = -4142
    $rng.VerticalAlignment = -4108
    foreach ($b in 7,8,9,10,11,12) { $rng.Borders.Item($b).LineStyle = 1; $rng.Borders.Item($b).Weight = 2 }
    foreach ($c in 'B','C','D','E','F','J','K','M','N') { $ws.Range("${c}13:${c}$endRow").HorizontalAlignment = -4108 }
    foreach ($c in 'G','H','I','L','O') { $ws.Range("${c}13:${c}$endRow").HorizontalAlignment = -4131 }
    foreach ($c in 'C','D','E','G','H','I','L','O') { $ws.Range("${c}13:${c}$endRow").WrapText = $true }
    $ws.Range("H13:I$endRow").VerticalAlignment = -4160
    $ws.Range("G13:G$endRow").VerticalAlignment = -4160
    $pc = $ws.Range("F13:F$endRow"); $pc.Font.Bold = $true; $pc.Font.Color = 192
    # validations
    $v = $ws.Range("F13:F$endRow").Validation; $v.Delete(); $v.Add(3, 1, 1, $cfg.valP) | Out-Null
    $v = $ws.Range("K13:K$endRow").Validation; $v.Delete(); $v.Add(3, 1, 1, 'Pass,Fail,N/A,Not test') | Out-Null

    # merges first, then autofit (merged cells are ignored by AutoFit -> heights follow column I)
    foreach ($g in $groups) {
      $s = 13 + $g[0]; $n = $g[1]; $e = $s + $n - 1
      if ($n -gt 1) {
        foreach ($c in 'B','C','D','E','F','G','H','L','O') { $ws.Range("${c}${s}:${c}${e}").Merge() }
      }
      $ws.Range("B${e}:O${e}").Borders.Item(9).Weight = -4138
    }
    $ws.Range("13:$endRow").Rows.AutoFit() | Out-Null
    $base = @{}; for ($rr = 13; $rr -le $endRow; $rr++) { $base[$rr] = [double]$ws.Rows.Item($rr).RowHeight }
    # measure merged-text height with helper columns Z..AF on the same sheet (same widths, unmerged)
    $helpW = @(58, 27, 30, 24, 17, 12, 12)
    for ($k = 0; $k -lt 7; $k++) { $ws.Columns.Item(26 + $k).ColumnWidth = $helpW[$k] }
    $hz = $ws.Range("Z13:AF$endRow"); $hz.Font.Name = $cfg.font; $hz.Font.Size = 9; $hz.WrapText = $true
    foreach ($g in $groups) {
      $s = 13 + $g[0]; $n = $g[1]
      $ws.Cells.Item($s, 26).Value2 = [string]$g[3]; $ws.Cells.Item($s, 27).Value2 = [string]$g[2]; $ws.Cells.Item($s, 28).Value2 = [string]$g[4]
      $ws.Cells.Item($s, 29).Value2 = [string]$g[5]; $ws.Cells.Item($s, 30).Value2 = [string]$g[6]
      $ws.Cells.Item($s, 31).Value2 = [string]$ws.Range("C$s").Value2; $ws.Cells.Item($s, 32).Value2 = [string]$ws.Range("D$s").Value2
      $ws.Rows.Item($s).AutoFit() | Out-Null
      $need = [double]$ws.Rows.Item($s).RowHeight + 4
      $ws.Range("Z${s}:AF${s}").ClearContents() | Out-Null
      $rowH = @(); for ($i = 0; $i -lt $n; $i++) { $rowH += $base[$s + $i] + 3 }
      $sum = ($rowH | Measure-Object -Sum).Sum
      if ($sum -lt $need) { $add = ($need - $sum) / $n; for ($i = 0; $i -lt $n; $i++) { $rowH[$i] += $add } }
      for ($i = 0; $i -lt $n; $i++) { $ws.Rows.Item($s + $i).RowHeight = [Math]::Min(409, [Math]::Max(18, $rowH[$i])) }
    }
    $hz.Clear() | Out-Null
    for ($k = 0; $k -lt 7; $k++) { $ws.Columns.Item(26 + $k).ColumnWidth = 8 }
  }

  # Read Me
  $rm = $wb.Worksheets.Item($cfg.readme)
  $h4 = $rm.Range('B4'); $hColor = $h4.Font.Color
  $lastR = $rm.UsedRange.Row + $rm.UsedRange.Rows.Count - 1
  $rm.Range("B4:B$([Math]::Max($lastR,200))").ClearContents() | Out-Null
  $rm.Range("B4:B$([Math]::Max($lastR,200))").Font.Bold = $false
  $row = 2
  foreach ($ln in $cfg.readmeLines) {
    $style = $ln[0]; $text = $ln[1]
    foreach ($k in $counts.Keys) { $text = $text.Replace("{CNT:$k}", ($cfg.cntFmt -f $counts[$k][0], $counts[$k][1])) }
    $cell = $rm.Range("B$row")
    $cell.Value2 = $text
    if ($style -ne 'title') {
      $cell.Font.Name = $cfg.font; $cell.Font.Size = 10; $cell.WrapText = $true
      if ($style -eq 'h') { $cell.Font.Bold = $true; $cell.Font.Color = $hColor } else { $cell.Font.Bold = $false; $cell.Font.Color = 0 }
    }
    $row++
  }
  $rm.Columns.Item(2).ColumnWidth = 120
  $rm.Range("B4:B$row").Rows.AutoFit() | Out-Null

  $wb.Worksheets.Item($cfg.readme).Activate()
  $wb.Save()
  $counts.GetEnumerator() | ForEach-Object { "{0}: TC {1}, rows {2}" -f $_.Key, $_.Value[0], $_.Value[1] }
}
finally {
  $x.ScreenUpdating = $true
  $wb.Close($true)
  $x.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
