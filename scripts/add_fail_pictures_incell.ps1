# add_fail_pictures_incell.ps1 : 'FAIL내역' 시트의 캡처를 각 결함 행의 E열부터 '셀에 배치'(Excel 365 그림 셀 값, Range.InsertPictureInCell)로 넣는다.
#   캡처가 여러 장이면 F열, G열 … 로 한 칸씩 옮겨 넣는다. 기존 떠 있는 그림(표 아래 방식)은 지운다. (2026-10-08 사용자 요청)
#   InsertPictureInCell 을 지원하지 않는 Excel 이면 같은 셀에 맞춰 떠 있는 그림으로 대신 넣는다.
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json, [double]$ColWidth = 48, [double]$RowHeight = 190)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $true   # InsertPictureInCell 은 Excel 이 보이는 상태에서만 저장까지 된다 (2026-10-08 실측: 숨김이면 오류 없이 빈 셀)
try {
    $wb = $xl.Workbooks.Open($Path)
    $ws = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq "FAIL내역") { $ws = $s } }
    if (-not $ws) { throw "FAIL내역 sheet not found" }
    $ws.Activate(); $xl.WindowState = -4137; $xl.ActiveWindow.Zoom = 60   # 시트가 화면에 보여야 InsertPictureInCell 이 실제로 들어간다 (2026-10-08 실측)
    for ($s = $ws.Shapes.Count; $s -ge 1; $s--) { $ws.Shapes.Item($s).Delete() }
    $rowOf = @{}; $r = 3
    while ($ws.Cells.Item($r, 1).Value2) { $rowOf[[string]$ws.Cells.Item($r, 1).Value2] = $r; $r++ }
    $maxN = 1
    foreach ($f in $d.fail_rows) { if ($f.image_paths.Count -gt $maxN) { $maxN = $f.image_paths.Count } }
    for ($c = 5; $c -lt 5 + $maxN; $c++) { $ws.Columns.Item($c).ColumnWidth = $ColWidth }
    $ws.Cells.Item(1, 5).Value2 = "캡처 등 상세내역 (셀에 배치, 여러 장이면 F·G열)"
    $n = 0; $inCell = 0; $floating = 0
    foreach ($f in $d.fail_rows) {
        if (-not $rowOf.ContainsKey([string]$f.id)) { continue }
        $rr = $rowOf[[string]$f.id]
        $imgs = @($f.image_paths | Where-Object { Test-Path $_ })
        if ($imgs.Count -eq 0) { $ws.Cells.Item($rr, 5).Value2 = "(캡처 없음)"; continue }
        $ws.Rows($rr).RowHeight = $RowHeight
        for ($k = 0; $k -lt $imgs.Count; $k++) {
            $cell = $ws.Cells.Item($rr, 5 + $k)
            $cell.Value2 = $null
            $cell.Select() | Out-Null; Start-Sleep -Milliseconds 300
            $ok = $false
            try { $cell.InsertPictureInCell($imgs[$k]) | Out-Null; $ok = $true; $inCell++ } catch { $ok = $false; if ($inCell -eq 0 -and $floating -eq 0) { "in-cell failed: $($_.Exception.Message)" } }
            if (-not $ok) {
                $shp = $ws.Shapes.AddPicture($imgs[$k], 0, -1, $cell.Left + 2, $cell.Top + 2, -1, -1)
                $shp.LockAspectRatio = -1; $shp.Width = $cell.Width - 4
                if ($shp.Height -gt $cell.Height - 4) { $shp.Height = $cell.Height - 4 }
                $shp.Placement = 1; $floating++
            }
            $n++
        }
    }
    Start-Sleep -Seconds 2; $ws.Range("A3").Select() | Out-Null
    $wb.Save(); $wb.Close($false)
    "fail pictures: $n (in-cell $inCell, floating fallback $floating), columns E..$([char](68 + $maxN)) width $ColWidth, row height $RowHeight"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
