# add_logic_diff_sheet.ps1 : 결과서에 '변경내역(기존vs변경)' 시트를 만들고(payload.logic_diff), '테스트데이터' 시트 끝에
#   payload.testdata_add 구간을 덧붙인다 (2026-10-03, v1.02). 사용자 지시: 테스트 문서에는 기존 로직 vs 변경 로직 차이를 반드시 명시.
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$HDR = 2050429; $OLD = 15921906; $NEW = 14348258
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
function HeadRow($ws, $row, $cells) {
    for ($i = 0; $i -lt $cells.Count; $i++) {
        $c = $ws.Cells.Item($row, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215
        $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.VerticalAlignment = -4108; $c.Borders.LineStyle = 1; $c.WrapText = $true
    }
}
function DataRow($ws, $row, $cells, $fills) {
    for ($i = 0; $i -lt $cells.Count; $i++) {
        $c = $ws.Cells.Item($row, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1
        $c.VerticalAlignment = -4160; $c.WrapText = $true
        if ($fills -and $fills[$i]) { $c.Interior.Color = $fills[$i] }
    }
}
try {
    $wb = $xl.Workbooks.Open($Path)
    # ---------------------------------------------------------------- 1. 변경내역 시트
    $ld = $d.logic_diff
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $ld.sheet) { $s.Delete() } }
    $pos = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $ld.after) { $pos = $s } }
    if (-not $pos) { $pos = $wb.Worksheets.Item($wb.Worksheets.Count) }
    $ws = $wb.Worksheets.Add([Type]::Missing, $pos); $ws.Name = $ld.sheet
    $ws.Cells.Item(1, 1).Value2 = $ld.title; $ws.Cells.Item(1, 1).Font.Size = 14; $ws.Cells.Item(1, 1).Font.Bold = $true
    $ws.Cells.Item(2, 1).Value2 = $ld.note; $ws.Cells.Item(2, 1).Font.Italic = $true
    HeadRow $ws 4 $ld.head
    $r = 5
    foreach ($row in $ld.rows) {
        DataRow $ws $r @($row) @($null, $OLD, $NEW, $null, $null)
        $ws.Cells.Item($r, 1).Font.Bold = $true
        $r++
    }
    $ws.Columns.Item(1).ColumnWidth = 26; $ws.Columns.Item(2).ColumnWidth = 48; $ws.Columns.Item(3).ColumnWidth = 48
    $ws.Columns.Item(4).ColumnWidth = 40; $ws.Columns.Item(5).ColumnWidth = 44
    $ws.Rows("5:$($r - 1)").EntireRow.AutoFit() | Out-Null
    $ws.Cells.Item($r + 1, 1).Value2 = "범례: 회색 = 변경 전, 초록 = 변경 후. 재테스트 결과는 '결함재테스트(1003)' 시트, 화면은 '결함재테스트증적(1003)' 시트."
    $ws.Cells.Item($r + 1, 1).Font.Italic = $true
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false; $xl.ActiveWindow.Zoom = 90
    "logic diff sheet: $($ld.rows.Count) rows"

    # ---------------------------------------------------------------- 2. 테스트데이터 시트 끝에 구간 추가
    $ta = $d.testdata_add
    $wt = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $ta.sheet) { $wt = $s } }
    if ($wt) {
        # 2026-10-07: 같은 제목의 구간이 이미 있으면 그 행부터 끝까지 지우고 다시 쓴다(재실행 시 중복 방지)
        $ur = $wt.UsedRange; $last = $ur.Row + $ur.Rows.Count - 1
        for ($i = 1; $i -le $last; $i++) { if ([string]$wt.Cells.Item($i, 1).Value2 -eq $ta.section) { $wt.Rows("$($i - 1):$last").Clear() | Out-Null; $last = $i - 2; break } }
        $r = $last + 2
        $c = $wt.Cells.Item($r, 1); $c.Value2 = $ta.section; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $r++
        HeadRow $wt $r $ta.head; $r++
        foreach ($row in $ta.rows) { DataRow $wt $r @($row) $null; $r++ }
        $wt.Rows("$($r - $ta.rows.Count):$($r - 1)").EntireRow.AutoFit() | Out-Null
        $wt.Cells.Item($r, 1).Value2 = $ta.note; $wt.Cells.Item($r, 1).Font.Italic = $true
        "testdata section added at row $($r - $ta.rows.Count - 2)"
    } else { "testdata sheet not found - skipped" }

    $wb.Worksheets.Item(1).Activate()
    $wb.Save(); $wb.Close($false)
    "saved"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
