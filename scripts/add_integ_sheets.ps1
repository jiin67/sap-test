# add_integ_sheets.ps1 : 결과서에 통합 테스트용 새 시트 2장을 추가한다 (2026-09-30, v0.05).
#   1) '프로그램별테스트케이스' 시트를 그대로 복사해 같은 양식의 '통합테스트케이스' 시트를 만들고 payload.integ.rows 를 쓴다.
#   2) '테스트증적(캡처)' 와 같은 방식으로 '통합테스트증적(캡처)' 시트를 만들고 payload.integ.evidence 캡처를 넣는다.
#   fill_template_subs.ps1 → add_testdata_sheet*.ps1 다음에 실행한다.
param([Parameter(Mandatory = $true)][string]$Path,
      [Parameter(Mandatory = $true)][string]$Json,
      [string]$Key = 'integ',              # payload 안의 키 (integ / integ2 …)
      [string]$After = '')                 # 이 이름의 시트 뒤에 케이스 시트를 둔다 (없으면 프로그램별테스트케이스 뒤)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$ig = $d.$Key
$RED = 13551615; $GREEN = 14348258

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false

function Set-Cell($ws, $row, $col, $val) { $c = $ws.Cells.Item($row, $col); $c.Value2 = $val; return $c }
function Mark($c, $result) {
    if ($result -eq "FAIL") { $c.Interior.Color = $RED; $c.Font.Bold = $true }
    elseif ($result -eq "PASS") { $c.Interior.Color = $GREEN }
}

try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $ig.sheet -or $s.Name -eq $ig.ev_sheet) { $s.Delete() } }

    # ------------------------------------------------ 1. 통합테스트케이스 (프로그램별테스트케이스 복사)
    $src = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -like "프로그램별*") { $src = $s } }
    if (-not $src) { $src = $wb.Worksheets.Item(5) }
    $pos = $src
    if ($After) { foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $After) { $pos = $s } } }
    $src.Copy([Type]::Missing, $pos)
    $ws = $wb.Worksheets.Item($pos.Index + 1)
    $ws.Name = $ig.sheet

    $firstRow = 5
    $oldN = $d.prog_rows.Count            # 복사본에 들어 있는 기존 케이스 행 수 (원본 = 프로그램별테스트케이스)
    $n = $ig.rows.Count
    if ($oldN -gt $n) {
        $ws.Rows("$($firstRow + $n):$($firstRow + $oldN - 1)").Delete() | Out-Null
    } elseif ($n -gt $oldN) {
        $ws.Rows("$($firstRow + $oldN):$($firstRow + $n - 1)").Insert(-4121) | Out-Null
        $ws.Rows($firstRow).Copy() | Out-Null
        $ws.Rows("$($firstRow + $oldN):$($firstRow + $n - 1)").PasteSpecial(-4122) | Out-Null
        $xl.CutCopyMode = 0
    }
    foreach ($p in $ig.header.PSObject.Properties) { $ws.Range($p.Name).Value2 = $p.Value }
    for ($i = 0; $i -lt $n; $i++) {
        $r = $firstRow + $i
        $row = $ig.rows[$i]
        Set-Cell $ws $r 1 $row.id     | Out-Null
        Set-Cell $ws $r 2 $row.case   | Out-Null
        Set-Cell $ws $r 3 $row.expect | Out-Null
        Set-Cell $ws $r 4 $row.sev    | Out-Null
        Set-Cell $ws $r 5 $row.data   | Out-Null
        $ws.Cells.Item($r, 7).Interior.ColorIndex = -4142
        $ws.Cells.Item($r, 7).Font.Bold = $false
        if ($row.result -ne "") {
            $dt = if ($row.date) { $row.date } elseif ($ig.date) { $ig.date } else { $d.date }   # 2026-10-07: 행별 테스트일 우선
            Set-Cell $ws $r 6 $dt | Out-Null
            Mark (Set-Cell $ws $r 7 $row.result) $row.result
        } else {
            Set-Cell $ws $r 6 "" | Out-Null; Set-Cell $ws $r 7 "" | Out-Null
        }
        for ($cc = 8; $cc -le 11; $cc++) { Set-Cell $ws $r $cc "" | Out-Null }
        $ws.Rows($r).VerticalAlignment = -4160
    }
    $ws.Rows("$firstRow`:$($firstRow + $n - 1)").EntireRow.AutoFit() | Out-Null
    $legendRow = $firstRow + $n + 1
    Set-Cell $ws $legendRow 1 $ig.legend.blank | Out-Null
    Set-Cell $ws ($legendRow + 1) 1 $ig.legend.note | Out-Null
    "integ case sheet: $n rows (copied layout from '$($src.Name)', removed $($oldN - $n) extra rows)"

    # ------------------------------------------------ 2. 통합테스트증적(캡처) (맨 끝)
    $shEv = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $shEv.Name = $ig.ev_sheet
    $shEv.Cells.Item(1, 1).Value2 = $ig.ev_sheet
    $shEv.Cells.Item(1, 1).Font.Size = 14
    $shEv.Cells.Item(1, 1).Font.Bold = $true
    $shEv.Columns.Item(1).ColumnWidth = 18
    $shEv.Columns.Item(2).ColumnWidth = 120
    $lastTc = ""; $rowPtr = 3
    foreach ($e in $ig.evidence) {
        if ($e.tc -ne $lastTc) {
            $shEv.Cells.Item($rowPtr, 1).Value2 = $e.tc
            $shEv.Cells.Item($rowPtr, 1).Font.Bold = $true
            $lastTc = $e.tc
        }
        $shEv.Cells.Item($rowPtr, 2).Value2 = $e.caption
        $shEv.Cells.Item($rowPtr, 2).Font.Bold = $true
        $capTop = $shEv.Rows($rowPtr).Top + $shEv.Rows($rowPtr).Height + 2
        if (Test-Path $e.path) {
            $shp = $shEv.Shapes.AddPicture($e.path, 0, -1, $shEv.Columns.Item(2).Left, $capTop, -1, -1)
            $rowsNeeded = [int][Math]::Ceiling($shp.Height / 15.0) + 3
        } else { $rowsNeeded = 3 }
        $rowPtr = $rowPtr + $rowsNeeded
    }
    $shEv.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    "integ evidence sheet: $($ig.evidence.Count) pictures"

    $wb.Worksheets.Item(1).Activate()
    $wb.Save(); $wb.Close($false)
    ($wb = $null) | Out-Null
    "saved"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
