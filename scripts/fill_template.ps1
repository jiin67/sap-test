# fill_template.ps1 : fill the company unit-test-case template via Excel COM.
# Sheets are addressed by index so no Korean literals are needed in this script.
#   1 표지 / 2 개정이력 / 3 작성가이드 / 4 공통테스트케이스 / 5 프로그램별테스트케이스 / 6 FAIL내역
param(
    [string]$Src  = "C:\sap-test\results\_tpl_src.xls",
    [string]$Json = "C:\sap-test\results\_fill_data.json",
    [string]$Dst  = "C:\sap-test\results\SAP_KR_단위테스트케이스및결과서_ZCRMR052_S_v0.01_20260909.xlsx"
)

$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json

Remove-Item $Dst -ErrorAction SilentlyContinue
Copy-Item $Src $Dst -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

function Set-Cell($ws, $row, $col, $val) {
    $c = $ws.Cells.Item($row, $col)
    $c.Value2 = $val
    return $c
}

try {
    $wb = $xl.Workbooks.Open($Dst)

    # =========================================================== 5. 프로그램별테스트케이스
    $ws = $wb.Worksheets.Item(5)

    foreach ($p in $d.prog_header.PSObject.Properties) {
        $ws.Range($p.Name).Value2 = $p.Value
    }

    $rows = $d.prog_rows
    $n = $rows.Count
    $firstRow = 5
    $sampleLast = 7

    # make room: template ships 3 sample rows (5..7)
    if ($n -gt 3) {
        $need = $n - 3
        $ws.Rows("$($sampleLast + 1):$($sampleLast + $need)").Insert(-4121) | Out-Null   # xlDown
        $ws.Rows($firstRow).Copy() | Out-Null
        $ws.Rows("$($firstRow + 1):$($firstRow + $n - 1)").PasteSpecial(-4122) | Out-Null # xlPasteFormats
        $xl.CutCopyMode = 0
    }

    for ($i = 0; $i -lt $n; $i++) {
        $r = $firstRow + $i
        $row = $rows[$i]
        Set-Cell $ws $r 1 $row.id      | Out-Null
        Set-Cell $ws $r 2 $row.case    | Out-Null
        Set-Cell $ws $r 3 $row.expect  | Out-Null
        Set-Cell $ws $r 4 $row.sev     | Out-Null
        Set-Cell $ws $r 5 $row.data    | Out-Null
        if ($row.result -ne "") {
            Set-Cell $ws $r 6 $d.date   | Out-Null
            $c = Set-Cell $ws $r 7 $row.result
            if ($row.result -eq "FAIL") {
                $c.Interior.Color = 13551615      # light red
                $c.Font.Bold = $true
            } else {
                $c.Interior.Color = 14348258      # light green
            }
        } else {
            Set-Cell $ws $r 6 "" | Out-Null
            Set-Cell $ws $r 7 "" | Out-Null
        }
        # reviewer / customer columns stay empty
        for ($cc = 8; $cc -le 11; $cc++) { Set-Cell $ws $r $cc "" | Out-Null }
        $ws.Rows($r).VerticalAlignment = -4160     # xlTop
    }
    $ws.Rows("$firstRow`:$($firstRow + $n - 1)").EntireRow.AutoFit() | Out-Null

    $legendRow = $firstRow + $n + 1
    Set-Cell $ws $legendRow 1 $d.legend.blank | Out-Null
    Set-Cell $ws ($legendRow + 1) 1 $d.legend.note | Out-Null
    "prog sheet: $n rows written"

    # =========================================================== 4. 공통테스트케이스
    $ws = $wb.Worksheets.Item(4)
    $ws.Range("C1").Value2 = $d.prog_header.B1
    $ws.Range("E1").Value2 = $d.prog_header.D1
    $ws.Range("C2").Value2 = $d.prog_header.B2
    $ws.Range("E2").Value2 = $d.prog_header.D2

    $map = @{}
    foreach ($p in $d.common.PSObject.Properties) { $map[$p.Name] = $p.Value }

    $hit = 0; $miss = 0
    for ($r = 5; $r -le 200; $r++) {
        $id = $ws.Cells.Item($r, 2).Text
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $id = $id.Trim()
        if (-not $map.ContainsKey($id)) { $miss++; continue }
        $e = $map[$id]
        Set-Cell $ws $r 5 $e.target | Out-Null
        Set-Cell $ws $r 6 $e.note   | Out-Null
        if ($e.result -ne "") {
            Set-Cell $ws $r 7 $d.date | Out-Null
            $c = Set-Cell $ws $r 8 $e.result
            $c.Interior.Color = 14348258
        } else {
            Set-Cell $ws $r 7 "" | Out-Null
            Set-Cell $ws $r 8 "" | Out-Null
        }
        for ($cc = 9; $cc -le 12; $cc++) { Set-Cell $ws $r $cc "" | Out-Null }
        $hit++
    }
    "common sheet: $hit filled, $miss unmatched"

    # =========================================================== 6. FAIL내역
    $ws = $wb.Worksheets.Item(6)
    # clear the two sample rows (3,4) but keep the guidance row 2
    $ws.Range("A3:E10").ClearContents() | Out-Null
    $r = 3
    foreach ($f in $d.fail_rows) {
        Set-Cell $ws $r 1 $f.id     | Out-Null
        Set-Cell $ws $r 2 $f.case   | Out-Null
        Set-Cell $ws $r 3 $f.tester | Out-Null
        Set-Cell $ws $r 4 $f.detail | Out-Null
        $ws.Rows($r).VerticalAlignment = -4160
        $ws.Rows($r).RowHeight = 150
        $r++
    }
    # place the FAIL evidence pictures under the table
    $picTop = $ws.Rows($r + 1).Top
    $left = $ws.Columns.Item(1).Left + 5
    foreach ($f in $d.fail_rows) {
        foreach ($ip in $f.image_paths) {
            if (Test-Path $ip) {
                $shp = $ws.Shapes.AddPicture($ip, 0, -1, $left, $picTop, -1, -1)
                $picTop = $picTop + $shp.Height + 12
            }
        }
    }
    "fail sheet: $($d.fail_rows.Count) rows + pictures"

    # =========================================================== 새 시트 : 테스트증적
    $shEv = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
    $shEv.Name = $d.evidence_sheet
    $shEv.Cells.Item(1, 1).Value2 = $d.evidence_sheet
    $shEv.Cells.Item(1, 1).Font.Size = 14
    $shEv.Cells.Item(1, 1).Font.Bold = $true
    $shEv.Columns.Item(1).ColumnWidth = 18
    $shEv.Columns.Item(2).ColumnWidth = 120

    $top = $shEv.Rows(3).Top
    $lastTc = ""
    $rowPtr = 3
    foreach ($e in $d.evidence) {
        if ($e.tc -ne $lastTc) {
            $shEv.Cells.Item($rowPtr, 1).Value2 = $e.tc
            $shEv.Cells.Item($rowPtr, 1).Font.Bold = $true
            $lastTc = $e.tc
        }
        $shEv.Cells.Item($rowPtr, 2).Value2 = $e.caption
        $shEv.Cells.Item($rowPtr, 2).Font.Bold = $true
        $capTop = $shEv.Rows($rowPtr).Top + $shEv.Rows($rowPtr).Height + 2
        if (Test-Path $e.path) {
            $shp = $shEv.Shapes.AddPicture($e.path, 0, -1,
                       $shEv.Columns.Item(2).Left, $capTop, -1, -1)
            $rowsNeeded = [int][Math]::Ceiling($shp.Height / 15.0) + 3
        } else {
            $rowsNeeded = 3
        }
        $rowPtr = $rowPtr + $rowsNeeded
    }
    $shEv.Activate()
    $xl.ActiveWindow.DisplayGridlines = $false
    "evidence sheet: $($d.evidence.Count) pictures"

    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    $wb.Close($true)
    "saved: $Dst"
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
}
