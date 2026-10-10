# add_fail_rows.ps1 : 이미 작성된 결과서의 'FAIL내역' 시트 표에 payload.fail_rows 중 아직 없는 ID 의 행을 덧붙인다 (2026-10-03, v1.02).
#   기존 행(사용자 수정분)은 건드리지 않는다. 행 서식은 마지막 기존 행을 복사한다. 그림은 add_fail_pictures.ps1 로 따로 넣는다.
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $ws = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq "FAIL내역") { $ws = $s } }
    if (-not $ws) { throw "FAIL내역 sheet not found" }
    $r = 3; $have = @{}
    while ($ws.Cells.Item($r, 1).Value2) { $have[[string]$ws.Cells.Item($r, 1).Value2] = $true; $r++ }
    $last = $r - 1
    $added = 0
    foreach ($f in $d.fail_rows) {
        if ($have.ContainsKey([string]$f.id)) { continue }
        $ws.Rows($last).Copy() | Out-Null
        $ws.Rows($r).PasteSpecial(-4122) | Out-Null          # 서식만
        $xl.CutCopyMode = 0
        $ws.Cells.Item($r, 1).Value2 = $f.id
        $ws.Cells.Item($r, 2).Value2 = $f.case
        $ws.Cells.Item($r, 3).Value2 = $f.tester
        $ws.Cells.Item($r, 4).Value2 = $f.detail
        $ws.Cells.Item($r, 5).Value2 = "아래 캡처 참조"
        $ws.Rows($r).VerticalAlignment = -4160
        $ws.Rows($r).WrapText = $true
        $ws.Rows($r).EntireRow.AutoFit() | Out-Null
        $r++; $added++
    }
    $wb.Save(); $wb.Close($false)
    "fail rows added: $added (table rows 3..$($r - 1))"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
