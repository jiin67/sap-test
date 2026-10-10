# add_fail_pictures.ps1 : 이미 작성된(사용자가 수정한) 결과서의 'FAIL내역' 시트 표 아래에 payload.fail_rows 의 캡처를 넣는다.
#   파일을 다시 만들지 않고 그 자리에서 그림만 추가한다. 기존 그림이 있으면 지우고 다시 넣는다.
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
    for ($s = $ws.Shapes.Count; $s -ge 1; $s--) { $ws.Shapes.Item($s).Delete() }
    # 표의 마지막 사용 행 찾기 (3행부터 A열이 비는 곳)
    $r = 3
    while ($ws.Cells.Item($r, 1).Value2) { $r++ }
    $picTop = $ws.Rows($r + 1).Top
    $left = $ws.Columns.Item(1).Left + 5
    $n = 0
    foreach ($f in $d.fail_rows) {
        foreach ($ip in $f.image_paths) {
            if (Test-Path $ip) {
                $lbl = $ws.Shapes.AddTextbox(1, $left, $picTop, 400, 18)
                $lbl.TextFrame2.TextRange.Text = "$($f.id)  $($f.case)"
                $lbl.TextFrame2.TextRange.Font.Bold = -1
                $lbl.Line.Visible = 0
                $picTop = $picTop + 20
                $shp = $ws.Shapes.AddPicture($ip, 0, -1, $left, $picTop, -1, -1)
                $picTop = $picTop + $shp.Height + 14
                $n++
            }
        }
    }
    $wb.Save(); $wb.Close($false)
    "fail pictures added: $n (table rows 3..$($r - 1))"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
