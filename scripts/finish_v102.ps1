# finish_v102.ps1 : 결과서의 개정이력 행(payload.revision_add)과 표지(payload.cover_add)를 쓴다 (2026-10-03, v1.02).
#   fill_template_subs.ps1 의 revision/cover 처리와 같은 방식(텍스트 서식 유지). 사용자가 고친 다른 셀은 건드리지 않는다.
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $rev = $wb.Worksheets.Item(2)      # 개정이력
    $prev = $rev.Rows(6)
    $prev.Copy() | Out-Null; $rev.Rows(7).PasteSpecial(-4122) | Out-Null; $xl.CutCopyMode = 0
    foreach ($p in $d.revision_add.PSObject.Properties) { $c = $rev.Range($p.Name); $c.NumberFormat = "@"; $c.Value2 = $p.Value }
    $rev.Rows(7).WrapText = $true; $rev.Rows(7).VerticalAlignment = -4160; $rev.Rows(7).EntireRow.AutoFit() | Out-Null
    $cov = $wb.Worksheets.Item(1)      # 표지
    foreach ($p in $d.cover_add.PSObject.Properties) { $cov.Range($p.Name).Value2 = $p.Value }
    $cov.Activate()
    $wb.Save(); $wb.Close($false)
    "revision row 7 + cover written"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
