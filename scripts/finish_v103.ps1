# finish_v103.ps1 : 개정이력 지정 행(payload.revision_add 의 셀 주소 그대로)과 표지(payload.cover_add)를 쓴다 (2026-10-07, v1.03).
#   finish_v102.ps1 과 같되 복사할 서식 행을 -FromRow, 쓸 행을 -ToRow 로 받는다.
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json, [int]$FromRow = 7, [int]$ToRow = 8)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $rev = $wb.Worksheets.Item(2)
    $rev.Rows($FromRow).Copy() | Out-Null; $rev.Rows($ToRow).PasteSpecial(-4122) | Out-Null; $xl.CutCopyMode = 0
    foreach ($p in $d.revision_add.PSObject.Properties) { $c = $rev.Range($p.Name); $c.NumberFormat = "@"; $c.Value2 = $p.Value }
    $rev.Rows($ToRow).WrapText = $true; $rev.Rows($ToRow).VerticalAlignment = -4160; $rev.Rows($ToRow).EntireRow.AutoFit() | Out-Null
    $cov = $wb.Worksheets.Item(1)
    foreach ($p in $d.cover_add.PSObject.Properties) { $cov.Range($p.Name).Value2 = $p.Value }
    $cov.Activate()
    $wb.Save(); $wb.Close($false)
    "revision row $ToRow + cover written"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
