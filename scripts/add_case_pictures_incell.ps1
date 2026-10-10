# add_case_pictures_incell.ps1 : 각 증적 시트의 캡처를 대응하는 테스트케이스 시트의 표 오른쪽(L열부터)에 '셀에 배치'로 넣는다 (2026-10-08 사용자 요청).
#   케이스 ID(A열, 5행~) 와 증적의 tc 를 맞춰 같은 행에, 여러 장이면 L, M, N … 으로 한 칸씩 옮겨 넣는다. 캡션은 셀 메모(Comment)로 단다.
#   payload: evidence(프로그램별테스트케이스) + integ*.sheet/evidence. 증적 시트 자체는 그대로 둔다.
#   InsertPictureInCell 은 Excel 이 보이고 시트가 활성·셀 선택 상태여야 실제로 들어간다(숨김이면 빈 셀로 저장됨).
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Json, [int]$StartCol = 12, [double]$ColWidth = 48, [double]$RowHeight = 190)
$ErrorActionPreference = "Stop"
$d = Get-Content -Path $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$sets = @(@{ sheet = "프로그램별테스트케이스"; ev = $d.evidence })
foreach ($k in 'integ', 'integ2', 'integ3', 'integ4', 'integ5') { if ($d.$k) { $sets += @{ sheet = $d.$k.sheet; ev = $d.$k.evidence } } }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $true
$total = 0
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($set in $sets) {
        $ws = $null
        foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $set.sheet) { $ws = $s } }
        if (-not $ws) { "sheet not found: $($set.sheet)"; continue }
        $ws.Activate(); $xl.WindowState = -4137; $xl.ActiveWindow.Zoom = 60
        # 케이스 ID -> 행 (5행부터, A열이 ID 패턴일 때만)
        $rowOf = @{}; $r = 5
        while ($r -lt 500) { $v = [string]$ws.Cells.Item($r, 1).Value2; if (-not $v -and -not $ws.Cells.Item($r + 1, 1).Value2) { break }; if ($v -match '^[A-Z]-\d{2}$') { $rowOf[$v] = $r }; $r++ }
        $byTc = @{}
        foreach ($e in $set.ev) { if (-not $byTc.ContainsKey($e.tc)) { $byTc[$e.tc] = @() }; $byTc[$e.tc] += $e }
        $maxN = 1; foreach ($k in $byTc.Keys) { if ($byTc[$k].Count -gt $maxN) { $maxN = $byTc[$k].Count } }
        for ($c = $StartCol; $c -lt $StartCol + $maxN; $c++) { $ws.Columns.Item($c).ColumnWidth = $ColWidth }
        $h = $ws.Cells.Item(3, $StartCol); $h.Value2 = "증적 캡처 (셀에 배치 - 여러 장이면 오른쪽 열로)"; $h.Font.Bold = $true
        $n = 0
        foreach ($tc in ($byTc.Keys | Sort-Object)) {
            if (-not $rowOf.ContainsKey($tc)) { "  [$($set.sheet)] no row for $tc"; continue }
            $rr = $rowOf[$tc]
            if ($ws.Rows($rr).RowHeight -lt $RowHeight) { $ws.Rows($rr).RowHeight = $RowHeight }
            $k = 0
            foreach ($e in $byTc[$tc]) {
                if (-not (Test-Path $e.path)) { continue }
                $cell = $ws.Cells.Item($rr, $StartCol + $k)
                $cell.Value2 = $null
                $cell.Select() | Out-Null; Start-Sleep -Milliseconds 250
                try { $cell.InsertPictureInCell($e.path) | Out-Null } catch { "  insert failed $tc #$k : $($_.Exception.Message)"; continue }
                try { if ($cell.Comment) { $cell.Comment.Delete() }; [void]$cell.AddComment($e.caption) } catch { }
                $k++; $n++
            }
        }
        $total += $n
        "[$($set.sheet)] $n pictures into columns $StartCol..$($StartCol + $maxN - 1) for $($byTc.Keys.Count) cases"
        $ws.Range("A1").Select() | Out-Null
    }
    Start-Sleep -Seconds 2
    $wb.Worksheets.Item(1).Activate()
    $wb.Save(); $wb.Close($false)
    "total in-cell pictures: $total"
}
finally { $xl.Visible = $false; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
