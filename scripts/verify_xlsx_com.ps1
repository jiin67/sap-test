# verify_xlsx_com.ps1 : read-only sanity dump of a filled DRM workbook via Excel COM
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    "sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join " | ")
    $ws = $wb.Worksheets.Item(1); "cover C17=" + $ws.Range("C17").Text
    $ws = $wb.Worksheets.Item(2); "revision B5=" + $ws.Range("B5").Text + " C5=" + $ws.Range("C5").Text + " D5=" + $ws.Range("D5").Text + " E5=" + $ws.Range("E5").Text
    $ws = $wb.Worksheets.Item(5)
    "prog header B1=" + $ws.Range("B1").Text + " | D1=" + $ws.Range("D1").Text + " | D2=" + $ws.Range("D2").Text
    $n = 0; $p = 0; $f = 0; $b = 0
    for ($r = 5; $r -le 200; $r++) {
        $id = $ws.Cells.Item($r, 1).Text
        if ([string]::IsNullOrWhiteSpace($id)) { break }
        $n++; $res = $ws.Cells.Item($r, 7).Text
        if ($res -eq "PASS") { $p++ } elseif ($res -eq "FAIL") { $f++ } else { $b++ }
    }
    "prog rows=$n PASS=$p FAIL=$f blank=$b ; last id row=" + ($r - 1) + " ; row5=" + $ws.Cells.Item(5,1).Text + " / " + $ws.Cells.Item(5,2).Text + " / " + $ws.Cells.Item(5,7).Text
    "legend A" + ($r + 1) + "=" + $ws.Cells.Item($r + 1, 1).Text.Substring(0, [Math]::Min(60, $ws.Cells.Item($r + 1, 1).Text.Length))
    $ws = $wb.Worksheets.Item(4)
    "common header C1=" + $ws.Range("C1").Text + " | E2=" + $ws.Range("E2").Text
    $y = 0; $cp = 0; $cf = 0; $filled = 0
    for ($r = 5; $r -le 130; $r++) {
        $id = $ws.Cells.Item($r, 2).Text
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($ws.Cells.Item($r, 5).Text -eq "Y") { $y++ }
        if ($ws.Cells.Item($r, 6).Text -ne "") { $filled++ }
        $res = $ws.Cells.Item($r, 8).Text
        if ($res -eq "PASS") { $cp++ } elseif ($res -eq "FAIL") { $cf++ }
    }
    "common Y=$y noteFilled=$filled PASS=$cp FAIL=$cf"
    $ws = $wb.Worksheets.Item(6)
    "fail A3=" + $ws.Cells.Item(3,1).Text + " | B3=" + $ws.Cells.Item(3,2).Text + " | shapes=" + $ws.Shapes.Count
    $ev = $wb.Worksheets.Item($wb.Worksheets.Count)
    "evidence sheet=" + $ev.Name + " shapes=" + $ev.Shapes.Count
    $wb.Close($false)
} finally {
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect()
}
