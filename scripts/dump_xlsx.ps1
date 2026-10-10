# dump_xlsx.ps1 : DRM 결과서의 시트 내용(텍스트)을 JSON 으로 덤프한다 → build_plain_xlsx.py 입력 (2026-10-03).
#   형식: { sheets: [ { name, rows: [[cell,...],...], widths: [..], merges: ["C5:F5",..] } ] }
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    $sheets = @()
    foreach ($s in @($wb.Worksheets)) {
        $ur = $s.UsedRange
        $nr = $ur.Row + $ur.Rows.Count - 1
        $nc = $ur.Column + $ur.Columns.Count - 1
        $vals = $s.Range($s.Cells.Item(1, 1), $s.Cells.Item($nr, $nc)).Value2   # 날짜는 일련번호로 오므로 아래에서 NumberFormat 을 보고 yyyy-MM-dd 로 바꾼다 (2026-10-07)
        $rows = @()
        for ($r = 1; $r -le $nr; $r++) {
            $line = @()
            for ($c = 1; $c -le $nc; $c++) {
                $v = if ($nr -eq 1 -and $nc -eq 1) { $vals } else { $vals[$r, $c] }
                if ($null -eq $v) { $line += "" }
                elseif (($v -is [double]) -and $v -gt 40000 -and $v -lt 70000 -and ($s.Cells.Item($r, $c).NumberFormat -match "[yd]")) { $line += [DateTime]::FromOADate($v).ToString("yyyy-MM-dd") }
                else { $line += [string]$v }
            }
            $rows += , $line
        }
        $widths = @(); for ($c = 1; $c -le $nc; $c++) { $widths += [double]$s.Columns.Item($c).ColumnWidth }
        $merges = @()
        foreach ($cell in @($ur.Cells)) {
            if ($cell.MergeCells -and $cell.MergeArea.Cells.Item(1, 1).Address($false, $false) -eq $cell.Address($false, $false)) {
                $merges += $cell.MergeArea.Address($false, $false)
            }
        }
        $sheets += [ordered]@{ name = $s.Name; rows = $rows; widths = $widths; merges = $merges }
        "dumped [$($s.Name)] rows=$nr cols=$nc merges=$($merges.Count)"
    }
    $wb.Close($false)
    $json = [ordered]@{ sheets = $sheets } | ConvertTo-Json -Depth 6 -Compress
    [IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding $false))
    "saved $Out"
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
