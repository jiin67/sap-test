# dump_xlsx_cells.ps1 : read every sheet of a (DRM) workbook via Excel COM, read-only, and write cell values to JSON.
#   Output: { sheets: [ { name, rows: [[...]], widths: [...], fills: {"r,c": colorInt} } ] }
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    $res = @{ sheets = @() }
    foreach ($ws in @($wb.Worksheets)) {
        $ur = $ws.UsedRange
        $r0 = $ur.Row; $c0 = $ur.Column
        $nr = $ur.Rows.Count + $r0 - 1; $nc = $ur.Columns.Count + $c0 - 1
        $vals = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($nr, $nc)).Value2
        $rows = @()
        for ($r = 1; $r -le $nr; $r++) {
            $row = @()
            for ($c = 1; $c -le $nc; $c++) {
                $v = if ($nr -eq 1 -and $nc -eq 1) { $vals } else { $vals[$r, $c] }
                if ($null -eq $v) { $row += "" } else { $row += [string]$v }
            }
            $rows += , $row
        }
        $widths = @(); for ($c = 1; $c -le $nc; $c++) { $widths += [double]$ws.Columns.Item($c).ColumnWidth }
        # merged areas (template headers use them)
        $merges = @()
        foreach ($cell in @($ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item([Math]::Min($nr, 60), $nc)).Cells)) {
            if ($cell.MergeCells -and $cell.MergeArea.Cells.Item(1, 1).Address() -eq $cell.Address()) { $merges += $cell.MergeArea.Address($false, $false) }
        }
        $res.sheets += @{ name = $ws.Name; rows = $rows; widths = $widths; merges = $merges }
        "sheet $($ws.Name): $nr x $nc, merges $($merges.Count)"
    }
    $wb.Close($false)
    [IO.File]::WriteAllText($Out, ($res | ConvertTo-Json -Depth 6 -Compress), (New-Object Text.UTF8Encoding($false)))
    "saved $Out"
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
