# t171.ps1 : run one S20 flow on ZCRM171S (-Mode S) or ZCRM171 (-Mode A), settle view, period 2026.09, then print the condensed log.
param([Parameter(Mandatory = $true)][string]$Tag, [Parameter(Mandatory = $true)][ValidateSet('S', 'A')][string]$Mode,
      [Parameter(ValueFromRemainingArguments = $true)][string[]]$Steps)
Set-Location C:\sap-test
$tc = if ($Mode -eq 'S') { 'ZCRM171S' } else { 'ZCRM171' }
$all = @("TC:$tc", 'SELECT:radPA_NEW', 'SET:txtPA_SPMON=2026.09', 'BTN:8', 'WAITIDLE:300') + $Steps
cscript //nologo scripts\S20_flow.vbs $Tag @all
Get-Content -Encoding Unicode "logs\S20_$Tag.txt" |
    Select-String "==== STEP \d+ (BTN|POPBTN|POPKEY|EXEC|SELROWS|CELL|KEY)|sbar=\[[A-Z]\]|#ROW|selected rows|exec rc|modifyCell|ERR|FAILED|exit code|MESSTXT" |
    ForEach-Object { $_.Line.Substring(0, [Math]::Min(230, $_.Line.Length)) } | Get-Unique
