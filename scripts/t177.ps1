# t177.ps1 : run one S20 flow on ZCRM177A (-Mode A) or ZCRM177AS (-Mode S) for the ZCRMR177_1 test.
#   -Subs adds the hidden 구독여부 column to the current ALV view (layout dialog, not saved).
#   Extra S20 steps can follow.
param([Parameter(Mandatory = $true, Position = 0)][string]$Tag,
    [Parameter(Mandatory = $true, Position = 1)][ValidateSet('A', 'S')][string]$Mode,
    [Parameter()][string]$From = '2026.09.19', [Parameter()][string]$To = '2026.09.19',
    [Parameter()][string]$Cslt = '', [Parameter()][string]$CsltTo = '',
    [switch]$Comp, [switch]$Subs,
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Steps)
Set-Location C:\sap-test
[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)
$tc = if ($Mode -eq 'S') { 'ZCRM177AS' } else { 'ZCRM177A' }
$add = 'usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING'

$all = @("TC:$tc", "SET:ctxtSO_INDAT-LOW=$From", "SET:ctxtSO_INDAT-HIGH=$To")
if ($Cslt) { $all += "SET:txtSO_CSLTN-LOW=$Cslt" }
if ($CsltTo) { $all += "SET:txtSO_CSLTN-HIGH=$CsltTo" }
if ($Comp) { $all += 'CHK:chkP_COMP=X' }
$all += @('BTN:8', 'WAITIDLE:300')
if ($Subs) { $all += @('GRIDMENU:&MB_VARIANT,&COL0', 'WAITIDLE:60', 'POPSELCOL:구독여부', "POPBTN:$add", 'WAITIDLE:30', 'POPKEY:0', 'WAITIDLE:60') }
$all += $Steps

cscript //nologo scripts\S20_flow.vbs $Tag @all
Get-Content -Encoding Unicode "logs\S20_$Tag.txt" |
    Select-String "==== STEP \d+ (BTN|POPBTN|POPKEY|SELROWS|CELL|KEY|TBTN)|sbar=\[[A-Z]\]|#COLS|#TITLES|#ROW|POPSELCOL|grids=|exit code|MESSTXT|STEP FAILED" |
    ForEach-Object { $_.Line }
