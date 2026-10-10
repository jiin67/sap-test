# db171.ps1 : read-only snapshot of ZCRM_RECPTF (SPMON 202609) via S12; -Shot also saves a screenshot results\shots\S12_ZCRM_RECPTF_<tag>.png
param([Parameter(Mandatory = $true)][string]$Tag, [switch]$Shot)
Set-Location C:\sap-test
if ($Shot) { $env:S12_SHOT = "1" } else { $env:S12_SHOT = "" }
cscript //nologo scripts\S12_se16n_cols.vbs ZCRM_RECPTF 50 $Tag "SPMON,KUNNR_SVC,SUBS_TYPE,KUNNRT,LIFNR,OK_QTY,NETWR,ANETWR,CFDATE,BELNR,BUDAT" "정산기간=202609" | Out-Null
$env:S12_SHOT = ""
$f = "results\raw\S12_ZCRM_RECPTF_$Tag.tsv"
"--- ZCRM_RECPTF 202609 [$Tag]"
Get-Content -Encoding Unicode $f | Select-Object -Skip 4
