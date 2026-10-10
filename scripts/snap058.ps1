# snap058.ps1 : read-only snapshot of ZCRM_ASSPTFEEH1 (SPMON 2026.09) for one dealer, used for the D-08 checks.
param([Parameter(Mandatory = $true)][string]$Tag, [string]$Kunnr = "100184")
Set-Location C:\sap-test
$env:S12_SHOT = "1"
cscript //nologo scripts\S12_se16n_cols.vbs ZCRM_ASSPTFEEH1 20 $Tag "SPMON,KUNNR,SUBS_TYPE,ASSPTFEE,CFDATE,CFDATE1,BUDAT,BELNR" "정산기간=202609" "고객=$Kunnr" | Out-Null
$env:S12_SHOT = ""
Get-Content -Encoding Unicode "results\raw\S12_ZCRM_ASSPTFEEH1_$Tag.tsv" | Select-Object -Skip 4
