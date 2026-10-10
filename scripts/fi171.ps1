# fi171.ps1 : read-only check of FI documents (BKPF header incl. reversal link, BSEG items) via S12, with screenshots.
param([Parameter(Mandatory = $true)][string]$Tag, [Parameter(Mandatory = $true)][string]$Belnr, [string]$Gjahr = "2026")
Set-Location C:\sap-test
$env:S12_SHOT = "1"
cscript //nologo scripts\S12_se16n_cols.vbs BKPF 20 "${Tag}" "BUKRS,BELNR,GJAHR,BLART,BLDAT,BUDAT,USNAM,TCODE,STBLG,STJAH,BKTXT,XBLNR" "전표 번호=$Belnr" "회계연도=$Gjahr" | Out-Null
cscript //nologo scripts\S12_se16n_cols.vbs BSEG 40 "${Tag}" "BUKRS,BELNR,GJAHR,BUZEI,BSCHL,KOART,SHKZG,WRBTR,MWSKZ,LIFNR,HKONT,ZTERM,KOSTL,AUFNR,SGTXT" "전표 번호=$Belnr" "회계연도=$Gjahr" | Out-Null
$env:S12_SHOT = ""
foreach ($t in 'BKPF', 'BSEG') {
    "--- $t [$Tag]"
    Get-Content -Encoding Unicode "logs\S12_${t}_$Tag.txt" | Select-String WARN
    Get-Content -Encoding Unicode "results\raw\S12_${t}_$Tag.tsv" | Select-Object -Skip 4
}
