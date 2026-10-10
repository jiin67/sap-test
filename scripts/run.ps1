# run.ps1 : run a VBS test script and capture its log with the right encodings.
#   .\scripts\run.ps1 -Script scripts\S01_compare_query.vbs -ScriptArgs @("2026.01.01","2026.09.09")
param(
    [Parameter(Mandatory = $true)][string]$Script,
    [string[]]$ScriptArgs = @(),
    [string]$LogDir = "C:\sap-test\logs"
)

[Console]::OutputEncoding = [Text.Encoding]::GetEncoding(949)

$stamp = Get-Date -Format "yyyyMMdd_HHmm"
$name = [IO.Path]::GetFileNameWithoutExtension($Script)
$console = Join-Path $LogDir "$($name)_console_$stamp.txt"

Write-Host "run: cscript //nologo $Script $($ScriptArgs -join ' ')"
& cscript //nologo $Script @ScriptArgs 2>&1 | Out-File -Encoding utf8 $console
$code = $LASTEXITCODE
Write-Host "exit code: $code"
Write-Host "console log: $console"

switch ($code) {
    0  { Write-Host "OK" }
    99 { Write-Host "GUARD: 대상 시스템이 KDQ/KDD 가 아니어서 중단됨" }
    default { if ($code -ge 10) { Write-Host "시나리오 실패" } else { Write-Host "연결 단계 실패" } }
}

# UTF-16 VBS logs -> readable copy
Get-ChildItem $LogDir -Filter "$name*.txt" | ForEach-Object {
    Write-Host "log: $($_.FullName)"
}
exit $code
