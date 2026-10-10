# run_kdq.ps1 : concatenate lib_kdq.vbs + <body>.vbs into a scratch script, run it, print its UTF-16 log as UTF-8
param([Parameter(Mandatory=$true)][string]$Body, [string]$LogName = "")
$ErrorActionPreference = "Continue"
if ($LogName -eq "") { $LogName = [IO.Path]::GetFileNameWithoutExtension($Body) }
$lib = Get-Content "C:\sap-test\scripts\lib_kdq.vbs" -Raw -Encoding UTF8
$bod = Get-Content "C:\sap-test\scripts\$Body" -Raw -Encoding UTF8
$tmp = "C:\sap-test\scripts\_combined_$LogName.vbs"
[IO.File]::WriteAllText($tmp, $lib + "`r`n" + $bod, [Text.Encoding]::ASCII)
cscript //nologo C:\sap-test\scripts\cleanup_kdq.vbs
cscript //nologo $tmp $LogName
"EXIT=$LASTEXITCODE"
$ic = "C:\Program Files\Git\usr\bin\iconv.exe"
if (Test-Path "C:\sap-test\logs\$LogName.txt") { & $ic -f UTF-16LE -t UTF-8 "C:\sap-test\logs\$LogName.txt" }
