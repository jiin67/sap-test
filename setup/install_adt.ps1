# install_adt.ps1 - SAP ADT (ABAP Development Tools) setup for a freshly installed Windows PC
#
# Run in PowerShell (as Administrator):
#   Set-ExecutionPolicy -Scope Process Bypass -Force
#   .\install_adt.ps1
#
# Installs:
#   1. Eclipse Temurin JDK 21              (winget)
#   2. Microsoft Visual C++ 2013 x64 runtime (winget)  - required by ADT's SAP JCo native library
#   3. Microsoft Visual C++ 2015-2022 x64/x86 (winget) - required by SAP GUI
#   4. Git, Python 3.12                     (winget, optional: -SkipExtras to skip)
#   5. Eclipse IDE for Java Developers      (zip -> C:\SAP_ADT\eclipse)
#   6. ABAP Development Tools plug-in       (p2 director from https://tools.hana.ondemand.com/latest)
# Checks (cannot be downloaded publicly):
#   7. SAP GUI for Windows 7.70/8.00 - install from the company package / SAP Software Center
#
# ASCII only on purpose (avoids CP949/UTF-8 issues in Windows PowerShell 5.1).

param(
    [string]$InstallRoot   = 'C:\SAP_ADT',
    [string[]]$EclipseVersions = @('2026-09', '2026-06', '2026-03', '2025-12', '2025-09'),
    [string]$AdtRepo       = 'https://tools.hana.ondemand.com/latest',
    [string[]]$Features    = @(),       # explicit p2 IUs, e.g. 'com.sap.adt.tools.abap.feature.group'
    [switch]$SkipExtras,                # skip Git / Python
    [switch]$SkipEclipse,               # only runtimes
    [int]$XmxMB            = 4096
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
$log = Join-Path $InstallRoot ("install_adt_{0:yyyyMMdd_HHmm}.log" -f (Get-Date))
function Log([string]$m) { $l = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $m; Write-Host $l; Add-Content -Path $log -Value $l }

# ---------- 0. prerequisites ----------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Log 'WARN: not running as Administrator - VC++ runtimes may prompt for UAC.' }
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Log 'ERROR: winget not found. Install "App Installer" from Microsoft Store, then re-run.'
    exit 1
}

function WingetInstall([string]$id) {
    $has = winget list --id $id -e --accept-source-agreements 2>$null | Select-String -SimpleMatch $id
    if ($has) { Log "OK   $id already installed"; return }
    Log "INST $id"
    winget install --id $id -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
    if ($LASTEXITCODE -ne 0) { Log "WARN $id winget exit $LASTEXITCODE" } else { Log "OK   $id" }
}

# ---------- 1-4. runtimes via winget ----------
WingetInstall 'EclipseAdoptium.Temurin.21.JDK'
WingetInstall 'Microsoft.VCRedist.2013.x64'
WingetInstall 'Microsoft.VCRedist.2015+.x64'
WingetInstall 'Microsoft.VCRedist.2015+.x86'
if (-not $SkipExtras) {
    WingetInstall 'Git.Git'
    WingetInstall 'Python.Python.3.12'
}
if ($SkipEclipse) { Log 'Done (runtimes only).'; exit 0 }

# ---------- 5. Eclipse ----------
$eclipseDir = Join-Path $InstallRoot 'eclipse'
$eclipsec   = Join-Path $eclipseDir 'eclipsec.exe'
$usedVer    = $null
if (Test-Path $eclipsec) {
    Log "OK   Eclipse already at $eclipseDir"
} else {
    $zip = Join-Path $InstallRoot 'eclipse.zip'
    foreach ($v in $EclipseVersions) {
        $urls = @(
            "https://download.eclipse.org/technology/epp/downloads/release/$v/R/eclipse-java-$v-R-win32-x86_64.zip",
            "https://www.eclipse.org/downloads/download.php?file=/technology/epp/downloads/release/$v/R/eclipse-java-$v-R-win32-x86_64.zip&r=1"
        )
        foreach ($u in $urls) {
            Log "GET  $u"
            & curl.exe -fL --retry 3 -o $zip $u 2>$null
            if ($LASTEXITCODE -eq 0 -and (Get-Item $zip).Length -gt 100MB) { $usedVer = $v; break }
        }
        if ($usedVer) { break }
    }
    if (-not $usedVer) { Log 'ERROR: Eclipse download failed. Download "Eclipse IDE for Java Developers" (Windows x86_64 zip) manually and unzip to C:\SAP_ADT\eclipse, then re-run.'; exit 2 }
    Log "OK   Eclipse $usedVer downloaded, extracting..."
    Expand-Archive -Path $zip -DestinationPath $InstallRoot -Force
    Remove-Item $zip -Force
}

# eclipse.ini: memory + UTF-8
$ini = Join-Path $eclipseDir 'eclipse.ini'
$t = Get-Content $ini
$t = $t | ForEach-Object { if ($_ -match '^-Xmx') { "-Xmx${XmxMB}m" } else { $_ } }
if (-not ($t -match '^-Dfile.encoding=')) { $t += '-Dfile.encoding=UTF-8' }
Set-Content -Path $ini -Value $t -Encoding ASCII
Log "OK   eclipse.ini -Xmx${XmxMB}m"

# ---------- 6. ADT plug-in ----------
$p2 = @('-nosplash', '-application', 'org.eclipse.equinox.p2.director')
if ($Features.Count -eq 0) {
    Log "LIST $AdtRepo"
    $list = & $eclipsec @p2 -repository $AdtRepo -list 2>&1
    $list | Out-File (Join-Path $InstallRoot 'adt_repo_list.txt') -Encoding ASCII
    $ius = $list | ForEach-Object { ($_ -split '=')[0].Trim() } |
           Where-Object { $_ -match '^com\.sap\.adt\..*\.feature\.group$' -and $_ -match 'abap' -and $_ -notmatch 'source' } |
           Sort-Object -Unique
    $Features = @($ius)
}
if ($Features.Count -eq 0) {
    Log 'WARN: ABAP feature not detected automatically (see adt_repo_list.txt).'
    Log "      Install manually: Eclipse > Help > Install New Software > $AdtRepo > ABAP Development Tools"
} else {
    Log ("INST ADT features: " + ($Features -join ', '))
    $repos = $AdtRepo
    if ($usedVer) { $repos = "$AdtRepo,https://download.eclipse.org/releases/$usedVer" }
    & $eclipsec @p2 -repository $repos -installIU ($Features -join ',') -destination $eclipseDir -profile SDKProfile 2>&1 |
        Tee-Object -FilePath (Join-Path $InstallRoot 'adt_install.txt') | Out-Null
    if ($LASTEXITCODE -ne 0) { Log "WARN: p2 director exit $LASTEXITCODE (see adt_install.txt). Use Help > Install New Software instead." }
    else { Log 'OK   ADT installed' }
}

# desktop shortcut
$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Eclipse ADT.lnk'
$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut($lnk); $s.TargetPath = (Join-Path $eclipseDir 'eclipse.exe'); $s.WorkingDirectory = $eclipseDir; $s.Save()
Log "OK   shortcut $lnk"

# ---------- 7. SAP GUI check ----------
$gui = @('C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe', 'C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe') | Where-Object { Test-Path $_ }
if ($gui) { Log "OK   SAP GUI found: $($gui[0])" }
else { Log 'TODO SAP GUI for Windows not installed - install from the company package (needed for SAP Logon entries and GUI transactions in ADT).' }

Log "Log file: $log"
