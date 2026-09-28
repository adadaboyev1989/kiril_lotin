<#
    Kirill-Lotin ni o'chirish: Word va Excel qo'shimchalarini olib tashlaydi.
#>
[CmdletBinding()]
param(
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
$StateDir = Join-Path $env:APPDATA 'KirillLotin'
$StateFile = Join-Path $StateDir 'installed.txt'

$paths = @(
    (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\KirillLotin.dotm'),
    (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\KirillLotin.xlam')
)
if (Test-Path $StateFile) {
    foreach ($line in Get-Content -Path $StateFile -Encoding UTF8) {
        $i = $line.IndexOf('=')
        if ($i -gt 0) { $paths += $line.Substring($i + 1) }
    }
}

foreach ($p in @(@('WINWORD', 'Word'), @('EXCEL', 'Excel'))) {
    while (Get-Process -Name $p[0] -ErrorAction SilentlyContinue) {
        Write-Host "$($p[1]) ochiq. Iltimos, uni yoping va Enter tugmasini bosing..." -ForegroundColor Yellow
        [void](Read-Host)
    }
}

$removed = 0
foreach ($p in ($paths | Select-Object -Unique)) {
    if ($p -and (Test-Path $p)) {
        Remove-Item -Path $p -Force
        Write-Host "O'chirildi: $p" -ForegroundColor Green
        $removed++
    }
}
if (Test-Path $StateDir) { Remove-Item -Path $StateDir -Recurse -Force }
# Sozlamalar (apostrof turi) - HKCU\Software\VB and VBA Program Settings\KirillLotin
$settings = 'HKCU:\Software\VB and VBA Program Settings\KirillLotin'
if (Test-Path $settings) { Remove-Item -Path $settings -Recurse -Force }

if ($removed -eq 0) {
    Write-Host "Kirill-Lotin o'rnatilmagan edi." -ForegroundColor Yellow
} else {
    Write-Host "Kirill-Lotin o'chirildi." -ForegroundColor Green
}

if (-not $NoPause) {
    Write-Host ''
    Write-Host 'Chiqish uchun Enter tugmasini bosing...'
    [void](Read-Host)
}
