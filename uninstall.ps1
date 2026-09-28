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

$wordDir = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$excelDir = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$paths = @(
    (Join-Path $wordDir 'KirillLotin.dotm'), (Join-Path $wordDir 'KirillLotin.dot'),
    (Join-Path $excelDir 'KirillLotin.xlam'), (Join-Path $excelDir 'KirillLotin.xla')
)
# O'rnatuvchi yozib qo'ygan haqiqiy joylar (masalan, lokallashtirilgan STARTUP papkasi)
if (Test-Path $StateFile) {
    foreach ($line in Get-Content -Path $StateFile -Encoding UTF8) {
        $i = $line.IndexOf('=')
        if ($i -gt 0) {
            $p = $line.Substring($i + 1)
            $paths += $p
            $dir = Split-Path -Parent $p
            foreach ($n in @('KirillLotin.dotm', 'KirillLotin.dot', 'KirillLotin.xlam', 'KirillLotin.xla')) {
                $paths += (Join-Path $dir $n)
            }
        }
    }
}

foreach ($p in @(@('WINWORD', 'Word'), @('EXCEL', 'Excel'))) {
    while (Get-Process -Name $p[0] -ErrorAction SilentlyContinue) {
        if ($NoPause) { Write-Host "$($p[1]) ochiq. Avval uni yoping." -ForegroundColor Yellow; exit 1 }
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
# O'rnatuvchi qo'shgan ishonchli joylar
foreach ($ver in @('12.0', '14.0', '15.0', '16.0')) {
    foreach ($app in @('Word', 'Excel')) {
        $tl = "HKCU:\Software\Microsoft\Office\$ver\$app\Security\Trusted Locations\KirillLotin"
        if (Test-Path $tl) { Remove-Item -Path $tl -Recurse -Force }
    }
}
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
