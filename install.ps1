<#
    Kirill-Lotin o'rnatuvchisi
    --------------------------
    Word va Excel qo'shimchalarini Office avtomatik yuklaydigan papkalarga
    joylaydi (Word STARTUP va Excel XLSTART). Bu papkalar Office uchun
    "ishonchli joy", shuning uchun makroslar ogohlantirishsiz ishlaydi.

    Office 2007 va yangiroq (2007/2010/2013/2016/2019/2021/365, 32/64 bit):
        addins\ dagi tayyor .dotm / .xlam nusxalanadi. Oldin har biri
        Word/Excel da bir marta ochib tekshiriladi; ishlamasa, qo'shimcha
        Word/Excel ning o'zi yordamida qayta yig'iladi (zaxira usul).
    Office 2003 va eskiroq (lenta yo'q):
        qo'shimcha Word/Excel yordamida .dot / .xla formatida yig'iladi,
        tugmalar asboblar panelida chiqadi.

    PowerShell 2.0 (Windows 7) va yangiroq versiyalarda ishlaydi.
    Jurnal: %TEMP%\KirillLotin-install.log

    Ishga tushirish: install.bat faylini ikki marta bosing.
    Parametrlar:  -NoPause    oxirida Enter kutilmaydi (EXE o'rnatuvchi uchun)
                  -SkipCheck  Office'da tekshirmasdan nusxalash (eng tez)
#>
[CmdletBinding()]
param(
    [switch]$NoPause,
    [switch]$SkipCheck
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = Join-Path $Root 'src'
$AddinDir = Join-Path $Root 'addins'
$StateDir = Join-Path $env:APPDATA 'KirillLotin'
$StateFile = Join-Path $StateDir 'installed.txt'
$LogFile = Join-Path $env:TEMP 'KirillLotin-install.log'
$CheckTimeout = 90
$BuildTimeout = 180

function Write-Log([string]$m) {
    try { Add-Content -Path $LogFile -Value ('[' + (Get-Date -Format 'HH:mm:ss') + '] ' + $m) } catch { }
}
function Write-Step([string]$m) { Write-Log $m; Write-Host $m -ForegroundColor Cyan }
function Write-Ok([string]$m)   { Write-Log $m; Write-Host $m -ForegroundColor Green }
function Write-Bad([string]$m)  { Write-Log $m; Write-Host $m -ForegroundColor Yellow }

function Test-ProgId([string]$progId) {
    return $null -ne [type]::GetTypeFromProgID($progId)
}

function Wait-AppClosed([string]$proc, [string]$name) {
    while (Get-Process -Name $proc -ErrorAction SilentlyContinue) {
        if ($NoPause) { throw "$name ochiq. Avval $name ni yoping." }
        Write-Bad "$name ochiq. Iltimos, barcha $name oynalarini yoping va Enter tugmasini bosing..."
        [void](Read-Host)
    }
}

# Office asosiy versiyasi reestrdan (Office'ni ishga tushirmasdan):
# 11 = 2003, 12 = 2007, 14 = 2010, 15 = 2013, 16 = 2016/2019/2021/365
function Get-OfficeMajor([string]$progId) {
    try {
        $cur = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\$progId\CurVer" -ErrorAction Stop).'(default)'
        if ($cur -match '\.(\d+)$') { return [int]$Matches[1] }
    } catch { }
    return 0
}

function Get-DefaultStartup([string]$app, [int]$major) {
    if ($app -eq 'Word') {
        if ($major -gt 0) {
            try {
                $key = "HKCU:\Software\Microsoft\Office\$major.0\Word\Options"
                $p = (Get-ItemProperty -Path $key -Name 'STARTUP-PATH' -ErrorAction Stop).'STARTUP-PATH'
                if ($p) { return [Environment]::ExpandEnvironmentVariables($p) }
            } catch { }
        }
        return Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
    }
    return Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
}

# Skript blokini alohida jarayonda vaqt chegarasi bilan bajarish.
# Office yashirin oynada "osilib" qolsa, jarayon to'xtatiladi.
function Invoke-Timed([scriptblock]$Block, [object[]]$ArgList, [int]$Seconds, [string]$Proc) {
    $job = Start-Job -ScriptBlock $Block -ArgumentList $ArgList
    try {
        if (-not (Wait-Job -Job $job -Timeout $Seconds)) {
            Stop-Job -Job $job
            Get-Process -Name $Proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            throw (New-Object System.TimeoutException "$Seconds soniya ichida javob bo'lmadi")
        }
        return (Receive-Job -Job $job -ErrorAction Stop)
    } finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

# Jobdan qaytgan "startup|versiya|javob" qatorini ajratish
function Split-Result($raw) {
    $text = [string]($raw | Select-Object -Last 1)
    $parts = $text -split '\|', 3
    while ($parts.Count -lt 3) { $parts += '' }
    return @{ Startup = $parts[0]; Version = $parts[1]; Answer = $parts[2] }
}

# ----------------------------------------------------------------------
#  Office ichida bajariladigan bloklar (alohida jarayonda).
#  Word metodlari PowerShell 2.0 da [ref] talab qiladi, yangi
#  versiyalarda esa oddiy qiymat - ikkalasi ham sinab ko'riladi.
# ----------------------------------------------------------------------
$TestWord = {
    param($path)
    $ErrorActionPreference = 'Stop'
    $w = New-Object -ComObject Word.Application
    $no = $false; $yes = $true; $zero = 0
    try {
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $startup = [string]$w.StartupPath
        $ver = [string]$w.Version
        try { $d = $w.Documents.Open($path, $no, $yes, $no) }
        catch { $d = $w.Documents.Open([ref]$path, [ref]$no, [ref]$yes, [ref]$no) }
        try { $answer = [string]$w.Run('KL_Ping') }
        finally { try { $d.Close($zero) } catch { $d.Close([ref]$zero) } }
        "$startup|$ver|$answer"
    } finally {
        try { $w.Quit($zero) } catch { try { $w.Quit([ref]$zero) } catch { } }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    }
}

$TestExcel = {
    param($path)
    $ErrorActionPreference = 'Stop'
    $x = New-Object -ComObject Excel.Application
    try {
        $x.Visible = $false
        $x.DisplayAlerts = $false
        $startup = [string]$x.StartupPath
        $ver = [string]$x.Version
        $wb = $x.Workbooks.Open($path)
        try { $answer = [string]$x.Run("'" + $wb.Name + "'!KL_Ping") }
        finally { $wb.Close($false) }
        "$startup|$ver|$answer"
    } finally {
        try { $x.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
    }
}

$BuildWord = {
    param($src, $out, $legacy)
    $ErrorActionPreference = 'Stop'
    $w = New-Object -ComObject Word.Application
    $zero = 0
    # 1 = wdFormatTemplate (.dot), 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
    if ($legacy) { $fmt = 1 } else { $fmt = 15 }
    try {
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $startup = [string]$w.StartupPath
        $ver = [string]$w.Version
        $d = $w.Documents.Add()
        $comps = $d.VBProject.VBComponents
        [void]$comps.Import((Join-Path $src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $src 'KLWord.bas'))
        try { $d.SaveAs2($out, $fmt) }
        catch {
            try { $d.SaveAs($out, $fmt) } catch { $d.SaveAs([ref]$out, [ref]$fmt) }
        }
        try { $d.Close($zero) } catch { $d.Close([ref]$zero) }
        "$startup|$ver|built"
    } finally {
        try { $w.Quit($zero) } catch { try { $w.Quit([ref]$zero) } catch { } }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    }
}

$BuildExcel = {
    param($src, $out, $legacy)
    $ErrorActionPreference = 'Stop'
    $x = New-Object -ComObject Excel.Application
    # 18 = xlAddIn (.xla), 55 = xlOpenXMLAddIn (.xlam)
    if ($legacy) { $fmt = 18 } else { $fmt = 55 }
    try {
        $x.Visible = $false
        $x.DisplayAlerts = $false
        $x.EnableEvents = $false
        $startup = [string]$x.StartupPath
        $ver = [string]$x.Version
        $wb = $x.Workbooks.Add()
        $comps = $wb.VBProject.VBComponents
        [void]$comps.Import((Join-Path $src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $src 'KLExcel.bas'))
        $codeName = [string]$wb.CodeName
        if (-not $codeName) { $codeName = 'ThisWorkbook' }
        $comps.Item($codeName).CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $src 'ThisWorkbook.vba')))
        $wb.SaveAs($out, $fmt)
        $wb.Close($false)
        "$startup|$ver|built"
    } finally {
        try { $x.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
    }
}

# VBA loyihasiga dasturiy kirishni vaqtincha yoqish (faqat zaxira usulda)
function Enable-VbomAccess([int]$major, [string]$appKey) {
    if ($major -le 0) { $major = 16 }
    $key = "HKCU:\Software\Microsoft\Office\$major.0\$appKey\Security"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    $old = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty -Path $key -Name AccessVBOM -Value 1 -Type DWord
    return @{ Key = $key; Old = $old }
}

function Restore-VbomAccess($state) {
    if ($null -eq $state) { return }
    if ($null -eq $state.Old) {
        Remove-ItemProperty -Path $state.Key -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $state.Key -Name AccessVBOM -Value $state.Old -Type DWord
    }
}

# Zaxira usulda yig'ilgan faylga lenta (Ribbon) XML ni qo'shish.
# System.IO.Packaging (WindowsBase, .NET 3.0+) - PowerShell 2.0 da ham bor.
function Add-RibbonXml([string]$package, [string]$xmlFile) {
    Add-Type -AssemblyName WindowsBase
    $relType = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    $uri = New-Object System.Uri('/customUI/customUI.xml', [System.UriKind]::Relative)
    $pkg = [System.IO.Packaging.Package]::Open($package, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        foreach ($r in @($pkg.GetRelationshipsByType($relType))) { $pkg.DeleteRelationship($r.Id) }
        if ($pkg.PartExists($uri)) { $pkg.DeletePart($uri) }
        $part = $pkg.CreatePart($uri, 'application/xml', [System.IO.Packaging.CompressionOption]::Normal)
        $bytes = [System.IO.File]::ReadAllBytes($xmlFile)
        $stream = $part.GetStream()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
        [void]$pkg.CreateRelationship($uri, [System.IO.Packaging.TargetMode]::Internal, $relType, 'rIdKL')
    } finally {
        $pkg.Close()
    }
}

# Faylni baytma-bayt yozish: yuklab olingan arxivdagi "Internetdan olingan"
# belgisi (Zone.Identifier) nusxaga o'tmaydi, Office makroslarni bloklamaydi.
function Write-FileCopy([string]$from, [string]$to) {
    [System.IO.File]::WriteAllBytes($to, [System.IO.File]::ReadAllBytes($from))
}

# Oldingi versiyalar qoldirgan nusxalarni olib tashlash
function Remove-OldCopies($app, [int]$major) {
    $dirs = @(Get-DefaultStartup $app.Name $major)
    if (Test-Path $StateFile) {
        foreach ($line in Get-Content -Path $StateFile -Encoding UTF8) {
            $i = $line.IndexOf('=')
            if ($i -gt 0 -and $line.Substring(0, $i) -eq $app.Name.ToLower()) {
                $dirs += (Split-Path -Parent $line.Substring($i + 1))
            }
        }
    }
    foreach ($dir in ($dirs | Select-Object -Unique)) {
        foreach ($name in @($app.File, $app.LegacyFile)) {
            $p = Join-Path $dir $name
            if (Test-Path $p) {
                Remove-Item -Path $p -Force
                Write-Log "  eski nusxa o'chirildi: $p"
            }
        }
    }
}

# ----------------------------------------------------------------------
#  Bitta ilova uchun o'rnatish
# ----------------------------------------------------------------------
function Install-Addin($app) {
    Write-Step "$($app.Name) uchun o'rnatilmoqda..."
    Wait-AppClosed $app.Proc $app.Name

    $major = Get-OfficeMajor $app.ProgId
    $legacy = ($major -gt 0 -and $major -lt 12)
    Write-Log "  $($app.ProgId): asosiy versiya $major, eski format: $legacy"
    if ($legacy) { $file = $app.LegacyFile } else { $file = $app.File }

    Remove-OldCopies $app $major

    $tmp = Join-Path $env:TEMP ('KirillLotin_' + [guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($file))
    $prebuilt = Join-Path $AddinDir $app.File
    $startup = $null
    $ok = $false
    $timedOut = $false
    try {
        if (-not $legacy) {
            if (-not (Test-Path $prebuilt)) { throw "Fayl topilmadi: $prebuilt" }
            Write-FileCopy $prebuilt $tmp
            if ($SkipCheck) {
                $ok = $true
            } else {
                Write-Step "  $($app.Name) da tekshirilmoqda (bir necha soniya)..."
                try {
                    $res = Split-Result (Invoke-Timed $app.Test @($tmp) $CheckTimeout $app.Proc)
                    $startup = $res.Startup
                    Write-Log "  $($app.Name) versiyasi: $($res.Version), STARTUP: $($res.Startup)"
                    if ($res.Answer -like 'OK*Shahar*') {
                        $ok = $true
                        Write-Ok "  Tekshiruv: $($res.Answer)"
                    } else {
                        Write-Bad "  Kutilmagan javob: '$($res.Answer)'"
                    }
                } catch [System.TimeoutException] {
                    $timedOut = $true
                    Write-Bad "  $($app.Name) tekshiruvga javob bermadi."
                } catch {
                    Write-Bad "  Tekshiruv xatosi: $($_.Exception.Message)"
                }
            }
        }

        if (-not $ok) {
            if ($legacy) {
                Write-Step "  Office 2003 yoki eskiroq: qo'shimcha $($app.Name) yordamida yig'ilmoqda..."
            } else {
                Write-Step "  Zaxira usul: qo'shimcha $($app.Name) yordamida qayta yig'ilmoqda..."
            }
            Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
            $vbom = Enable-VbomAccess $major $app.Name
            try {
                $res = Split-Result (Invoke-Timed $app.Build @($Src, $tmp, $legacy) $BuildTimeout $app.Proc)
                if ($res.Startup) { $startup = $res.Startup }
                Write-Log "  $($app.Name) versiyasi: $($res.Version), STARTUP: $($res.Startup)"
                if (-not (Test-Path $tmp)) { throw "Fayl saqlanmadi." }
                $ok = $true
            } catch {
                $msg = $_.Exception.Message
                if ($timedOut -and -not $legacy) {
                    # Office umuman javob bermayapti - tayyor fayl tekshiruvsiz o'rnatiladi
                    Write-Bad "  Zaxira usul ham ishlamadi ($msg). Tayyor fayl tekshiruvsiz o'rnatiladi."
                    Write-FileCopy $prebuilt $tmp
                } else {
                    throw "Qo'shimchani yig'ib bo'lmadi: $msg. $($app.Name) > Parametrlar > Ishonch markazi > Makros parametrlari > 'VBA loyihasi obyekt modeliga ishonish' ni yoqib, qayta urinib ko'ring."
                }
            } finally {
                Restore-VbomAccess $vbom
            }
            if ($ok -and -not $legacy) {
                try { Add-RibbonXml $tmp (Join-Path $Src 'customUI.xml') }
                catch { Write-Bad "  Lenta yorlig'ini qo'shib bo'lmadi ($($_.Exception.Message)); tezkor tugmalar ishlaydi." }
            }
        }

        if (-not $startup) { $startup = Get-DefaultStartup $app.Name $major }
        if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
        $dest = Join-Path $startup $file
        Write-FileCopy $tmp $dest
        Write-Ok "  $($app.Name): o'rnatildi -> $dest"
        return $dest
    } finally {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ----------------------------------------------------------------------
#  ASOSIY QISM
# ----------------------------------------------------------------------
try { Set-Content -Path $LogFile -Value ('Kirill-Lotin o''rnatish jurnali, ' + (Get-Date)) } catch { }
Write-Log ('PowerShell ' + $PSVersionTable.PSVersion + ', Windows ' + [Environment]::OSVersion.Version + ', ' + ([IntPtr]::Size * 8) + ' bit')

Write-Host ''
Write-Host '=============================================' -ForegroundColor White
Write-Host '  Kirill-Lotin  -  Word va Excel uchun o''rnatish' -ForegroundColor White
Write-Host '=============================================' -ForegroundColor White
Write-Host ''

$apps = @(
    @{ Name = 'Word';  Proc = 'WINWORD'; ProgId = 'Word.Application';  File = 'KirillLotin.dotm'; LegacyFile = 'KirillLotin.dot'; Test = $TestWord;  Build = $BuildWord },
    @{ Name = 'Excel'; Proc = 'EXCEL';   ProgId = 'Excel.Application'; File = 'KirillLotin.xlam'; LegacyFile = 'KirillLotin.xla'; Test = $TestExcel; Build = $BuildExcel }
)

$installed = @()
$failed = $false
foreach ($app in $apps) {
    if (-not (Test-ProgId $app.ProgId)) {
        Write-Bad "Microsoft $($app.Name) topilmadi - o'tkazib yuborildi."
        continue
    }
    try {
        $dest = Install-Addin $app | Select-Object -Last 1
        $installed += ($app.Name.ToLower() + '=' + $dest)
    } catch {
        $failed = $true
        Write-Bad ("  $($app.Name): xatolik - " + $_.Exception.Message)
    }
}

if (-not (Test-Path $StateDir)) { New-Item -ItemType Directory -Path $StateDir -Force | Out-Null }
if ($installed.Count -gt 0) {
    Set-Content -Path $StateFile -Value $installed -Encoding UTF8
}

Write-Host ''
if ($installed.Count -gt 0 -and -not $failed) {
    Write-Ok "Tayyor! Word yoki Excel ni oching - 'Kirill-Lotin' tugmalari paydo bo'ladi."
    Write-Ok "Tezkor tugmalar: Alt+Shift+L (Kirill -> Lotin), Alt+Shift+K (Lotin -> Kirill)."
} elseif ($installed.Count -gt 0) {
    Write-Bad "Qisman o'rnatildi. Yuqoridagi xabarlarni ko'ring."
} else {
    Write-Bad "O'rnatib bo'lmadi."
}
Write-Host "Jurnal: $LogFile"
try { Copy-Item -Path $LogFile -Destination (Join-Path $StateDir 'install.log') -Force } catch { }

if (-not $NoPause) {
    Write-Host ''
    Write-Host 'Chiqish uchun Enter tugmasini bosing...'
    [void](Read-Host)
}
if ($installed.Count -eq 0) { exit 1 }
if ($failed) { exit 2 }
exit 0
