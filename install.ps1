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
                  -Diagnose   o'rnatmasdan, faqat tashxis hisobotini tuzish
#>
[CmdletBinding()]
param(
    [switch]$NoPause,
    [switch]$SkipCheck,
    [switch]$Diagnose
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = Join-Path $Root 'src'
$AddinDir = Join-Path $Root 'addins'
$StateDir = Join-Path $env:APPDATA 'KirillLotin'
$StateFile = Join-Path $StateDir 'installed.txt'
$LogFile = Join-Path $env:TEMP 'KirillLotin-install.log'
$ReportFile = Join-Path $env:TEMP 'KirillLotin-tashxis.txt'
$Warnings = New-Object System.Collections.ArrayList
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
    if ($ArgList -and $ArgList.Count -gt 0) {
        $job = Start-Job -ScriptBlock $Block -ArgumentList $ArgList
    } else {
        $job = Start-Job -ScriptBlock $Block
    }
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
#  Office xavfsizlik sozlamalari (Ishonch markazi)
# ----------------------------------------------------------------------
function Get-RegValue([string]$path, [string]$name) {
    try { return (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name } catch { return $null }
}

function Add-Warning([string]$m) {
    if ($Warnings -contains $m) { return }
    [void]$Warnings.Add($m)
    Write-Bad "  DIQQAT: $m"
}

# Office bir marta muammo qilgan qo'shimchani "O'chirilgan elementlar"
# (Disabled Items) ro'yxatiga qo'shadi va boshqa yuklamaydi. Bizning
# fayllarimizga tegishli yozuvlar o'chiriladi.
function Clear-DisabledItems([int]$major, [string]$app) {
    $n = 0
    foreach ($sub in @('DisabledItems', 'StartupItems')) {
        $key = "HKCU:\Software\Microsoft\Office\$major.0\$app\Resiliency\$sub"
        if (-not (Test-Path $key)) { continue }
        $item = Get-Item -Path $key
        foreach ($name in $item.GetValueNames()) {
            $data = $item.GetValue($name)
            if ($data -is [byte[]]) {
                $text = [Text.Encoding]::Unicode.GetString($data).ToLower()
                if ($text.Contains('kirilllotin')) {
                    Remove-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue
                    $n++
                }
            }
        }
    }
    if ($n -gt 0) { Write-Ok "  $app 'O'chirilgan elementlar' ro'yxatidan $n ta yozuv olib tashlandi." }
}

# STARTUP/XLSTART papkasini aniq "ishonchli joy" sifatida qo'shish
# (odatda u allaqachon ishonchli, lekin sozlamalardan olib tashlangan bo'lishi mumkin)
function Add-TrustedLocation([int]$major, [string]$app, [string]$folder) {
    if ($major -lt 12) { return }
    $key = "HKCU:\Software\Microsoft\Office\$major.0\$app\Security\Trusted Locations\KirillLotin"
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    Set-ItemProperty -Path $key -Name 'Path' -Value ($folder.TrimEnd('\') + '\')
    Set-ItemProperty -Path $key -Name 'AllowSubFolders' -Value 1 -Type DWord
    Set-ItemProperty -Path $key -Name 'Description' -Value 'Kirill-Lotin'
    Write-Log "  ishonchli joy qo'shildi: $folder"
}

# Qo'shimchani to'sishi mumkin bo'lgan sozlamalarni aniqlash
function Test-OfficeSecurity([int]$major, [string]$app) {
    if ($major -lt 12) { return }
    $v = "$major.0"
    $roots = @(
        @("HKCU:\Software\Microsoft\Office\$v\$app\Security", 'sizning sozlamangiz'),
        @("HKCU:\Software\Policies\Microsoft\Office\$v\$app\Security", 'guruh siyosati'),
        @("HKLM:\Software\Policies\Microsoft\Office\$v\$app\Security", 'guruh siyosati (kompyuter)')
    )
    $where = "$app > Fayl > Parametrlar > Ishonch markazi > Ishonch markazi parametrlari"
    foreach ($r in $roots) {
        $key = $r[0]; $src = $r[1]
        $x = Get-RegValue $key 'DisableAllAddins'
        if ($x -and [int]$x -ne 0) { Add-Warning "${app}: 'Barcha ilova qo'shimchalarini o'chirish' yoqilgan ($src). O'chiring: $where > Qo'shimchalar." }
        $x = Get-RegValue $key 'RequireAddinSig'
        if ($x -and [int]$x -ne 0) { Add-Warning "${app}: 'Ilova qo'shimchalari ishonchli nashriyotchi imzosiga ega bo'lishi shart' yoqilgan ($src). O'chiring: $where > Qo'shimchalar." }
        $x = Get-RegValue ($key + '\Trusted Locations') 'AllLocationsDisabled'
        if ($x -and [int]$x -ne 0) { Add-Warning "${app}: 'Barcha ishonchli joylarni o'chirish' yoqilgan ($src). O'chiring: $where > Ishonchli joylar." }
        $x = Get-RegValue $key 'VBAWarnings'
        if ($x) { Write-Log "  $app VBAWarnings = $x ($src)" }
    }
    foreach ($k in @("HKCU:\Software\Policies\Microsoft\Office\$v\Common\Security\Trusted Locations", "HKLM:\Software\Policies\Microsoft\Office\$v\Common\Security\Trusted Locations")) {
        $x = Get-RegValue $k 'Allow User Locations'
        if ($null -ne $x -and [int]$x -eq 0) { Add-Warning "Guruh siyosati foydalanuvchi ishonchli joylarini taqiqlagan - STARTUP papkasidagi makroslar ishlamasligi mumkin. Tizim administratoriga murojaat qiling." }
    }
    foreach ($k in @("HKCU:\Software\Policies\Microsoft\Office\$v\Common", "HKLM:\Software\Policies\Microsoft\Office\$v\Common")) {
        $x = Get-RegValue $k 'VBAOff'
        if ($x -and [int]$x -ne 0) { Add-Warning "Office'da VBA (makroslar) guruh siyosati bilan o'chirilgan - qo'shimcha ishlay olmaydi. Tizim administratoriga murojaat qiling." }
    }
}

# O'rnatilgandan keyin: Word qo'shimchani STARTUP dan haqiqatan yuklayaptimi?
# Avval Word o'zi yuklagan qo'shimchalar ichidan qidiriladi; topilmasa
# (avtomatlashtirishda Word STARTUP ni o'tkazib yuborishi mumkin) fayl
# global shablon sifatida qo'lda ulanadi - xavfsizlik sozlamalari bunda ham amal qiladi.
$VerifyWord = {
    param($path)
    $ErrorActionPreference = 'Stop'
    $w = New-Object -ComObject Word.Application
    $zero = 0; $yes = $true
    try {
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $found = ''; $inst = ''; $answer = ''
        foreach ($a in $w.AddIns) {
            if ([string]$a.Name -like 'KirillLotin*') { $found = [string]$a.Path + '\' + [string]$a.Name; $inst = [string]$a.Installed }
        }
        if (-not $found) {
            try {
                try { $a = $w.AddIns.Add($path, $yes) } catch { $a = $w.AddIns.Add([ref]$path, [ref]$yes) }
                $found = $path; $inst = 'qo''lda: ' + [string]$a.Installed
            } catch { $answer = 'XATO: ' + $_.Exception.Message }
        }
        if ($found) {
            try { $answer = [string]$w.Run('KL_Ping') } catch { $answer = 'XATO: ' + $_.Exception.Message }
        }
        "$found|$inst|$answer"
    } finally {
        try { $w.Quit($zero) } catch { try { $w.Quit([ref]$zero) } catch { } }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    }
}

# Natija: @{ Status = 'ok' | 'notloaded' | 'blocked' | 'unknown'; Message = ... }
function Test-WordLoadsAddin([string]$path) {
    Write-Step "Word qo'shimchani ishga tushishda yuklashi tekshirilmoqda..."
    try {
        $r = Split-Result (Invoke-Timed $VerifyWord @($path) $CheckTimeout 'WINWORD')
        Write-Log "  Word AddIns: '$($r.Startup)', Installed=$($r.Version), javob: $($r.Answer)"
        if (-not $r.Startup) {
            return @{ Status = 'notloaded'; Message = "Word qo'shimchani yuklamadi ($($r.Answer)). Word > Fayl > Parametrlar > Qo'shimchalar > Boshqarish: 'O'chirilgan elementlar' va 'Word qo'shimchalari' ni tekshiring." }
        }
        if ($r.Answer -like 'OK*') {
            Write-Ok "  Word qo'shimchani yuklaydi va makroslar ishlaydi."
            return @{ Status = 'ok'; Message = '' }
        }
        return @{ Status = 'blocked'; Message = "Word qo'shimchani yuklaydi, lekin makroslar ishlamayapti ($($r.Answer)). Ishonch markazi sozlamalarini tekshiring." }
    } catch {
        Write-Log "  Word tekshiruvi bajarilmadi: $($_.Exception.Message)"
        return @{ Status = 'unknown'; Message = '' }
    }
}

# Tashxis hisoboti (install.ps1 -Diagnose)
function Write-Diagnostics() {
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Kirill-Lotin tashxis hisoboti - ' + (Get-Date))
    [void]$lines.Add('PowerShell ' + $PSVersionTable.PSVersion + ', Windows ' + [Environment]::OSVersion.Version + ', ' + ([IntPtr]::Size * 8) + ' bitli jarayon')
    foreach ($app in $apps) {
        [void]$lines.Add('')
        [void]$lines.Add("=== $($app.Name) ===")
        if (-not (Test-ProgId $app.ProgId)) { [void]$lines.Add('  o''rnatilmagan'); continue }
        $major = Get-OfficeMajor $app.ProgId
        [void]$lines.Add("  Office versiyasi: $major.0")
        $startup = Get-DefaultStartup $app.Name $major
        foreach ($f in @($app.File, $app.LegacyFile)) {
            $p = Join-Path $startup $f
            if (Test-Path $p) { [void]$lines.Add("  Fayl bor: $p ($((Get-Item $p).Length) bayt)") }
        }
        if (Test-Path $StateFile) {
            foreach ($line in Get-Content -Path $StateFile -Encoding UTF8) {
                if ($line -like ($app.Name.ToLower() + '=*')) {
                    $p = $line.Substring($line.IndexOf('=') + 1)
                    if (Test-Path $p) { $e = 'fayl bor' } else { $e = 'FAYL YO''Q' }
                    [void]$lines.Add("  O'rnatilgan joy: $p ($e)")
                }
            }
        }
        $v = "$major.0"
        foreach ($k in @("HKCU:\Software\Microsoft\Office\$v\$($app.Name)\Security", "HKCU:\Software\Policies\Microsoft\Office\$v\$($app.Name)\Security", "HKLM:\Software\Policies\Microsoft\Office\$v\$($app.Name)\Security")) {
            foreach ($n in @('VBAWarnings', 'DisableAllAddins', 'RequireAddinSig', 'AccessVBOM')) {
                $x = Get-RegValue $k $n
                if ($null -ne $x) { [void]$lines.Add("  $k : $n = $x") }
            }
            $x = Get-RegValue ($k + '\Trusted Locations') 'AllLocationsDisabled'
            if ($null -ne $x) { [void]$lines.Add("  $k\Trusted Locations : AllLocationsDisabled = $x") }
        }
        $tl = "HKCU:\Software\Microsoft\Office\$v\$($app.Name)\Security\Trusted Locations"
        if (Test-Path $tl) {
            foreach ($sub in Get-ChildItem -Path $tl) {
                [void]$lines.Add("  Ishonchli joy: " + (Get-RegValue $sub.PSPath 'Path'))
            }
        }
        foreach ($sub in @('DisabledItems', 'StartupItems')) {
            $key = "HKCU:\Software\Microsoft\Office\$v\$($app.Name)\Resiliency\$sub"
            if (Test-Path $key) {
                $item = Get-Item -Path $key
                foreach ($name in $item.GetValueNames()) {
                    $data = $item.GetValue($name)
                    if ($data -is [byte[]]) {
                        $text = [Text.Encoding]::Unicode.GetString($data) -replace '[^\u0020-\u007E\u0400-\u04FF]', ' '
                        [void]$lines.Add("  $sub : " + $text.Trim())
                    }
                }
            }
        }
        Test-OfficeSecurity $major $app.Name
    }
    [void]$lines.Add('')
    [void]$lines.Add('=== Word ichidan tekshiruv ===')
    if (Test-ProgId 'Word.Application') {
        if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
            [void]$lines.Add('  Word ochiq - tekshiruv o''tkazib yuborildi (Word ni yopib qayta ishga tushiring).')
        } else {
            $wp = ''
            if (Test-Path $StateFile) { foreach ($line in Get-Content -Path $StateFile -Encoding UTF8) { if ($line -like 'word=*') { $wp = $line.Substring(5) } } }
            $v = Test-WordLoadsAddin $wp
            if ($v.Message) { Add-Warning $v.Message }
        }
    }
    [void]$lines.Add('')
    [void]$lines.Add('=== Topilgan muammolar ===')
    if ($Warnings.Count -eq 0) { [void]$lines.Add('  Muammo topilmadi.') }
    foreach ($w in $Warnings) { [void]$lines.Add("  - $w") }
    [void]$lines.Add('')
    [void]$lines.Add('=== O''rnatish jurnali ===')
    if (Test-Path (Join-Path $StateDir 'install.log')) { foreach ($l in Get-Content (Join-Path $StateDir 'install.log')) { [void]$lines.Add($l) } }
    Set-Content -Path $ReportFile -Value $lines -Encoding UTF8
    Write-Host ''
    Write-Host "Tashxis hisoboti: $ReportFile"
    try { Start-Process notepad.exe -ArgumentList ('"' + $ReportFile + '"') } catch { }
}

# ----------------------------------------------------------------------
#  Bitta ilova uchun o'rnatish
# ----------------------------------------------------------------------
$Built = @{}
function Install-Addin($app, [bool]$forceBuild = $false) {
    Write-Step "$($app.Name) uchun o'rnatilmoqda..."
    Wait-AppClosed $app.Proc $app.Name

    $major = Get-OfficeMajor $app.ProgId
    $legacy = ($major -gt 0 -and $major -lt 12)
    Write-Log "  $($app.ProgId): asosiy versiya $major, eski format: $legacy"
    if ($legacy) { $file = $app.LegacyFile } else { $file = $app.File }

    Remove-OldCopies $app $major
    if ($major -gt 0) { Clear-DisabledItems $major $app.Name }

    $tmp = Join-Path $env:TEMP ('KirillLotin_' + [guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($file))
    $prebuilt = Join-Path $AddinDir $app.File
    $startup = $null
    $ok = $false
    $timedOut = $false
    try {
        if (-not $legacy -and -not $forceBuild) {
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
                $Built[$app.Name] = $true
            } catch {
                $msg = $_.Exception.Message
                if ($timedOut -and -not $legacy -and -not $forceBuild) {
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
        if ($major -gt 0) {
            try { Add-TrustedLocation $major $app.Name $startup } catch { Write-Log "  ishonchli joy qo'shilmadi: $($_.Exception.Message)" }
            Test-OfficeSecurity $major $app.Name
        }
        return $dest
    } finally {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ----------------------------------------------------------------------
#  ASOSIY QISM
# ----------------------------------------------------------------------
$apps = @(
    @{ Name = 'Word';  Proc = 'WINWORD'; ProgId = 'Word.Application';  File = 'KirillLotin.dotm'; LegacyFile = 'KirillLotin.dot'; Test = $TestWord;  Build = $BuildWord },
    @{ Name = 'Excel'; Proc = 'EXCEL';   ProgId = 'Excel.Application'; File = 'KirillLotin.xlam'; LegacyFile = 'KirillLotin.xla'; Test = $TestExcel; Build = $BuildExcel }
)

if ($Diagnose) {
    Write-Diagnostics
    if (-not $NoPause) { Write-Host 'Chiqish uchun Enter tugmasini bosing...'; [void](Read-Host) }
    exit 0
}

try { Set-Content -Path $LogFile -Value ('Kirill-Lotin o''rnatish jurnali, ' + (Get-Date)) } catch { }
Write-Log ('PowerShell ' + $PSVersionTable.PSVersion + ', Windows ' + [Environment]::OSVersion.Version + ', ' + ([IntPtr]::Size * 8) + ' bit')

Write-Host ''
Write-Host '=============================================' -ForegroundColor White
Write-Host '  Kirill-Lotin  -  Word va Excel uchun o''rnatish' -ForegroundColor White
Write-Host '=============================================' -ForegroundColor White
Write-Host ''


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

# Word qo'shimchani haqiqatan yuklayaptimi (Excel avtomatlashtirishda XLSTART ni yuklamaydi)
if (-not $SkipCheck -and ($installed -like 'word=*') -and -not (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
    $wordPath = (@($installed -like 'word=*')[0]).Substring(5)
    $check = Test-WordLoadsAddin $wordPath
    if (($check.Status -eq 'notloaded' -or $check.Status -eq 'blocked') -and -not $Built['Word']) {
        # Tayyor fayl bu Office'da ishlamadi - qo'shimchalarni Office'ning o'zi yordamida qayta yig'amiz
        Write-Bad "  Tayyor fayl Word ishga tushganda ishlamadi. Qo'shimchalar Office yordamida qayta yig'ilmoqda..."
        foreach ($app in $apps) {
            if (($installed -like ($app.Name.ToLower() + '=*')) -and -not $Built[$app.Name]) {
                try { [void](Install-Addin $app $true | Select-Object -Last 1) }
                catch { Write-Bad ("  $($app.Name): qayta yig'ib bo'lmadi - " + $_.Exception.Message) }
            }
        }
        $check = Test-WordLoadsAddin $wordPath
    }
    if ($check.Message) { Add-Warning $check.Message }
}
Remove-Item -Path $ReportFile -Force -ErrorAction SilentlyContinue
if ($Warnings.Count -gt 0) {
    $txt = @('Kirill-Lotin o''rnatildi, lekin quyidagi Office sozlamalari uning ishlashiga xalaqit berishi mumkin:', '')
    foreach ($w in $Warnings) { $txt += "- $w"; $txt += '' }
    $txt += 'Sozlamani o''zgartirgach, Word/Excel ni qayta ishga tushiring.'
    $txt += "Batafsil jurnal: $LogFile"
    Set-Content -Path $ReportFile -Value $txt -Encoding UTF8
}

Write-Host ''
if ($installed.Count -gt 0 -and -not $failed -and $Warnings.Count -gt 0) {
    Write-Bad "O'rnatildi, lekin Office sozlamalari qo'shimchani to'sishi mumkin (yuqoridagi DIQQAT xabarlari)."
    Write-Bad "Batafsil: $ReportFile"
} elseif ($installed.Count -gt 0 -and -not $failed) {
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
if ($Warnings.Count -gt 0) { exit 3 }
exit 0
