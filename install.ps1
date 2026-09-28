<#
    Kirill-Lotin o'rnatuvchisi
    --------------------------
    addins\ papkasidagi tayyor qo'shimchalarni Office avtomatik yuklaydigan
    papkalarga nusxalaydi:

      Word  -> %APPDATA%\Microsoft\Word\STARTUP\KirillLotin.dotm
      Excel -> %APPDATA%\Microsoft\Excel\XLSTART\KirillLotin.xlam

    Bu papkalar Office uchun "ishonchli joy" hisoblanadi, shuning uchun
    makroslar ogohlantirishsiz ishlaydi.

    Nusxalashdan oldin har bir qo'shimcha Word/Excel da bir marta ochib
    tekshiriladi (vaqt chegarasi bilan). Agar tayyor fayl ishlamasa, zaxira
    usulda qo'shimcha Office'ning o'zi yordamida qayta yig'iladi.

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
$CheckTimeout = 90
$BuildTimeout = 180

function Write-Step([string]$m) { Write-Host $m -ForegroundColor Cyan }
function Write-Ok([string]$m)   { Write-Host $m -ForegroundColor Green }
function Write-Bad([string]$m)  { Write-Host $m -ForegroundColor Yellow }

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

# Office versiyasi reestrdan olinadi (Office'ni ishga tushirmasdan)
function Get-OfficeVersion([string]$progId) {
    try {
        $cur = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\$progId\CurVer" -ErrorAction Stop).'(default)'
        if ($cur -match '\.(\d+)$') { return "$($Matches[1]).0" }
    } catch { }
    return '16.0'
}

function Get-StartupFolder([string]$app) {
    if ($app -eq 'Word') {
        $ver = Get-OfficeVersion 'Word.Application'
        try {
            $p = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ver\Word\Options" -Name 'STARTUP-PATH' -ErrorAction Stop).'STARTUP-PATH'
            if ($p) { return [Environment]::ExpandEnvironmentVariables($p) }
        } catch { }
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

# ----------------------------------------------------------------------
#  Tekshiruv: qo'shimchani ochib, KL_Ping makrosini chaqirish
# ----------------------------------------------------------------------
$TestWord = {
    param($path)
    $ErrorActionPreference = 'Stop'
    $w = New-Object -ComObject Word.Application
    try {
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $d = $w.Documents.Open($path, $false, $true, $false)
        try { [string]$w.Run('KL_Ping') } finally { $d.Close(0) }
    } finally {
        try { $w.Quit(0) } catch { }
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
        $wb = $x.Workbooks.Open($path)
        try { [string]$x.Run("'" + $wb.Name + "'!KL_Ping") } finally { $wb.Close($false) }
    } finally {
        try { $x.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
    }
}

# ----------------------------------------------------------------------
#  Zaxira usul: qo'shimchani Office yordamida yig'ish
# ----------------------------------------------------------------------
$BuildWord = {
    param($src, $out)
    $ErrorActionPreference = 'Stop'
    $w = New-Object -ComObject Word.Application
    try {
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $d = $w.Documents.Add()
        $comps = $d.VBProject.VBComponents
        [void]$comps.Import((Join-Path $src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $src 'KLWord.bas'))
        $d.SaveAs2($out, 15)      # wdFormatXMLTemplateMacroEnabled
        $d.Close(0)
    } finally {
        try { $w.Quit(0) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    }
}

$BuildExcel = {
    param($src, $out)
    $ErrorActionPreference = 'Stop'
    $x = New-Object -ComObject Excel.Application
    try {
        $x.Visible = $false
        $x.DisplayAlerts = $false
        $x.EnableEvents = $false
        $wb = $x.Workbooks.Add()
        $comps = $wb.VBProject.VBComponents
        [void]$comps.Import((Join-Path $src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $src 'KLExcel.bas'))
        $codeName = [string]$wb.CodeName
        if (-not $codeName) { $codeName = 'ThisWorkbook' }
        $comps.Item($codeName).CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $src 'ThisWorkbook.vba')))
        $wb.SaveAs($out, 55)      # xlOpenXMLAddIn
        $wb.Close($false)
    } finally {
        try { $x.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
    }
}

# VBA loyihasiga dasturiy kirishni vaqtincha yoqish (faqat zaxira usulda)
function Enable-VbomAccess([string]$ver, [string]$appKey) {
    $key = "HKCU:\Software\Microsoft\Office\$ver\$appKey\Security"
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

# Zaxira usulda yig'ilgan faylga lenta (Ribbon) XML ni qo'shish
function Add-RibbonXml([string]$package, [string]$xmlFile) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $relType = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    $relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $zip = [System.IO.Compression.ZipFile]::Open($package, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $old = $zip.GetEntry('customUI/customUI.xml')
        if ($old) { $old.Delete() }
        $bytes = [System.IO.File]::ReadAllBytes($xmlFile)
        $stream = $zip.CreateEntry('customUI/customUI.xml').Open()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()

        $relsEntry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsText = $reader.ReadToEnd()
        $reader.Close()
        [xml]$rels = $relsText
        $exists = $false
        foreach ($r in $rels.DocumentElement.ChildNodes) {
            if ($r -is [System.Xml.XmlElement] -and $r.GetAttribute('Type') -eq $relType) { $exists = $true }
        }
        if (-not $exists) {
            $node = $rels.CreateElement('Relationship', $relNs)
            $node.SetAttribute('Id', 'rIdKL')
            $node.SetAttribute('Type', $relType)
            $node.SetAttribute('Target', 'customUI/customUI.xml')
            [void]$rels.DocumentElement.AppendChild($node)
        }
        $relsEntry.Delete()
        $writer = New-Object System.IO.StreamWriter($zip.CreateEntry('_rels/.rels').Open(), (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write($rels.OuterXml)
        $writer.Close()
    } finally {
        $zip.Dispose()
    }
}

# ----------------------------------------------------------------------
#  Bitta ilova uchun o'rnatish
# ----------------------------------------------------------------------
function Install-Addin($app) {
    Write-Step "$($app.Name) uchun o'rnatilmoqda..."
    Wait-AppClosed $app.Proc $app.Name

    $prebuilt = Join-Path $AddinDir $app.File
    if (-not (Test-Path $prebuilt)) { throw "Fayl topilmadi: $prebuilt" }

    $startup = Get-StartupFolder $app.Name
    if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
    $dest = Join-Path $startup $app.File
    # Eski nusxa tekshiruvga xalaqit bermasligi uchun avval olib tashlanadi
    if (Test-Path $dest) { Remove-Item -Path $dest -Force }

    $ext = [IO.Path]::GetExtension($app.File)
    $tmp = Join-Path $env:TEMP ('KirillLotin_' + [guid]::NewGuid().ToString('N') + $ext)
    Copy-Item -Path $prebuilt -Destination $tmp -Force
    try {
        $ok = $true
        if (-not $SkipCheck) {
            Write-Step "  $($app.Name) da tekshirilmoqda (bir necha soniya)..."
            try {
                $answer = [string](Invoke-Timed $app.Test @($tmp) $CheckTimeout $app.Proc)
                if ($answer -like 'OK*Shahar*') {
                    Write-Ok "  Tekshiruv: $answer"
                } else {
                    $ok = $false
                    Write-Bad "  Kutilmagan javob: '$answer'"
                }
            } catch [System.TimeoutException] {
                Write-Bad "  $($app.Name) tekshiruvga javob bermadi - tekshiruvsiz o'rnatiladi."
            } catch {
                $ok = $false
                Write-Bad "  Tekshiruv xatosi: $($_.Exception.Message)"
            }
        }

        if (-not $ok) {
            Write-Step "  Zaxira usul: qo'shimcha $($app.Name) yordamida qayta yig'ilmoqda..."
            Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
            $vbom = Enable-VbomAccess (Get-OfficeVersion $app.ProgId) $app.Name
            try {
                Invoke-Timed $app.Build @($Src, $tmp) $BuildTimeout $app.Proc | Out-Null
            } catch {
                throw "Qo'shimchani yig'ib bo'lmadi: $($_.Exception.Message). Ehtimol, $($app.Name) > Fayl > Parametrlar > Ishonch markazi > Makros parametrlari > 'VBA loyihasi obyekt modeliga ishonish' ni yoqish kerak."
            } finally {
                Restore-VbomAccess $vbom
            }
            Add-RibbonXml $tmp (Join-Path $Src 'customUI.xml')
        }

        Copy-Item -Path $tmp -Destination $dest -Force
        Write-Ok "  $($app.Name): o'rnatildi -> $dest"
        return $dest
    } finally {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
    }
}

# ----------------------------------------------------------------------
#  ASOSIY QISM
# ----------------------------------------------------------------------
Write-Host ''
Write-Host '=============================================' -ForegroundColor White
Write-Host '  Kirill-Lotin  -  Word va Excel uchun o''rnatish' -ForegroundColor White
Write-Host '=============================================' -ForegroundColor White
Write-Host ''

$apps = @(
    @{ Name = 'Word';  Proc = 'WINWORD'; ProgId = 'Word.Application';  File = 'KirillLotin.dotm'; Test = $TestWord;  Build = $BuildWord },
    @{ Name = 'Excel'; Proc = 'EXCEL';   ProgId = 'Excel.Application'; File = 'KirillLotin.xlam'; Test = $TestExcel; Build = $BuildExcel }
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

if ($installed.Count -gt 0) {
    if (-not (Test-Path $StateDir)) { New-Item -ItemType Directory -Path $StateDir -Force | Out-Null }
    Set-Content -Path $StateFile -Value $installed -Encoding UTF8
}

Write-Host ''
if ($installed.Count -gt 0 -and -not $failed) {
    Write-Ok "Tayyor! Word yoki Excel ni oching - lentada 'Kirill-Lotin' yorlig'i paydo bo'ladi."
    Write-Ok "Tezkor tugmalar: Alt+Shift+L (Kirill -> Lotin), Alt+Shift+K (Lotin -> Kirill)."
} elseif ($installed.Count -gt 0) {
    Write-Bad "Qisman o'rnatildi. Yuqoridagi xabarlarni ko'ring."
} else {
    Write-Bad "O'rnatib bo'lmadi."
}

if (-not $NoPause) {
    Write-Host ''
    Write-Host 'Chiqish uchun Enter tugmasini bosing...'
    [void](Read-Host)
}
if ($installed.Count -eq 0) { exit 1 }
if ($failed) { exit 2 }
exit 0
