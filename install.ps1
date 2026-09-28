<#
    Kirill-Lotin o'rnatuvchisi
    --------------------------
    Microsoft Word va Excel uchun qo'shimchalarni yasaydi va ularni
    avtomatik yuklanadigan papkalarga joylaydi:

      Word  -> %APPDATA%\Microsoft\Word\STARTUP\KirillLotin.dotm
      Excel -> %APPDATA%\Microsoft\Excel\XLSTART\KirillLotin.xlam

    Bu papkalar Office uchun "ishonchli joy" hisoblanadi, shuning uchun
    makroslar ogohlantirishsiz ishlaydi.

    Ishga tushirish: install.bat faylini ikki marta bosing.
#>
[CmdletBinding()]
param(
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = Join-Path $Root 'src'
$StateDir = Join-Path $env:APPDATA 'KirillLotin'
$StateFile = Join-Path $StateDir 'installed.txt'
$FileBase = 'KirillLotin'

function Write-Step([string]$m) { Write-Host $m -ForegroundColor Cyan }
function Write-Ok([string]$m)   { Write-Host $m -ForegroundColor Green }
function Write-Bad([string]$m)  { Write-Host $m -ForegroundColor Yellow }

function Release-Com($o) {
    if ($null -ne $o) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { }
    }
}

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

function Get-OfficeVersion([string]$progId) {
    $app = New-Object -ComObject $progId
    try { return [string]$app.Version }
    finally {
        try { $app.Quit() } catch { }
        Release-Com $app
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# VBA loyihasiga dasturiy kirishni vaqtincha yoqish (o'rnatish tugagach qaytariladi)
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

# Office faylining ichiga lenta (Ribbon) XML ni joylash
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
            $node.SetAttribute('Id', 'klCustomUI')
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

function New-TempPath([string]$ext) {
    return Join-Path $env:TEMP ($FileBase + '_' + [guid]::NewGuid().ToString('N') + $ext)
}

# ----------------------------------------------------------------------
#  WORD
# ----------------------------------------------------------------------
function Install-Word {
    Write-Step "Word uchun qo'shimcha yasalmoqda..."
    $ver = Get-OfficeVersion 'Word.Application'
    $vbom = Enable-VbomAccess $ver 'Word'
    $word = $null; $doc = $null
    $tmp = New-TempPath '.dotm'
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $startup = [string]$word.StartupPath
        if ([string]::IsNullOrEmpty($startup)) { $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP' }

        $doc = $word.Documents.Add()
        try { $comps = $doc.VBProject.VBComponents }
        catch { throw "VBA loyihasiga kirish taqiqlangan. Word > Fayl > Parametrlar > Ishonch markazi > Makros parametrlari > 'VBA loyihasi obyekt modeliga ishonish' ni yoqing va qayta urinib ko'ring." }
        [void]$comps.Import((Join-Path $Src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $Src 'KLWord.bas'))

        # Tezkor tugmalar: Alt+Shift+L va Alt+Shift+K
        try {
            $word.CustomizationContext = $doc
            $alt = 1024; $shift = 256
            [void]$word.KeyBindings.Add(2, 'KL_WordToLatin', $word.BuildKeyCode($alt, $shift, 76))
            [void]$word.KeyBindings.Add(2, 'KL_WordToCyrillic', $word.BuildKeyCode($alt, $shift, 75))
        } catch {
            Write-Bad "  Tezkor tugmalarni o'rnatib bo'lmadi (lenta tugmalari baribir ishlaydi)."
        }

        # 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
        try { $doc.SaveAs2($tmp, 15) } catch { $doc.SaveAs([ref]$tmp, [ref]15) }
        try { $doc.Close(0) } catch { $doc.Close([ref]0) }
        Release-Com $doc; $doc = $null
        try { $word.Quit(0) } catch { $word.Quit([ref]0) }
        Release-Com $word; $word = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 1

        Add-RibbonXml $tmp (Join-Path $Src 'customUI.xml')

        if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
        $dest = Join-Path $startup "$FileBase.dotm"
        Copy-Item -Path $tmp -Destination $dest -Force
        Write-Ok "  Word: o'rnatildi -> $dest"
        return $dest
    } finally {
        if ($doc) { try { $doc.Close(0) } catch { }; Release-Com $doc }
        if ($word) { try { $word.Quit(0) } catch { }; Release-Com $word }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Restore-VbomAccess $vbom
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

# ----------------------------------------------------------------------
#  EXCEL
# ----------------------------------------------------------------------
function Install-Excel {
    Write-Step "Excel uchun qo'shimcha yasalmoqda..."
    $ver = Get-OfficeVersion 'Excel.Application'
    $vbom = Enable-VbomAccess $ver 'Excel'
    $excel = $null; $wb = $null
    $tmp = New-TempPath '.xlam'
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $startup = [string]$excel.StartupPath
        if ([string]::IsNullOrEmpty($startup)) { $startup = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART' }

        $wb = $excel.Workbooks.Add()
        try { $comps = $wb.VBProject.VBComponents }
        catch { throw "VBA loyihasiga kirish taqiqlangan. Excel > Fayl > Parametrlar > Ishonch markazi > Makros parametrlari > 'VBA loyihasi obyekt modeliga ishonish' ni yoqing va qayta urinib ko'ring." }
        [void]$comps.Import((Join-Path $Src 'KLCore.bas'))
        [void]$comps.Import((Join-Path $Src 'KLExcel.bas'))

        $codeName = [string]$wb.CodeName
        if ([string]::IsNullOrEmpty($codeName)) { $codeName = 'ThisWorkbook' }
        $code = [System.IO.File]::ReadAllText((Join-Path $Src 'ThisWorkbook.vba'))
        $comps.Item($codeName).CodeModule.AddFromString($code)

        # 55 = xlOpenXMLAddIn (.xlam)
        $wb.SaveAs($tmp, 55)
        $wb.Close($false)
        Release-Com $wb; $wb = $null
        $excel.Quit()
        Release-Com $excel; $excel = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 1

        Add-RibbonXml $tmp (Join-Path $Src 'customUI.xml')

        if (-not (Test-Path $startup)) { New-Item -ItemType Directory -Path $startup -Force | Out-Null }
        $dest = Join-Path $startup "$FileBase.xlam"
        Copy-Item -Path $tmp -Destination $dest -Force
        Write-Ok "  Excel: o'rnatildi -> $dest"
        return $dest
    } finally {
        if ($wb) { try { $wb.Close($false) } catch { }; Release-Com $wb }
        if ($excel) { try { $excel.Quit() } catch { }; Release-Com $excel }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Restore-VbomAccess $vbom
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
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

$installed = @()
$failed = $false

if (Test-ProgId 'Word.Application') {
    try {
        Wait-AppClosed 'WINWORD' 'Word'
        $installed += "word=" + (Install-Word | Select-Object -Last 1)
    }
    catch { $failed = $true; Write-Bad ("  Word: xatolik - " + $_.Exception.Message) }
} else {
    Write-Bad 'Microsoft Word topilmadi - o''tkazib yuborildi.'
}

if (Test-ProgId 'Excel.Application') {
    try {
        Wait-AppClosed 'EXCEL' 'Excel'
        $installed += "excel=" + (Install-Excel | Select-Object -Last 1)
    }
    catch { $failed = $true; Write-Bad ("  Excel: xatolik - " + $_.Exception.Message) }
} else {
    Write-Bad 'Microsoft Excel topilmadi - o''tkazib yuborildi.'
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
