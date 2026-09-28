; =====================================================================
;  Kirill-Lotin - Word va Excel uchun o'rnatuvchi (NSIS 3)
;
;  Yig'ish:  makensis installer/KirillLotin.nsi
;  Natija:   dist/KirillLotin-Setup.exe
;
;  O'rnatuvchi fayllarni %LOCALAPPDATA%\Programs\KirillLotin ga
;  ochadi va install.ps1 ni ishga tushiradi. install.ps1 Word va
;  Excel qo'shimchalarini yasab, STARTUP / XLSTART papkalariga
;  joylaydi. Administrator huquqi kerak emas.
; =====================================================================
Unicode true

!define APPNAME   "Kirill-Lotin"
!define APPVER    "1.0.0"
!define PUBLISHER "Kirill-Lotin"
!define UNINSTKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\KirillLotin"

Name "${APPNAME}"
Caption "${APPNAME} ${APPVER} - Word va Excel uchun"
OutFile "..\dist\KirillLotin-Setup.exe"
InstallDir "$LOCALAPPDATA\Programs\KirillLotin"
RequestExecutionLevel user
SetCompressor /SOLID lzma
ShowInstDetails show
ShowUninstDetails show
BrandingText "${APPNAME} ${APPVER}"

!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "x64.nsh"

!define MUI_ICON   "KirillLotin.ico"
!define MUI_UNICON "KirillLotin.ico"
!define MUI_ABORTWARNING

!define MUI_WELCOMEPAGE_TITLE "Kirill-Lotin o'rnatuvchisi"
!define MUI_WELCOMEPAGE_TEXT "Bu dastur Microsoft Word va Excel ga o'zbek tilidagi matnni kirill alifbosidan lotin alifbosiga va lotindan kirillga o'giradigan qo'shimchani o'rnatadi.$\r$\n$\r$\nDavom etishdan oldin Word va Excel oynalarini yoping.$\r$\n$\r$\n$_CLICK"
!define MUI_FINISHPAGE_TITLE "O'rnatish tugadi"
!define MUI_FINISHPAGE_TEXT "Word yoki Excel ni oching - lentada $\"Kirill-Lotin$\" yorlig'i paydo bo'ladi.$\r$\n$\r$\nTezkor tugmalar:$\r$\n   Alt+Shift+L  -  Kirill > Lotin$\r$\n   Alt+Shift+K  -  Lotin > Kirill$\r$\n$\r$\nExcel formulalari: =LOTINGA(A1), =KIRILGA(A1)"
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\README.md"
!define MUI_FINISHPAGE_SHOWREADME_TEXT "Qo'llanmani ochish"
!define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Uzbek"

VIProductVersion "${APPVER}.0"
VIAddVersionKey /LANG=${LANG_UZBEK} "ProductName"     "${APPNAME}"
VIAddVersionKey /LANG=${LANG_UZBEK} "ProductVersion"  "${APPVER}"
VIAddVersionKey /LANG=${LANG_UZBEK} "FileVersion"     "${APPVER}"
VIAddVersionKey /LANG=${LANG_UZBEK} "FileDescription" "Kirill-Lotin: Word va Excel uchun o'rnatuvchi"
VIAddVersionKey /LANG=${LANG_UZBEK} "CompanyName"     "${PUBLISHER}"
VIAddVersionKey /LANG=${LANG_UZBEK} "LegalCopyright"  "${PUBLISHER}"

Var PowerShell

; 64 bitli Windows da 64 bitli PowerShell ishlatiladi (Office qaysi razryadda bo'lmasin COM ishlaydi)
!macro SetPowerShell
  ${If} ${RunningX64}
    StrCpy $PowerShell "$WINDIR\sysnative\WindowsPowerShell\v1.0\powershell.exe"
  ${Else}
    StrCpy $PowerShell "$SYSDIR\WindowsPowerShell\v1.0\powershell.exe"
  ${EndIf}
  ${IfNot} ${FileExists} "$PowerShell"
    StrCpy $PowerShell "powershell.exe"
  ${EndIf}
!macroend

; Word va Excel yopilganini tekshirish
!macro CheckOfficeFn PREFIX
Function ${PREFIX}CheckOffice
  retry:
    nsExec::ExecToStack '"$PowerShell" -NoProfile -NonInteractive -Command "if (Get-Process WINWORD,EXCEL -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"'
    Pop $0
    Pop $1
    ${If} $0 == 1
      MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Microsoft Word yoki Excel ochiq.$\r$\n$\r$\nIltimos, barcha Word va Excel oynalarini yoping (ishingizni saqlab) va $\"Retry$\" (Qayta urinish) tugmasini bosing." /SD IDCANCEL IDRETRY retry
      Abort
    ${EndIf}
FunctionEnd
!macroend
!insertmacro CheckOfficeFn ""
!insertmacro CheckOfficeFn "un."

Function .onInit
  !insertmacro SetPowerShell
FunctionEnd

Function un.onInit
  !insertmacro SetPowerShell
FunctionEnd

Section "Kirill-Lotin" SecMain
  SectionIn RO
  Call CheckOffice

  SetOutPath "$INSTDIR"
  File "..\install.ps1"
  File "..\uninstall.ps1"
  File "..\README.md"
  File "KirillLotin.ico"
  SetOutPath "$INSTDIR\src"
  File "..\src\KLCore.bas"
  File "..\src\KLWord.bas"
  File "..\src\KLExcel.bas"
  File "..\src\ThisWorkbook.vba"
  File "..\src\customUI.xml"
  SetOutPath "$INSTDIR"

  DetailPrint "Word va Excel qo'shimchalari yasalmoqda, bu 1-2 daqiqa davom etadi..."
  nsExec::ExecToLog '"$PowerShell" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$INSTDIR\install.ps1" -NoPause'
  Pop $0

  ${If} $0 == "error"
    MessageBox MB_OK|MB_ICONSTOP "PowerShell ni ishga tushirib bo'lmadi. O'rnatish to'xtatildi." /SD IDOK
    Abort "PowerShell topilmadi."
  ${ElseIf} $0 == 1
    MessageBox MB_OK|MB_ICONSTOP "Qo'shimchani o'rnatib bo'lmadi.$\r$\n$\r$\nSabablari o'rnatuvchi oynasidagi xabarlarda ko'rsatilgan (Tafsilotlar)." /SD IDOK
    Abort "O'rnatish muvaffaqiyatsiz tugadi."
  ${ElseIf} $0 == 2
    MessageBox MB_OK|MB_ICONEXCLAMATION "Qo'shimcha faqat qisman o'rnatildi (Word yoki Excel uchun xatolik bo'ldi).$\r$\nTafsilotlarni o'rnatuvchi oynasida ko'ring." /SD IDOK
  ${EndIf}

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr   HKCU "${UNINSTKEY}" "DisplayName"     "Kirill-Lotin (Word va Excel uchun)"
  WriteRegStr   HKCU "${UNINSTKEY}" "DisplayVersion"  "${APPVER}"
  WriteRegStr   HKCU "${UNINSTKEY}" "Publisher"       "${PUBLISHER}"
  WriteRegStr   HKCU "${UNINSTKEY}" "DisplayIcon"     "$INSTDIR\KirillLotin.ico"
  WriteRegStr   HKCU "${UNINSTKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr   HKCU "${UNINSTKEY}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
  WriteRegStr   HKCU "${UNINSTKEY}" "QuietUninstallString" "$\"$INSTDIR\Uninstall.exe$\" /S"
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTKEY}" "NoRepair" 1
  WriteRegDWORD HKCU "${UNINSTKEY}" "EstimatedSize" 200
SectionEnd

Section "Uninstall"
  Call un.CheckOffice

  ${If} ${FileExists} "$INSTDIR\uninstall.ps1"
    nsExec::ExecToLog '"$PowerShell" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$INSTDIR\uninstall.ps1" -NoPause'
    Pop $0
  ${EndIf}

  RMDir /r "$INSTDIR\src"
  Delete "$INSTDIR\install.ps1"
  Delete "$INSTDIR\uninstall.ps1"
  Delete "$INSTDIR\README.md"
  Delete "$INSTDIR\KirillLotin.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINSTKEY}"
SectionEnd
