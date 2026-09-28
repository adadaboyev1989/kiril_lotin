; =====================================================================
;  Kirill-Lotin - Word va Excel uchun o'rnatuvchi (NSIS 3)
;
;  Yig'ish:  python3 installer/build_addins.py
;            makensis installer/KirillLotin.nsi
;  Natija:   dist/KirillLotin-Setup.exe
;
;  O'rnatuvchi fayllarni %LOCALAPPDATA%\Programs\KirillLotin ga
;  ochadi va install.ps1 ni ishga tushiradi. install.ps1 tayyor
;  qo'shimchalarni (addins\) Word STARTUP va Excel XLSTART
;  papkalariga nusxalaydi. Administrator huquqi kerak emas.
; =====================================================================
Unicode true

!define APPNAME   "Kirill-Lotin"
!define APPVER    "1.3.0"
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
!include "nsDialogs.nsh"

!define MUI_ICON   "KirillLotin.ico"
!define MUI_UNICON "KirillLotin.ico"
!define MUI_ABORTWARNING

!define MUI_WELCOMEPAGE_TITLE "Kirill-Lotin o'rnatuvchisi"
!define MUI_WELCOMEPAGE_TEXT "Bu dastur Microsoft Word va Excel ga o'zbek tilidagi matnni kirill alifbosidan lotin alifbosiga va lotindan kirillga o'giradigan qo'shimchani o'rnatadi.$\r$\n$\r$\nDavom etishdan oldin Word va Excel oynalarini yoping.$\r$\n$\r$\n$_CLICK"
!define MUI_FINISHPAGE_TITLE "O'rnatish tugadi"
!define MUI_FINISHPAGE_TEXT "Word yoki Excel ni oching - lentada $\"Kirill-Lotin$\" yorlig'i paydo bo'ladi.$\r$\n$\r$\nTezkor tugmalar:$\r$\n   Alt+Shift+L  -  Kirill > Lotin$\r$\n   Alt+Shift+K  -  Lotin > Kirill$\r$\n$\r$\nExcel formulalari: =LOTINGA(A1), =KIRILGA(A1)$\r$\n$\r$\nOffice 2003 da tugmalar $\"Kirill-Lotin$\" asboblar panelida chiqadi."
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\README.md"
!define MUI_FINISHPAGE_SHOWREADME_TEXT "Qo'llanmani ochish"
!define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED

!insertmacro MUI_PAGE_WELCOME
Page custom OfficePageCreate OfficePageLeave
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
UninstPage custom un.OfficePageCreate un.OfficePageLeave
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

;  Office dasturlarini aniqlash va majburan yopish
; ---------------------------------------------------------------------
;  $OpenApps  - ochiq Office dasturlari ro'yxati (matn)
;  $OfficeBusy - 1: Word yoki Excel ochiq (o'rnatishga xalaqit beradi)
Var OpenApps
Var OfficeBusy
Var OfficeLabel
Var OfficeKillBtn

!macro _DetectApp EXE TITLE BLOCKS
  nsExec::ExecToStack '"$SYSDIR\cmd.exe" /c ""$SYSDIR\tasklist.exe" /NH /FI "IMAGENAME eq ${EXE}" | "$SYSDIR\find.exe" /I "${EXE}" >nul"'
  Pop $0
  Pop $1
  ${If} $0 == 0
    StrCpy $OpenApps "$OpenApps$\r$\n      - ${TITLE}"
    ${If} ${BLOCKS} == 1
      StrCpy $OfficeBusy 1
    ${EndIf}
  ${EndIf}
!macroend

!macro _KillApp EXE
  nsExec::ExecToStack '"$SYSDIR\taskkill.exe" /F /T /IM ${EXE}'
  Pop $0
  Pop $1
!macroend

!macro OfficeFns PREFIX
Function ${PREFIX}DetectOffice
  StrCpy $OpenApps ""
  StrCpy $OfficeBusy 0
  !insertmacro _DetectApp WINWORD.EXE  "Microsoft Word"       1
  !insertmacro _DetectApp EXCEL.EXE    "Microsoft Excel"      1
  !insertmacro _DetectApp POWERPNT.EXE "Microsoft PowerPoint" 0
  !insertmacro _DetectApp OUTLOOK.EXE  "Microsoft Outlook"    0
  !insertmacro _DetectApp MSACCESS.EXE "Microsoft Access"     0
  !insertmacro _DetectApp MSPUB.EXE    "Microsoft Publisher"  0
  !insertmacro _DetectApp ONENOTE.EXE  "Microsoft OneNote"    0
  !insertmacro _DetectApp VISIO.EXE    "Microsoft Visio"      0
  !insertmacro _DetectApp WINPROJ.EXE  "Microsoft Project"    0
FunctionEnd

; Barcha ochiq Office dasturlarini majburan yopadi
Function ${PREFIX}KillOffice
  DetailPrint "Office dasturlari majburan yopilmoqda..."
  !insertmacro _KillApp WINWORD.EXE
  !insertmacro _KillApp EXCEL.EXE
  !insertmacro _KillApp POWERPNT.EXE
  !insertmacro _KillApp OUTLOOK.EXE
  !insertmacro _KillApp MSACCESS.EXE
  !insertmacro _KillApp MSPUB.EXE
  !insertmacro _KillApp ONENOTE.EXE
  !insertmacro _KillApp VISIO.EXE
  !insertmacro _KillApp WINPROJ.EXE
  Sleep 1500
FunctionEnd

; O'rnatish/o'chirish bosqichidagi oxirgi tekshiruv (sahifadan keyin
; kimdir Word/Excel ni qayta ochgan bo'lsa)
Function ${PREFIX}CheckOffice
  retry:
    Call ${PREFIX}DetectOffice
    ${If} $OfficeBusy == 1
      MessageBox MB_YESNOCANCEL|MB_ICONEXCLAMATION "Quyidagi Office dasturlari ochiq:$OpenApps$\r$\n$\r$\nHa  -  barchasini MAJBURAN YOPISH (saqlanmagan o'zgarishlar yo'qoladi)$\r$\nYo'q  -  o'zim yopdim, qayta tekshirish$\r$\nBekor qilish  -  to'xtatish" /SD IDCANCEL IDYES force IDNO retry
      Abort
    force:
      Call ${PREFIX}KillOffice
      Goto retry
    ${EndIf}
FunctionEnd

; --- "Office dasturlari ochiq" sahifasi ------------------------------
Function ${PREFIX}OfficePageUpdate
  Call ${PREFIX}DetectOffice
  GetDlgItem $0 $HWNDPARENT 1
  ${If} $OfficeBusy == 1
    ${NSD_SetText} $OfficeLabel "Quyidagi Office dasturlari ochiq:$OpenApps$\r$\n$\r$\nDavom etish uchun Word va Excel yopilishi kerak.$\r$\n$\r$\n1) Ishingizni saqlab, dasturlarni o'zingiz yoping va $\"Qayta tekshirish$\" tugmasini bosing, yoki$\r$\n2) $\"Majburan yopish$\" tugmasini bosing - ro'yxatdagi barcha Office dasturlari darhol yopiladi (saqlanmagan o'zgarishlar yo'qoladi)."
    EnableWindow $0 0
    EnableWindow $OfficeKillBtn 1
  ${Else}
    ${NSD_SetText} $OfficeLabel "Word va Excel yopiq. Davom etish uchun keyingi bosqichga o'ting."
    EnableWindow $0 1
    EnableWindow $OfficeKillBtn 0
  ${EndIf}
FunctionEnd

Function ${PREFIX}OfficePageForce
  Pop $0
  MessageBox MB_YESNO|MB_ICONEXCLAMATION "Ochiq Office dasturlari majburan yopiladi:$OpenApps$\r$\n$\r$\nSaqlanmagan o'zgarishlar YO'QOLADI. Davom etilsinmi?" IDNO skip
  Call ${PREFIX}KillOffice
  skip:
  Call ${PREFIX}OfficePageUpdate
FunctionEnd

Function ${PREFIX}OfficePageRecheck
  Pop $0
  Call ${PREFIX}OfficePageUpdate
FunctionEnd

Function ${PREFIX}OfficePageCreate
  Call ${PREFIX}DetectOffice
  ${If} $OfficeBusy == 0
    Abort                               ; hammasi yopiq - sahifa ko'rsatilmaydi
  ${EndIf}
  !insertmacro MUI_HEADER_TEXT "Office dasturlari ochiq" "Davom etishdan oldin Word va Excel yopilishi kerak."
  nsDialogs::Create 1018
  Pop $1
  ${NSD_CreateLabel} 0 0 100% 70% ""
  Pop $OfficeLabel
  ${NSD_CreateButton} 0 78% 48% 16u "Majburan yopish"
  Pop $OfficeKillBtn
  ${NSD_OnClick} $OfficeKillBtn ${PREFIX}OfficePageForce
  ${NSD_CreateButton} 52% 78% 48% 16u "Qayta tekshirish"
  Pop $1
  ${NSD_OnClick} $1 ${PREFIX}OfficePageRecheck
  Call ${PREFIX}OfficePageUpdate
  nsDialogs::Show
FunctionEnd

Function ${PREFIX}OfficePageLeave
  Call ${PREFIX}DetectOffice
  ${If} $OfficeBusy == 1
    Call ${PREFIX}OfficePageUpdate
    Abort
  ${EndIf}
FunctionEnd
!macroend
!insertmacro OfficeFns ""
!insertmacro OfficeFns "un."

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
  SetOutPath "$INSTDIR\addins"
  File "..\addins\KirillLotin.dotm"
  File "..\addins\KirillLotin.xlam"
  SetOutPath "$INSTDIR"

  DetailPrint "Word va Excel qo'shimchalari o'rnatilmoqda va tekshirilmoqda..."
  nsExec::ExecToLog '"$PowerShell" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$INSTDIR\install.ps1" -NoPause'
  Pop $0

  ${If} $0 == "error"
    MessageBox MB_OK|MB_ICONSTOP "PowerShell ni ishga tushirib bo'lmadi. O'rnatish to'xtatildi.$\r$\n$\r$\nWindows XP/Vista da avval Windows PowerShell 2.0 (yoki yangiroq) ni o'rnating." /SD IDOK
    Abort "PowerShell topilmadi."
  ${ElseIf} $0 == 1
    MessageBox MB_OK|MB_ICONSTOP "Qo'shimchani o'rnatib bo'lmadi.$\r$\n$\r$\nSabablari o'rnatuvchi oynasidagi xabarlarda ko'rsatilgan (Tafsilotlar).$\r$\nBatafsil jurnal: $TEMP\KirillLotin-install.log" /SD IDOK
    Abort "O'rnatish muvaffaqiyatsiz tugadi."
  ${ElseIf} $0 == 2
    MessageBox MB_OK|MB_ICONEXCLAMATION "Qo'shimcha faqat qisman o'rnatildi (Word yoki Excel uchun xatolik bo'ldi).$\r$\nTafsilotlarni o'rnatuvchi oynasida ko'ring.$\r$\nBatafsil jurnal: $TEMP\KirillLotin-install.log" /SD IDOK
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
  RMDir /r "$INSTDIR\addins"
  Delete "$INSTDIR\install.ps1"
  Delete "$INSTDIR\uninstall.ps1"
  Delete "$INSTDIR\README.md"
  Delete "$INSTDIR\KirillLotin.ico"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINSTKEY}"
SectionEnd
