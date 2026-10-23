Unicode true
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "x64.nsh"
Name "Wordplai and Secret Tunnel Lite"
OutFile "${OUTPUT}"
RequestExecutionLevel user
SetCompressor /SOLID lzma
VIProductVersion "0.1.0.0"
VIAddVersionKey "ProductName" "Wordplai and Secret Tunnel Lite"
VIAddVersionKey "FileDescription" "Optional two-application preview installer"
VIAddVersionKey "FileVersion" "0.1.0-preview.1"
VIAddVersionKey "LegalCopyright" "See component notices"
!define MUI_ICON "${ICON}"
!define MUI_WELCOMEPAGE_TEXT "Install Wordplai for Microsoft Word and the optional Secret Tunnel Lite companion.$\r$\n$\r$\nSave and close Word before continuing. Each app keeps its own settings and uninstaller. No account credentials are included."
!define MUI_FINISHPAGE_TEXT "Both component installers completed.$\r$\n$\r$\nOpen Word to use Wordplai. Open Secret Tunnel Lite from Start to configure folders, zrok and GitHub access.$\r$\n$\r$\nEach app can be removed independently from Windows Apps."
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_LANGUAGE "English"

Function .onInit
  ${IfNot} ${IsNativeAMD64}
    MessageBox MB_OK|MB_ICONEXCLAMATION "This preview requires Windows x64 (Intel or AMD)." /SD IDOK
    SetErrorLevel 2
    Abort
  ${EndIf}
  FindWindow $0 "OpusApp"
  ${If} $0 != 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "Save your documents and close Word before installing these apps." /SD IDOK
    SetErrorLevel 2
    Abort
  ${EndIf}
FunctionEnd

Section "Install both applications"
  InitPluginsDir
  SetOutPath "$PLUGINSDIR"
  File /oname=wordplai-setup.exe "${WORD_INSTALLER}"
  File /oname=secret-tunnel-lite-setup.exe "${LITE_INSTALLER}"
  DetailPrint "Installing Wordplai…"
  ClearErrors
  ExecWait '"$PLUGINSDIR\wordplai-setup.exe" /S' $0
  ${If} ${Errors}
    StrCpy $0 1
  ${EndIf}
  ${If} $0 != 0
    MessageBox MB_OK|MB_ICONSTOP "Wordplai did not complete (exit $0). Secret Tunnel Lite was not started. Inspect the component installer before retrying." /SD IDOK
    SetErrorLevel 1
    Abort
  ${EndIf}
  DetailPrint "Installing Secret Tunnel Lite…"
  ClearErrors
  ExecWait '"$PLUGINSDIR\secret-tunnel-lite-setup.exe" /S' $0
  ${If} ${Errors}
    StrCpy $0 1
  ${EndIf}
  ${If} $0 != 0
    MessageBox MB_OK|MB_ICONSTOP "Secret Tunnel Lite did not complete (exit $0). Wordplai remains installed. Inspect the component installer before retrying; no user data was removed." /SD IDOK
    SetErrorLevel 1
    Abort
  ${EndIf}
  SetErrorLevel 0
SectionEnd
