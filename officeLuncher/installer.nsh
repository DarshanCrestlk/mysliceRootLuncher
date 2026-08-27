!include "LogicLib.nsh"

!macro preInit
  SetRegView 64
  ReadEnvStr $0 PROGRAMDATA
  StrCmp $0 "" 0 +2
  StrCpy $0 "C:\ProgramData"
  WriteRegExpandStr HKLM "${INSTALL_REGISTRY_KEY}" InstallLocation "$0\myslice\mysliceLTS\launcher"
  WriteRegExpandStr HKCU "${INSTALL_REGISTRY_KEY}" InstallLocation "$0\myslice\mysliceLTS\launcher"
!macroend

!macro customInstall
  DetailPrint "Running MySlice setup (protocol, share, Office catalog)..."
  ExecWait '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$INSTDIR\resources\install.ps1" -ExePath "$INSTDIR\mysliceLTS.exe" -ManifestSource "$INSTDIR\resources\manifest.xml"' $0
  ${If} $0 == 0
    MessageBox MB_OK|MB_ICONINFORMATION "MySlice LTS setup complete. Restart Word/Excel if they are open."
  ${Else}
    MessageBox MB_OK|MB_ICONEXCLAMATION "MySlice setup had a problem.$\r$\n$\r$\nRun the installer as Administrator."
  ${EndIf}
!macroend

!macro customUnInstall
  IfFileExists "$INSTDIR\resources\uninstall.ps1" 0 skip_myslice_uninstall
    ExecWait '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$INSTDIR\resources\uninstall.ps1"' $0
  skip_myslice_uninstall:
!macroend
