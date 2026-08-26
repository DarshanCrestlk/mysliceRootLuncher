!macro customUnInstall
  ; Electron uninstaller (same cleanup as MySlice Uninstall.exe)
  IfFileExists "$INSTDIR\MySlice.exe" 0 skip_myslice_uninstall
    ExecWait '"$INSTDIR\MySlice.exe" --uninstall --silent'
  skip_myslice_uninstall:
!macroend
