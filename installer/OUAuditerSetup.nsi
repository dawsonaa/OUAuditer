Name "OU Auditer"
OutFile "OUAuditerInstaller.exe"
InstallDir "$PROFILE\OU Auditer"
RequestExecutionLevel admin

Var PROFILE_PATH
Var LENGTH

Page directory
Page instfiles

Section "Install OU Auditer"

    SetOutPath $INSTDIR
    File "..\icon.ico"
    File "..\OUAuditer.ps1"
    File "..\LICENSE"
    File "..\README.md"

    StrCpy $PROFILE_PATH $INSTDIR
    StrLen $LENGTH "OU Auditer"
    IntOp $LENGTH $LENGTH + 1 ; Account for the trailing slash/backslash
    StrCpy $PROFILE_PATH $PROFILE_PATH -$LENGTH

    CreateDirectory "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OU Auditer"

    CreateDirectory "$INSTDIR\images"
    setOutPath "$INSTDIR\images"
    File /r "..\images\*.*"

    CreateShortCut "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OU Auditer\OU Auditer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\OUAuditer.ps1"` \
        "$INSTDIR\icon.ico" 0
    CreateShortCut "$PROFILE_PATH\Desktop\OU Auditer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\OUAuditer.ps1"` \
        "$INSTDIR\icon.ico" 0

    ; Unblock all files in the installation directory
    nsExec::Exec 'powershell -Command "Get-ChildItem -Path ''$INSTDIR\*'' -Recurse | Unblock-File"'

SectionEnd