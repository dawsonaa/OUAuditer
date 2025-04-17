; Define basic information about the installer
Name "OU Auditer"
OutFile "OUAuditerInstaller.exe"
InstallDir "$PROFILE\OU Auditer"  ; Set installation directory to current user's home
RequestExecutionLevel admin  ; Request admin rights

; Declare variables
Var PROFILE_PATH
Var LENGTH

; Pages to display
Page directory  ; Let the user choose the install directory
Page instfiles  ; Show the installation progress

; Default section for installation
Section "Install OU Auditer"

    ; Set the output path to the installation directory
    SetOutPath $INSTDIR
    File "..\icon.ico"  ; Adjusted path to look in root
    File "..\OU Auditer.ps1"  ; Adjusted path to look in root
    File "..\LICENSE"  ; Adjusted path to look in root
    File "..\README.md"  ; Adjusted path to look in root

    ; Extract the profile-specific base path from $INSTDIR
    StrCpy $PROFILE_PATH $INSTDIR
    ; Find the length of "OU Auditer" (adjust for your actual folder name length)
    StrLen $LENGTH "OU Auditer"
    IntOp $LENGTH $LENGTH + 1 ; Account for the trailing slash/backslash
    ; Trim the application folder from $INSTDIR
    StrCpy $PROFILE_PATH $PROFILE_PATH -$LENGTH

    ; Create directories in the user's profile-specific Start Menu programs folder
    CreateDirectory "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OU Auditer"

    CreateDirectory "$INSTDIR\images"
    setOutPath "$INSTDIR\images"
    File /r "..\images\*.*"

    ; Create Start Menu and Desktop shortcuts
    CreateShortCut "$PROFILE_PATH\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\OU Auditer\OU Auditer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\OU Auditer.ps1"` \
        "$INSTDIR\icon.ico" 0
    CreateShortCut "$PROFILE_PATH\Desktop\OU Auditer.lnk" \
        "$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" `-ExecutionPolicy Bypass -File "$INSTDIR\OU Auditer.ps1"` \
        "$INSTDIR\icon.ico" 0

    ; Optional: Run a PowerShell command to unblock all files in the installation directory
    nsExec::Exec 'powershell -Command "Get-ChildItem -Path ''$INSTDIR\*'' -Recurse | Unblock-File"'

SectionEnd