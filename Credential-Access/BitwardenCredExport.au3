#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <GuiToolBar.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Exports Bitwarden vault credentials to disk using UI automation & key presses.
    Requires user's master password.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== Bitwarden credential export ===========" & @CRLF & @CRLF)

; Bitwarden window checks
_CheckbitwardenWindow()

; If all checks pass, attempt to export creds to json file on disk
_ExportBitwardenCreds()

Func _CheckBitwardenWindow()

    ; Process check
    ConsoleWrite("----- Process Check -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking for Bitwarden process" & @CRLF & @CRLF)
    $bitwardenPID = ProcessExists("Bitwarden.exe")

    If Not $bitwardenPID = 0 Then
        ConsoleWrite("  [i] Bitwarden.exe is running (PID: " & $bitwardenPID & ")" & @CRLF)
    ElseIf $bitwardenPID = 0 Then
        ConsoleWrite("  [X] Bitwarden.exe is not running. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get bitwarden window handle
    Global $bitwardenHandle = _GetHwndFromPID($bitwardenPID)

    ; Window active check
    If WinActive($bitwardenHandle) Then
        ConsoleWrite("  [X] Bitwarden window is currently in the foreground. Try again later?" & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get bitwarden Window title
    $bitwardenWindowTitle = WinGetTitle($bitwardenHandle)
    ConsoleWrite("  [i] Bitwarden window title: " & $bitwardenWindowTitle & @CRLF & @CRLF)

EndFunc ;==>_CheckBitwardenWindow

Func _ExportBitwardenCreds()

    ; Cred export
    ConsoleWrite("----- Credential Export -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Attempting Bitwarden credential export" & @CRLF & @CRLF)

    ; Click on Bitwarden tray icon
    _ClickonBitwardenTrayIcon()
    ; Activate Bitwarden window
    WinActivate($bitwardenHandle)

    ; Unlock Bitwarden with user's master password
    ConsoleWrite("  [*] Unlocking Bitwarden" & @CRLF & @CRLF)
    ; Master password
    $masterPass = String("[ENTER MASTER PASSWORD HERE]")
    Send($masterPass)
    Send("{ENTER}")
    ; Give script some time to input password & open main screen
    Sleep(1500)

    ; Open export menu using keyboard shortcuts
    ConsoleWrite("  [*] Opening export menu" & @CRLF & @CRLF)
    Send("{ALT}")
    Send("{ENTER}")
    Send("E")
    ; Enter master pass to enable export
    Send($masterPass)
    Send("{ENTER}")
    Sleep(1000)
    ; Confirm export
    Send("{ENTER}")
    Sleep(1000)

    ; Send output file location (current user's temp directory)
    ConsoleWrite("  [*] Exporting credentials to json file" & @CRLF & @CRLF)
    $outputFile = @TempDir & "\bitwarden-export.json"
    Send($outputFile)
    Send("{ENTER}")
    Sleep(1000)

    ; Lock and minimize Bitwarden window
    ConsoleWrite("  [*] Locking vault and minimizing window" & @CRLF & @CRLF)
    WinActivate($bitwardenHandle)
    Send("^l")
    WinSetState($bitwardenHandle, "", @SW_MINIMIZE)

    ; Print output file location
    ConsoleWrite("  [i] Bitwarden export file written to: " & $outputFile & @CRLF & @CRLF)

    ; Read export file
    ConsoleWrite("[*] Reading output file" & @CRLF & @CRLF)
    Sleep(2000)
    If FileExists($outputFile) Then
        Local $sfileRead = FileRead($outputFile)
        ConsoleWrite("[+] Output file contents:" & @CRLF & @CRLF & $sfileRead)
    EndIf

    ; Delete export file
    ConsoleWrite(@CRLF & @CRLF & "[*] Deleting output file" & @CRLF & @CRLF)
    FileDelete($outputFile)
    ConsoleWrite("  [i] Export file at '" & $outputFile & "' has been deleted" & @CRLF & @CRLF)

    ConsoleWrite("[*] Done" & @CRLF & @CRLF)

EndFunc ;==>_ExportBitwardenCreds()

#cs ----------------------------------------------------------------------------

Util functions:
    _GetHwndFromPID()
    _ClickonBitwardenTrayIcon()

#ce ----------------------------------------------------------------------------

; Gets window handle from PID
; Adapted from - https://www.autoitscript.com/wiki/FAQ#How_can_I_get_a_window_handle_when_all_I_have_is_a_PID.3F
Func _GetHwndFromPID($PID)

    $hWnd = 0
    $stPID = DllStructCreate("int")
    Do
        $winlist2 = WinList()
        For $i = 1 To $winlist2[0][0]
            If $winlist2[$i][0] <> "" Then
                DllCall("user32.dll", "int", "GetWindowThreadProcessId", "hwnd", $winlist2[$i][1], "ptr", DllStructGetPtr($stPID))
                If DllStructGetData($stPID, 1) = $PID Then
                    $hWnd = $winlist2[$i][1]
                    ExitLoop
                EndIf
            EndIf
        Next
        Sleep(100)
    Until $hWnd <> 0
    Return $hWnd

EndFunc ;==>_GetHwndFromPID

; Search for Bitwarden icon in toolbar and click on it if found
; Adapted from - https://www.autoitscript.com/forum/topic/198426-get-text-from-specific-toolbar-tray-icon/?do=findComment&comment=1423626
; Credits - Subz (https://www.autoitscript.com/forum/profile/101464-subz/)
Func _ClickonBitwardenTrayIcon()

    Global $g_aTaskBar[0]
    Global $g_iInstance = 1
    Global $g_hWnd, $g_iCount, $g_iCmdID, $g_iSearch, $g_sSearch = "Bitwarden"
    While 1
        $g_hWnd = ControlGetHandle("[CLASS:Shell_TrayWnd]", "", "[CLASS:ToolbarWindow32;INSTANCE:" & $g_iInstance& "]")
        ;~ Get number of buttons
        $g_iCount = _GUICtrlToolbar_ButtonCount($g_hWnd)
            If $g_iCount > 1 Then ExitLoop
        $g_iInstance += 1
    WEnd
    ;~ Loop through buttons
    For $i = 1 To $g_iCount
        $g_iCmdID = _GUICtrlToolbar_IndexToCommand($g_hWnd, $i)
        ;~ Check during the loop
        If StringInStr(_GUICtrlToolbar_GetButtonText($g_hWnd, $g_iCmdID), $g_sSearch) Then _GUICtrlToolbar_ClickButton($g_hWnd,$g_iCmdID, "left", False, 1)
        ;~ Add the item to an array
        _ArrayAdd($g_aTaskBar, _GUICtrlToolbar_GetButtonText($g_hWnd, $g_iCmdID))
    Next
EndFunc ;==>_ClickonBitwardenTrayIcon
