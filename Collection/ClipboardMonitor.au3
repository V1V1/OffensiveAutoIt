#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Periodically monitors the clipboard for text and prints the content to stdout.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== Clipboard monitor ===========" & @CRLF & @CRLF)

;~ Commandline arguments check
_CheckArguments()

; Monitor clipboard
_GetClipboardData()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _GetClipboardData()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide an interval to check the clipoard for content (in seconds)" & @CRLF & @CRLF)
        ConsoleWrite("[i] ClipboardMonitor.exe 5" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] ClipboardMonitor.exe 5" & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckArguments

Func _GetClipboardData()

    Global $monitorArg = Number($CmdLine[1])

    ; Ensure interval is a number
    ConsoleWrite("[i] Interval: " & $monitorArg & @CRLF & @CRLF)

    If $monitorArg = 0 Then
        ConsoleWrite("[X] Interval must be a numeric value. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Start monitor
    ConsoleWrite("[*] Checking clipboard for content every " & $monitorArg & " seconds" & @CRLF & @CRLF)

    ; Monitor loop
    $hTimer = TimerInit()
    While 1
        If TimerDiff($hTimer) > ($monitorArg * 1000) Then
            ConsoleWrite(@CRLF & "----- Clipboard content (" & _NowTime() & " | " & _NowDate() & ") -----" & @CRLF & @CRLF)
            $clipboardData = ClipGet()

            If @error = 0 Then
                ConsoleWrite($clipboardData & @CRLF & @CRLF)
            ElseIf @error = 1 Then
                ConsoleWrite("[i] Clipboard is empty." & @CRLF & @CRLF)
            ElseIf @error = 2 Then
                ConsoleWrite("[i] Clipboard contains a non-text entry." & @CRLF & @CRLF)
            ElseIf @error >= 3 Then
                ConsoleWrite("[X] Cannot access the clipboard." & @CRLF & @CRLF)
            EndIf
            $hTimer = TimerInit()

        EndIf
    WEnd

EndFunc ;==>_GetClipboardData
