#NoTrayIcon
#Region
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion

#include <Process.au3>
#include <Memory.au3>
#include <WinAPI.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Dumps process memory to file on disk using the MiniDumpWriteDump API.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== Process MiniDump ===========" & @CRLF & @CRLF)

;~ Commandline arguments check
_CheckArguments()

; Check if process is running
_CheckProcess()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _CheckProcess()
    _ProcessMiniDump()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a process ID" & @CRLF & @CRLF)
        ConsoleWrite("[i] MiniDump.exe pid" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] MiniDump.exe pid" & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckArguments

Func _CheckProcess()

    ; Process check
    ConsoleWrite("----- Process Check -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking for target process" & @CRLF & @CRLF)

    Global $targetPID = ProcessExists($CmdLine[1])

    If Not $targetPID = 0 Then
        Global $targetProcName = _ProcessGetName($targetPID)
        ConsoleWrite("[i] Target process is running (" & $targetProcName & ")" & @CRLF & @CRLF)

        ; Dump target process memory
        _ProcessMiniDump()

    ElseIf $targetPID = 0 Then
        ConsoleWrite("[X] Target process is not running. Exiting." & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckProcess


Func _ProcessMiniDump()

    ; Dump memory
    ConsoleWrite("----- Process memory dump -----" & @CRLF & @CRLF)

    ; Attempt to access process
    ConsoleWrite("[*] Attempting to open handle to target process." & @CRLF & @CRLF)

    ; OpenProcess
    $hProcess = _WinAPI_OpenProcess( _
        $PROCESS_ALL_ACCESS, _
        0, _
        $targetPID, _
        True)

    If $hProcess = 0 Then
        ConsoleWrite("[X] Failed to open handle to process. Exiting" & @CRLF & @CRLF)
        Exit
    ElseIf Not $hProcess = 0 Then
        ConsoleWrite("[i] Process handle acquired - proceeding with minidump." & @CRLF & @CRLF)
    EndIf

    ; Process minidump
    ConsoleWrite("[*] Dumping process memory." & @CRLF & @CRLF)

    ; Output file name
    Local $dumpOutputFile = _GenerateOutputFile()
    ConsoleWrite("[i] Process dump will be written to: " & $dumpOutputFile & @CRLF & @CRLF)

    ; CreateFile
    $hFile = _WinAPI_CreateFile( _
        $dumpOutputFile, _
        1)

    ; MiniDumpWriteDump
    ; Adapted from - https://www.autoitscript.com/forum/topic/184516-process-dumping-doesnt-work/
    ; Credits - Terenz (https://www.autoitscript.com/forum/profile/80076-terenz/)
    $minidumpCall = DllCall("dbghelp.dll", "bool", "MiniDumpWriteDump", _
        "handle", $hProcess, _
        "dword", $targetPID, _
        "handle", $hFile, _
        "dword", 0x00000002, _ ;MiniDumpWithFullMemory
        "dword", 0, _
        "dword", 0, _
        "dword", 0)

    $dumpProcess = $minidumpCall[0]

    ConsoleWrite("[*] Done" & @CRLF & @CRLF)

EndFunc

#cs ----------------------------------------------------------------------------

Util functions:
    _GenerateOutputFile()
    _RandomString()

#ce ----------------------------------------------------------------------------

; Generate output file name
Func _GenerateOutputFile()
    Local $fileExtension = _RandomString()
    Local $outputFile = "C:\Windows\Temp\" & $targetPID & "-" & $targetProcName & "." & $fileExtension
    Return $outputFile
EndFunc ;==>_GenerateOutputFile

; Adapted from - https://www.autoitscript.com/forum/topic/41556-random-string-generator/
; Credits - ValdeZ (https://www.autoitscript.com/forum/profile/21781-valdez/)
Func _RandomString()
    $iLength = 3
    $aChars = StringSplit("abcdefghijklmnopqrstuvwxyz", "")
    $sString = ""
    While $iLength > StringLen($sString)
        If $sString < "A" Then $sString = ""
        $sString &= $aChars[Random(1, $aChars[0],1)]
    WEnd
    Return $sString
EndFunc   ;==>_RandomString
