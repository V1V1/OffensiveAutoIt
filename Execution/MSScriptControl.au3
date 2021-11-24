#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Executes JScript and VBScript using the MSScriptControl COM object.

 Notes:
	This is a direct port of scriptcontrol_bin.nim in byt3bl33d3r's OffensiveNim
    repo (https://github.com/byt3bl33d3r/OffensiveNim/blob/master/src/scriptcontrol_bin.nim)

#ce --------------------------------------------------------------------------------

; NOTE: MSScriptControl component is 32-bit only.
If @AutoItX64 Then
    ConsoleWriteError('[X] MSScriptControl only supports 32 bit.' & @CRLF)
    Exit
EndIf

; JScript
$objJs = ObjCreate("MSScriptControl.ScriptControl")
$objJs.Language = "JScript"
$exp = "Math.pow(5, 2) * Math.PI"
$answer = $objJs.eval($exp)
$msg = "" & $exp & " = " & $answer & ""

; VBScript
$objVbs = ObjCreate("MSScriptControl.ScriptControl")
$objVbs.Language = "VBScript"
$objVbs.AllowUI = True
$title = "Windows COM for AutoIt"
$vbsCode = 'MsgBox("This is a VBScript message box." & vbCRLF & vbCRLF & "This is JScript code:" & vbCRLF & "' & $msg & '", vbOKOnly, "' & $title & '")'
$objVbs.Eval($vbsCode)
