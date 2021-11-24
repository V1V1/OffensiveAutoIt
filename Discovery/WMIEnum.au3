#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Enumerates basic info about a Windows host using WMI.

 Reference:
    https://github.com/kyleavery/WMIEnum
    Credits - kyleavery (https://github.com/kyleavery)

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== WMIEnum ===========" & @CRLF & @CRLF)

; WMI queries adapted from - https://www.autoitscript.com/forum/topic/195947-get-capability-of-empty-cddvd-drive/?tab=comments#comment-1405046
; Credits - FrancescoDiMuro (https://www.autoitscript.com/forum/profile/99495-francescodimuro/)

; System info
ConsoleWrite("------ System info ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_ComputerSystem"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")
    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("Hostname  :  " & $objWMI_Item.Name & @CRLF)
        ConsoleWrite("Domain    :  " & $objWMI_Item.Domain & @CRLF)
    Next
    $strWMI_Query = "SELECT * FROM Win32_OperatingSystem"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")
    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("OS        :  " & $objWMI_Item.Version & @CRLF & @CRLF)
    Next
EndIf

; AV
ConsoleWrite(@CRLF & "------ AntiVirus ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\securitycenter2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT displayName FROM AntiVirusProduct"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("AntiVirusProduct  :  " & $objWMI_Item.DisplayName & @CRLF & @CRLF)
    Next

EndIf

; User info
ConsoleWrite(@CRLF & "------ User info ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_UserAccount"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("Name    :  " & $objWMI_Item.Name & @CRLF)
        ConsoleWrite("Domain  :  " & $objWMI_Item.Domain & @CRLF)
        ConsoleWrite("SID     :  " & $objWMI_Item.SID & @CRLF & @CRLF)
    Next

EndIf

; Group info
ConsoleWrite(@CRLF & "------ Group info ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_Group"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("Name    :  " & $objWMI_Item.Name & @CRLF)
        ConsoleWrite("Domain  :  " & $objWMI_Item.Domain & @CRLF)
        ConsoleWrite("SID     :  " & $objWMI_Item.SID & @CRLF & @CRLF)
    Next

EndIf

; Drive info
ConsoleWrite(@CRLF & "------ Drive info ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_LogicalDisk"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")
    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite("DeviceId   :  " & $objWMI_Item.DeviceId & @CRLF)
        ConsoleWrite("DriveType  :  " & $objWMI_Item.DriveType & @CRLF)
        ConsoleWrite("Size       :  " & $objWMI_Item.Size & @CRLF & @CRLF)
    Next

EndIf

; Network info
ConsoleWrite(@CRLF & "------ Network info ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_NetworkadApterConfiguration"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    For $objWMI_Item In $objWMI_QueryResult
        Local $ipAdd = $objWMI_Item.IPAddress
        Local $ipGateway = $objWMI_Item.DefaultIPGateway
        If IsArray($ipAdd) And IsArray($ipGateway) Then
            ConsoleWrite("IPAddress      :  " & $ipAdd[0] & @CRLF)
            ConsoleWrite("Gateway        :  " & $ipGateway[0] & @CRLF)
            ConsoleWrite("DHCPEnabled    :  " & $objWMI_Item.DHCPEnabled & @CRLF)
        EndIf
        ConsoleWrite("DNSDomain      :  " & $objWMI_Item.DNSDomain & @CRLF)
        ConsoleWrite("ServiceName    :  " & $objWMI_Item.ServiceName & @CRLF)
        ConsoleWrite("Description    :  " & $objWMI_Item.Description & @CRLF & @CRLF)
    Next

EndIf

; Processes
ConsoleWrite(@CRLF & "------ Process list ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * from win32_process"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    ConsoleWrite("PID , Name" & @CRLF)
    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite($objWMI_Item.Handle & " , " & $objWMI_Item.Name & @CRLF)
    Next

EndIf

; Services
ConsoleWrite(@CRLF & "------ Services ------" & @CRLF & @CRLF)
$objWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If IsObj($objWMI) Then
    $strWMI_Query = "SELECT * FROM Win32_Service"
    $objWMI_QueryResult = $objWMI.ExecQuery($strWMI_Query, "WQL")

    ConsoleWrite("Name , State , Mode, Path" & @CRLF)
    For $objWMI_Item In $objWMI_QueryResult
        ConsoleWrite($objWMI_Item.Name & " , " & $objWMI_Item.State & " , " & $objWMI_Item.StartMode & " , " & $objWMI_Item.PathName & @CRLF)
    Next

EndIf
