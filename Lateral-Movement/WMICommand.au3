#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1) | (https://twitter.com/_theVIVI)

 Script Function:
	Executes a command on a remote computer using WMI.

#ce --------------------------------------------------------------------------------

ConsoleWrite(@CRLF & "=========== WMICommand ===========" & @CRLF & @CRLF)

;~ Commandline arguments check
_CheckArguments()

; Execute remote command via WMI
_ExecuteCommand()

#cs ----------------------------------------------------------------------------

Main functions:
    _CheckArguments()
    _ExecuteCommand()

#ce ----------------------------------------------------------------------------

Func _CheckArguments()

    If $CmdLine[0] <= 0 Then
        ConsoleWrite("[i] WMICommand.exe TARGET USER PASSWORD COMMAND" & @CRLF & @CRLF)
        ConsoleWrite("Example: WMICommand.exe 192.168.60.101 DOMAIN\user Password123 C:\Windows\System32\calc.exe" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] <= 3 Then
        ConsoleWrite("[X] Too few arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] WMICommand.exe TARGET USER PASSWORD COMMAND" & @CRLF & @CRLF)
        Exit

    ElseIf $CmdLine[0] > 4 Then
        ConsoleWrite("[X] Too many arguments provided." & @CRLF & @CRLF)
        ConsoleWrite("[i] WMICommand.exe TARGET USER PASSWORD COMMAND" & @CRLF & @CRLF)
        Exit

    EndIf

EndFunc ;==>_CheckArguments

Func _ExecuteCommand()

    ; Command details
    Local $targetHost = String($CmdLine[1])
    Local $userName = String($CmdLine[2])
    Local $userPass = String($CmdLine[3])
    Local $command = String($CmdLine[4])

    ConsoleWrite("[*] Attempting remote WMI command with these details:" & @CRLF & @CRLF)

    ConsoleWrite("[+] Target host :  " & $targetHost & @CRLF)
    ConsoleWrite("[+] Username    :  " & $userName & @CRLF)
    ConsoleWrite("[+] Password    :  " & $userPass & @CRLF)
    ConsoleWrite("[+] Command     :  " & $command & @CRLF & @CRLF)

    ConsoleWrite("[*] Executing WMI command..." & @CRLF & @CRLF)

    ; Adapted from - https://www.autoitscript.com/forum/topic/31929-multiple-domains-wmi-and-alternate-credentials/
    ; Credits - Bob Hoss (https://www.autoitscript.com/forum/profile/15292-bob-hoss/)

    ; Create The WbemScripting.SWbemLocator Object
    $objSWbemLocator = ObjCreate("WbemScripting.SWbemLocator")
    ; Error Handling For Com Object
    $objError = ObjEvent("AutoIt.Error","WMI_Connect_Error")

    ; Attempt WMI connection with supplied creds
    $objWMIService = $objSWbemLocator.ConnectServer($targetHost, "root\cimv2", $userName, $userPass)

    ; If connection fails - exit
    If Not IsObj($objWMIService) Then
        ConsoleWrite("[X] Couldn't connect to WMI." & @CRLF & @CRLF)
        Exit

    EndIf

    ; Connected to WMI O.K.
    If IsObj($objWMIService) Then

        ; Adapted from - https://www.autoitscript.com/forum/topic/65870-remote-execute/?do=findComment&comment=489418
        ; Credits - ptrex (https://www.autoitscript.com/forum/profile/6305-ptrex/)

        ; Obtain the Win32_Process class of object.
        $objProcess = $objWMIService.Get("Win32_Process")
        $objProgram = $objProcess.Methods_("Create").InParameters.SpawnInstance_()
        $objProgram.CommandLine = $command

        ; Execute the program now at the command line.
        $strShell = $objWMIService.ExecMethod("Win32_Process", "Create", $objProgram)
        ConsoleWrite ("[+] Created: '" & $command & "' on '" & $targetHost & "'" & @CRLF & @CRLF)

        ConsoleWrite("[*] Done." & @CRLF & @CRLF)

    EndIf

EndFunc ;==>_ExecuteCommand

; Error Handling Function For WbemScripting.SWbemLocator Object - Forces Return To Code Instead Of Bombing Out
Func WMI_Connect_Error()
     Return("Caught Error")
EndFunc
