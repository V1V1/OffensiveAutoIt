#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <ScreenCapture.au3>

#cs --------------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         VIVI (https://github.com/V1V1)

 Script Function:
	Takes screenshot and saves it in the user's temp directory.

#ce --------------------------------------------------------------------------------

; Title
ConsoleWrite(@CRLF & "=========== Screen capture ===========" & @CRLF & @CRLF)

; Take screenshot
_ScreenCapture()

#cs ----------------------------------------------------------------------------

Main functions:
    _ScreenCapture()

#ce ----------------------------------------------------------------------------

Func _ScreenCapture()

    ; Output file name
    Local $screencapOutputFile = _GenerateOutputFile()
    ConsoleWrite("[i] Screenshot will be saved to: " & $screencapOutputFile & @CRLF & @CRLF)

    ; Take screenshot
    $hBmp = _ScreenCapture_Capture("")
    _ScreenCapture_SaveImage($screencapOutputFile, $hBmp)

    ConsoleWrite("[*] Done" & @CRLF & @CRLF)

EndFunc ;==>_ScreenCapture

#cs ----------------------------------------------------------------------------

Util functions:
    _GenerateOutputFile()

#ce ----------------------------------------------------------------------------

Func _GenerateOutputFile()
    ; File format (sc-YEAR-MONTH-DAY_HOUR-MINUTE-SECOND.png)
    Local $timestamp = @YEAR & "-" & @MON & "-" & @MDAY & "_" & @HOUR & "-" & @MIN & "-" & @SEC
    Local $outputFile = @TempDir & "\sc-" & $timestamp & ".png"
    Return $outputFile
EndFunc ;==>_GenerateOutputFile
