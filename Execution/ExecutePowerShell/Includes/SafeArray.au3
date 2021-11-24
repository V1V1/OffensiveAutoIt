#include-once

Global Const $tagSAFEARRAYBOUND = _
	"ulong  cElements;"  & _ ; The number of elements in the dimension.
	"long   lLbound;"        ; The lower bound of the dimension.

Global Const $tagSAFEARRAY = _
	"ushort cDims;"      & _ ; The number of dimensions.
	"ushort fFeatures;"  & _ ; Flags, see below.
	"ulong  cbElements;" & _ ; The size of an array element.
	"ulong  cLocks;"     & _ ; The number of times the array has been locked without a corresponding unlock.
	"ptr    pvData;"     & _ ; The data.
	$tagSAFEARRAYBOUND       ; One $tagSAFEARRAYBOUND for each dimension.

; fFeatures flags
Global Const $FADF_AUTO        = 0x0001 ; An array that is allocated on the stack.
Global Const $FADF_STATIC      = 0x0002 ; An array that is statically allocated.
Global Const $FADF_EMBEDDED    = 0x0004 ; An array that is embedded in a structure.
Global Const $FADF_FIXEDSIZE   = 0x0010 ; An array that may not be resized or reallocated.
Global Const $FADF_RECORD      = 0x0020 ; An array that contains records. When set, there will be a pointer to the IRecordInfo interface at negative offset 4 in the array descriptor.
Global Const $FADF_HAVEIID     = 0x0040 ; An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safearray descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
Global Const $FADF_HAVEVARTYPE = 0x0080 ; An array that has a variant type. The variant type can be retrieved with SafeArrayGetVartype.
Global Const $FADF_BSTR        = 0x0100 ; An array of BSTRs.
Global Const $FADF_UNKNOWN     = 0x0200 ; An array of IUnknown*.
Global Const $FADF_DISPATCH    = 0x0400 ; An array of IDispatch*.
Global Const $FADF_VARIANT     = 0x0800 ; An array of VARIANTs.
Global Const $FADF_RESERVED    = 0xF008 ; Bits reserved for future use.



; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

; Safearray functions
; Copied from AutoItObject.au3 by the AutoItObject-Team: monoceres, trancexx, Kip, ProgAndy
; https://www.autoitscript.com/forum/index.php?showtopic=110379

Func SafeArrayCreate($vType, $cDims, $rgsabound)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "ptr", "SafeArrayCreate", "dword", $vType, "uint", $cDims, 'struct*', $rgsabound)
	If @error Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc

Func SafeArrayDestroy($pSafeArray)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "int", "SafeArrayDestroy", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc

Func SafeArrayAccessData($pSafeArray, ByRef $pArrayData)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "int", "SafeArrayAccessData", "ptr", $pSafeArray, 'ptr*', 0)
	If @error Then Return SetError(1, 0, 1)
	$pArrayData = $aCall[2]
	Return $aCall[0]
EndFunc

Func SafeArrayUnaccessData($pSafeArray)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "int", "SafeArrayUnaccessData", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc

Func SafeArrayGetUBound($pSafeArray, $iDim, ByRef $iBound)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "int", "SafeArrayGetUBound", "ptr", $pSafeArray, 'uint', $iDim, 'long*', 0)
	If @error Then Return SetError(1, 0, 1)
	$iBound = $aCall[3]
	Return $aCall[0]
EndFunc

Func SafeArrayGetLBound($pSafeArray, $iDim, ByRef $iBound)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "int", "SafeArrayGetLBound", "ptr", $pSafeArray, 'uint', $iDim, 'long*', 0)
	If @error Then Return SetError(1, 0, 1)
	$iBound = $aCall[3]
	Return $aCall[0]
EndFunc

Func SafeArrayGetDim($pSafeArray)
	Local $aResult = DllCall("OleAut32.dll", "uint", "SafeArrayGetDim", "ptr", $pSafeArray)
	If @error Then Return SetError(1, 0, 0)
	Return $aResult[0]
EndFunc

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Func SafeArrayCopy( $pSafeArrayIn, ByRef $pSafeArrayOut )
	Local $aRet = DllCall( "OleAut32.dll", "int", "SafeArrayCopy", "ptr", $pSafeArrayIn, "ptr*", 0 )
	If @error Then Return SetError(1,0,1)
	$pSafeArrayOut = $aRet[2]
	Return $aRet[0]
EndFunc

Func SafeArrayDestroyData( $pSafeArray )
	Local $aRet = DllCall( "OleAut32.dll", "int", "SafeArrayDestroyData", "ptr", $pSafeArray )
	If @error Then Return SetError(1,0,1)
	Return $aRet[0]
EndFunc

Func SafeArrayGetVartype( $pSafeArray, ByRef $vt )
	Local $aRet = DllCall( "OleAut32.dll", "int", "SafeArrayGetVartype", "ptr", $pSafeArray, "ptr*", 0 )
	If @error Then Return SetError(1,0,1)
	$vt = $aRet[2]
	Return $aRet[0]
EndFunc

; Modified Junkew
;~ HRESULT SafeArrayGetElement(  _In_  SAFEARRAY *psa,  _In_  LONG      *rgIndices,  _Out_ void      *pv);
Func SafeArrayGetElement($pSA, $rgIndices, $pv)
    Local $aCall = DllCall("OleAut32.dll", "long", "SafeArrayGetElement", "ptr", $pSA, "long*", $rgIndices, 'uint_ptr*', $pv)
    If @error Then Return SetError(1, 0, 0)
    Return $aCall[3]
EndFunc