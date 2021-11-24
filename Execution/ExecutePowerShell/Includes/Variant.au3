#include-once

; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

; Copied from AutoItObject.au3 by the AutoItObject-Team: monoceres, trancexx, Kip, ProgAndy
; https://www.autoitscript.com/forum/index.php?showtopic=110379

Global Const $tagVARIANT = "word vt;word r1;word r2;word r3;ptr data; ptr"
; The structure takes up 16/24 bytes when running 32/64 bit
; Space for the data element at the end represents 2 pointers
; This is 8 bytes running 32 bit and 16 bytes running 64 bit

Global Const $VT_EMPTY            = 0  ; 0x0000
Global Const $VT_NULL             = 1  ; 0x0001
Global Const $VT_I2               = 2  ; 0x0002
Global Const $VT_I4               = 3  ; 0x0003
Global Const $VT_R4               = 4  ; 0x0004
Global Const $VT_R8               = 5  ; 0x0005
Global Const $VT_CY               = 6  ; 0x0006
Global Const $VT_DATE             = 7  ; 0x0007
Global Const $VT_BSTR             = 8  ; 0x0008
Global Const $VT_DISPATCH         = 9  ; 0x0009
Global Const $VT_ERROR            = 10 ; 0x000A
Global Const $VT_BOOL             = 11 ; 0x000B
Global Const $VT_VARIANT          = 12 ; 0x000C
Global Const $VT_UNKNOWN          = 13 ; 0x000D
Global Const $VT_DECIMAL          = 14 ; 0x000E
Global Const $VT_I1               = 16 ; 0x0010
Global Const $VT_UI1              = 17 ; 0x0011
Global Const $VT_UI2              = 18 ; 0x0012
Global Const $VT_UI4              = 19 ; 0x0013
Global Const $VT_I8               = 20 ; 0x0014
Global Const $VT_UI8              = 21 ; 0x0015
Global Const $VT_INT              = 22 ; 0x0016
Global Const $VT_UINT             = 23 ; 0x0017
Global Const $VT_VOID             = 24 ; 0x0018
Global Const $VT_HRESULT          = 25 ; 0x0019
Global Const $VT_PTR              = 26 ; 0x001A
Global Const $VT_SAFEARRAY        = 27 ; 0x001B
Global Const $VT_CARRAY           = 28 ; 0x001C
Global Const $VT_USERDEFINED      = 29 ; 0x001D
Global Const $VT_LPSTR            = 30 ; 0x001E
Global Const $VT_LPWSTR           = 31 ; 0x001F
Global Const $VT_RECORD           = 36 ; 0x0024
Global Const $VT_INT_PTR          = 37 ; 0x0025
Global Const $VT_UINT_PTR         = 38 ; 0x0026
Global Const $VT_FILETIME         = 64 ; 0x0040
Global Const $VT_BLOB             = 65 ; 0x0041
Global Const $VT_STREAM           = 66 ; 0x0042
Global Const $VT_STORAGE          = 67 ; 0x0043
Global Const $VT_STREAMED_OBJECT  = 68 ; 0x0044
Global Const $VT_STORED_OBJECT    = 69 ; 0x0045
Global Const $VT_BLOB_OBJECT      = 70 ; 0x0046
Global Const $VT_CF               = 71 ; 0x0047
Global Const $VT_CLSID            = 72 ; 0x0048
Global Const $VT_VERSIONED_STREAM = 73 ; 0x0049
Global Const $VT_BSTR_BLOB        = 0xFFF
Global Const $VT_VECTOR           = 0x1000
Global Const $VT_ARRAY            = 0x2000
Global Const $VT_BYREF            = 0x4000
Global Const $VT_RESERVED         = 0x8000
Global Const $VT_ILLEGAL          = 0xFFFF
Global Const $VT_ILLEGALMASKED    = 0xFFF
Global Const $VT_TYPEMASK         = 0xFFF

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


;Global Const $tagVARIANT = "word vt;word r1;word r2;word r3;ptr data; ptr"
; The structure takes up 16/24 bytes when running 32/64 bit
; Space for the data element at the end represents 2 pointers
; This is 8 bytes running 32 bit and 16 bytes running 64 bit

#cs
DECIMAL structure
https://msdn.microsoft.com/en-us/library/windows/desktop/ms221061(v=vs.85).aspx

From oledb.h:
typedef struct tagDEC {
    USHORT wReserved;			; vt,     2 bytes
    union {								; r1,     2 bytes
        struct {
            BYTE scale;
            BYTE sign;
        };
        USHORT signscale;
    };
    ULONG Hi32;						; r2, r3, 4 bytes
    union {								; data,   8 bytes
        struct {
#ifdef _MAC
            ULONG Mid32;
            ULONG Lo32;
#else
            ULONG Lo32;
            ULONG Mid32;
#endif
        };
        ULONGLONG Lo64;
    };
} DECIMAL;
#ce

Global Const $tagDEC = "word wReserved;byte scale;byte sign;uint Hi32;uint Lo32;uint Mid32"


; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

; Variant functions
; Copied from AutoItObject.au3 by the AutoItObject-Team: monoceres, trancexx, Kip, ProgAndy
; https://www.autoitscript.com/forum/index.php?showtopic=110379

; #FUNCTION# ====================================================================================================================
; Name...........: VariantClear
; Description ...: Clears the value of a variant
; Syntax.........: VariantClear($pvarg)
; Parameters ....: $pvarg       - the VARIANT to clear
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: VariantFree
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221165.aspx
; Example .......:
; ===============================================================================================================================
Func VariantClear($pvarg)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "long", "VariantClear", "ptr", $pvarg)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: VariantCopy
; Description ...: Copies a VARIANT to another
; Syntax.........: VariantCopy($pvargDest, $pvargSrc)
; Parameters ....: $pvargDest   - Destionation variant
;                  $pvargSrc    - Source variant
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: VariantRead
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221697.aspx
; Example .......:
; ===============================================================================================================================
Func VariantCopy($pvargDest, $pvargSrc)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "long", "VariantCopy", "ptr", $pvargDest, 'ptr', $pvargSrc)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: VariantInit
; Description ...: Initializes a variant.
; Syntax.........: VariantInit($pvarg)
; Parameters ....: $pvarg       - the VARIANT to initialize
; Return values .: Success      - 0
;                  Failure      - nonzero
; Author ........: Prog@ndy
; Modified.......:
; Remarks .......:
; Related .......: VariantClear
; Link ..........: http://msdn.microsoft.com/en-us/library/ms221402.aspx
; Example .......:
; ===============================================================================================================================
Func VariantInit($pvarg)
	; Author: Prog@ndy
	Local $aCall = DllCall("OleAut32.dll", "long", "VariantInit", "ptr", $pvarg)
	If @error Then Return SetError(1, 0, 1)
	Return $aCall[0]
EndFunc

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Func VariantChangeType( $pVarDest, $pVarSrc, $wFlags, $vt )
	Local $aRet = DllCall( "OleAut32.dll", "long", "VariantChangeType", "ptr", $pVarDest, "ptr", $pVarSrc, "word", $wFlags, "word", $vt )
	If @error Then Return SetError(1,0,1)
	Return $aRet[0]
EndFunc

Func VariantChangeTypeEx( $pVarDest, $pVarSrc, $lcid, $wFlags, $vt )
	Local $aRet = DllCall( "OleAut32.dll", "long", "VariantChangeTypeEx", "ptr", $pVarDest, "ptr", $pVarSrc, "word", $lcid, "word", $wFlags, "word", $vt )
	If @error Then Return SetError(1,0,1)
	Return $aRet[0]
EndFunc

Func VarAdd( $pVarLeft, $pVarRight, $pVarResult )
	Local $aRet = DllCall( "OleAut32.dll", "long", "VarAdd", "ptr", $pVarLeft, "ptr", $pVarRight, "ptr", $pVarResult )
	If @error Then Return SetError(1,0,1)
	Return $aRet[0]
EndFunc

Func VarSub( $pVarLeft, $pVarRight, $pVarResult )
	Local $aRet = DllCall( "OleAut32.dll", "long", "VarSub", "ptr", $pVarLeft, "ptr", $pVarRight, "ptr", $pVarResult )
	If @error Then Return SetError(1,0,1)
	Return $aRet[0]
EndFunc


; >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

; BSTR (basic string) functions
; Copied from AutoItObject.au3 by the AutoItObject-Team: monoceres, trancexx, Kip, ProgAndy
; https://www.autoitscript.com/forum/index.php?showtopic=110379

Func SysAllocString( $str )
	Local $aRet = DllCall( "OleAut32.dll", "ptr", "SysAllocString", "wstr", $str )
	If @error Then Return SetError(1, 0, 0)
	Return $aRet[0]
EndFunc

Func SysFreeString( $pBSTR )
	If Not $pBSTR Then Return SetError(1, 0, 0)
	DllCall( "OleAut32.dll", "none", "SysFreeString", "ptr", $pBSTR )
	If @error Then Return SetError(2, 0, 0)
EndFunc

Func SysReadString( $pBSTR, $iLen = -1 )
	If Not $pBSTR Then Return SetError(1, 0, "")
	If $iLen < 1 Then $iLen = SysStringLen( $pBSTR )
	If $iLen < 1 Then Return SetError(2, 0, "")
	Return DllStructGetData( DllStructCreate( "wchar[" & $iLen & "]", $pBSTR ), 1 )
EndFunc

Func SysStringLen( $pBSTR )
	If Not $pBSTR Then Return SetError(1, 0, 0)
	Local $aRet = DllCall( "OleAut32.dll", "uint", "SysStringLen", "ptr", $pBSTR )
	If @error Then Return SetError(2, 0, 0)
	Return $aRet[0]
EndFunc

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
