; *************************************************************************
;
; .NET Common Language Runtime (CLR) UDF v3.0 - into the AutoIt process
;
; Developers : Danyfirex / Larsj / Trancexx / Junkew
;
; *************************************************************************

#include-once
#include "SafeArray.au3"
#include "Variant.au3"

; *************************************************************************
; Declare Binding Flags : InvokeMember, InvokeMember_2, InvokeMember_3
; *************************************************************************
Global Const $BindingFlags_Default = 0x0000
Global Const $BindingFlags_IgnoreCase = 0x0001
Global Const $BindingFlags_DeclaredOnly = 0x0002
Global Const $BindingFlags_Instance = 0x0004
Global Const $BindingFlags_Static = 0x0008
Global Const $BindingFlags_Public = 0x0010
Global Const $BindingFlags_NonPublic = 0x0020
Global Const $BindingFlags_FlattenHierarchy = 0x0040
Global Const $BindingFlags_InvokeMethod = 0x0100
Global Const $BindingFlags_CreateInstance = 0x0200
Global Const $BindingFlags_GetField = 0x0400
Global Const $BindingFlags_SetField = 0x0800
Global Const $BindingFlags_GetProperty = 0x1000
Global Const $BindingFlags_SetProperty = 0x2000
Global Const $BindingFlags_PutDispProperty = 0x4000
Global Const $BindingFlags_PutRefDispProperty = 0x8000
Global Const $BindingFlags_ExactBinding = 0x00010000
Global Const $BindingFlags_SuppressChangeType = 0x00020000
Global Const $BindingFlags_OptionalParamBinding = 0x00040000
Global Const $BindingFlags_IgnoreReturn = 0x01000000
; 0x158 = $BindingFlags_Static + $BindingFlags_Public + $BindingFlags_FlattenHierarchy + $BindingFlags_InvokeMethod

; *************************************************************************
; Declare Interfaces Variables for the CLSID's / IID's / Tags
; *************************************************************************
Global Const $sCLSID_CLRMetaHost = "{9280188d-0e8e-4867-b30c-7fa83884e8de}"
Global Const $tCLSID_CLRMetaHost = _WinAPI_CLSIDFromString($sCLSID_CLRMetaHost)
Global Const $sIID_ICLRMetaHost = "{d332db9e-b9b3-4125-8207-a14884f53216}"
Global Const $tIID_ICLRMetaHost = _WinAPI_CLSIDFromString($sIID_ICLRMetaHost)
Global Const $sTag_ICLRMetaHost = _
		"GetRuntime hresult(wstr;struct*;ptr);" & _
		"GetVersionFromFile hresult(ptr;ptr;ptr);" & _
		"EnumerateInstalledRuntimes hresult(ptr);" & _
		"EnumerateLoadedRuntimes hresult(ptr;ptr);" & _
		"RequestRuntimeLoadedNotification hresult(ptr,ptr,ptr):" & _
		"QueryLegacyV2RuntimeBinding hresult(ptr;ptr);" & _
		"ExitProcess hresult(int);"

Global Const $sIID_ICLRRuntimeInfo = "{BD39D1D2-BA2F-486a-89B0-B4B0CB466891}"
Global Const $tIID_ICLRRuntimeInfo = _WinAPI_CLSIDFromString($sIID_ICLRRuntimeInfo)
Global Const $sTag_ICLRRuntimeInfo = _
		"GetVersionString hresult(ptr;ptr);" & _
		"GetRuntimeDirectory hresult(ptr;ptr);" & _
		"IsLoaded hresult(ptr;ptr);" & _
		"LoadErrorString hresult(ptr;ptr;ptr;ptr);" & _
		"LoadLibrary hresult(ptr;ptr);" & _
		"GetProcAddress hresult(ptr;ptr);" & _
		"GetInterface hresult(ptr;ptr;ptr);" & _
		"IsLoadable hresult(Bool*);" & _
		"SetDefaultStartupFlags hresult(ptr;ptr);" & _
		"GetDefaultStartupFlags hresult(ptr;ptr;ptr);" & _
		"BindAsLegacyV2Runtime hresult();" & _
		"IsStarted hresult(ptr;ptr);"

Global Const $sCLSID_CLRRuntimeHost = "{90F1A06E-7712-4762-86B5-7A5EBA6BDB02}"
Global Const $tCLSID_CLRRuntimeHost = _WinAPI_CLSIDFromString($sCLSID_CLRRuntimeHost)
Global Const $sIID_ICLRRuntimeHost = "{90F1A06C-7712-4762-86B5-7A5EBA6BDB02}"
Global Const $tIID_ICLRRuntimeHost = _WinAPI_CLSIDFromString($sIID_ICLRRuntimeHost)
Global Const $sTag_ICLRRuntimeHost = _
		"Start hresult();" & _
		"Stop hresult();" & _
		"SetHostControl hresult(ptr);" & _
		"GetCLRControl hresult(ptr*);" & _
		"UnloadAppDomain hresult(ptr;ptr);" & _
		"ExecuteInAppDomain hresult(ptr;ptr;ptr);" & _
		"GetCurrentAppDomainId hresult(ptr);" & _
		"ExecuteApplication hresult(ptr;ptr;ptr;ptr;ptr;ptr);" & _
		"ExecuteInDefaultAppDomain hresult(wstr;wstr;wstr;wstr;ptr*);"

Global Const $sCLSID_CorRuntimeHost = "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}"
Global Const $tCLSID_CorRuntimeHost = _WinAPI_CLSIDFromString($sCLSID_CorRuntimeHost)
Global Const $sIID_ICorRuntimeHost = "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"
Global Const $tIID_ICorRuntimeHost = _WinAPI_CLSIDFromString($sIID_ICorRuntimeHost)
Global Const $sTag_ICorRuntimeHost = _
		"CreateLogicalThreadState hresult();" & _
		"DeleteLogicalThreadState hresult();" & _
		"SwitchInLogicalThreadState hresult();" & _
		"SwitchOutLogicalThreadState hresult();" & _
		"LocksHeldByLogicalThread hresult();" & _
		"MapFile hresult();" & _
		"GetConfiguration hresult();" & _
		"Start hresult();" & _
		"Stop hresult();" & _
		"CreateDomain hresult();" & _
		"GetDefaultDomain hresult(ptr*);" & _
		"EnumDomains hresult();" & _
		"NextDomain hresult();" & _
		"CloseEnum hresult();" & _
		"CreateDomainEx hresult();" & _
		"CreateDomainSetup hresult();" & _
		"CreateEvidence hresult();" & _
		"UnloadDomain hresult();" & _
		"CurrentDomain hresult();"

Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Global Const $sTag_IDispatch = _
		"GetTypeInfoCount hresult(dword*);" & _ ; Retrieves the number of type information interfaces that an object provides (either 0 or 1).
		"GetTypeInfo hresult(dword;dword;ptr*);" & _ ; Gets the type information for an object.
		"GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _ ; Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
		"Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);" ; Provides access to properties and methods exposed by an object.

Global Const $sIID_IAppDomain = "{05F696DC-2B29-3663-AD8B-C4389CF2A713}"
Global Const $sTag_IAppDomain = _
		$sTag_IDispatch & _
		"get_ToString hresult();" & _
		"Equals hresult();" & _
		"GetHashCode hresult();" & _
		"GetType hresult(ptr*);" & _
		"InitializeLifetimeService hresult();" & _
		"GetLifetimeService hresult();" & _
		"get_Evidence hresult();" & _
		"add_DomainUnload hresult();" & _
		"remove_DomainUnload hresult();" & _
		"add_AssemblyLoad hresult();" & _
		"remove_AssemblyLoad hresult();" & _
		"add_ProcessExit hresult();" & _
		"remove_ProcessExit hresult();" & _
		"add_TypeResolve hresult();" & _
		"remove_TypeResolve hresult();" & _
		"add_ResourceResolve hresult();" & _
		"remove_ResourceResolve hresult();" & _
		"add_AssemblyResolve hresult();" & _
		"remove_AssemblyResolve hresult();" & _
		"add_UnhandledException hresult();" & _
		"remove_UnhandledException hresult();" & _
		"DefineDynamicAssembly hresult();" & _
		"DefineDynamicAssembly_2 hresult();" & _
		"DefineDynamicAssembly_3 hresult();" & _
		"DefineDynamicAssembly_4 hresult();" & _
		"DefineDynamicAssembly_5 hresult();" & _
		"DefineDynamicAssembly_6 hresult();" & _
		"DefineDynamicAssembly_7 hresult();" & _
		"DefineDynamicAssembly_8 hresult();" & _
		"DefineDynamicAssembly_9 hresult();" & _
		"CreateInstance hresult(bstr;bstr;object*);" & _
		"CreateInstanceFrom hresult();" & _
		"CreateInstance_2 hresult();" & _
		"CreateInstanceFrom_2 hresult();" & _
		"CreateInstance_3 hresult(bstr;bstr;bool;int;ptr;ptr;ptr;ptr;ptr;ptr*);" & _
		"CreateInstanceFrom_3 hresult();" & _
		"Load hresult();" & _
		"Load_2 hresult();" & _
		"Load_3 hresult();" & _
		"Load_4 hresult();" & _
		"Load_5 hresult();" & _
		"Load_6 hresult();" & _
		"Load_7 hresult();" & _
		"ExecuteAssembly hresult();" & _
		"ExecuteAssembly_2 hresult();" & _
		"ExecuteAssembly_3 hresult();" & _
		"get_FriendlyName hresult();" & _
		"get_BaseDirectory hresult();" & _
		"get_RelativeSearchPath hresult();" & _
		"get_ShadowCopyFiles hresult();" & _
		"GetAssemblies hresult(ptr*);" & _ ; replaced "GetAssemblies hresult();" & _
		"AppendPrivatePath hresult();" & _
		"ClearPrivatePath ) = 0; hresult();" & _
		"SetShadowCopyPath hresult();" & _
		"ClearShadowCopyPath ) = 0; hresult();" & _
		"SetCachePath hresult();" & _
		"SetData hresult();" & _
		"GetData hresult();" & _
		"SetAppDomainPolicy hresult();" & _
		"SetThreadPrincipal hresult();" & _
		"SetPrincipalPolicy hresult();" & _
		"DoCallBack hresult();" & _
		"get_DynamicDirectory hresult();"

Global Const $sIID_IType = "{BCA8B44D-AAD6-3A86-8AB7-03349F4F2DA2}"
Global Const $sTag_IType = _
    $sTag_IDispatch & _
    "get_ToString hresult(bstr*);" & _
    "Equals hresult(variant;short*);" & _
    "GetHashCode hresult(int*);" & _
    "GetType hresult(ptr);" & _
    "get_MemberType hresult(ptr);" & _
    "get_name hresult(bstr*);" & _
    "get_DeclaringType hresult(ptr);" & _
    "get_ReflectedType hresult(ptr);" & _
    "GetCustomAttributes hresult(ptr;short;ptr);" & _
    "GetCustomAttributes_2 hresult(short;ptr);" & _
    "IsDefined hresult(ptr;short;short*);" & _
    "get_Guid hresult(ptr);" & _
    "get_Module hresult(ptr);" & _
    "get_Assembly hresult(ptr*);" & _
    "get_TypeHandle hresult(ptr);" & _
    "get_FullName hresult(bstr*);" & _
    "get_Namespace hresult(bstr*);" & _
    "get_AssemblyQualifiedName hresult(bstr*);" & _
    "GetArrayRank hresult(int*);" & _
    "get_BaseType hresult(ptr);" & _
    "GetConstructors hresult(ptr;ptr);" & _
    "GetInterface hresult(bstr;short;ptr);" & _
    "GetInterfaces hresult(ptr);" & _
    "FindInterfaces hresult(ptr;variant;ptr);" & _
    "GetEvent hresult(bstr;ptr;ptr);" & _
    "GetEvents hresult(ptr);" & _
    "GetEvents_2 hresult(int;ptr);" & _
    "GetNestedTypes hresult(int;ptr);" & _
    "GetNestedType hresult(bstr;ptr;ptr);" & _
    "GetMember hresult(bstr;ptr;ptr;ptr);" & _
    "GetDefaultMembers hresult(ptr);" & _
    "FindMembers hresult(ptr;ptr;ptr;variant;ptr);" & _
    "GetElementType hresult(ptr);" & _
    "IsSubclassOf hresult(ptr;short*);" & _
    "IsInstanceOfType hresult(variant;short*);" & _
    "IsAssignableFrom hresult(ptr;short*);" & _
    "GetInterfaceMap hresult(ptr;ptr);" & _
    "GetMethod hresult(bstr;ptr;ptr;ptr;ptr;ptr);" & _
    "GetMethod_2 hresult(bstr;ptr;ptr);" & _
    "GetMethods hresult(int;ptr);" & _
    "GetField hresult(bstr;ptr;ptr);" & _
    "GetFields hresult(int;ptr);" & _
    "GetProperty hresult(bstr;ptr;ptr);" & _
    "GetProperty_2 hresult(bstr;ptr;ptr;ptr;ptr;ptr;ptr);" & _
    "GetProperties hresult(ptr;ptr);" & _
    "GetMember_2 hresult(bstr;ptr;ptr);" & _
    "GetMembers hresult(int;ptr*);" & _ ; Replaced "GetMembers hresult(int;ptr);" & _
    "InvokeMember hresult(bstr;ptr;ptr;variant;ptr;ptr;ptr;ptr;variant*);" & _
    "get_UnderlyingSystemType hresult(ptr);" & _
    "InvokeMember_2 hresult(bstr;int;ptr;variant;ptr;ptr;variant*);" & _
    "InvokeMember_3 hresult(bstr;int;ptr;variant;ptr;variant*);" & _
    "GetConstructor hresult(ptr;ptr;ptr;ptr;ptr;ptr);" & _
    "GetConstructor_2 hresult(ptr;ptr;ptr;ptr;ptr);" & _
    "GetConstructor_3 hresult(ptr;ptr);" & _
    "GetConstructors_2 hresult(ptr);" & _
    "get_TypeInitializer hresult(ptr);" & _
    "GetMethod_3 hresult(bstr;ptr;ptr;ptr;ptr;ptr;ptr);" & _
    "GetMethod_4 hresult(bstr;ptr;ptr;ptr);" & _
    "GetMethod_5 hresult(bstr;ptr;ptr);" & _
    "GetMethod_6 hresult(bstr;ptr);" & _
    "GetMethods_2 hresult(ptr);" & _
    "GetField_2 hresult(bstr;ptr);" & _
    "GetFields_2 hresult(ptr);" & _
    "GetInterface_2 hresult(bstr;ptr);" & _
    "GetEvent_2 hresult(bstr;ptr);" & _
    "GetProperty_3 hresult(bstr;ptr;ptr;ptr;ptr);" & _
    "GetProperty_4 hresult(bstr;ptr;ptr;ptr);" & _
    "GetProperty_5 hresult(bstr;ptr;ptr);" & _
    "GetProperty_6 hresult(bstr;ptr;ptr);" & _
    "GetProperty_7 hresult(bstr;ptr);" & _
    "GetProperties_2 hresult(ptr);" & _
    "GetNestedTypes_2 hresult(ptr);" & _
    "GetNestedType_2 hresult(bstr;ptr);" & _
    "GetMember_3 hresult(bstr;ptr);" & _
    "GetMembers_2 hresult(ptr);" & _
    "get_Attributes hresult(ptr);" & _
    "get_IsNotPublic hresult(short*);" & _
    "get_IsPublic hresult(short*);" & _
    "get_IsNestedPublic hresult(short*);" & _
    "get_IsNestedPrivate hresult(short*);" & _
    "get_IsNestedFamily hresult(short*);" & _
    "get_IsNestedAssembly hresult(short*);" & _
    "get_IsNestedFamANDAssem hresult(short*);" & _
    "get_IsNestedFamORAssem hresult(short*);" & _
    "get_IsAutoLayout hresult(short*);" & _
    "get_IsLayoutSequential hresult(short*);" & _
    "get_IsExplicitLayout hresult(short*);" & _
    "get_IsClass hresult(short*);" & _
    "get_IsInterface hresult(short*);" & _
    "get_IsValueType hresult(short*);" & _
    "get_IsAbstract hresult(short*);" & _
    "get_IsSealed hresult(short*);" & _
    "get_IsEnum hresult(short*);" & _
    "get_IsSpecialName hresult(short*);" & _
    "get_IsImport hresult(short*);" & _
    "get_IsSerializable hresult(short*);" & _
    "get_IsAnsiClass hresult(short*);" & _
    "get_IsUnicodeClass hresult(short*);" & _
    "get_IsAutoClass hresult(short*);" & _
    "get_IsArray hresult(short*);" & _
    "get_IsByRef hresult(short*);" & _
    "get_IsPointer hresult(short*);" & _
    "get_IsPrimitive hresult(short*);" & _
    "get_IsCOMObject hresult(short*);" & _
    "get_HasElementType hresult(short*);" & _
    "get_IsContextful hresult(short*);" & _
    "get_IsMarshalByRef hresult(short*);" & _
    "Equals_2 hresult(ptr;short*);"


Global Const $sIID_IAssembly = "{17156360-2F1A-384A-BC52-FDE93C215C5B}"
Global Const $sTag_IAssembly = _
		$sTag_IDispatch & _
		"get_ToString hresult(bstr*);" & _ ; Replaced "get_ToString hresult();" & _
		"Equals hresult();" & _
		"GetHashCode hresult();" & _
		"GetType hresult(ptr*);" & _
		"get_CodeBase hresult();" & _
		"get_EscapedCodeBase hresult();" & _
		"GetName hresult();" & _
		"GetName_2 hresult();" & _
		"get_FullName hresult(bstr*);" & _ ; Replaced "get_FullName hresult();" & _
		"get_EntryPoint hresult();" & _
		"GetType_2 hresult(bstr;ptr*);" & _
		"GetType_3 hresult();" & _
		"GetExportedTypes hresult();" & _
		"GetTypes hresult(ptr*);" & _ ; Replaced "GetTypes hresult();" & _
		"GetManifestResourceStream hresult();" & _
		"GetManifestResourceStream_2 hresult();" & _
		"GetFile hresult();" & _
		"GetFiles hresult();" & _
		"GetFiles_2 hresult();" & _
		"GetManifestResourceNames hresult();" & _
		"GetManifestResourceInfo hresult();" & _
		"get_Location hresult();" & _
		"get_Evidence hresult();" & _
		"GetCustomAttributes hresult();" & _
		"GetCustomAttributes_2 hresult();" & _
		"IsDefined hresult();" & _
		"GetObjectData hresult();" & _
		"add_ModuleResolve hresult();" & _
		"remove_ModuleResolve hresult();" & _
		"GetType_4 hresult();" & _
		"GetSatelliteAssembly hresult();" & _
		"GetSatelliteAssembly_2 hresult();" & _
		"LoadModule hresult();" & _
		"LoadModule_2 hresult();" & _
		"CreateInstance hresult(bstr;variant*);" & _
		"CreateInstance_2 hresult(bstr;bool;variant*);" & _
		"CreateInstance_3 hresult(bstr;bool;int;ptr;ptr;ptr;ptr;variant*);" & _
		"GetLoadedModules hresult();" & _
		"GetLoadedModules_2 hresult();" & _
		"GetModules hresult();" & _
		"GetModules_2 hresult();" & _
		"GetModule hresult();" & _
		"GetReferencedAssemblies hresult();" & _
		"get_GlobalAssemblyCache hresult();"

Global Const $sIID_IObjectHandle = "{C460E2B4-E199-412A-8456-84DC3E4838C3}"
Global Const $sTag_IObjectHandle = _
		"Unwrap hresult(variant*);"

; *************************************************************************
; DllCall MSCorEE.DLL (The module containing the .NET Framework Functions
; *************************************************************************

Global Const $hMSCorEE = DllOpen("MSCorEE.DLL")
Global $oErrorHandler = ObjEvent("AutoIt.Error", "__CLR_ErrorHandler")

; *************************************************************************
; _CLR_GetDefaultDomain() : The AppDomain contains the InProc session state
; *************************************************************************
Func _CLR_GetDefaultDomain()
	Local $aCall = DllCall($hMSCorEE, "long", "CLRCreateInstance", "struct*", $tCLSID_CLRMetaHost, "struct*", $tIID_ICLRMetaHost, "ptr*", 0)
	If @error Or $aCall[0] Then Return SetError(1, 0, 0)
	If $aCall[0] = 0 Then ;$S_OK = 0
		Local $pClrHost = $aCall[3]
		Local $oClrHost = ObjCreateInterface($pClrHost, $sIID_ICLRMetaHost, $sTag_ICLRMetaHost)
;~ 		ConsoleWrite(">oClrHost: " & IsObj($oClrHost) & @CRLF)

		Local $sNETFrameworkVersion = "v4.0.30319"
		Local $tCLRRuntimeInfo = DllStructCreate("ptr")
		$oClrHost.GetRuntime($sNETFrameworkVersion, $tIID_ICLRRuntimeInfo, DllStructGetPtr($tCLRRuntimeInfo))
		Local $pCLRRuntimeInfo = DllStructGetData($tCLRRuntimeInfo, 1)
;~ 		ConsoleWrite(">pCLRRuntimeInfo: " & $pCLRRuntimeInfo & @CRLF)

		Local $oCLRRuntimeInfo = ObjCreateInterface($pCLRRuntimeInfo, $sIID_ICLRRuntimeInfo, $sTag_ICLRRuntimeInfo)
;~ 		ConsoleWrite(">oCLRRuntimeInfo: " & IsObj($oCLRRuntimeInfo) & @CRLF)
		Local $isIsLoadable = 0
		$oCLRRuntimeInfo.IsLoadable($isIsLoadable)
;~ 		ConsoleWrite(">IsLoadable: " & $isIsLoadable & @CRLF)

		If $isIsLoadable Then
			Local $tCLRRuntimeHost = DllStructCreate("ptr")
			$oCLRRuntimeInfo.GetInterface(DllStructGetPtr($tCLSID_CLRRuntimeHost), DllStructGetPtr($tIID_ICLRRuntimeHost), DllStructGetPtr($tCLRRuntimeHost))
			Local $pCLRRuntimeHost = DllStructGetData($tCLRRuntimeHost, 1)
;~ 			ConsoleWrite(">pCLRRuntimeHost: " & $pCLRRuntimeHost & @CRLF)
			Local $oCLRRuntimeHost = ObjCreateInterface($pCLRRuntimeHost, $sIID_ICLRRuntimeHost, $sTag_ICLRRuntimeHost)
;~ 			ConsoleWrite(">oCLRRuntimeHost: " & IsObj($oCLRRuntimeHost) & @CRLF)

			$oCLRRuntimeHost.Start()
			Local $tCorRuntimeHost = DllStructCreate("ptr")
			$oCLRRuntimeInfo.GetInterface(DllStructGetPtr($tCLSID_CorRuntimeHost), DllStructGetPtr($tIID_ICorRuntimeHost), DllStructGetPtr($tCorRuntimeHost))
			Local $pCorRuntimeHost = DllStructGetData($tCorRuntimeHost, 1)
;~ 			;ConsoleWrite("$pCorRuntimeHost = " & $pCorRuntimeHost & @CRLF)

			Local $oCorRuntimeHost = ObjCreateInterface($pCorRuntimeHost, $sIID_ICorRuntimeHost, $sTag_ICorRuntimeHost)
;~ 			ConsoleWrite("IsObj( $oCorRuntimeHost ) = " & IsObj($oCorRuntimeHost) & @CRLF)

			$oCorRuntimeHost.Start()
			Local $pAppDomain = 0
			$oCorRuntimeHost.GetDefaultDomain($pAppDomain)
;~ 			ConsoleWrite("$pAppDomain = " & Ptr($pAppDomain) & @CRLF)

			Local $oAppDomain = ObjCreateInterface($pAppDomain, $sIID_IAppDomain, $sTag_IAppDomain)
;~ 			ConsoleWrite("IsObj( $oAppDomain ) = " & IsObj($oAppDomain) & @CRLF)
			Return $oAppDomain
		EndIf
	EndIf
EndFunc   ;==>_CLR_GetDefaultDomain

; *************************************************************************
; _CLR_LoadLibrary() : Loads a .NET Assembly enabling Invoking it's Members
; *************************************************************************
Func _CLR_LoadLibrary($sAssemblyName, $AppDomain = 0)
	Local $oAppDomain = _CLR_GetDefaultDomain()
	Local $pType = 0
	$oAppDomain.GetType($pType)
;~ 	ConsoleWrite("$pType = " & Ptr($pType) & @CRLF)

	Local $oType = ObjCreateInterface($pType, $sIID_IType, $sTag_IType)
;~ 	ConsoleWrite("IsObj( $oType ) = " & IsObj($oType) & @CRLF)

	Local $pAssembly = 0
	$oType.get_Assembly($pAssembly)
;~ 	ConsoleWrite("$pAssembly = " & Ptr($pAssembly) & @CRLF)

	Local $oAssembly = ObjCreateInterface($pAssembly, $sIID_IAssembly, $sTag_IAssembly)
;~ 	ConsoleWrite("IsObj( $oAssembly ) = " & IsObj($oAssembly) & @CRLF)

	Local $pAssemblyType = 0
	$oAssembly.GetType($pAssemblyType)
;~ 	ConsoleWrite("$pAssemblyType = " & Ptr($pAssemblyType) & @CRLF)

	Local $oAssemblyType = ObjCreateInterface($pAssemblyType, $sIID_IType, $sTag_IType)
;~ 	ConsoleWrite("IsObj( $oAssemblyType ) = " & IsObj($oAssemblyType) & @CRLF)

	; args := ComObjArray(0xC, 1),  args[0] := "XPTable.dll"
	Local $tSafeArrayBound = DllStructCreate($tagSAFEARRAYBOUND)
	Local $pSafeArray, $pSafeArrayData
	DllStructSetData($tSafeArrayBound, "cElements", 1)
	DllStructSetData($tSafeArrayBound, "lLbound", 0)
	$pSafeArray = SafeArrayCreate($VT_VARIANT, 1, $tSafeArrayBound)
	SafeArrayAccessData($pSafeArray, $pSafeArrayData)
	DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_BSTR)
	DllStructSetData(DllStructCreate("ptr", $pSafeArrayData + 8), 1, SysAllocString($sAssemblyName))
	SafeArrayUnaccessData($pSafeArray)

	Local $pObject = 0
	$oAssemblyType.InvokeMember_3("LoadFrom", 0x158, 0, 0, $pSafeArray, $pObject)
;~ 	ConsoleWrite("-$pObject = " & Ptr($pObject) & @CRLF)

	If Not Ptr($pObject) Then
		; args := ComObjArray(0xC, 1),  args[0] := "System"
		Local $tSafeArrayBound = DllStructCreate($tagSAFEARRAYBOUND)
		Local $pSafeArray, $pSafeArrayData
		DllStructSetData($tSafeArrayBound, "cElements", 1)
		DllStructSetData($tSafeArrayBound, "lLbound", 0)
		$pSafeArray = SafeArrayCreate($VT_VARIANT, 1, $tSafeArrayBound)
		SafeArrayAccessData($pSafeArray, $pSafeArrayData)
		DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_BSTR)
		DllStructSetData(DllStructCreate("ptr", $pSafeArrayData + 8), 1, SysAllocString($sAssemblyName))
		SafeArrayUnaccessData($pSafeArray)

		$oAssemblyType.InvokeMember_3("LoadWithPartialName", 0x158, 0, 0, $pSafeArray, $pObject)

;~ 		ConsoleWrite("-$pAsmProvider = " & Ptr($pObject) & @CRLF)

	EndIf
	$oObject = ObjCreateInterface($pObject, $sIID_IAssembly, $sTag_IAssembly)
;~ 	ConsoleWrite("-IsObj( $oObject ) = " & IsObj($oObject) & @CRLF)

	Return $oObject
EndFunc   ;==>_CLR_LoadLibrary


; *************************************************************************
; _CLR_CreateObject() : Creates an Instance to a .NET COM Object + Parameters
; *************************************************************************
Func _CLR_CreateObject(ByRef $oAssembly, $sTypeName, $v3 = Default, $v4 = Default, $v5 = Default, $v6 = Default, $v7 = Default, $v8 = Default, $v9 = Default)
	Local $aParams = [ $v3, $v4, $v5, $v6, $v7, $v8, $v9 ], $oObject = 0

	If @NumParams = 2 Then
		$oAssembly.CreateInstance_2($sTypeName, True, $oObject)
		Return $oObject
	EndIf

	Local $iArgs = @NumParams - 2, $aArgs[$iArgs]
	For $i = 0 To $iArgs - 1
		$aArgs[$i] = $aParams[$i]
	Next

	; static Array_Empty := ComObjArray(0xC,0), null := ComObject(13,0)
	Local $pSAEmpty, $tSAB = DllStructCreate($tagSAFEARRAYBOUND)
	DllStructSetData($tSAB, "cElements", 0)
	DllStructSetData($tSAB, "lLbound", 0)
	$pSAEmpty = SafeArrayCreate($VT_VARIANT, 0, $tSAB)

	$oAssembly.CreateInstance_3($sTypeName, True, 0, 0, CreateSafeArray($aArgs), 0, $pSAEmpty, $oObject)
	Return $oObject
EndFunc   ;==>_CLR_CreateObject

; *************************************************************************
; _CLR_CompileAssembly : Initiates the .NET Compiler Engine
; *************************************************************************
Func _CLR_CompileAssembly($sCode, $sReferences, $sProviderAssembly, $sProviderType, $AppDomain = 0, $sFileName = "", $sCompilerOptions = "")
	Local $oAppDomain = _CLR_GetDefaultDomain()
	Local $pType = 0
	$oAppDomain.GetType($pType)
;~ 	ConsoleWrite("$pType = " & Ptr($pType) & @CRLF)

	Local $oType = ObjCreateInterface($pType, $sIID_IType, $sTag_IType)
;~ 	ConsoleWrite("IsObj( $oType ) = " & IsObj($oType) & @CRLF)

	Local $pAssembly
	$oType.get_Assembly($pAssembly)
;~ 	ConsoleWrite("$pAssembly = " & Ptr($pAssembly) & @CRLF)

	Local $oAssembly = ObjCreateInterface($pAssembly, $sIID_IAssembly, $sTag_IAssembly)
;~ 	ConsoleWrite("IsObj( $oAssembly ) = " & IsObj($oAssembly) & @CRLF)

	Local $pAssemblyType
	$oAssembly.GetType($pAssemblyType)
;~ 	ConsoleWrite("$pAssemblyType = " & Ptr($pAssemblyType) & @CRLF)

	Local $oAssemblyType = ObjCreateInterface($pAssemblyType, $sIID_IType, $sTag_IType)
;~ 	ConsoleWrite("IsObj( $oAssemblyType ) = " & IsObj($oAssemblyType) & @CRLF)

	; args := ComObjArray(0xC, 1),  args[0] := "System"
	Local $tSafeArrayBound = DllStructCreate($tagSAFEARRAYBOUND)
	Local $pSafeArray, $pSafeArrayData
	DllStructSetData($tSafeArrayBound, "cElements", 1)
	DllStructSetData($tSafeArrayBound, "lLbound", 0)
	$pSafeArray = SafeArrayCreate($VT_VARIANT, 1, $tSafeArrayBound)
	SafeArrayAccessData($pSafeArray, $pSafeArrayData)
	DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_BSTR)
	DllStructSetData(DllStructCreate("ptr", $pSafeArrayData + 8), 1, SysAllocString("System"))
	SafeArrayUnaccessData($pSafeArray)


	Local $pAsmProvider = 0
	$oAssemblyType.InvokeMember_3("LoadWithPartialName", 0x158, 0, 0, $pSafeArray, $pAsmProvider)
;~ 	ConsoleWrite("$pAsmProvider = " & Ptr($pAsmProvider) & @CRLF)


	Local $oAsmProvider = _CLR_LoadLibrary("System")
;~ 	ConsoleWrite("IsObj( $oAsmProvider ) = " & IsObj($oAsmProvider) & @CRLF)

	; codeProvider := asmProvider.CreateInstance("Microsoft.CSharp.CSharpCodeProvider")
	Local $oCodeProvider = 0
	$oAsmProvider.CreateInstance($sProviderType, $oCodeProvider)
;~ 	ConsoleWrite("IsObj( $oCodeProvider ) = " & IsObj($oCodeProvider) & @CRLF)

	; codeCompiler := codeProvider.CreateCompiler()
	Local $oCodeCompiler = $oCodeProvider.CreateCompiler()
;~ 	ConsoleWrite("IsObj( $oCodeCompiler ) = " & IsObj($oCodeCompiler) & @CRLF)

	Local $oAsmSystem = $oAsmProvider
	Local $aReferences = StringSplit($sReferences, "|")

	; Convert | delimited list of references into an array.
	; References = "System.dll | System.Management.dll | System.Windows.Forms.dll"
	; StringSplit, Refs, References, |, %A_Space%%A_Tab%
	; aRefs := ComObjArray(8, Refs0)
	; Loop % Refs0
	; 	aRefs[A_Index-1] := Refs%A_Index%

	Local $tSafeArray1Bound = DllStructCreate($tagSAFEARRAYBOUND)
	Local $pSafeArray1, $pSafeArray1Data
	DllStructSetData($tSafeArray1Bound, "cElements", $aReferences[0])
	DllStructSetData($tSafeArray1Bound, "lLbound", 0)
	$pSafeArray1 = SafeArrayCreate($VT_BSTR, 1, $tSafeArray1Bound)
	SafeArrayAccessData($pSafeArray1, $pSafeArray1Data)
	For $i = 1 To $aReferences[0]
		DllStructSetData(DllStructCreate("ptr", $pSafeArray1Data), 1, SysAllocString(StringStripWS($aReferences[$i], 8)))
		$pSafeArray1Data += @AutoItX64 ? 8 : 4
	Next

;~ 	DllStructSetData(DllStructCreate("ptr", $pSafeArray1Data), 1, SysAllocString("System.dll"))
;~ 	$pSafeArray1Data += @AutoItX64 ? 8 : 4
;~ 	DllStructSetData(DllStructCreate("ptr", $pSafeArray1Data), 1, SysAllocString("System.Management.dll"))
;~ 	$pSafeArray1Data += @AutoItX64 ? 8 : 4
;~ 	DllStructSetData(DllStructCreate("ptr", $pSafeArray1Data), 1, SysAllocString("System.Windows.Forms.dll"))
	SafeArrayUnaccessData($pSafeArray1)

	; prms := CLR_CreateObject(asmSystem, "System.CodeDom.Compiler.CompilerParameters", aRefs)

	; args := ComObjArray(0xC, 1),  args[0] := "XPTable.dll"
	Local $tSafeArray2Bound = DllStructCreate($tagSAFEARRAYBOUND)
	Local $pSafeArray2, $pSafeArray2Data
	DllStructSetData($tSafeArray2Bound, "cElements", 1)
	DllStructSetData($tSafeArray2Bound, "lLbound", 0)
	$pSafeArray2 = SafeArrayCreate($VT_VARIANT, 1, $tSafeArray2Bound)
	SafeArrayAccessData($pSafeArray2, $pSafeArray2Data)
	DllStructSetData(DllStructCreate("word", $pSafeArray2Data), 1, $VT_BSTR + $VT_ARRAY)
	DllStructSetData(DllStructCreate("ptr", $pSafeArray2Data + 8), 1, $pSafeArray1)
	SafeArrayUnaccessData($pSafeArray2)

	; static Array_Empty := ComObjArray(0xC,0), null := ComObject(13,0)
	Local $pSAEmpty, $tSAB = DllStructCreate($tagSAFEARRAYBOUND)
	DllStructSetData($tSAB, "cElements", 0)
	DllStructSetData($tSAB, "lLbound", 0)
	$pSAEmpty = SafeArrayCreate($VT_VARIANT, 0, $tSAB)

	Local $oPrms = 0
	$oAsmSystem.CreateInstance_3("System.CodeDom.Compiler.CompilerParameters", True, 0, 0, $pSafeArray2, 0, $pSAEmpty, $oPrms)
;~ 	ConsoleWrite("IsObj( $oPrms ) = " & IsObj($oPrms) & @CRLF)
;~ 	ConsoleWrite("$sCode = " & @CRLF & $sCode)

;~ 	; Set parameters for compiler.
	$oPrms.OutputAssembly = $sFileName
	$oPrms.GenerateInMemory = ($sFileName = "")
	$oPrms.GenerateExecutable = (StringRight($sFileName, 4) = ".exe")
	$oPrms.CompilerOptions = $sCompilerOptions
	$oPrms.IncludeDebugInformation = True


	; Compile!
	; compilerRes := codeCompiler.CompileAssemblyFromSource(prms, Code)
	Local $oCompilerRes = $oCodeCompiler.CompileAssemblyFromSource($oPrms, $sCode)
;~ 	ConsoleWrite("IsObj( $oCompilerRes ) = " & IsObj($oCompilerRes) & @CRLF)


	If $oCompilerRes.Errors.Count() Then
;~ 		ConsoleWrite("!" & $oCompilerRes.Errors.Count() & @CRLF)
		Return 0
	EndIf

	If $sFileName Then
		Local $sPathToAssembly = $oCompilerRes.PathToAssembly()
;~ 		ConsoleWrite("$sPathToAssembly = " & $sPathToAssembly & @CRLF)
		Return $sPathToAssembly
	Else
		Local $pCodeAssembly = $oCompilerRes.CompiledAssembly()
;~ 		ConsoleWrite("$pCodeAssembly = " & Ptr($pCodeAssembly) & @CRLF)
		Local $oCodeAssembly = ObjCreateInterface($pCodeAssembly, $sIID_IAssembly, $sTag_IAssembly)
;~ 		ConsoleWrite("IsObj( $oCodeAssembly ) = " & IsObj($oCodeAssembly) & @CRLF)
		Return $oCodeAssembly
	EndIf
EndFunc   ;==>_CLR_CompileAssembly

; *************************************************************************
; _CLR_CompileVB : Compile to an Assembly.dll using VB.Net Source Code
; *************************************************************************
Func _CLR_CompileVB($sCode, $sReferences = "", $AppDomain = 0, $sFileName = "", $sCompilerOptions = "")
	Return _CLR_CompileAssembly($sCode, $sReferences, "System", "Microsoft.VisualBasic.VBCodeProvider", $AppDomain, $sFileName, $sCompilerOptions)
EndFunc   ;==>_CLR_CompileVB

; *************************************************************************
; _CLR_CompileVB : Compile to an Assembly.dll using C# Source Code
; *************************************************************************
Func _CLR_CompileCSharp($sCode, $sReferences = "", $AppDomain = 0, $sFileName = "", $sCompilerOptions = "")
	Return _CLR_CompileAssembly($sCode, $sReferences, "System", "Microsoft.CSharp.CSharpCodeProvider", $AppDomain, $sFileName, $sCompilerOptions)
EndFunc   ;==>_CLR_CompileCSharp

; *************************************************************************
; CreateSafeArray : .NET CLR needs SafeArrays ;-)
; *************************************************************************
Func CreateSafeArray( $aArgs )
	Local $tSafeArrayBound = DllStructCreate($tagSAFEARRAYBOUND)
	Local $iArgs = UBound( $aArgs ), $pSafeArray, $pSafeArrayData
	DllStructSetData($tSafeArrayBound, "cElements", $iArgs)
	$pSafeArray = SafeArrayCreate($VT_VARIANT, 1, $tSafeArrayBound)

	SafeArrayAccessData($pSafeArray, $pSafeArrayData)

	For $i = 0 To $iArgs - 1
		Switch VarGetType( $aArgs[$i] )
			Case "Bool"
				DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_BOOL)
				DllStructSetData(DllStructCreate("short", $pSafeArrayData + 8), 1, $aArgs[$i])
			Case "Double"
				DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_R8)
				DllStructSetData(DllStructCreate("double", $pSafeArrayData + 8), 1, $aArgs[$i])
			Case "Int32"
				DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_I4)
				DllStructSetData(DllStructCreate("int", $pSafeArrayData + 8), 1, $aArgs[$i])
			Case "String"
				DllStructSetData(DllStructCreate("word", $pSafeArrayData), 1, $VT_BSTR)
				DllStructSetData(DllStructCreate("ptr", $pSafeArrayData + 8), 1, SysAllocString($aArgs[$i]))
		EndSwitch
		$pSafeArrayData += @AutoItX64 ? 24 : 16
	Next

	SafeArrayUnaccessData($pSafeArray)

	Return $pSafeArray
EndFunc   ;==>CreateSafeArray

; *************************************************************************
; _WinAPI_CLSIDFromString : .NET CLR needs SafeArrays ;-)
; *************************************************************************
Func _WinAPI_CLSIDFromString($sGUID)
	Local $tGUID = DllStructCreate('ulong Data1;ushort Data2;ushort Data3;byte Data4[8]')
	Local $iRet = DllCall('ole32.dll', 'uint', 'CLSIDFromString', 'wstr', $sGUID, 'ptr', DllStructGetPtr($tGUID))
	If (@error) Or ($iRet[0]) Then
		Return SetError(@error, @extended, 0)
	EndIf
	Return $tGUID
EndFunc   ;==>_WinAPI_CLSIDFromString

; *************************************************************************
; __CLR_ErrorHandler : Intercept .NET COM Errors
; *************************************************************************
Func __CLR_ErrorHandler($oError)
	; Do anything here.
	ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
			@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
			@TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
			@TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
			@TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
			@TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
			@TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
			@TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
			@TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
			@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>__CRL_ErrorHandler
