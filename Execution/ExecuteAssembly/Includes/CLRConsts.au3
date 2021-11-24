#include-once
; Moved this stuff here to make the script easier to read/modify

Global Const $S_OK = 0

Global Const $sTag_CLRMetaHost = _
        "GetRuntime hresult(wstr;struct*;ptr);" & _
        "GetVersionFromFile hresult(ptr;ptr;ptr);" & _
        "EnumerateInstalledRuntimes hresult(ptr);" & _
        "EnumerateLoadedRuntimes hresult(ptr;ptr);" & _
        "RequestRuntimeLoadedNotification hresult(ptr,ptr,ptr):" & _
        "QueryLegacyV2RuntimeBinding hresult(ptr;ptr);" & _
        "ExitProcess hresult(int);"

Global Const $sTagCLRRuntimeInfo = "GetVersionString hresult(ptr;ptr);" & _
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

Global Const $sTag_CLRRuntimeInfo = _
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

Global Const $sTag_CLRRuntimeHost = _
        "Start hresult();" & _
        "Stop hresult();" & _
        "SetHostControl hresult(ptr);" & _
        "GetCLRControl hresult(ptr*);" & _
        "UnloadAppDomain hresult(ptr;ptr);" & _
        "ExecuteInAppDomain hresult(ptr;ptr;ptr);" & _
        "GetCurrentAppDomainId hresult(ptr);" & _
        "ExecuteApplication hresult(ptr;ptr;ptr;ptr;ptr;ptr);" & _
        "ExecuteInDefaultAppDomain hresult(wstr;wstr;wstr;wstr;ptr*);"

Global Const $sTagEnumUnknown = "Next hresult(ulong;ptr*;ulong); Skip hresult(ptr); Reset hresult(); Clone hresult(ptr);"
Global Const $sIID_IDispatch = "{00020400-0000-0000-C000-000000000046}"
Global Const $sTag_IDispatch = _
        "GetTypeInfoCount hresult(dword*);" & _ ; Retrieves the number of type information interfaces that an object provides (either 0 or 1).
        "GetTypeInfo hresult(dword;dword;ptr*);" & _ ; Gets the type information for an object.
        "GetIDsOfNames hresult(ptr;ptr;dword;dword;ptr);" & _ ; Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
        "Invoke hresult(dword;ptr;dword;word;ptr;ptr;ptr;ptr);" ; Provides access to properties and methods exposed by an object.

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
        "get_EntryPoint hresult(ptr*);" & _
        "GetType_2 ptr(bstr);" & _
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
        "GetMembers hresult(int;ptr);" & _
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

Local Const $sIID_MethodInfo = "{FFCC1B5D-ECB8-38DD-9B01-3DC8ABC2AA5F}"
Local Const $sTag_IMethodInfo = "GetTypeInfoCount hresult();" & _
        "GetTypeInfo hresult();" & _
        "GetIDsOfNames hresult();" & _
        "Invoke hresult();" & _
        "ToString hresult();" & _
        "Equals hresult();" & _
        "GetHashCode hresult();" & _
        "GetType hresult();" & _
        "MemberType hresult();" & _
        "name hresult(bstr*);" & _
        "DeclaringType hresult();" & _
        "ReflectedType hresult();" & _
        "GetCustomAttributes hresult();" & _
        "GetCustomAttributes_2 hresult();" & _
        "IsDefined hresult();" & _
        "GetParameters hresult();" & _
        "GetMethodImplementationFlags hresult();" & _
        "MethodHandle hresult();" & _
        "Attributes hresult();" & _
        "CallingConvention hresult();" & _
        "Invoke_2 hresult();" & _
        "IsPublic hresult();" & _
        "IsPrivate hresult();" & _
        "IsFamily hresult();" & _
        "IsAssembly hresult();" & _
        "IsFamilyAndAssembly hresult();" & _
        "IsFamilyOrAssembly hresult();" & _
        "IsStatic hresult();" & _
        "IsFinal hresult();" & _
        "IsVirtual hresult();" & _
        "IsHideBySig hresult();" & _
        "IsAbstract hresult();" & _
        "IsSpecialName hresult();" & _
        "IsConstructor hresult();" & _
        "Invoke_3 hresult(variant;ptr;variant*);" & _
        "returnType hresult();" & _
        "ReturnTypeCustomAttributes hresult();" & _
        "GetBaseDefinition hresult();"

Local Const $sIID_AppDomain = "{05F696DC-2B29-3663-AD8B-C4389CF2A713}"
Local Const $sTag_AppDomain = _
        "GetTypeInfoCount hresult();" & _
        "GetTypeInfo hresult();" & _
        "GetIDsOfNames hresult();" & _
        "Invoke hresult();" & _
        "get_ToString hresult();" & _
        "Equals hresult();" & _
        "GetHashCode hresult();" & _
        "GetType hresult();" & _
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
        "CreateInstance hresult();" & _
        "CreateInstanceFrom hresult();" & _
        "CreateInstance_2 hresult();" & _
        "CreateInstanceFrom_2 hresult();" & _
        "CreateInstance_3 hresult();" & _
        "CreateInstanceFrom_3 hresult();" & _
        "Load hresult();" & _
        "Load_2 hresult(bstr;ptr*);" & _
        "Load_3 hresult(ptr;ptr*);" & _
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
        "GetAssemblies hresult();" & _
        "AppendPrivatePath hresult();" & _
        "ClearPrivatePath hresult();" & _
        "SetShadowCopyPath hresult();" & _
        "ClearShadowCopyPath hresult();" & _
        "SetCachePath hresult();" & _
        "SetData hresult();" & _
        "GetData hresult();" & _
        "SetAppDomainPolicy hresult();" & _
        "SetThreadPrincipal hresult();" & _
        "SetPrincipalPolicy hresult();" & _
        "DoCallBack hresult();" & _
        "get_DynamicDirectory hresult();"

Local Const $sCLSID_CorRuntimeHost = "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}"
Local Const $sIID_ICorRuntimeHost = "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}"
Local $tCLSID_CorRuntimeHost = _WinAPI_CLSIDFromString($sCLSID_CorRuntimeHost)
Local $tIID_ICorRuntimeHost = _WinAPI_CLSIDFromString($sIID_ICorRuntimeHost)
Local Const $sTag_ICorRuntimeHost = _
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

#Region CLSID & IID
Global Const $sCLSID_CLRMetaHost = "{9280188d-0e8e-4867-b30c-7fa83884e8de}"
Global Const $sIID_ICLRMetaHost = "{d332db9e-b9b3-4125-8207-a14884f53216}"
Global Const $sIID_ICLRRuntimeInfo = "{BD39D1D2-BA2F-486a-89B0-B4B0CB466891}"
Global Const $sCLSID_CLRRuntimeHost = "{90F1A06E-7712-4762-86B5-7A5EBA6BDB02}"
Global Const $sIID_ICLRRuntimeHost = "{90F1A06C-7712-4762-86B5-7A5EBA6BDB02}"
Global Const $sIID_IEnumUnknown = "{00000100-0000-0000-C000-000000000046}"

Global $tCLSID_CLRMetaHost = _WinAPI_CLSIDFromString($sCLSID_CLRMetaHost)
Global $tIID_ICLRMetaHost = _WinAPI_CLSIDFromString($sIID_ICLRMetaHost)
Global $tIID_ICLRRuntimeInfo = _WinAPI_CLSIDFromString($sIID_ICLRRuntimeInfo)
Global $tCLSID_CLRRuntimeHost = _WinAPI_CLSIDFromString($sCLSID_CLRRuntimeHost)
Global $tIID_ICLRRuntimeHost = _WinAPI_CLSIDFromString($sIID_ICLRRuntimeHost)
Global $tIID_IEnumUnknown = _WinAPI_CLSIDFromString($sIID_IEnumUnknown)
#EndRegion CLSID & IID

Func _WinAPI_CLSIDFromString($sGUID)
    Local $tGUID = DllStructCreate('ulong Data1;ushort Data2;ushort Data3;byte Data4[8]')
    Local $iRet = DllCall('ole32.dll', 'uint', 'CLSIDFromString', 'wstr', $sGUID, 'ptr', DllStructGetPtr($tGUID))
    If (@error) Or ($iRet[0]) Then
        Return SetError(@error, @extended, 0)
    EndIf
    Return $tGUID
EndFunc   ;==>_WinAPI_CLSIDFromString
