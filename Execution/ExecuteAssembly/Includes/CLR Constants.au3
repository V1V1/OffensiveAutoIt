; ---------------------------------
; Global Constants binding flags :
; ---------------------------------
    ;
    ;     Specifies no binding flag.
    ;
    Global Const $Default = 0

    ;
    ;     Specifies that the case of the member name should not be considered when binding.
    ;
    Global Const $IgnoreCase = 1

    ;
    ;     Specifies that only members declared at the level of the supplied type's
    ;     hierarchy should be considered. Inherited members are not considered.
    ;
    Global Const $DeclaredOnly = 2

    ;
    ;     Specifies that instance members are to be included in the search.
    ;
    Global Const $Instance = 4

    ;
    ;     Specifies that static members are to be included in the search.
    ;
    Global Const $Static = 8

    ;
    ;     Specifies that public members are to be included in the search.
    ;
    Global Const $Public = 16

    ;
    ;     Specifies that non-public members are to be included in the search.
    ;
    Global Const $NonPublic = 32

    ;
    ;     Specifies that public and protected static members up the hierarchy should
    ;     be returned. Private static members in inherited classes are not returned.
    ;     Static members include fields, methods, events, and properties. Nested types
    ;     are not returned.
    ;
    Global Const $FlattenHierarchy = 64

    ;
    ;     Specifies that a method is to be invoked. This must not be a constructor
    ;     or a type initializer.
    ;
    Global Const $InvokeMethod = 256

    ;
    ;     Specifies that Reflection should create an instance of the specified type.
    ;     Calls the constructor that matches the given arguments. The supplied member
    ;     name is ignored. If the type of lookup is not specified, (Instance | Public)
    ;     will apply. It is not possible to call a type initializer.
    ;
    Global Const $CreateInstance = 512

    ;
    ;     Specifies that the value of the specified field should be returned.
    ;
    Global Const $GetField = 1024

    ;
    ;     Specifies that the value of the specified field should be set.
    ;
    Global Const $SetField = 2048

    ;
    ;     Specifies that the value of the specified property should be returned.
    ;
    Global Const $GetProperty = 4096

    ;
    ;     Specifies that the value of the specified property should be set. For COM
    ;     properties, specifying this binding flag is equivalent to specifying PutDispProperty
    ;     and PutRefDispProperty.
    ;
    Global Const $SetProperty = 8192

    ;
    ;     Specifies that the PROPPUT member on a COM object should be invoked. PROPPUT
    ;     specifies a property-setting function that uses a value. Use PutDispProperty
    ;     if a property has both PROPPUT and PROPPUTREF and you need to distinguish
    ;     which one is called.
    ;
    Global Const $PutDispProperty = 16384

    ;
    ;     Specifies that the PROPPUTREF member on a COM object should be invoked. PROPPUTREF
    ;     specifies a property-setting function that uses a reference instead of a
    ;     value. Use PutRefDispProperty if a property has both PROPPUT and PROPPUTREF
    ;     and you need to distinguish which one is called.
    ;
    Global Const $PutRefDispProperty = 32768

    ;
    ;     Specifies that types of the supplied arguments must exactly match the types
    ;     of the corresponding formal parameters. Reflection throws an exception if
    ;     the caller supplies a non-null Binder object, since that implies that the
    ;     caller is supplying BindToXXX implementations that will pick the appropriate
    ;     method.
    ;
    Global Const $ExactBinding = 65536

    ;
    ;     Not implemented.
    ;
    Global Const $SuppressChangeType = 131072

    ;
    ;     Returns the set of members whose parameter count matches the number of supplied
    ;     arguments. This binding flag is used for methods with parameters that have
    ;     default values and methods with variable arguments (varargs). This flag should
    ;     only be used with System.Type.InvokeMember(System.String,System.Reflection.BindingFlags,System.Reflection.Binder,System.Object,System.Object[],System.Reflection.ParameterModifier[],System.Globalization.CultureInfo,System.String[]).
    ;
    Global Const $OptionalParamBinding = 262144

    ;
    ;     Used in COM interop to specify that the return value of the member can be
    ;     ignored.
    ;
    Global Const $IgnoreReturn = 16777216