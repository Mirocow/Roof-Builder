Attribute VB_Name = "mHandler"

'** Exception Handling Module
'** Ideas/credits for this module go out to two projects: Ulli's GPF Interceptor:
'** http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=51606&lngWId=1
'** and Thushan Fernando's wExceptionHandler project.
'** http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41471&lngWId=1
'** Both great foundations for this module, from two talented programmers..
'** I took the best of these two, and made this fleshed out hybrid.
'** Every error I could find is declared and described, and added events for
'** context dump, email and error log notification, application auto-restart,
'** and routine tracking structure.
'** Thanks to Jim for finding an issue with InIde, this has been fixed
'** This will not handle -all- gpf's.. this is a limitation of setunhandledexception api
'** but it will handle most of them, (so try to keep errors out of your code, ok?)

Option Explicit

'//handler constants
Private Const MAX_PARAMS            As Long = 15
Private Const EXECUTE_HANDLER       As Long = 1
Private Const CONTINUE_EXECUTION    As Long = -1

'//shellexecute constant
Private Const SW_NORMAL             As Long = 1

'//I'm pretty sure this is every error there is in vb..
'//errors enum
Public Enum eException
    NoGoSub = &H3
    InvalidProceedure = &H5
    Overflow = &H6
    OutOfMemory = &H7
    SubscriptOutOfRange = &H9
    ArrayLocked = &HA
    DivisionbyZero = &HB
    TypeMismatch = &HD
    OutOfStringSpace = &HE
    ExpressionTooComplex = &H10
    RequestFailed = &H11
    UserInterrupt = &H12
    ResumeWithoutError = &H14
    OutOfStackSpace = &H1C
    FunctionNotDefined = &H23
    DllOverloading = &H2F
    ErrorLoadingLibrary = &H30
    BadDllCallingConvention = &H31
    InternalError = &H33
    BadFileName = &H34
    FileNotFound = &H35
    BadFileMode = &H36
    FileAlreadyOpen = &H37
    IOError = &H39
    FileExists = &H3A
    BadRecordLength = &H3B
    DiskFull = &H3D
    InputEndLine = &H3E
    BadRecordNumber = &H3F
    TooManyFiles = &H43
    DeviceUnavailable = &H44
    PermissionDenied = &H46
    DiskNotReady = &H47
    RenameFailed = &H4A
    AccessError = &H4B
    PathNotFound = &H4C
    ObjectNotSet = &H5B
    LoopNotInitialized = &H5C
    InvalidPattern = &H5D
    InvalidNull = &H5E
    CouldNotLoadDll = &H12A
    InvalidName = &H140
    InvalidFileFormat = &H141
    FailToCreateFile = &H142
    BadModule = &H143
    InvalidResource = &H145
    DataValueNotFound = &H147
    IllegalParameter = &H148
    RegistryAccessFailure = &H145
    ComponentNotRegistered = &H14F
    ComponentNotFound = &H151
    ComponentFailure = &H152
    ObjectAlreadyLoaded = &H168
    CanNotLoadObject = &H169
    ControlNotFound = &H16B
    ObjectWasUnloaded = &H16C
    FailedToUnload = &H16D
    FileOutOfDate = &H170
    OwnershipFailure = &H173
    InvalidProperty = &H17C
    InvalidPropertyArray = &H17C
    InvalidPropertyIndex = &H17D
    NotRunTimeProperty = &H17E
    PropertyInvalidReadOnly = &H17F
    MissingArrayIndex = &H181
    PropertySetNotPermitted = &H183
    PropertyGetFailed = &H189
    PropertyWriteFailed = &H18A
    CanNotShowModal = &H190
    CloseTopModal = &H192
    ObjectPermissionDenied = &H1A3
    PropertyNotFound = &H1A6
    MethodNotFound = &H1A7
    ObjectRequired = &H1A8
    InvalidObjectUse = &H1A9
    CanNotCreateObject = &H1AD
    NoOleAutomation = &H1AE
    ClassNotFound = &H1B0
    ObjectNoSupport = &H1B6
    AutomationError = &H1B8
    LibraryConnectionLost = &H1BA
    NoDefaultValue = &H1BB
    ObjectNotSupportAction = &H1BD
    ObjectNotSupportNamed = &H1BE
    ObjectNotSupportLocale = &H1BF
    NamedArgumentNotFound = &H1C0
    ArgumentNotOptional = &H1C1
    WrongNumberOfArguments = &H1C2
    ObjectNotCollection = &H1C3
    InvalidOrdinal = &H1C4
    DllFunctionNotFound = &H1C5
    ResourceNotFound = &H1C6
    ResourceLocked = &H1C7
    CollectionKeyInUse = &H1C9
    VariableTypeNotSupported = &H1CA
    EventsNotSupported = &H1CB
    InvalidClipboardFormat = &H1CC
    InvalidDataFormat = &H1CC
    FailAutodrawImage = &H1E0
    InvalidPicture = &H1E1
    PrinterError = &H1E2
    PrintDriverNotSupportProperty = &H1E3
    PrinterInformationFailure = &H1E4
    InvalidPictureType = &H1E5
    PrintImageMismatch = &H1E6
    CanNotSaveToTempDirectory = &H2DF
    SearchTextNotFound = &H2E8
    ReplacementTooLong = &H2EA
    ClassNoPropertyName = &H3E8
    ClassNoMethodName = &H3E9
    MissingArgument = &H3EA
    RequiredNumberOfArguments = &H3EB
    UnableToSetPropertyName = &H3ED
    UnableToGetPropertyName = &H3EE
    NoMemory = &H7919
    NoObject = &H791C
    ClassNotSet = &H792A
    ActivateObjectFailed = &H7933
    CreateEmbeddedFailed = &H7938
    ErrorSavingToFile = &H793C
    ErrorLoadingFile = &H793D
End Enum

'//context structure
Private Type tContext
    FltF0                               As Double
    FltF1                               As Double
    FltF2                               As Double
    FltF3                               As Double
    FltF4                               As Double
    FltF5                               As Double
    FltF6                               As Double
    FltF7                               As Double
    FltF8                               As Double
    FltF9                               As Double
    FltF10                              As Double
    FltF11                              As Double
    FltF12                              As Double
    FltF13                              As Double
    FltF14                              As Double
    FltF15                              As Double
    FltF16                              As Double
    FltF17                              As Double
    FltF18                              As Double
    FltF19                              As Double
    FltF20                              As Double
    FltF21                              As Double
    FltF22                              As Double
    FltF23                              As Double
    FltF24                              As Double
    FltF25                              As Double
    FltF26                              As Double
    FltF27                              As Double
    FltF28                              As Double
    FltF29                              As Double
    FltF30                              As Double
    FltF31                              As Double
    IntV0                               As Double
    IntT0                               As Double
    IntT1                               As Double
    IntT2                               As Double
    IntT3                               As Double
    IntT4                               As Double
    IntT5                               As Double
    IntT6                               As Double
    IntT7                               As Double
    IntS0                               As Double
    IntS1                               As Double
    IntS2                               As Double
    IntS3                               As Double
    IntS4                               As Double
    IntS5                               As Double
    IntFp                               As Double
    IntA0                               As Double
    IntA1                               As Double
    IntA2                               As Double
    IntA3                               As Double
    IntA4                               As Double
    IntA5                               As Double
    IntT8                               As Double
    IntT9                               As Double
    IntT10                              As Double
    IntT11                              As Double
    IntRa                               As Double
    IntT12                              As Double
    IntAt                               As Double
    IntGp                               As Double
    IntSp                               As Double
    IntZero                             As Double
    Fpcr                                As Double
    SoftFpcr                            As Double
    Fir                                 As Double
    Psr                                 As Long
    ContextFlags                        As Long
    Fill(4)                             As Long
End Type

'//record structure
Private Type tRecord
    ExceptionCode                       As Long
    ExceptionFlags                      As Long
    pExceptionRecord                    As Long
    ExceptionAddress                    As Long
    NumberParameters                    As Long
    ExceptionInformation(MAX_PARAMS)    As Long
End Type

'//record pointers
Private Type tPointers
    pExceptionRecord                    As tRecord
    ContextRecord                       As tContext
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As tRecord, _
                                                                     ByVal LPEXCEPTION_RECORD As Long, _
                                                                     ByVal lngBytes As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Function Exception_Handler(ByRef tpePtrs As tPointers) As Long
'//exception handler

Dim tpeRecord       As tRecord
Dim cRecord         As New Collection
Dim vItem           As Variant
Dim sTemp           As String
Dim iCount          As Integer
Dim cTemp           As New Collection

On Error Resume Next
    
    '//initialize objects
    tpeRecord = tpePtrs.pExceptionRecord
    Set cRecord = New Collection
    
    '//get records
    Do Until tpeRecord.pExceptionRecord = 0
        CopyMemory tpeRecord, tpeRecord.pExceptionRecord, Len(tpeRecord)
    Loop
    
    '//format record data
    With tpeRecord
        cRecord.Add Get_Exception(.ExceptionCode)
        cRecord.Add "Description: " & Get_Description(.ExceptionCode)
        cRecord.Add "Address: " & Hex_Format(.ExceptionAddress)
        cRecord.Add "Time/Date: " & Time & " [" & Date & "]"
    End With
    
    '//ignore and log event /or/ raise handler
    If cLog.EResume Then
        '//temporary storage
        Set cTemp = New Collection
        
        '//add record
        For Each vItem In cRecord
            cTemp.Add vItem & vbNewLine
        Next vItem
        
        '//log error location
        If Len(cLog.ELocale) > 0 Then
            With frmEvent.txtData
                cTemp.Add "Error Location: " & vbNewLine & cLog.ELocale & vbNewLine
            End With
        End If
    
        '//log error description
        If Len(cLog.EData) > 0 Then
            With frmEvent.txtData
                cTemp.Add "Event Description: " & vbNewLine & cLog.EData & vbNewLine
            End With
        End If
        
        '//dump context record
        If cLog.EDump Then
            With frmEvent.txtData
                cTemp.Add "Exception Data:" & vbNewLine
                iCount = 0
                For Each vItem In Context_Dump(tpePtrs)
                    sTemp = sTemp & vItem & Chr$(32) & Chr$(32)
                    If iCount = 6 Then
                        cTemp.Add sTemp & vbNewLine
                        sTemp = vbNullString
                        iCount = 0
                    End If
                    iCount = iCount + 1
                Next vItem
            End With
        End If
        Exception_Handler = CONTINUE_EXECUTION
    Else

        '//display notification
        With frmEvent.txtData
            For Each vItem In cRecord
                .Text = .Text & vItem & vbNewLine
            Next vItem
        End With
        
        '//log error location
        If Len(cLog.ELocale) > 0 Then
            With frmEvent.txtData
                .Text = .Text & vbNewLine & "Error Location: " & vbNewLine & cLog.ELocale & vbNewLine
            End With
        End If
    
        '//log error description
        If Len(cLog.EData) > 0 Then
            With frmEvent.txtData
                .Text = .Text & vbNewLine & "Event Description: " & vbNewLine & cLog.EData & vbNewLine
            End With
        End If
        
        '//dump context record
        If cLog.EDump Then
            With frmEvent.txtData
                .Text = .Text & vbNewLine & "Exception Data:" & vbNewLine
                For Each vItem In Context_Dump(tpePtrs)
                    sTemp = sTemp & vItem & Chr$(32) & Chr$(32)
                    If iCount = 6 Then
                        .Text = .Text & sTemp & vbNewLine
                        sTemp = vbNullString
                        iCount = 0
                    End If
                    iCount = iCount + 1
                Next vItem
            End With
        End If
        
        '//show message
        frmEvent.Show vbModal
        Exception_Handler = CONTINUE_EXECUTION
    End If

On Error GoTo 0

End Function

Private Function Context_Dump(ByRef tpePtrs As tPointers) As Collection
'//dump context structure

Dim tpeContext      As tContext
Dim cTemp           As New Collection
Dim vItem           As Variant

On Error Resume Next

    tpeContext = tpePtrs.ContextRecord
    Set cTemp = New Collection
    
    With tpeContext
        cTemp.Add Hex_Format(.ContextFlags)
        cTemp.Add .Fill
        cTemp.Add Hex_Format(.IntFp)
        cTemp.Add Hex_Format(.IntGp)
        cTemp.Add Hex_Format(.IntRa)
        cTemp.Add Hex_Format(.IntZero)
        cTemp.Add Hex_Format(.Psr)
        cTemp.Add Hex_Format(.SoftFpcr)
        cTemp.Add Hex_Format(.Fir)
        cTemp.Add Hex_Format(.Fpcr)
        cTemp.Add Hex_Format(.FltF0)
        cTemp.Add Hex_Format(.FltF1)
        cTemp.Add Hex_Format(.FltF2)
        cTemp.Add Hex_Format(.FltF3)
        cTemp.Add Hex_Format(.FltF4)
        cTemp.Add Hex_Format(.FltF5)
        cTemp.Add Hex_Format(.FltF6)
        cTemp.Add Hex_Format(.FltF7)
        cTemp.Add Hex_Format(.FltF8)
        cTemp.Add Hex_Format(.FltF9)
        cTemp.Add Hex_Format(.FltF10)
        cTemp.Add Hex_Format(.FltF11)
        cTemp.Add Hex_Format(.FltF12)
        cTemp.Add Hex_Format(.FltF13)
        cTemp.Add Hex_Format(.FltF14)
        cTemp.Add Hex_Format(.FltF15)
        cTemp.Add Hex_Format(.FltF16)
        cTemp.Add Hex_Format(.FltF17)
        cTemp.Add Hex_Format(.FltF18)
        cTemp.Add Hex_Format(.FltF19)
        cTemp.Add Hex_Format(.FltF20)
        cTemp.Add Hex_Format(.FltF21)
        cTemp.Add Hex_Format(.FltF22)
        cTemp.Add Hex_Format(.FltF23)
        cTemp.Add Hex_Format(.FltF24)
        cTemp.Add Hex_Format(.FltF25)
        cTemp.Add Hex_Format(.FltF26)
        cTemp.Add Hex_Format(.FltF27)
        cTemp.Add Hex_Format(.FltF28)
        cTemp.Add Hex_Format(.FltF29)
        cTemp.Add Hex_Format(.FltF30)
        cTemp.Add Hex_Format(.FltF31)
        cTemp.Add Hex_Format(.IntS0)
        cTemp.Add Hex_Format(.IntS1)
        cTemp.Add Hex_Format(.IntS2)
        cTemp.Add Hex_Format(.IntS3)
        cTemp.Add Hex_Format(.IntS4)
        cTemp.Add Hex_Format(.IntS5)
        cTemp.Add Hex_Format(.IntT0)
        cTemp.Add Hex_Format(.IntT1)
        cTemp.Add Hex_Format(.IntT2)
        cTemp.Add Hex_Format(.IntT3)
        cTemp.Add Hex_Format(.IntT4)
        cTemp.Add Hex_Format(.IntT5)
        cTemp.Add Hex_Format(.IntT6)
        cTemp.Add Hex_Format(.IntT7)
        cTemp.Add Hex_Format(.IntT8)
        cTemp.Add Hex_Format(.IntT9)
        cTemp.Add Hex_Format(.IntT10)
        cTemp.Add Hex_Format(.IntT11)
        cTemp.Add Hex_Format(.IntT12)
        cTemp.Add Hex_Format(.IntA0)
        cTemp.Add Hex_Format(.IntA1)
        cTemp.Add Hex_Format(.IntA2)
        cTemp.Add Hex_Format(.IntA3)
        cTemp.Add Hex_Format(.IntA4)
        cTemp.Add Hex_Format(.IntA5)
    End With
    
    Set Context_Dump = cTemp
    Set cTemp = Nothing
    
On Error GoTo 0
    
End Function

Private Function Hex_Format(ByVal lNum As Long) As String
'//format hex

On Error Resume Next

    Hex_Format = Format$(Right$("00000000" & Hex$(lNum), 8), "<@@\-@@\-@@\-@@")
    
On Error GoTo 0
    
End Function

Public Function Get_Description(ExceptionType As eException) As String
'//error description string

On Error Resume Next

    Select Case CInt(ExceptionType)
    Case NoGoSub
        Get_Description = "Return without GoSub."
    Case InvalidProceedure
        Get_Description = "Invalid procedure call."
    Case Overflow
        Get_Description = "Overflow."
    Case OutOfMemory
        Get_Description = "Out of memory."
    Case SubscriptOutOfRange
        Get_Description = "Subscript out of range."
    Case ArrayLocked
        Get_Description = "This array is fixed or temporarily locked."
    Case DivisionbyZero
        Get_Description = "Division by zero."
    Case TypeMismatch
        Get_Description = "Type mismatch."
    Case OutOfStringSpace
        Get_Description = "Out of string space."
    Case ExpressionTooComplex
        Get_Description = "Expression too complex."
    Case RequestFailed
        Get_Description = "Can't perform requested operation."
    Case UserInterrupt
        Get_Description = "User interrupt occurred."
    Case ResumeWithoutError
        Get_Description = "Resume without error."
    Case OutOfStackSpace
        Get_Description = "Out of stack space."
    Case FunctionNotDefined
        Get_Description = "Sub, function, or property not defined."
    Case DllOverloading
        Get_Description = "Too many DLL application clients."
    Case ErrorLoadingLibrary
        Get_Description = "Error in loading DLL."
    Case BadDllCallingConvention
        Get_Description = "Bad DLL calling convention."
    Case InternalError
        Get_Description = "Internal Error."
    Case BadFileName
        Get_Description = "Bad file name or number."
    Case FileNotFound
        Get_Description = "File Not found."
    Case BadFileMode
        Get_Description = "Bad file mode."
    Case FileAlreadyOpen
        Get_Description = "File already open."
    Case IOError
        Get_Description = "Device I/O error."
    Case FileExists
        Get_Description = "File already exists."
    Case BadRecordLength
        Get_Description = "Bad record length."
    Case DiskFull
        Get_Description = "Disk full."
    Case InputEndLine
        Get_Description = "Input past end of line."
    Case BadRecordNumber
        Get_Description = "Bad record number."
    Case TooManyFiles
        Get_Description = "Too many files."
    Case DeviceUnavailable
        Get_Description = "Device unavailable."
    Case PermissionDenied
        Get_Description = "Permission denied."
    Case DiskNotReady
        Get_Description = "Disk Not ready."
    Case RenameFailed
        Get_Description = "Can't rename with different drive."
    Case AccessError
        Get_Description = "Path/File access error."
    Case PathNotFound
        Get_Description = "Path Not found."
    Case ObjectNotSet
        Get_Description = "Object variable or With block variable not set."
    Case LoopNotInitialized
        Get_Description = "For Loop not initialized"
    Case InvalidPattern
        Get_Description = "Invalid pattern string."
    Case InvalidNull
        Get_Description = "Invalid use of Null."
    Case CouldNotLoadDll
        Get_Description = "System DLL could not be loaded."
    Case InvalidName
        Get_Description = "Can't use character device names in specified file names."
    Case InvalidFileFormat
        Get_Description = "Invalid file format"
    Case FailToCreateFile
        Get_Description = "Can't create necessary temporary file."
    Case BadModule
        Get_Description = "Can't load module; invalid format."
    Case InvalidResource
        Get_Description = "Invalid format in resource file."
    Case DataValueNotFound
        Get_Description = "Data value named was not found."
    Case IllegalParameter
        Get_Description = "Illegal parameter; can't write arrays."
    Case RegistryAccessFailure
        Get_Description = "Could not access system registry."
    Case ComponentNotRegistered
        Get_Description = "ActiveX component not correctly registered."
    Case ComponentNotFound
        Get_Description = "ActiveX component not found."
    Case ComponentFailure
        Get_Description = "ActiveX component did not correctly run."
    Case ObjectAlreadyLoaded
        Get_Description = "Object already loaded."
    Case CanNotLoadObject
        Get_Description = "Can't load or unload this object."
    Case ControlNotFound
        Get_Description = "Specified ActiveX control not found."
    Case ObjectWasUnloaded
        Get_Description = "Object was unloaded."
    Case FailedToUnload
        Get_Description = "Unable to unload within this context."
    Case FileOutOfDate
        Get_Description = "The specified file is out of date. This program requires a newer version."
    Case OwnershipFailure
        Get_Description = "The specified object can't be used as an owner form for Show."
    Case InvalidProperty
        Get_Description = "Invalid property value."
    Case InvalidPropertyArray
        Get_Description = "Invalid property-array."
    Case InvalidPropertyIndex
        Get_Description = "Invalid property-array index."
    Case NotRunTimeProperty
        Get_Description = "Property Set can't be executed at run time."
    Case PropertyInvalidReadOnly
        Get_Description = "Property Set can't be used with a read-only property."
    Case MissingArrayIndex
        Get_Description = "Need property-array index."
    Case PropertySetNotPermitted
        Get_Description = "Property Set not permitted."
    Case PropertyGetFailed
        Get_Description = "Property Get can't be executed at run time."
    Case PropertyWriteFailed
        Get_Description = "Property Get can't be executed on write-only property."
    Case CanNotShowModal
        Get_Description = "Form already displayed; can't show modally."
    Case CloseTopModal
        Get_Description = "Code must close topmost modal form first."
    Case ObjectPermissionDenied
        Get_Description = "Permission to use object denied."
    Case PropertyNotFound
        Get_Description = "Property not found."
    Case MethodNotFound
        Get_Description = "Property or method not found."
    Case ObjectRequired
        Get_Description = "Object required."
    Case InvalidObjectUse
        Get_Description = "Invalid object use."
    Case CanNotCreateObject
        Get_Description = "ActiveX component can't create object or return reference to this object."
    Case NoOleAutomation
        Get_Description = "Class doesn't support Automation."
    Case ClassNotFound
        Get_Description = "File name or class name not found during Automation operation."
    Case ObjectNoSupport
        Get_Description = "Object doesn't support this property or method."
    Case AutomationError
        Get_Description = "OLE Automation error."
    Case LibraryConnectionLost
        Get_Description = "Connection to type library or object library for remote process has been lost."
    Case NoDefaultValue
        Get_Description = "Automation object doesn't have a default value."
    Case ObjectNotSupportAction
        Get_Description = "Object doesn't support this action."
    Case ObjectNotSupportNamed
        Get_Description = "Object doesn't support named arguments."
    Case ObjectNotSupportLocale
        Get_Description = "Object doesn't support current locale settings."
    Case NamedArgumentNotFound
        Get_Description = "Named Not Argument."
    Case ArgumentNotOptional
        Get_Description = "Argument not optional or invalid property assignment."
    Case WrongNumberOfArguments
        Get_Description = "Wrong number of arguments or invalid property assignment."
    Case ObjectNotCollection
        Get_Description = "Object not a collection."
    Case InvalidOrdinal
        Get_Description = "Invalid ordinal."
    Case DllFunctionNotFound
        Get_Description = "Specified DLL function not found."
    Case ResourceNotFound
        Get_Description = "Code Not resource."
    Case ResourceLocked
        Get_Description = "Code resource lock error."
    Case CollectionKeyInUse
        Get_Description = "This key is already associated with an element of this Collection."
    Case VariableTypeNotSupported
        Get_Description = "Variable uses a type not supported in Visual Basic."
    Case EventsNotSupported
        Get_Description = "This component doesn't support events."
    Case InvalidClipboardFormat
        Get_Description = "Invalid clipboard format."
    Case InvalidDataFormat
        Get_Description = "Specified format doesn't match format of data."
    Case FailAutodrawImage
        Get_Description = "Can't create AutoRedraw image."
    Case InvalidPicture
        Get_Description = "Invalid picture."
    Case PrinterError
        Get_Description = "Printer error."
    Case PrintDriverNotSupportProperty
        Get_Description = "Printer driver does not support specified property."
    Case PrinterInformationFailure
        Get_Description = "Problem getting printer information from the system. Make sure the printer is set up correctly."
    Case InvalidPictureType
        Get_Description = "Invalid picture type."
    Case PrintImageMismatch
        Get_Description = "Can't print form image to this type of printer."
    Case CanNotSaveToTempDirectory
        Get_Description = "Can't save file to Temp directory."
    Case SearchTextNotFound
        Get_Description = "Search text not found."
    Case ReplacementTooLong
        Get_Description = "Replacements too long."
    Case ClassNoPropertyName
        Get_Description = "Classname does not have propertyname property."
    Case ClassNoMethodName
        Get_Description = "Classname does not have methodname method."
    Case MissingArgument
        Get_Description = "Missing required argument argument name."
    Case RequiredNumberOfArguments
        Get_Description = "Invalid number of arguments."
    Case UnableToSetPropertyName
        Get_Description = "Unable to set the propertyname property of the classname."
    Case UnableToGetPropertyName
        Get_Description = "Unable to get the propertyname property of the classname."
    Case OutOfMemory
        Get_Description = "Out of memory."
    Case NoObject
        Get_Description = "No object."
    Case ClassNotSet
        Get_Description = "Class is not set."
    Case ActivateObjectFailed
        Get_Description = "Unable to activate object."
    Case CreateEmbeddedFailed
        Get_Description = "Unable to create embedded object."
    Case ErrorSavingToFile
        Get_Description = "Error saving to file."
    Case ErrorLoadingFile
        Get_Description = "Error loading from file."
    Case Else
        Get_Description = "Exception Error " & ExceptionType
    End Select

On Error GoTo 0

End Function

Public Function Get_Exception(ExceptionType As eException) As String
'//error name string

On Error Resume Next

    Select Case CInt(ExceptionType)
    Case NoGoSub
        Get_Exception = "Error: 3 No GoSub"
    Case InvalidProceedure
        Get_Exception = "Error: 5 Invalid Proceedure"
    Case Overflow
        Get_Exception = "Error: 6 Overflow"
    Case OutOfMemory
        Get_Exception = "Error: 7 Out of Memory"
    Case SubscriptOutOfRange
        Get_Exception = "Error: 9 Subscript Out of Range"
    Case ArrayLocked
        Get_Exception = "Error: 10 Array is Locked"
    Case DivisionbyZero
        Get_Exception = "Error: 11 Division by Zero"
    Case TypeMismatch
        Get_Exception = "Error: 13 Type Mismatch"
    Case OutOfStringSpace
        Get_Exception = "Error: 14 String too Large"
    Case ExpressionTooComplex
        Get_Exception = "Error: 16 Expression is too Complex"
    Case RequestFailed
        Get_Exception = "Error: 17 Request Failure"
    Case UserInterrupt
        Get_Exception = "Error: 18 User Interrupt"
    Case ResumeWithoutError
        Get_Exception = "Error: 20 Handler Fault"
    Case OutOfStackSpace
        Get_Exception = "Error: 28 Out of Stack Space"
    Case FunctionNotDefined
        Get_Exception = "Error: 35 Invalid Routine Call"
    Case DllOverloading
        Get_Exception = "Error: 47 Dll Request Overload"
    Case ErrorLoadingLibrary
        Get_Exception = "Error: 48 Dll Failed to Load"
    Case BadDllCallingConvention
        Get_Exception = "Error: 49 Dll passed bad Paramater"
    Case InternalError
        Get_Exception = "Error: 51 Internal Error"
    Case BadFileName
        Get_Exception = "Error: 52 Bad File Name"
    Case FileNotFound
        Get_Exception = "Error: 53 File Not Found"
    Case BadFileMode
        Get_Exception = "Error: 54 Bad File Mode"
    Case FileAlreadyOpen
        Get_Exception = "Error: 55 File Already Loaded"
    Case IOError
        Get_Exception = "Error: 57 Storage Device Failure"
    Case FileExists
        Get_Exception = "Error: 58 File Already Exists"
    Case BadRecordLength
        Get_Exception = "Error: 59 Invalid Record Length"
    Case DiskFull
        Get_Exception = "Error: 61 Disk is Full"
    Case InputEndLine
        Get_Exception = "Error: 62 Invalid Input Length"
    Case BadRecordNumber
        Get_Exception = "Error: 63 Invalid Record Number"
    Case TooManyFiles
        Get_Exception = "Error: 67 Too Many Files Open"
    Case DeviceUnavailable
        Get_Exception = "Error: 68 Device is Unavailable"
    Case PermissionDenied
        Get_Exception = "Error: 70 Access is Denied"
    Case DiskNotReady
        Get_Exception = "Error: 71 Disk is not Ready for Access"
    Case RenameFailed
        Get_Exception = "Error: 74 Can not Rename File"
    Case AccessError
        Get_Exception = "Error: 75 File Access Error"
    Case PathNotFound
        Get_Exception = "Error: 76 Invalid Path"
    Case ObjectNotSet
        Get_Exception = "Error: 91 Variable not Initialized"
    Case LoopNotInitialized
        Get_Exception = "Error: 94 For Loop Not Initialized"
    Case InvalidPattern
        Get_Exception = "Error: 95 Custom Error"
    Case InvalidNull
        Get_Exception = "Error: 94 Invalid Use of Null"
    Case CouldNotLoadDll
        Get_Exception = "Error: 298 Could Not Load Library"
    Case InvalidName
        Get_Exception = "Error: 320 Invalid File Name"""
    Case InvalidFileFormat
        Get_Exception = "Error: 321 Invalid Format"
    Case FailToCreateFile
        Get_Exception = "Error: 322 Temporary Resource Allocation Failure"
    Case BadModule
        Get_Exception = "Error: 323 Invalid Module Format"
    Case InvalidResource
        Get_Exception = "Error: 325 Invalid Resource Format"
    Case DataValueNotFound
        Get_Exception = "Error: 327 Data Name Not Found"
    Case IllegalParameter
        Get_Exception = "Error: 328 Illegal Array Paramater"
    Case RegistryAccessFailure
        Get_Exception = "Error: 335 Registry Access Failure"
    Case ComponentNotRegistered
        Get_Exception = "Error: 336 Library Not Registered"
    Case ComponentNotFound
        Get_Exception = "Error: 337 Library Not Found"
    Case ComponentFailure
        Get_Exception = "Error: 338 Library Operation Failure"
    Case ObjectAlreadyLoaded
        Get_Exception = "Error: 360 Object is Already Loaded"
    Case CanNotLoadObject
        Get_Exception = "Error: 361 Object Failed to Load"
    Case ControlNotFound
        Get_Exception = "Error: 363 Library Not Found"
    Case ObjectWasUnloaded
        Get_Exception = "Error: 364 Object was Unloaded"
    Case FailedToUnload
        Get_Exception = "Error: 365 Invalid Calling Context"
    Case FileOutOfDate
        Get_Exception = "Error: 368 File Out of Date"
    Case OwnershipFailure
        Get_Exception = "Error: 371 Invalid Owner Form"
    Case InvalidProperty
        Get_Exception = "Error: 380 Invalid Property Value"
    Case InvalidPropertyArray
        Get_Exception = "Error: 380 Invalid Property Value"
    Case InvalidPropertyIndex
        Get_Exception = "Error: 381 Invalid Property Index"
    Case NotRunTimeProperty
        Get_Exception = "Error: 382 Propert Execution Failure"
    Case PropertyInvalidReadOnly
        Get_Exception = "Error: 383 Property Can Not be Read-Only"
    Case MissingArrayIndex
        Get_Exception = "Error: 385 Missing Property Index"
    Case PropertySetNotPermitted
        Get_Exception = "Error: 387 Property Set Not Allowed"
    Case PropertyGetFailed
        Get_Exception = "Error: 393 Property Get Not Allowed"
    Case PropertyWriteFailed
        Get_Exception = "Error: 394 Property Write-Only"
    Case CanNotShowModal
        Get_Exception = "Error: 400 Form Modal Failure"
    Case CloseTopModal
        Get_Exception = "Error: 402 Form Modal Error"
    Case ObjectPermissionDenied
        Get_Exception = "Error: 419 Object Access Denied"
    Case PropertyNotFound
        Get_Exception = "Error: 422 Invalid Property Reference"
    Case MethodNotFound
        Get_Exception = "Error: 423 Method Not Found"
    Case ObjectRequired
        Get_Exception = "Error: 424 Object Required"
    Case InvalidObjectUse
        Get_Exception = "Error: 425 Invalid Object Use"
    Case CanNotCreateObject
        Get_Exception = "Error: 429 Library Return Failure"
    Case NoOleAutomation
        Get_Exception = "Error: 430 Automation Not Supported"
    Case ClassNotFound
        Get_Exception = "Error: 432 Class Not Found"
    Case ObjectNoSupport
        Get_Exception = "Error: 438 Propert Not Supported"
    Case AutomationError
        Get_Exception = "Error: 440 Aotomation Error"
    Case LibraryConnectionLost
        Get_Exception = "Error: 442 Remote Process Failure"
    Case NoDefaultValue
        Get_Exception = "Error: 443 Object has No Default Value"
    Case ObjectNotSupportAction
        Get_Exception = "Error: 445 Object Action Not Supported"
    Case ObjectNotSupportNamed
        Get_Exception = "Error: 446 Named Arguments Not Supported"
    Case ObjectNotSupportLocale
        Get_Exception = "Error: 447 Local Settings Not Supported"
    Case NamedArgumentNotFound
        Get_Exception = "Error: 448 Missing Arguement"
    Case ArgumentNotOptional
        Get_Exception = "Error: 449 Argument Not Optional"
    Case WrongNumberOfArguments
        Get_Exception = "Error: 450 Wrong Number of Arguments"
    Case ObjectNotCollection
        Get_Exception = "Error: 451 Object Not Collection"
    Case InvalidOrdinal
        Get_Exception = "Error: 452 Invalid Number"
    Case DllFunctionNotFound
        Get_Exception = "Error: 453 Library Not Found"
    Case ResourceNotFound
        Get_Exception = "Error: 454 Invalid Code Resource"
    Case ResourceLocked
        Get_Exception = "Error: 455 Resource Locked"
    Case CollectionKeyInUse
        Get_Exception = "Error: 457 Key In Use"
    Case VariableTypeNotSupported
        Get_Exception = "Error: 458 Type Not Supported"
    Case EventsNotSupported
        Get_Exception = "Error: 459 Event Not Supported"
    Case InvalidClipboardFormat
        Get_Exception = "Error: 460 In valid Clipboard Format"
    Case InvalidDataFormat
        Get_Exception = "Error: 461 Invalid Picture"
    Case FailAutodrawImage
        Get_Exception = "Error: 480 Failed to Draw Picture"
    Case InvalidPicture
        Get_Exception = "Error: 481 Invalid Picture"
    Case PrinterError
        Get_Exception = "Error: 482 Printer Error"
    Case PrintDriverNotSupportProperty
        Get_Exception = "Error: 483 Printer Driver Does Not Support Method"
    Case PrinterInformationFailure
        Get_Exception = "Error: 484 Printer Information Failure"
    Case InvalidPictureType
        Get_Exception = "Error: 485 Invalid Picture Type"
    Case PrintImageMismatch
        Get_Exception = "Error: 486 Can Not Print Image"
    Case CanNotSaveToTempDirectory
        Get_Exception = "Error: 735 Can Not Save File"
    Case SearchTextNotFound
        Get_Exception = "Error: 744 Text Not Found"
    Case ReplacementTooLong
        Get_Exception = "Error: 746 Element Too Long"
    Case ClassNoPropertyName
        Get_Exception = "Error: 1000 Invalid Property Name"
    Case ClassNoMethodName
        Get_Exception = "Error: 1001 Invalid Method Name"
    Case MissingArgument
        Get_Exception = "Error: 1002 Missing Argument"
    Case RequiredNumberOfArguments
        Get_Exception = "Error: 1003 Invalid Number of Arguments"
    Case UnableToSetPropertyName
        Get_Exception = "Error: 1005 Property Set Failure"
    Case UnableToGetPropertyName
        Get_Exception = "Error: 1006 Property Get Failure"
    Case OutOfMemory
        Get_Exception = "Error: 31001 Out of Memory"
    Case NoObject
        Get_Exception = "Error: 31004 No Object"
    Case ClassNotSet
        Get_Exception = "Error: 31018 Class Not Initialized"
    Case ActivateObjectFailed
        Get_Exception = "Error: 31027 Object Failed to Load"
    Case CreateEmbeddedFailed
        Get_Exception = "Error: 31032 Failed to Create Embedded"
    Case ErrorSavingToFile
        Get_Exception = "Error: 31036 Error Saving File"
    Case ErrorLoadingFile
        Get_Exception = "Error: 31037 Error Loading File"
    Case Else
        Get_Exception = "Error: " & ExceptionType
    End Select

On Error GoTo 0

End Function

Public Sub Application_Restart()
'//restart the application

Dim frm     As Form
Dim lHwnd   As Long
Dim sPath   As String

On Error GoTo Handler

    For Each frm In Forms
        Unload frm
    Next frm

Handler:
    sPath = App.Path & Chr$(92) & App.EXEName & ".exe"
    lHwnd = GetDesktopWindow()
    ShellExecute lHwnd, "Open", sPath, vbNullString, vbNullString, SW_NORMAL
    End

End Sub

'//whew, that was a lot of typing.. ~;o)


