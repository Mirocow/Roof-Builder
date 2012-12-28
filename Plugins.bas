Attribute VB_Name = "Plugins"
Option Explicit

Public Type P
    Dll As Object
    About As String
    Pname As String
    ERR As String
    ERRDescription As String
    Ver As String
    guid As String
End Type
Public Plgs() As P

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal szLib As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal szFnc As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal szModule As String) As Long
Private Declare Function LoadTypeLibEx Lib "oleaut32" (ByVal szFile As Long, ByVal REGKIND As Long, pptlib As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (tGuid As Any, ByVal lpszString As String, ByVal lMax As Long) As Long
Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal dlen As Long)

Private Type IUnknown
    QueryInterface          As Long
    AddRef                  As Long
    Release                 As Long
End Type

Private Type IClassFactory
    IUnk                    As IUnknown
    CreateInstance          As Long
    Lock                    As Long
End Type

Private Type ITypeInfo
    IUnk                    As IUnknown
    GetTypeAttr             As Long
    GetTypeComp             As Long
    GetFuncDesc             As Long
    GetVarDesc              As Long
    GetNames                As Long
    GetRefTypeOfImplType    As Long
    GetImplTypeFlags        As Long
    GetIDsOfNames           As Long
    Invoke                  As Long
    GetDocumentation        As Long
    GetDllEntry             As Long
    GetRefTypeInfo          As Long
    AddressOfMember         As Long
    CreateInstance          As Long
    GetMops                 As Long
    GetContainingTypeLib    As Long
    ReleaseTypeAttr         As Long
    ReleaseFuncDesc         As Long
    ReleaseVarDesc          As Long
End Type

Private Type ITypeLib
    IUnk                    As IUnknown
    GetTypeInfoCount        As Long
    GetTypeInfo             As Long
    GetTypeInfoType         As Long
    GetTypeInfoOfGuid       As Long
    GetLibAttr              As Long
    GetTypeComp             As Long
    GetDocumentation        As Long
    IsName                  As Long
    FindName                As Long
    ReleaseTLibAttr         As Long
End Type

Private Type TYPEATTR
    guid(15)                As Byte
    lcid                    As Long
    dwReserved              As Long
    memidConstructor        As Long
    memidDestructor         As Long
    pstrSchema              As Long
    cbSizeInstance          As Long
    TYPEKIND                As Long
    cFuncs                  As Integer
    cVars                   As Integer
    cImplTypes              As Integer
    cbSizeVft               As Integer
    cbAlignment             As Integer
    wTypeFlags              As Integer
    wMajorVerNum            As Integer
    wMinorVerNum            As Integer
    tdescAlias              As Long
    idldescType             As Long
End Type

Private Enum TYPEKIND
    TKIND_ENUM
    TKIND_RECORD
    TKIND_MODULE
    TKIND_INTERFACE
    TKIND_DISPATCH
    TKIND_COCLASS
    TKIND_ALIAS
    TKIND_UNION
    TKIND_MAX
End Enum

Private Enum HRESULT
    S_OK = 0
End Enum

Private Type CoClass
    name                As String
    guid()              As Byte
End Type

Private Type guid
    data1               As Long
    data2               As Integer
    data3               As Integer
    data4(7)            As Byte
End Type

Private Enum REGKIND
    REGKIND_DEFAULT
    REGKIND_REGISTER
    REGKIND_NONE
End Enum

Private Function LoadPlugin(ByVal strLibrary As String, ByVal strClassName As String, udtCoCls() As CoClass) As stdole.IUnknown
On Error GoTo ERR

    Dim newobj              As stdole.IUnknown
    Dim udtCF               As IClassFactory
    
    Dim classid             As guid
    Dim IID_ClassFactory    As guid
    Dim IID_IUnknown        As guid
    Dim lib                 As String
    Dim Obj                 As Long
    Dim vtbl                As Long
    Dim hModule             As Long
    Dim pFunc               As Long
    Dim i                   As Long

    With IID_ClassFactory
        .data1 = &H1
        .data4(0) = &HC0
        .data4(7) = &H46
    End With

    With IID_IUnknown
        .data4(0) = &HC0
        .data4(7) = &H46
    End With

    ' get all CoClasses from the type lib of
    ' the file, and find the GUID of the Prog ID
    If Not GetCoClasses(strLibrary, udtCoCls) Then
        Exit Function
    End If

    For i = 0 To UBound(udtCoCls)
        If StrComp(udtCoCls(i).name, strClassName, vbTextCompare) = 0 Then
            CpyMem classid, udtCoCls(i).guid(0), Len(classid)
            Exit For
        End If
    Next

    If i = UBound(udtCoCls) + 1 Then
        Exit Function
    End If

    ' load the file, if it isn't yet
    hModule = GetModuleHandle(strLibrary)
    If hModule = 0 Then
        hModule = LoadLibrary(strLibrary)
        If hModule = 0 Then Exit Function
    End If

    pFunc = GetProcAddress(hModule, "DllGetClassObject")
    If pFunc = 0 Then Exit Function

    ' call DllGetClassObject to get a
    ' class factory for the class ID
    If 0 <> CallPointer(pFunc, VarPtr(classid), VarPtr(IID_ClassFactory), VarPtr(Obj)) Then
        Exit Function
    End If

    ' IClassFactory VTable
    CpyMem vtbl, ByVal Obj, 4
    CpyMem udtCF, ByVal vtbl, Len(udtCF)

    ' create an instance of the object
    If 0 <> CallPointer(udtCF.CreateInstance, Obj, 0, VarPtr(IID_IUnknown), VarPtr(newobj)) Then
        ' Set IClassFactory = Nothing
        CallPointer udtCF.IUnk.Release, Obj
        Exit Function
    End If

    ' Set IClassFactory = Nothing
    CallPointer udtCF.IUnk.Release, Obj

    Set LoadPlugin = newobj
Exit Function
ERR:
MsgBox ERR.Description
End Function

Private Function GetCoClasses(ByVal strFile As String, udtCoClasses() As CoClass) As Boolean

    Dim hRes            As HRESULT
    Dim udtITypeLib     As ITypeLib
    Dim udtITypeInfo    As ITypeInfo
    Dim udtTypeAttr     As TYPEATTR

    Dim oTypeLib        As Long
    Dim oTypeInfo       As Long
    Dim pVTbl           As Long
    Dim pAttr           As Long

    Dim lngTypeInfos    As Long
    Dim lngCoCls        As Long
    Dim strTypeName     As String

    Dim i               As Long

    ' load the type lib of the file
    hRes = LoadTypeLibEx(StrPtr(strFile), REGKIND_NONE, oTypeLib)
    If hRes <> S_OK Then Exit Function

    ' ITypeLib's VTable
    CpyMem pVTbl, ByVal oTypeLib, 4
    CpyMem udtITypeLib, ByVal pVTbl, Len(udtITypeLib)

    lngTypeInfos = CallPointer(udtITypeLib.GetTypeInfoCount, oTypeLib)

    For i = 0 To lngTypeInfos - 1

        hRes = CallPointer(udtITypeLib.GetTypeInfo, oTypeLib, i, VarPtr(oTypeInfo))

        If hRes <> S_OK Then GoTo NextItem

        ' ITypeInfo's VTable
        CpyMem pVTbl, ByVal oTypeInfo, 4
        CpyMem udtITypeInfo, ByVal pVTbl, Len(udtITypeInfo)

        ' TYPEATTR struct, which describes the type
        CallPointer udtITypeInfo.GetTypeAttr, oTypeInfo, VarPtr(pAttr)
        CpyMem udtTypeAttr, ByVal pAttr, Len(udtTypeAttr)
        CallPointer udtITypeInfo.ReleaseTypeAttr, oTypeInfo, pAttr

        ' name of the type
        CallPointer udtITypeLib.GetDocumentation, oTypeLib, i, VarPtr(strTypeName), 0, 0, 0

        If udtTypeAttr.TYPEKIND = TKIND_COCLASS Then
            ReDim Preserve udtCoClasses(lngCoCls) As CoClass

            With udtCoClasses(lngCoCls)
                .guid = udtTypeAttr.guid
                .name = strTypeName
            End With

            lngCoCls = lngCoCls + 1
        End If

        ' Set ITypeInfo = Nothing
        CallPointer udtITypeInfo.IUnk.Release, oTypeInfo
        
        oTypeInfo = 0

NextItem:
    Next

    ' Set ITypeLib = Nothing
    CallPointer udtITypeLib.IUnk.Release, oTypeLib
    '
    GetCoClasses = True
End Function

Private Function CallPointer(ByVal fnc As Long, ParamArray params()) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    For i = UBound(params) To 0 Step -1
        AddPush pASM, CLng(params(i))   ' PUSH dword
    Next

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProc(VarPtr(btASM(0)), 0, 0, 0, 0)
End Function

Private Sub AddPush(pASM As Long, Lng As Long)
    AddByte pASM, &H68
    AddLong pASM, Lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, Lng As Long)
    CpyMem ByVal pASM, Lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, bt As Byte)
    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub

' http://www.aboutvb.de/khw/artikel/khwcreateguid.htm
Private Function GUID2Str(GUIDBytes() As Byte) As String

    Dim nTemp       As String
    Dim nGUID(15)   As Byte
    Dim nLength     As Long

    nTemp = Space$(78)
    CpyMem nGUID(0), GUIDBytes(0), 16
    nLength = StringFromGUID2(nGUID(0), nTemp, Len(nTemp))
    GUID2Str = Left$(StrConv(nTemp, vbFromUnicode), nLength - 1)
End Function

Sub GetPlugins()

    Dim i As Integer
    Dim plg As Object
    Dim Ver As String, Description As String

    On Error GoTo ERR 'Resume Next
            
    Dim strPlugin As String
    strPlugin = dir(App.Path & "\plugins\*.dll")

        While strPlugin <> ""
 
            Dim udtCoCls() As CoClass
            Set plg = LoadPlugin(App.Path & "\plugins\" & strPlugin, "CalcRoofList", udtCoCls)
            
            MsgBox "OK"
    
            If Not plg Is Nothing Then
                ReDim Preserve Plgs(i)
                Set Plgs(i).Dll = plg
                Set plg = Nothing
                        
                Plgs(i).Pname = strPlugin
                Plgs(i).About = Plgs(i).Dll.About
                Plgs(i).ERRDescription = Plgs(i).Dll.ERRDescription
                '        Plgs(i).GUID = GUID2Str(DllDescription(0).GUID)
        
                If Plgs(i).ERR = "" Then
            
                    Ver = ""
                    Ver = Plgs(i).Dll.RBLibVer
                    If Ver <> "" Then Plgs(i).Ver = Ver
            
                    Plgs(i).Dll.Load_Library Gl.MAXSLOPELISTS, Gl.MAXC, App.ProductName
                    Setup.Combo3.AddItem strPlugin
            
                Else
        
                    Setup.Combo3.AddItem strPlugin
            
                End If
                i = i + 1
            End If
    
            strPlugin = dir()
        Wend

        Exit Sub
ERR:
        Plgs(i).ERR = ERR.Description
'        Resume Next
MsgBox Plgs(i).ERR
End Sub

