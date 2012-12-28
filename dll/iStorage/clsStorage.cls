VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum tStorage
    clsCollection
    clsArray
    clsHash
    clsDirect
End Enum

Public Enum tType
    iInteger
    iong
    iSingle
    iDouble
    iCurrency
    iString
    iBoolean
    iByte
    iDate
    iEmpty
    iObject
End Enum

Private Storage As IContainer
Private mDriver As tStorage
Private mFileName As String

Public Property Let FileName(ByVal vData As String)
    mFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let Driver(ByVal vData As tStorage)
    'Set Storage = vData
    Select Case vData
    Case clsCollection
        Set Storage = New clsCollection
    Case clsArray
        Set Storage = New clsArray
    Case clsHash
        Set Storage = New clsHash
    Case clsDirect
        Set Storage = New clsDirect
    End Select
    mDriver = vData
End Property

Public Property Get Driver() As tStorage
    Driver = mDriver
End Property

Private Sub Class_Initialize()
Set Storage = Nothing
End Sub

Private Sub Class_Terminate()
Set Storage = Nothing
End Sub

Public Function Add(ByVal data As Variant, ByVal Key As Long) As Long
    Add = Storage.Add(data, Key)
End Function

Public Sub Clear()
    Storage.Clear
End Sub

Public Function Count() As Long
    Count = Storage.Count
End Function

Public Function Exists(ByVal Key As Long) As Boolean
    Exists = Storage.Exists(Key)
End Function

Public Property Let Item(ByVal Key As Long, ByVal RHS As Variant)
    Storage.Item(Key) = RHS
End Property

Public Property Get Item(ByVal Key As Long) As Variant
    Item = Storage.Item(Key)
End Property

Public Property Get GetType(ByVal Key As Long) As tType
    GetType = VarType(Storage.Item(Key))
End Property

Public Function Keys() As Long()
    Keys = Storage.Keys
End Function

Public Function Items() As Variant()
    Items = Storage.Items
End Function

Public Function Lookup(ByVal Key As Long, data As Variant) As Boolean
    Lookup = Storage.Lookup(Key, data)
End Function

Public Property Get Name() As String
    Name = Storage.Name
End Property

Public Function Remove(ByVal Key As Long) As Boolean
    Remove = Storage.Remove(Key)
End Function

Public Function Save() As Boolean
    Save = SaveDataIntoFile(Storage, Me.FileName)
End Function

Public Function Read() As Boolean
    Read = LoadDataFromFile(Storage, Me.FileName)
End Function

Public Function Del() As Boolean
'    Del = Kill(Me.FileName)
End Function
