VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' clsRecordset
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer
Private Const module As String = "clsRecordset"
Private mCtr As Long
Private rs As ADOR.Recordset


Private Sub Class_Initialize()
    ' Create a new recordset
    mCtr = 1
    Set rs = New ADOR.Recordset
    With rs
       .Fields.Append "ID", adInteger
       .Fields.Append "Data", adVariant, , adFldIsNullable
       .CursorType = adOpenDynamic
       .LockType = adLockOptimistic
       .Open
    End With
End Sub


Private Sub Class_Terminate()
    ' Close the recordset
    rs.Close
    Set rs = Nothing
End Sub


Private Function GotoRecord(ByVal Key As Long) As Boolean
    ' Hop to the specified record
    On Error GoTo Err_Init
    rs.MoveFirst
    rs.Find "ID = " & Key
    If rs.EOF Then
        GotoRecord = False
    Else
        GotoRecord = True
    End If
    Exit Function

Err_Init:
    If Err.Number = 3021 Then
        Exit Function
    Else
        MsgBox Err.Number & " - " & Err.Description
        Resume Next
    End If
End Function


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to this container
    On Error GoTo Err_Init
    IContainer_Add = 0
    If Key = 0 Then
        ' Assign a unique ID
        Key = mCtr
        mCtr = mCtr + 1
    End If
    If GotoRecord(Key) = True Then
        Exit Function
    End If
    rs.AddNew
    rs!ID = Key
    rs!Data = Data
    rs.Update
    IContainer_Add = Key
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Function


Private Sub IContainer_Clear()
    ' Clear all items from this container
    rs.MoveFirst
    Do While Not rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in this container
    IContainer_Count = rs.RecordCount
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists
    If GotoRecord(Key) = True Then
        IContainer_Exists = True
    Else
        IContainer_Exists = False
    End If
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal RHS As Variant)
    ' Set an item
    If GotoRecord(Key) = False Then
        rs.AddNew
        rs!ID = Key
    End If
    rs!Data = RHS
    rs.Update
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get an item
    If GotoRecord(Key) = False Then
        Exit Property
    End If
    IContainer_Item = rs!Data
End Property


Private Function IContainer_Keys() As Long()
    ' Return an array of all the keys in this container
    On Error GoTo Err_Init
    Dim arr() As Long
    Dim i As Long
    ReDim arr(-1 To rs.RecordCount - 1)
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        arr(i) = rs!ID
        rs.MoveNext
        i = i + 1
    Loop
    IContainer_Keys = arr
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Function


Private Function IContainer_Lookup(ByVal Key As Long, ByRef Data As Variant) As Boolean
    ' Lookup an item from this container
    On Error GoTo Err_Init
    IContainer_Lookup = True
    If GotoRecord(Key) = False Then
        IContainer_Lookup = False
        Exit Function
    End If
    If IsObject(Data) Then
        Set Data = rs!Data.OriginalValue
    Else
        Data = rs!Data.OriginalValue
    End If
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of this container
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an item from this container
    IContainer_Remove = False
    If GotoRecord(Key) = False Then
        Exit Function
    End If
    rs.Delete
    IContainer_Remove = True
End Function


Private Function IContainer_Test(ByVal table As IContainer) As Boolean
    ' Test this container
'    Dim Test As IContainer
'    Set Test = New IContainer
'    IContainer_Test = Test.Test(Me)
End Function


Private Function IContainer_GetTimings(ByVal obj As IContainer, ByVal Loops As Long) As clsTimings
'     Stress test this container
    Dim Test As IContainer
    Set Test = New IContainer
    Set IContainer_GetTimings = Test.GetTimings(Me, Loops)
End Function




