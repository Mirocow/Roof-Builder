VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' clsCollection
'--------------------------------------------------------------------------------
Option Explicit


' Private variables
Implements IContainer

Private Const module As String = "clsCollection"
Private mData As Collection
Private mCtr As Long
Private mKeys() As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Sub Class_Initialize()
    ' Initialize the collection
    IContainer_Clear
    ReDim mKeys(-1 To -1)
    mCtr = 1
End Sub


Private Sub Class_Terminate()
    ' Terminate the collection
    Set mData = Nothing
    Erase mKeys
End Sub


Private Sub AddKey(ByVal Key As Long)
    ' Add a key to the list of keys
    Dim max As Long
    max = UBound(mKeys, 1) + 1
    ReDim Preserve mKeys(-1 To max)
    mKeys(max) = Key
End Sub


Private Sub RemoveKey(ByVal Key As Long)
    ' Remove a key from the list of keys
    Dim i As Long
    Dim max As Long
    Dim Remaining As Long
    max = UBound(mKeys, 1)
    For i = 0 To max
        If mKeys(i) = Key Then
            'Remove this key
            Remaining = max - i + 1
            If Remaining > 1 Then
                CopyMemory mKeys(i), mKeys(i + 1), 4 * Remaining
            End If
            ReDim Preserve mKeys(-1 To max - 1)
            Exit For
        End If
    Next i
    Exit Sub
End Sub


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the collection
    On Error GoTo Err_Init
    
    IContainer_Add = 0
    If Key = 0 Then
        ' Assign a key
        Key = mCtr
        mCtr = mCtr + 1
    End If
    mData.Add Data, CStr(Key)
    AddKey Key
    IContainer_Add = Key
    
    Exit Function
    
Err_Init:
    If Err.Number <> 457 Then
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function


Private Sub IContainer_Clear()
    ' Empty out all members
    Set mData = New Collection
    ReDim mKeys(-1 To -1)
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in the array
    IContainer_Count = mData.Count
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not a key exists
    On Error GoTo Err_Init
    Dim v As Variant
    
    IContainer_Exists = False
    
    v = mData(CStr(Key))
    IContainer_Exists = True
    
    Exit Function

Err_Init:
    If Err.Number = 438 Or Err.Number = 450 Then
        IContainer_Exists = True
        Exit Function
    ElseIf Err.Number = 5 Then
        Exit Function
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal Data As Variant)
    ' Set a new value for the item
    mData.Remove CStr(Key)
    mData.Add Data, CStr(Key)
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get the value for the item
    On Error GoTo Err_Init
    IContainer_Item = mData(CStr(Key))
    Exit Property
Err_Init:
    If Err.Number = 5 Then
        ' Item not found
    ElseIf Err.Number = 438 Or Err.Number = 450 Then
        ' It's an object
        Set IContainer_Item = mData(CStr(Key))
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Property


Private Function IContainer_Keys() As Long()
    ' Return an array of all the keys
    IContainer_Keys = mKeys
End Function

Private Function IContainer_items() As Variant()
    ' Return an array of all the keys
'    IContainer_Keys = mKeys
End Function

Private Function IContainer_Lookup(ByVal Key As Long, Data As Variant) As Boolean
    ' Return an item
    On Error GoTo Err_Init
    Dim v As Variant
    IContainer_Lookup = False
    v = mData.Item(CStr(Key))
    
    If IsObject(v) Then
        If IsObject(Data) Then
            Set Data = v
        Else
            MsgBox "The data is an object, and you passed in a non-object variable! The data was not set."
        End If
    Else
        If IsObject(Data) Then
            MsgBox "The data is not an object, and you passed in an object variable! The data was not set."
        Else
            Data = v
        End If
    End If
    IContainer_Lookup = True
    
    Exit Function
    
Err_Init:
    If Err.Number = 5 Then
        Exit Function
    ElseIf Err.Number = 438 Or Err.Number = 450 Then
        Set v = mData.Item(CStr(Key))
        Resume Next
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of this module
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an entry from the collection
    mData.Remove CStr(Key)
    If Err.Number <> 0 Then
        IContainer_Remove = False
    Else
        IContainer_Remove = True
        RemoveKey Key
    End If
End Function



