VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' clsDictionary
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer

Private Const module As String = "clsDictionary"
Private mDict As Dictionary
Private mCtr As Long


Private Sub Class_Initialize()
    ' Create new dictionary
    mCtr = 1
    Set mDict = New Dictionary
End Sub


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the container
    If Key = 0 Then
        Key = mCtr
        mCtr = mCtr + 1
    End If
    
    If mDict.Exists(Key) = True Then
        IContainer_Add = 0
        Exit Function
    End If
    
    mDict.Add Key, Data
    IContainer_Add = Key
End Function


Private Sub IContainer_Clear()
    ' Remove all items from the container
    mDict.RemoveAll
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in the container
    IContainer_Count = mDict.Count
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists in the container
    IContainer_Exists = mDict.Exists(Key)
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal Value As Variant)
    ' Set an item
    mDict(Key) = Value
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get an item
    If mDict.Exists(Key) = True Then
        If IsObject(mDict(Key)) Then
            Set IContainer_Item = mDict(Key)
        Else
            IContainer_Item = mDict(Key)
        End If
    End If
End Property


Private Function IContainer_Keys() As Long()
    ' Return an array of all the keys in the array
    Dim arr() As Long
    Dim i As Long
    Dim max As Long
    Dim v As Variant
    
    v = mDict.Keys
    ReDim arr(-1 To mDict.Count - 1)
    max = UBound(v, 1)
    
    For i = 0 To max
        arr(i) = v(i)
    Next i
    
    IContainer_Keys = arr
End Function


Private Function IContainer_Lookup(ByVal Key As Long, ReturnData As Variant) As Boolean
    ' Lookup an item from the container
    IContainer_Lookup = False
    If mDict.Exists(Key) = False Then
        Exit Function
    End If
    IContainer_Lookup = True
    If IsObject(mDict(Key)) Then
        If IsObject(ReturnData) Then
            Set ReturnData = mDict(Key)
        Else
            MsgBox "The data is an object, and you passed in a non-object variable! The data was not set."
        End If
    Else
        If IsObject(ReturnData) Then
            MsgBox "The data is not an object, and you passed in an object variable! The data was not set."
        Else
            ReturnData = mDict(Key)
        End If
    End If
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of this container
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an item from this container
    If mDict.Exists(Key) Then
        mDict.Remove Key
        IContainer_Remove = True
    Else
        IContainer_Remove = False
    End If
End Function


Private Function IContainer_Test(ByVal table As IContainer) As Boolean
'    ' Test this container
'    Dim Test As IContainer
'    Set Test = New IContainer
'    IContainer_Test = Test.Test(Me)
End Function


Private Function IContainer_GetTimings(ByVal obj As IContainer, ByVal Loops As Long) As clsTimings
    ' Stress test this container
    Dim Test As IContainer
    Set Test = New IContainer
    Set IContainer_GetTimings = Test.GetTimings(Me, Loops)
End Function




