VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================
'= Name:        CCollectionEx
'= Programmer:  Andrew Davey
'= Purpose:     Improve built in collection object and allow implementation.
'===================================================================================
Option Explicit

Public Items As Collection
'#Const isDebug = True   ' Set to false on release build.

Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set Items = Nothing
End Sub

Public Sub Add(newObj As Object, Optional Index As Variant, Optional Before As Variant, Optional After As Variant)
    Dim bAfter As Boolean, bBefore As Boolean
    
    bAfter = Not IsMissing(After)
    bBefore = Not IsMissing(Before)
    
    If IsMissing(Index) Then
        If bAfter Then
            Items.Add newObj, , After
            Exit Sub ' Just in case bBefore = True also.
        ElseIf bBefore Then
            Items.Add newObj, , Before
        Else
            Items.Add newObj
        End If
    Else
        If bAfter Then
            Items.Add newObj, CStr(Index), After
            Exit Sub ' Just in case bBefore = True also.
        ElseIf bBefore Then
            Items.Add newObj, CStr(Index), Before
        Else
            Items.Add newObj, CStr(Index)
        End If
    End If
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error Resume Next
    Items.Remove CStr(Index)
'    If ERR.Number Then displayError
End Sub

Public Sub Clear()
    If Items Is Nothing Then Exit Sub    ' Shouldn't happen, but just in case.
    
    Dim i As Integer
    For i = Items.Count To 1 Step -1
        Items.Remove i
    Next i
End Sub

Public Property Get Item(ByVal Index As Variant) As Object
    On Error Resume Next
    Set Item = Items.Item(CStr(Index))
    If ERR.Number Then Set Item = Nothing
End Property

'Public Function NewEnum() As IUnknown
'    ' Allows the nifty For..Each structure to work.
'    NewEnum = Items.[_NewEnum]   ' Calls the hidden function.
'End Function

Public Property Get Count() As Long
    If Items Is Nothing Then Exit Property    ' Shouldn't happen, but just in case.

    Count = Items.Count
End Property


Public Function Exists(ByVal vKey As Variant) As Boolean

    On Error Resume Next
    Items.Item vKey
    Exists = (ERR.Number = 0)
    ERR.Clear

End Function


