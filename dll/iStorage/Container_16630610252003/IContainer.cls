VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' IContainer - Interface for all the container classes
'--------------------------------------------------------------------------------
Option Explicit

Private Const module As String = "IContainer"   '$ PROBHIDE DEAD

'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Public Property Get Name() As String
    ' Return the name of the container
End Property

Public Function Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the container
End Function

Public Property Get Item(ByVal Key As Long) As Variant  '$ PROBHIDE EXCESS_SCOPE
Attribute Item.VB_UserMemId = 0
    ' Get an item from the container
End Property

Public Property Let Item(ByVal Key As Long, ByVal Data As Variant)
    ' Set an item in the container
End Property

Public Function Lookup(ByVal Key As Long, ByRef Data As Variant) As Boolean
    ' Lookup an item from the container
End Function

Public Function Remove(ByVal Key As Long) As Boolean
    ' Remove an item from the container
End Function

Public Sub Clear() '$ PROBHIDE EXCESS_SCOPE
    ' Remove all items from the container
End Sub

Public Function Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists in the container
End Function

Public Function Count() As Long
    ' Return the item count
End Function

Public Function Keys() As Long()
    ' Return an array of all keys in the container
End Function

Public Function Items() As Variant()
    ' Return an array of all keys in the container
End Function

'$ END
Private Sub Class_Initialize()

End Sub
