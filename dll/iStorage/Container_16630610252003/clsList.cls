VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' clsList
' An 'unsafe' array optimized for speed.
' It assumes you know what you're doing, and won't add duplicate keys to the list.
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer

' Private declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Private variables
Private Const module As String = "clsList"
Private mData() As Long
Private mMax As Long
Private mCnt As Long


'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Private Sub Class_Initialize()
    ' Clear the container
    IContainer_Clear
End Sub


Private Sub Class_Terminate()
    ' Erase the container
    Erase mData
End Sub


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the  container
    If Key < 1 Then
        IContainer_Add = 0
        Exit Function
    End If
    If mMax <= mCnt Then
        mMax = mCnt * 2
        ReDim Preserve mData(-1 To mMax)
    End If
    mData(mCnt) = Key
    mCnt = mCnt + 1
    IContainer_Add = Key
End Function


Private Sub IContainer_Clear()
    ' Remove all items from the container
    mCnt = 0
    mMax = 2
    ReDim mData(-1 To mMax)
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in the container
    IContainer_Count = mCnt
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists
    IContainer_Exists = IContainer_Lookup(Key, Nothing)
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal RHS As Variant)
    ' Set an item in the container
    If IContainer_Lookup(Key, Nothing) = True Then
        IContainer_Remove Key
    End If
    IContainer_Add RHS, RHS
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get an item from the container
    IContainer_Item = IContainer_Lookup(Key, Nothing)
End Property


Private Function IContainer_Keys() As Long()
    ' Return a copy of the list
    Dim Copy() As Long
    ReDim Copy(-1 To mCnt - 1)
    If mCnt > 0 Then
        CopyMemory Copy(0), mData(0), 4 * mCnt
    End If
    IContainer_Keys = Copy
End Function


Private Function IContainer_Lookup(ByVal Key As Long, Data As Variant) As Boolean
    ' Return whether or not a child is in the list
    Dim i As Long
    Dim max As Long
    
    IContainer_Lookup = False
    max = UBound(mData, 1)
    For i = 0 To max
        If mData(i) = Key Then
            IContainer_Lookup = True
            Exit Function
        End If
    Next i
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of the container
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove a child from the list
    Dim i As Long

    ' Default to false
    IContainer_Remove = False

    ' Leave early if there are no elements in the array
    If mCnt = 0 Then
        Exit Function
    End If

    ' Search the array
    For i = 0 To mCnt - 1
        If mData(i) = Key Then
            If i < mCnt - 1 Then
                'Remove it and shift the rest up
                CopyMemory mData(i), mData(i + 1), 4 * (mCnt - i - 1)
                mData(mCnt - 1) = 0
            Else
                mData(i) = 0
            End If
            IContainer_Remove = True
            mCnt = mCnt - 1
            Exit For
        ElseIf mData(i) = 0 Then
            ' We've run out of numbers
            Exit For
        End If
    Next i

End Function


'$ END
