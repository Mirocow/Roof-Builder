VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' clsDirect
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer

' Private variables
Private Const module As String = "clsDirect"
Private mData() As Variant
Private mCount As Long
Private mMax As Long
Private mNextItem As Long


'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Private Sub Class_Initialize()
    ' Clear the list
    mNextItem = 1
    IContainer_Clear
End Sub


Private Sub Class_Terminate()
    ' Erase the list
    Erase mData
End Sub


Private Function FindSlot() As Long
    ' Find an unused slot number
    On Error GoTo Err_Init
    Dim i As Long
    Dim NewMax As Long
    Dim here As Long
    
    ' Default to not found
    FindSlot = -1

    ' Look for empty slot from here to end
    ' (so we don't reuse item numbers too quickly)
    For i = mNextItem To mMax
        If IsEmpty(mData(i)) Then
            FindSlot = i
            Exit For
        End If
    Next i
    
    ' Look for empty slot from 1 to here
    If FindSlot = -1 Then
        here = mNextItem - 1
        If here > mMax Then
            here = mMax
        End If
        For i = 1 To here
            If IsEmpty(mData(i)) Then
                FindSlot = i
                Exit For
            End If
        Next i
    End If
    
    ' Expand array
    If FindSlot = -1 Then
        If mMax < 2 Then
            NewMax = 2
        Else
            NewMax = mMax * 2
        End If
        FindSlot = mMax + 1
        Enlarge NewMax
    End If
    
    mNextItem = FindSlot + 1
                
    Exit Function

Err_Init:
    MsgBox Err.Number & " - " & Err.Description
    Resume Next
End Function


Private Sub Enlarge(ByVal NewMax As Long)
    ' Enlarge capacity of table
    If NewMax > mMax Then
        ReDim Preserve mData(-1 To NewMax)
        mMax = NewMax
    End If
End Sub


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an item to the list
    
    Key = FindSlot
    
    If IsObject(Data) Then
        Set mData(Key) = Data
    Else
        mData(Key) = Data
    End If
    
    mCount = mCount + 1
    IContainer_Add = Key
    
End Function


Private Sub IContainer_Clear()
    ' Remove all items from the list
    mCount = 0
    mMax = 2
    ReDim mData(-1 To mMax)
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of children in the list
    IContainer_Count = mCount
End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not an item exists
    If Key < 1 Or Key > mMax Then
        IContainer_Exists = False
    ElseIf IsEmpty(mData(Key)) Then
        IContainer_Exists = False
    Else
        IContainer_Exists = True
    End If
End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal RHS As Variant)
    ' Set an item
    If Key < 1 Or Key > mMax Then
        'skip it
    Else
        If IsObject(RHS) Then
            Set mData(Key) = RHS
        Else
            mData(Key) = RHS
        End If
    End If
End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    ' Get an item
    If Key < 1 Or Key > mMax Then
        IContainer_Item = Empty
    ElseIf IsEmpty(mData(Key)) Then
        IContainer_Item = Empty
    Else
        If IsObject(mData(Key)) Then
            Set IContainer_Item = mData(Key)
        Else
            IContainer_Item = mData(Key)
        End If
    End If
End Property


Private Function IContainer_Keys() As Long()
    ' Return an array of all the keys in the array
    Dim arr() As Long
    Dim i As Long
    Dim max As Long
    Dim ctr As Long
    
    ReDim arr(-1 To IContainer_Count - 1)
    max = UBound(mData, 1)
    For i = 0 To max
        If Not (IsEmpty(mData(i))) Then
            arr(ctr) = i
            ctr = ctr + 1
        End If
    Next i
    IContainer_Keys = arr
End Function

Private Function IContainer_items() As Variant()
    ' Return an array of all the keys in the array
    Dim arr() As Variant
    Dim i As Long
    Dim max As Long
    Dim ctr As Long
    
    ReDim arr(1 To IContainer_Count)
    max = UBound(mData, 1) - 1
    For i = 1 To max
        If Not (IsEmpty(mData(i))) Then
            ctr = ctr + 1
            arr(ctr) = mData(i)
        End If
    Next i
    IContainer_items = arr
End Function

Private Function IContainer_Lookup(ByVal Key As Long, Data As Variant) As Boolean
    ' Return an item from the list
    If Key < 1 Or Key > mMax Then
        IContainer_Lookup = False
    ElseIf IsEmpty(mData(Key)) Then
        IContainer_Lookup = False
    Else
        IContainer_Lookup = True
        If IsObject(mData(Key)) Then
            Set Data = mData(Key)
        Else
            Data = mData(Key)
        End If
    End If
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of this container
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an item from the list
    If Key < 1 Or Key > mMax Then
        IContainer_Remove = False
    ElseIf IsEmpty(mData(Key)) Then
        IContainer_Remove = False
    Else
        IContainer_Remove = True
        mData(Key) = Empty
        mCount = mCount - 1
    End If
End Function


'$ END

