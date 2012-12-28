VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
' clsArray
'--------------------------------------------------------------------------------
Option Explicit

Implements IContainer

' Name/value pair
Private Type typeNameValue
    Key  As Long
    Data As Variant
End Type


' Private variables
Private Const module As String = "clsArray"
Private mData() As typeNameValue
Private Const mHighValues As Long = 2147483647
Private mCtr As Long


'$ PROBHIDE NO_ERROR_HANDLER BEGIN

Private Sub Class_Initialize()
    ' Initialize the array
    mCtr = 1
    IContainer_Clear
End Sub


Private Sub Class_Terminate()
    ' Erase the array
    Erase mData
End Sub


Private Sub SetSize(ByVal NewSize As Long)
    ' Set the size of a table
    Dim max As Long
    Dim i As Long

    max = UBound(mData, 1)
    If NewSize > max Then
        ReDim Preserve mData(-1 To NewSize)
        For i = max + 1 To NewSize
            mData(i).Key = mHighValues
            mData(i).Data = Null
        Next i
    End If
End Sub


Private Function FindSlot(ByVal Key As Long, ByVal FindExisting As Boolean) As Long
    ' Find either a key or an empty slot
    Dim i As Long
    Dim max As Long
    Dim NewMax As Long
    Dim here As Long
    Static nextitem As Long
    
    FindSlot = -1

    ' Exit early if an invalid key was passed
    If Key = mHighValues Then
        Exit Function
    End If
    
    ' Determine the upper bound of the array
    max = UBound(mData, 1)
    
    ' Look for existing key
    If FindExisting = True Then
        For i = 0 To max
            If mData(i).Key = Key Then
                ' Found a match
                FindSlot = i
                Exit Function
            End If
        Next i
        ' Didn't find a match
        Exit Function
    End If
        
    ' Look for empty slot from here to end
    For i = nextitem To max
        If mData(i).Key = mHighValues Then
            FindSlot = i
            Exit For
        End If
    Next i
    
    ' Look for empty slot from 0 to here
    If FindSlot = -1 Then
        here = nextitem - 1
        If here > max Then
            here = max
        End If
        For i = 0 To here
            If mData(i).Key = mHighValues Then
                FindSlot = i
                Exit For
            End If
        Next i
    End If
    
    ' Expand array
    If FindSlot = -1 Then
        If max < 2 Then
            NewMax = 2
        Else
            NewMax = max * 2
        End If
        SetSize NewMax
        i = max + 1
        FindSlot = i
    End If
        
    nextitem = i + 1
End Function


Private Function IContainer_Add(ByVal Data As Variant, ByVal Key As Long) As Long
    ' Add an entry to the array
    
    Dim i As Long
    
    If Key = 0 Then
        Key = mCtr
        mCtr = mCtr + 1
    End If
        
    i = FindSlot(Key, True)
    If i >= 0 Then
        IContainer_Add = 0
        Exit Function
    End If
        
    ' Add to empty slot
    i = FindSlot(Key, False)
    mData(i).Key = Key
    If IsObject(Data) Then
        Set mData(i).Data = Data
    Else
        mData(i).Data = Data
    End If
    IContainer_Add = Key
    
End Function


Private Sub IContainer_Clear()
    ' Empty out the array
    ReDim mData(-1 To -1)
End Sub


Private Function IContainer_Count() As Long
    ' Return the number of items in the array
    Dim i As Long
    Dim max As Long
    Dim ctr As Long
    max = UBound(mData, 1)
    For i = 0 To max
        If Not (mData(i).Key = mHighValues) Then
            ctr = ctr + 1
        End If
    Next i
    IContainer_Count = ctr

End Function


Private Function IContainer_Exists(ByVal Key As Long) As Boolean
    ' Return whether or not a key exists
    Dim i As Long
    
    ' Find key
    i = FindSlot(Key, True)
    If i < 0 Then
        ' Key doesn't exist
        IContainer_Exists = False
    Else
        ' Key exists
        IContainer_Exists = True
    End If

End Function


Private Property Let IContainer_Item(ByVal Key As Long, ByVal Data As Variant)
    ' Set an item
    Dim i As Long
    
    ' Look for existing entry first
    i = FindSlot(Key, True)
    
    ' If not found, find a new slot
    If i < 0 Then
        i = FindSlot(Key, False)
    End If
    
    ' Set the key and data values
    mData(i).Key = Key
    If IsObject(Data) Then
        Set mData(i).Data = Data
    Else
        mData(i).Data = Data
    End If

End Property


Private Property Get IContainer_Item(ByVal Key As Long) As Variant
    'Retrieve an entry from the array
    Dim i As Long

    ' Find key
    i = FindSlot(Key, True)
    If i < 0 Then
        ' Key not found
        Exit Property
    End If
    ' Return entry
    If IsObject(mData(i).Data) Then
        Set IContainer_Item = mData(i).Data
    Else
        IContainer_Item = mData(i).Data
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
        If Not (mData(i).Key = mHighValues) Then
            arr(ctr) = mData(i).Key
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
        If Not (mData(i).Key = mHighValues) Then
             ctr = ctr + 1
             arr(ctr) = mData(i).Data
        End If
    Next i
    IContainer_items = arr
End Function

Private Function IContainer_Lookup(ByVal Key As Long, ReturnData As Variant) As Boolean
    'Retrieve an entry from the array
    Dim i As Long
    On Error GoTo Err_Init

    IContainer_Lookup = False
    
    ' Find key
    i = FindSlot(Key, True)
    
    If i >= 0 Then
        ' Return entry
        IContainer_Lookup = True
        If IsObject(mData(i).Data) Then
            If IsObject(ReturnData) Then
                Set ReturnData = mData(i).Data
            Else
                MsgBox "The data is an object, and you passed in a non-object variable! The data was not set."
            End If
        Else
            If IsObject(ReturnData) Then
                MsgBox "The data is not an object, and you passed in an object variable! The data was not set."
            Else
                ReturnData = mData(i).Data
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.Number <> 438 Then
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function


Private Property Get IContainer_Name() As String
    ' Return the name of this module
    IContainer_Name = module
End Property


Private Function IContainer_Remove(ByVal Key As Long) As Boolean
    ' Remove an entry from the array
    Dim i As Long
    
    ' Find key
    i = FindSlot(Key, True)
    
    If i < 0 Then
        ' Entry not found
        IContainer_Remove = False
    Else
        ' Delete entry
        IContainer_Remove = True
        mData(i).Key = mHighValues
        mData(i).Data = Null
    End If
    
End Function


'$ END
