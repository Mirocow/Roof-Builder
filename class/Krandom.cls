VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Krandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_holdrand As Long
Private rnd32 As Long

Public Property Let holdrand(ByVal value As Long)
m_holdrand = value
End Property

Public Property Get holdrand() As Long
holdrand = m_holdrand
End Property

Public Property Get GetRnd() As Long
GetRnd = rnd32
End Property

Function rand16() As Long
holdrand = holdrand * 2131 + 25011
'holdrand = ShiftToRight(holdrand, 16)
holdrand = holdrand And &H7FFF
rand16 = holdrand
End Function

Function rand32() As Long
Dim low As Long
Dim high As Long
low = rand16()
high = rand16()
rnd32 = ShiftToLeft(high, 16) Or low
rand32 = rnd32
End Function

Function ShiftToRight(value As Long, Shift As Integer) As Long
Dim i As Integer
For i = 1 To Shift
value = value / 2
Next
ShiftToRight = value
End Function

Function ShiftToLeft(value As Long, Shift As Integer) As Long
Dim i As Integer
For i = 1 To Shift
value = value * 2
Next
ShiftToLeft = value
End Function
