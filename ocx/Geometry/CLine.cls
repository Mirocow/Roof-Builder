VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_Key As String

Private m_BeginPoint As CPoint
Private m_EndPoint As CPoint

Public Property Get BeginPoint() As CPoint
    Set BeginPoint = m_BeginPoint
End Property

Public Property Let BeginPoint(NewP As CPoint)
    Set m_BeginPoint = NewP
End Property

Public Property Get EndPoint() As CPoint
    Set EndPoint = m_EndPoint
End Property

Public Property Let EndPoint(NewP As CPoint)
    Set m_EndPoint = NewP
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(NewKey As String)
    m_Key = NewKey
End Property
