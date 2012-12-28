VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'*********************************************************
'A simple class to hold a point
'*********************************************************
Private m_X As Single
Private m_Y As Single
Private m_Key As String
Private m_SomeObject As Object
Private m_Children As Integer
Private m_isPoint As Boolean

'-------- Объект для хранения данных

Public Property Get Obj() As Object
    Set Obj = m_SomeObject
End Property

Public Property Let Obj(newObj As Object)
   Set m_SomeObject = newObj
End Property

'------- Координаты

Public Property Get X() As Single
    X = m_X
End Property

Public Property Let X(newX As Single)
    m_X = newX
End Property

Public Property Get Y() As Single
    Y = m_Y
End Property

Public Property Let Y(newY As Single)
    m_Y = newY
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(newv As String)
    m_Key = newv
End Property

Public Property Get Children() As Integer
    Children = m_Children
End Property

Public Property Let Children(newv As Integer)
    m_Children = newv
End Property

Public Property Get isPoint() As Boolean
    isPoint = m_isPoint
End Property

Public Property Let isPoint(newv As Boolean)
    m_isPoint = newv
End Property
