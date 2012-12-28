VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarList As Integer 'local copy
Public Property Let List(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.List = 5
    mvarList = vData
End Property


Public Property Get List() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.List
    List = mvarList
End Property



