VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum cDriver
    cFile
    cSerialize
    cIni
    cCVS
End Enum

Private cObj As cStorage

'/****** Работа с Сериализацией *********/

'/****** Работа с INI файлом *********/

'/****** Работа с CVS файлом *********/

'/****** Работа с FILE бинарными файлами ******/

Private mvarSelectDriver As cDriver 'local copy

Public Property Let SelectDriver(ByVal vData As cDriver)
    mvarSelectDriver = vData
    selectDriverObject mvarSelectDriver
End Property


Public Property Get SelectDriver() As cDriver
    SelectDriver = mvarSelectDriver
End Property

Private Sub selectDriverObject(name As cDriver)
    Set cObj = Nothing
    Select Case name
    Case cFile
        Set cObj = New clsFile
    Case cSerialize
        Set cObj = New clsSerialize
    Case cIni
        MsgBox "Don`t use"
    Case cCVS
        MsgBox "Don`t use"
    Else
        MsgBox "Undefined"
    End Select
End Sub

Public Function SetData()
    ' Записать данные в хранилище
    cObj.SetValue Value, Key
End Function

Public Function Complite()
    ' Обработать данные
End Function







