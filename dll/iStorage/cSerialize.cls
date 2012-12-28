VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' &LABEL.ARRAY_COUNT{TYPE.SIZE.NAME,TYPE.SIZE.NAME}(DATA, DATA);
'
' LABEL - ID элемента
' ARRAY_COUNT - размер структуры (рекомендуется)
' TYPE - (обязательный)
' SIZE -
' NAME -
' DATA - Внутри DATA могут быть ссылки на индификаторы (Вложенные структуры)
'
' &32543255.1{i}(1);
' &l5346346.2{i.4.x_coord,i..x_coord}(1,l3534634);
' &l5346346.2{i..x_coord,i..x_coord}(1,l3534634);
' &l5346346.2{i..,i..x_coord}(1,l3534634);
'
' i - integer
' d - bouble
' f - float
' s - string
' подробнее http://188.134.2.31:2121/redmine/documents/4

Private Lables() As String

Private Function GetLabel() As String
    Dim max As Long
    max = UBound(Lables, 1) + 1
    ReDim Preserve Lables(-1 To max)
    Lables(Key) = Time(Now)
    GetLabel = "&" & Lables(Key)
End Function

Private Function Time(ByVal DateTime As String) As String
'Time = DateDiff("s", "01/1/1970 12:00:00 AM", DateTime)
Time = (Now - DateSerial(1970, 1, 1)) * 86400
End Function
