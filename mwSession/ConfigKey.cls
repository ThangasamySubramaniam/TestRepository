VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mKeyName As String
Private mKeyValue As String

Public Property Let KeyName(ByVal Data As String)
    mKeyName = UCase(Data)
End Property

Public Property Get KeyName() As String
    KeyName = UCase(mKeyName)
End Property

Public Property Let KeyValue(ByVal Data As String)
    mKeyValue = Data
End Property

Public Property Get KeyValue() As String
   On Error GoTo PropertyError
   KeyValue = mKeyValue
   Exit Property
PropertyError:
   MsgBox "here"
End Property



'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property


'Public Property Get A() As String
'    A = mA
'End Property
