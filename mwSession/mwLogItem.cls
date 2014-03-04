VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwLogItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' LogItem -
' 11/28/00 ms

Option Explicit

Private mMessageType As mwl_LogMessageType
Private mCriticality As mwl_LogMessageCriticality
Private mDateTime As Date
Private mMessage As String
Private moError As mwError
Private mIsLoggedToFile As Boolean




Public Property Set Error(ByVal vData As mwError)
   If moError Is Nothing Then
      Set moError = New mwError
   End If
   Set moError = vData
End Property


Public Property Get Error() As mwError
   If moError Is Nothing Then
      Set moError = New mwError
   End If
   Set Error = moError
End Property

Public Property Get MessageType() As mwl_LogMessageType
    MessageType = mMessageType
End Property

Public Property Let MessageType(ByVal Data As mwl_LogMessageType)
    mMessageType = Data
End Property

Public Property Get DateTime() As Date
    DateTime = mDateTime
End Property

Public Property Let Message(ByVal Data As String)
    mMessage = Data
End Property

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let IsLoggedToFile(ByVal Data As Boolean)
    mIsLoggedToFile = Data
End Property

Public Property Get IsLoggedToFile() As Boolean
    IsLoggedToFile = mIsLoggedToFile
End Property

Public Property Let Criticality(ByVal Data As mwl_LogMessageCriticality)
    mCriticality = Data
End Property

Public Property Get Criticality() As mwl_LogMessageCriticality
    Criticality = mCriticality
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


Private Sub Class_Initialize()
   mDateTime = Now()
End Sub
