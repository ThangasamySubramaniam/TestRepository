VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwError Class - Object contains information about an error that has occurred
' 11/24/2000 ms Copyright 2000 Maritime Systems Inc

Option Explicit

Private mmwErrorCode As mwe_ErrorCodes
Private mmwErrorDescription As String
Private mmwErrorSource As String
Private mVbErrorNumber As String
Private mVbErrorDescription As String
Private mVbErrorSource As String
Private mErrorDateTime As Date

Public Enum mwe_ErrorCodes
    mwe_NO_ERROR = 0
    mwe_GENERAL_ERROR = 1
End Enum





Public Property Let mwErrorCode(ByVal Data As mwe_ErrorCodes)
    mmwErrorCode = Data
End Property

Public Property Get mwErrorCode() As mwe_ErrorCodes
    mwErrorCode = mmwErrorCode
End Property

Public Property Let mwErrorDescription(ByVal Data As String)
    mmwErrorDescription = Data
End Property

Public Property Get mwErrorDescription() As String
    mwErrorDescription = mmwErrorDescription
End Property

Public Property Let mwErrorSource(ByVal Data As String)
    mmwErrorSource = Data
End Property

Public Property Get mwErrorSource() As String
    mwErrorSource = mmwErrorSource
End Property

Public Property Let VbErrorNumber(ByVal Data As String)
    mVbErrorNumber = Data
End Property

Public Property Get VbErrorNumber() As String
    VbErrorNumber = mVbErrorNumber
End Property

Public Property Let VbErrorSource(ByVal Data As String)
    mVbErrorSource = Data
End Property

Public Property Get VbErrorSource() As String
    VbErrorSource = mVbErrorSource
End Property

Public Property Get ErrorDateTime() As String
    ErrorDateTime = mErrorDateTime
End Property

Public Property Let VbErrorDescription(ByVal Data As String)
    mVbErrorDescription = Data
End Property

Public Property Get VbErrorDescription() As String
    VbErrorDescription = mVbErrorDescription
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

Private Sub Class_Initialize()
   mErrorDateTime = Now()
   mVbErrorNumber = err.Number
   mVbErrorDescription = err.Description
   mVbErrorSource = err.Source
End Sub
