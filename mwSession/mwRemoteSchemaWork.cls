VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwRemoteSchemaWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwRemoteSchemaWork - Utility class to build and maintain
' 25 Feb 2003 ms
Option Explicit

Private mIssueDate As Date
Private mDescription As String
Private mConnectionID As String
Private mSourceID As String
Private mSendNotification As Boolean
Private mChangeID As Long
Private mReference As String
Private mDescription As String
Private mBuildReference As String
Private moSchemaLines As Collection


Public Function MakeNewSchemaUpdate(SourceID As String) As Boolean




   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.MakeNewForm.", err.Number, err.Description
End Function

