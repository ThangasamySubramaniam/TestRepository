VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwRoleType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetRoleTypeName(RoleTypeKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select RoleTypeName from mwcRoleType where ID=" & RoleTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetRoleTypeName = ""
   Else
      GetRoleTypeName = loRs!RoleTypeName
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwRoleType.GetRoleTypeName.", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   GetRoleTypeName = ""
End Function

Public Function GetRoleTypeEmailAddress(RoleTypeKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT RoleTypeEmailAddress from mwcRoleType where ID=" & RoleTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetRoleTypeEmailAddress = ""
   Else
      GetRoleTypeEmailAddress = BlankNull(loRs!RoleTypeEmailAddress)
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwRoleType.GetRoleTypeEmailAddress.", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   GetRoleTypeEmailAddress = ""
End Function

