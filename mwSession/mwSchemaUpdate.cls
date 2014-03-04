VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwSchemaUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwSchemaUpdate
' 8 March 2003 ms
Option Explicit

Private Const DB_SCHEMA_SQL_SECTION_START_OFFSET = 2
Private Const DB_SCHEMA_SECTION_SQL_START_OFFSET = 5


Public Function RemoteSchemaUpdater() As Boolean
   ' FSO objects
   Dim fso As FileSystemObject
   Dim loFolder As Folder
   Dim loFile As file
   ' ConfigGroups stuff
   Dim loSchema As ConfigGroups
   Dim loGroup As ConfigGroup
   Dim loKey As ConfigKey
   ' Database
   Dim loCon As ADODB.Connection
   Dim loRsDB As Recordset
   Dim loRsSU As Recordset
   ' State variables
   Dim IsScriptError As Boolean
   Dim IsError As Boolean
   ' Local Variables
   Dim strFile As String
   Dim IsMwConnection As Boolean
   Dim strMoveFile As String
   Dim strConnectString As String
   Dim strDescription As String
   Dim strChangeDescription As String
   Dim strSendNotification As String
   Dim strSQL As String
   Dim strMsg As String
   Dim strConnectionID As String
   Dim strSourceID As String
   Dim strChangeID As String
   Dim strBuildReference As String
   Dim strReference As String
   Dim strError As String
   Dim i As Integer
   Dim j As Integer
   '
   On Error GoTo FunctionError
   '
   ' Check for db_update*.dat file
   '
   'Table of database connections...
   Set loRsDB = New Recordset
   loRsDB.CursorLocation = adUseClient
   loRsDB.Open "mwcDbConnection", goSession.DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdTable
   
   'Table of database connections...
   Set loRsSU = New Recordset
   loRsSU.CursorLocation = adUseClient
   loRsSU.Open "mwcSchemaUpdate", goSession.DBConnection, adOpenDynamic, adLockOptimistic, adCmdTable
   
   
   '
   ' Folder collection...looking for files to process...
   '
   Set fso = New FileSystemObject
   Set loFolder = fso.GetFolder(goSession.Workflow.PollingFolder)
   For Each loFile In loFolder.Files
      If UCase(fso.GetExtensionName(loFile.Path)) = "SCHEMA" Or UCase(fso.GetExtensionName(loFile.Path)) = "SCH" Then
         '
         ' State Flags
         '
         IsError = False
         IsScriptError = False
         RemoteSchemaUpdater = True
         '
         ' Have a schema file to process...
         '
         Set loSchema = goSession.LoadConfigGroupsFile(loFile.Path)
         If loSchema Is Nothing Then
            goSession.RaiseError "Error in mwSession.RemoteSchemaUpdater, Bad Schema File Received:" & loFile.Name
            strMoveFile = loFile.Path & "." & Format(Now(), "yyyymmddhhmmss")
            fso.MoveFile loFile.Path, strMoveFile
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
            GoTo NextFile
         End If
         '
         ' Get Header Information
         '
         strConnectionID = Trim(UCase(loSchema("Header").ConfigKeys.GetKeyValue("ConnectionID")))
         strSendNotification = UCase(loSchema("Header").ConfigKeys.GetKeyValue("SendNotification"))
         strDescription = UCase(loSchema("Header").ConfigKeys.GetKeyValue("Description"))
         strSourceID = loSchema("Header").ConfigKeys.GetKeyValue("SourceID")
         If Trim(strSourceID) <> "" Then
            ' Validate SourceID
            If Not IsValidateSource(strSourceID) Then
               goSession.RaiseError "Error in mwSession.RemoteSchemaUpdater, Invalid Source ID: " & strSourceID & " , in File:" & loFile.Name
               strMoveFile = loFile.Path & "." & Format(Now(), "yyyymmddhhmmss")
               fso.MoveFile loFile.Path, strMoveFile
               goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
               GoTo NextFile
            End If
         End If
         If Trim(strConnectionID) = "" Then
            strConnectionID = "MW"
         End If
         loRsDB.Filter = "ConnectionID='" & strConnectionID & "'"
         If loRsDB.RecordCount < 1 Then
            strError = "Error in mwSession.RemoteSchemaUpdater, ConnectionID missing in database. ConnectionID: " & strConnectionID & ", File:" & loFile.Name
            MsgBox strError
            goSession.Logger.LogIt mwl_Workstation, mwl_Critical, strError
            strMoveFile = strFile & "." & Format(Now(), "yyyymmddhhmmss")
            fso.MoveFile strFile, strMoveFile
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
            GoTo NextFile
         End If
         If strConnectionID = "MW" Or IsNull(loRsDB!DBConnectString) Then
            IsMwConnection = True
            Set loCon = goSession.DBConnection
         Else
            IsMwConnection = False
            strConnectString = loRsDB!DBConnectString
            Set loCon = New Connection
            loCon.Open strConnectString
         End If
         loRsDB.Filter = adFilterNone
         '
         ' Create Database Connection
         '
         '
         ' Process File...
         '
         strMsg = "Schema Update Received for " & strDescription & ", Issued: " & _
           loSchema("Header").ConfigKeys.GetKeyValue("IssueDate") & vbCrLf & _
           "Description: " & loSchema("Header").ConfigKeys.GetKeyValue("Description")
         goSession.Logger.LogIt mwl_Workstation, mwl_Information, strMsg
         i = MsgBox(strMsg & vbCrLf & vbCrLf & "This update will now be installed." & vbCrLf & _
           "Please make sure all other users are logged out of the system.", vbOKOnly, _
           "Remote Schema Update")
         '
         ' Loop through it...
         '
         For i = DB_SCHEMA_SQL_SECTION_START_OFFSET To loSchema.Count
            '
            ' Check whether script received
            '
            strChangeID = loSchema(i).ConfigKeys.GetKeyValue("ChangeID")
            strChangeDescription = loSchema(i).ConfigKeys.GetKeyValue("Description")
            strReference = (loSchema(i).ConfigKeys.GetKeyValue("Reference"))
            strBuildReference = UCase(loSchema(i).ConfigKeys.GetKeyValue("BuildReference"))
            '
            ' Check if ConnectionID/ChangeID record exists...
            '
            If strSourceID = "" Then
               If IsColumnInTable(loRsSU, "SourceID") Then
                  loRsSU.Filter = "ConnectionID='" & strConnectionID & "' AND ChangeID='" & strChangeID & "' and SourceID = null"
               Else
                  loRsSU.Filter = "ConnectionID='" & strConnectionID & "' AND ChangeID='" & strChangeID & "'"
               End If
            Else
               If IsColumnInTable(loRsSU, "SourceID") Then
                  loRsSU.Filter = "ConnectionID='" & strConnectionID & _
                    "' AND ChangeID='" & strChangeID & "' AND SourceID='" & strSourceID & "'"
               Else
                  loRsSU.Filter = "ConnectionID='" & strConnectionID & _
                    "' AND ChangeID='" & strChangeID & "'"
               End If
            End If
            If loRsSU.RecordCount <> 0 Then
               goSession.Logger.LogIt mwl_Workstation, mwl_Warning, " Script already processed (Connection/Source/Change): " & _
                 strConnectionID & "/" & strSourceID & "/" & strChangeID
            Else
               ' OK, let's do it...
               goSession.Logger.LogIt mwl_User_Defined, mwl_Information, "Starting Schema update (Connection/Source/Change): " & _
                 strConnectionID & "/" & strSourceID & "/" & strChangeID
               '
               ' MSDE/SQL Server - Begin Transaction
               '
               loCon.BeginTrans
               For j = DB_SCHEMA_SECTION_SQL_START_OFFSET To loSchema(i).ConfigKeys.Count
                  On Error Resume Next
                  strSQL = loSchema(i).ConfigKeys(j).KeyValue
                  If strSQL <> "" Then
                     loCon.Execute strSQL
                     If err Then
                        strMsg = "Error in " & strConnectionID & " Schema Update ! Script: " & _
                          strChangeID & " SQL: " & strSQL & ", Error: " & err.Number & "-" & err.Description
                        MsgBox strMsg, vbCritical
                        goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
                        IsScriptError = True
                        IsError = True
                     Else
                        strMsg = "Executed " & strConnectionID & " Schema Update ! Change ID: " & _
                          strChangeID & " SQL: " & strSQL
                        goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
                     End If
                  End If
                  On Error GoTo FunctionError
               Next j
               If IsScriptError Then
                  '
                  ' SQL Server - Rollback Transaction
                  '
                  loCon.RollbackTrans
               Else
                  '
                  ' Commit Transaction - SQL Server
                  '
                  loCon.CommitTrans
                  
                  '
                  ' Update mwcSchemaUpdate...
                  '
                  loRsSU.AddNew
                  loRsSU!ConnectionID = strConnectionID
                  If strSourceID <> "" Then
                     loRsSU!SourceID = strSourceID
                  End If
                  loRsSU!ChangeID = strChangeID
                  loRsSU!ChangeDate = Now()
                  loRsSU!BuildReference = strBuildReference
                  loRsSU!Reference = strReference
                  loRsSU!Description = strChangeDescription
                  loRsSU.Update
                  goSession.Logger.LogIt mwl_Workstation, mwl_Warning, "Successfully installed Schema Script: " & strChangeID
                  '
                  ' Create RSU Receipt...
                  '
                  If strSendNotification = "1" Or strSendNotification = "TRUE" Then
                     CreateRemoteSchemaReceipt strSourceID, strChangeID
                  End If

               End If
               IsScriptError = False
            End If
            loRsSU.Filter = adFilterNone
            
         Next i
         '
         ' Send Notification ?
         '
         If IsScriptError Then
            goSession.SendNotification strConnectionID & " Remote Schema Update Error Encountered for: " & goSession.Site.SiteName, "Script File: " & _
              loFile.Name & vbCrLf & _
              "ConnectionID= " & strConnectionID & vbCrLf & _
              "Last ChangeID= " & strChangeID & _
              "Description= " & strDescription
         End If
         
         '
         ' Archive schema update...
         '
         strMoveFile = loFile.Path & "." & Format(Now(), "yyyymmddhhmmss")
         fso.MoveFile loFile.Path, strMoveFile
         If IsError Then
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
            MsgBox " Errors occurred updating Database Schema through Revision: " & strChangeID & _
              vbCrLf & vbCrLf & "Maritime Workstation exit when all updates are installed." & vbCrLf & _
              "Restart system for successful updates to take effect. Thank you.", vbCritical
         Else
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchivedOK, True
            MsgBox " Database Schema has been updated through Revision: " & strChangeID & _
              vbCrLf & vbCrLf & "Maritime Workstation will exit when all updates are installed." & vbCrLf & _
              "Restart system for updates to take effect. Thank you.", vbInformation
         End If
         goSession.KillObject loSchema
      End If
NextFile:
   Next
   If IsMwConnection Then
      Set loCon = Nothing
   Else
      goSession.CloseConnection loCon
   End If
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.RemoteSchemaUpdater. ", err.Number, err.Description
   goSession.SendNotification "Error in Remote Schema Updater, on " & goSession.Site.SiteName, strError
   If Not loFile Is Nothing Then
      strMoveFile = loFile.Path & "." & Format(Now(), "yyyymmddhhmmss")
      fso.MoveFile loFile.Path, strMoveFile
      goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
      '
   End If
   
End Function



Private Function CreateRemoteSchemaReceipt(SourceID As String, ChangeID As String) As Boolean
   On Error GoTo FunctionError
   Dim loForm As mwForm
   Dim oRec As ConfigKeys
   On Error GoTo FunctionError
   '
   ' Create Remote record
   '
   Set loForm = goSession.OpenNewForm()
   If Not loForm.MakeNewForm("RSU_RECEIPT") Then
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loForm
      Exit Function
   End If
   Set oRec = loForm.DataForm.AddRecord
   With oRec
      .Add "Table", "mwcSchemaUpdateRemoteLog"
      .Add "Operation", "A"
      .Add "SourceID", SourceID
      .Add "mwcSitesKey", goSession.Site.SiteKey
      .Add "ChangeID", ChangeID
      .Add "ChangeDate", Now()
   End With
   '
   ' Write the message...
   '
   If Not loForm.DataForm.WriteDataForm() Then
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loForm
      goSession.KillObject oRec
      Exit Function
   End If
   '
   ' Transmit the message
   '
   loForm.DeleteAfterSend = True
   If Not loForm.Send() Then
      ' Error sending email
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loForm
      goSession.KillObject oRec
      Exit Function
   End If
   goSession.KillObject loForm
   goSession.KillObject oRec
   CreateRemoteSchemaReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.CreateRemoteSchemaReceipt. ", err.Number, err.Description
   CreateRemoteSchemaReceipt = False
End Function


Private Function IsValidateSource(SourceID As String) As Boolean
   Dim loRS As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError:
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select * from mwcSchemaUpdateSourceType where SourceID='" & SourceID & "'"
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.RecordCount > 0 Then
      IsValidateSource = True
   Else
      IsValidateSource = False
   End If
   goSession.CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.CreateRemoteSchemaReceipt. ", err.Number, err.Description
   IsValidateSource = False
   goSession.CloseRecordset loRS
End Function


