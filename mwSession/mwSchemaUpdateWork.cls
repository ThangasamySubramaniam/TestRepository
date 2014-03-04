VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwSchemaUpdateWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwSchemaUpdate
' 8 March 2003 ms
Option Explicit

Private Const DB_SCHEMA_SQL_SECTION_START_OFFSET = 2
Private Const DB_SCHEMA_SECTION_SQL_START_OFFSET = 5


Public Function RemoteSchemaUpdater(Optional IsReplicatedSchema As Boolean) As Boolean
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
   Dim loRs As Recordset
   
   ' DataWork
   Dim loDataWork As mwDataWork
   ' State variables
   Dim IsScriptError As Boolean
   ' Local Variables
   Dim strFile As String
   Dim IsMwConnection As Boolean
   Dim strMoveFile As String
   Dim strConnectString As String
   Dim strDescription As String
   Dim strChangeDescription As String
   Dim strSendNotification As String
   Dim strSQL As String
   Dim strOracleContinuationSQL As String
   Dim strMsg As String
   Dim strConnectionID As String
   Dim strSourceID As String
   Dim strChangeID As String
   Dim strBuildReference As String
   Dim strReference As String
   Dim strError As String
   Dim i As Integer
   Dim j As Integer
   Dim IsSqlServer As Boolean
   Dim IsOracle As Boolean
   Dim IsAccess As Boolean
   Dim strTargetSiteKey As String
   Dim strReplicateSchema As String
   Dim IsReplicateThisSchema As Boolean
   Dim IsError As Boolean
   Dim lofrmSchemaProgress As frmSchemaProgress
   Dim loCnt As Long
   
   '
   On Error GoTo FunctionError
   '
   ' Folder collection...looking for files to process...
   '
   Set fso = New FileSystemObject
   Set loFolder = fso.GetFolder(goSession.Workflow.PollingFolder)
   If loFolder.Files.Count < 1 Then
      RemoteSchemaUpdater = False
      Set fso = Nothing
      Exit Function
   End If
   '
   ' Check for db_update*.dat file
   '
   'Table of database connections...
   Set loDataWork = goSession.OpenNewDataWork
   Set loRsDB = New Recordset
   loRsDB.CursorLocation = adUseClient
   loRsDB.Open "mwcDbConnection", goCon, adOpenForwardOnly, adLockReadOnly, adCmdTable
   
   'Table of database connections...
   Set loRsSU = New Recordset
   loRsSU.CursorLocation = adUseClient
   loRsSU.Open "mwcSchemaUpdate", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   
   
   Set loRs = New Recordset
   loRs.Fields.Append "ID", adInteger, 4
   loRs.Fields.Append "Status", adVarWChar, 50
   loRs.Fields.Append "Path", adVarWChar, 200
   loRs.Fields.Append "Name", adVarWChar, 56
   loRs.Fields.Append "ChangeID", adVarWChar, 50
   loRs.Fields.Append "Description", adVarWChar, 255
   loRs.Fields.Append "Message", adVarWChar, 255
   loRs.Open
   
   loCnt = 0
   For Each loFile In loFolder.Files
      If UCase(fso.GetExtensionName(loFile.Path)) = "SCHEMA" Or UCase(fso.GetExtensionName(loFile.Path)) = "SCH" Then
      
         loCnt = loCnt + 1
         
         loRs.AddNew
         loRs!ID = loCnt
         loRs!Status = ""
         loRs!Path = loFile.Path
         loRs!Name = loFile.Name
         loRs!ChangeID = ""
         loRs!Description = ""
         loRs!Message = ""
         loRs.Update
      End If
   Next loFile
   If Not loRs.EOF Then
      loRs.MoveFirst
      loRs.Sort = "path"
      RemoteSchemaUpdater = True
   End If
   
   If loRs.RecordCount > 0 Then
   
      IsError = False
      
      If IsReplicatedSchema = False Then
         Set lofrmSchemaProgress = New frmSchemaProgress
         Load lofrmSchemaProgress
         
         If lofrmSchemaProgress.InitForm(loRs) Then
            lofrmSchemaProgress.Show vbModal
         End If
         
         Unload lofrmSchemaProgress
         Set lofrmSchemaProgress = Nothing
            
      Else
         Do While Not loRs.EOF And IsError = False
            
            IsError = ProcessSchema(loRs, IsReplicatedSchema)
            
            loRs.MoveNext
         Loop
      End If
   End If
   
   If IsMwConnection Then
      Set loCon = Nothing
   Else
      goSession.CloseConnection loCon
   End If
   goSession.KillObject fso
   
   CloseRecordset loRsDB
   CloseRecordset loRsSU
   CloseRecordset loRs
   
   Exit Function
FunctionError:
'Resume Next
   goSession.RaiseError "General Error in mwSession.mwSchemaUpdateWork.RemoteSchemaUpdater. ", Err.Number, Err.Description
   strError = "General Error in mwSession.mwSchemaUpdateWork.RemoteSchemaUpdater: " & Err.Number & " - " & Err.Description
   If Err.Number <> 0 Then
      goSession.SendNotification "Error in Remote Schema Updater, on " & goSession.Site.SiteName, strError
   End If
   If Not loFile Is Nothing Then
      strMoveFile = loRs!Path & "." & Format(Now(), "yyyymmddhhmmss")
      fso.MoveFile loRs!Path, strMoveFile
      goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
      '
   End If
   On Error Resume Next
   loCon.RollbackTrans
   
End Function
Public Function ProcessSchema(ByRef loRs As Recordset, Optional IsReplicatedSchema As Boolean) As Boolean
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
   
   ' DataWork
   Dim loDataWork As mwDataWork
   ' State variables
   Dim IsScriptError As Boolean
   ' Local Variables
   Dim strFile As String
   Dim IsMwConnection As Boolean
   Dim strMoveFile As String
   Dim strConnectString As String
   Dim strDescription As String
   Dim strChangeDescription As String
   Dim strSendNotification As String
   Dim strSQL As String
   Dim strOracleContinuationSQL As String
   Dim strOracleContinuationSQL2 As String
   Dim strMsg As String
   Dim strConnectionID As String
   Dim strSourceID As String
   Dim strChangeID As String
   Dim strBuildReference As String
   Dim strReference As String
   Dim strError As String
   Dim i As Integer
   Dim j As Integer
   Dim IsSqlServer As Boolean
   Dim IsOracle As Boolean
   Dim IsAccess As Boolean
   Dim strTargetSiteKey As String
   Dim strReplicateSchema As String
   Dim IsReplicateThisSchema As Boolean
   '
   On Error GoTo FunctionError
   '
   ' Folder collection...looking for files to process...
   '
   Set fso = New FileSystemObject
   
   '
   ' Check for db_update*.dat file
   '
   'Table of database connections...
   Set loDataWork = goSession.OpenNewDataWork
   Set loRsDB = New Recordset
   loRsDB.CursorLocation = adUseClient
   loRsDB.Open "mwcDbConnection", goCon, adOpenForwardOnly, adLockReadOnly, adCmdTable
   
   'Table of database connections...
   Set loRsSU = New Recordset
   loRsSU.CursorLocation = adUseClient
   loRsSU.Open "mwcSchemaUpdate", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   
   If Not loRs.EOF Then
      
      If UCase(fso.GetExtensionName(loRs!Path)) = "SCHEMA" Or UCase(fso.GetExtensionName(loRs!Path)) = "SCH" Then
         '
         ' State Flags
         '
         IsScriptError = False
         ProcessSchema = False
         '
         ' Have a schema file to process...
         '
         Set loSchema = goSession.LoadConfigGroupsFile(loRs!Path)
         If loSchema Is Nothing Then
            strMsg = "Error in mwSession.ProcessSchema, Bad Schema File Received:" & loRs!Name
            goSession.RaiseError strMsg
            strMoveFile = loRs!Path & "." & Format(Now(), "yyyymmddhhmmss")
            fso.MoveFile loRs!Path, strMoveFile
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
            loRs!Status = "Error"
            loRs!Message = Left(strMsg, 255)
            loRs.Update
            GoTo NextFile
         End If
         '
         ' Get Header Information
         '
         strConnectionID = Trim(UCase(loSchema("Header").ConfigKeys.GetKeyValue("ConnectionID")))
         strSendNotification = UCase(loSchema("Header").ConfigKeys.GetKeyValue("SendNotification"))
         strDescription = loSchema("Header").ConfigKeys.GetKeyValue("Description")
         strSourceID = loSchema("Header").ConfigKeys.GetKeyValue("SourceID")
         strTargetSiteKey = loSchema("Header").ConfigKeys.GetKeyValue("TargetSiteKey")
         strReplicateSchema = loSchema("Header").ConfigKeys.GetKeyValue("ReplicateSchema")
         
         strChangeID = loSchema(DB_SCHEMA_SQL_SECTION_START_OFFSET).ConfigKeys.GetKeyValue("ChangeID")
         loRs!ChangeID = strChangeID
         loRs!Description = strDescription
         loRs.Update
         
         If Len(strReplicateSchema) = 0 Or UCase(Left$(Trim(strReplicateSchema), 1)) = "Y" Then
            If IsReplicatedSchema = False And goSession.ThisSite.IsReplicateSchemaTransactions = True And goSession.Site.SiteType = SITE_TYPE_SHORE Then
               IsReplicateThisSchema = True
            Else
               IsReplicateThisSchema = False
            End If
         Else
            IsReplicateThisSchema = False
         End If
         
         
         If Trim(strSourceID) <> "" Then
            ' Validate SourceID
            If Not IsValidateSource(strSourceID) Then
               strMsg = "Error in mwSession.ProcessSchema, Invalid Source ID: " & strSourceID & " , in File:" & loRs!Name
               goSession.RaiseError strMsg
               strMoveFile = loRs!Path & "." & Format(Now(), "yyyymmddhhmmss")
               fso.MoveFile loRs!Path, strMoveFile
               goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
               loRs!Status = "Error"
               loRs!Message = Left(strMsg, 255)
               loRs.Update
               GoTo NextFile
            End If
         End If
         If Trim(strConnectionID) = "" Then
            strConnectionID = "MW"
         End If
         loRsDB.Filter = "ConnectionID='" & strConnectionID & "'"
         If loRsDB.RecordCount < 1 Then
            strError = "Error in mwSession.ProcessSchema, ConnectionID missing in database. ConnectionID: " & strConnectionID & ", File:" & loRs!Name
            If IsReplicatedSchema = False Then
               MsgBox strError
            End If
            
            goSession.Logger.LogIt mwl_Workstation, mwl_Critical, strError
            'strMoveFile = strFile & "." & Format(Now(), "yyyymmddhhmmss")
            'fso.MoveFile strFile, strMoveFile
            goSession.Workflow.ArchiveWorkflowItem loRs!Path, ArchiveError, True
            loRs!Status = "Error"
            loRs!Message = Left(strError, 255)
            loRs.Update
            GoTo NextFile
         End If
         If strConnectionID = "MW" Or IsNull(loRsDB!DBConnectString) Then
            IsMwConnection = True
            Set loCon = goCon
         Else
            IsMwConnection = False
            
'            strConnectString = loRsDB!DBConnectString
            
            Set loCon = loDataWork.OpenConnection(strConnectionID)
         End If
         loRsDB.Filter = adFilterNone
         '
         ' Set SQL Server flag...
         '
         IsSqlServer = loDataWork.IsSqlServer()
         IsOracle = loDataWork.IsOracle()
         IsAccess = loDataWork.IsAccess()
         '
         ' Process File...
         '
         
         If Len(strTargetSiteKey) > 0 And CStr(goSession.Site.SiteKey) <> strTargetSiteKey Then
            strMsg = "Schema Update Received for " & strDescription & ", Issued: " & _
              loSchema("Header").ConfigKeys.GetKeyValue("IssueDate") & vbCrLf & _
              "Description: " & loSchema("Header").ConfigKeys.GetKeyValue("Description")
            goSession.Logger.LogIt mwl_Workstation, mwl_Information, strMsg

            strMsg = "Schema " & strChangeID & " not processed beause it is marked for TargetSiteKey = " & strTargetSiteKey
            goSession.Logger.LogIt mwl_Workstation, mwl_Information, strMsg
            loRs!Status = "Info"
            loRs!Message = Left(strMsg, 255)
            loRs.Update
            
            If IsReplicateThisSchema = True Then
               CreateSchemaReplicationRecord loRs!Path, loRs!Name
            End If
            
         Else
            strMsg = "Schema Update Received for " & strDescription & ", Issued: " & _
              loSchema("Header").ConfigKeys.GetKeyValue("IssueDate") & vbCrLf & _
              "Description: " & loSchema("Header").ConfigKeys.GetKeyValue("Description")
            goSession.Logger.LogIt mwl_Workstation, mwl_Information, strMsg
            
'            If IsReplicatedSchema = False Then
'               i = MsgBox(strMsg & vbCrLf & vbCrLf & "This update will now be installed." & vbCrLf & _
'                 "Please make sure all other users are logged out of the system.", vbOKOnly, _
'                 "Remote Schema Update")
'            End If
            
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
                  strMsg = " Script already processed (Connection/Source/Change): " & strConnectionID & "/" & strSourceID & "/" & strChangeID
                  If Not goSession.Workflow.IsSilentErrors And IsReplicatedSchema = False Then
                     MsgBox strMsg, vbInformation, "Remote Schema Updater"
                  End If
                  loRs!Status = "Error"
                  loRs!Message = Left(strMsg, 255)
                  loRs.Update
                  goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
               Else
                  ' OK, let's do it...
                  goSession.Logger.LogIt mwl_User_Defined, mwl_Information, "Starting Schema update (Connection/Source/Change): " & _
                    strConnectionID & "/" & strSourceID & "/" & strChangeID
                  '
                  ' MSDE/SQL Server - Begin Transaction
                  '
                  
                  If IsReplicatedSchema = False Then
                     loCon.BeginTrans
                  End If
                  
                  j = DB_SCHEMA_SECTION_SQL_START_OFFSET
                  While j <= loSchema(i).ConfigKeys.Count
                     On Error Resume Next
                     strSQL = loSchema(i).ConfigKeys(j).KeyValue
                     If strSQL <> "" Then
                        If IsSqlServer Then
                           strSQL = loDataWork.ConvertToSQLSyntax(strSQL)
                        End If
                        If IsOracle Then
                           strSQL = loDataWork.ConvertToOracleSyntax(strSQL, strOracleContinuationSQL, strOracleContinuationSQL2)
                           If Len(strOracleContinuationSQL) > 0 Then
                              loSchema(i).ConfigKeys.Add loSchema(i).ConfigKeys(j).KeyName & "B", strOracleContinuationSQL
                           End If
                           If Len(strOracleContinuationSQL2) > 0 Then
                              loSchema(i).ConfigKeys.Add loSchema(i).ConfigKeys(j).KeyName & "C", strOracleContinuationSQL2
                           End If
                           
                        End If
                        
                        loCon.Execute strSQL
                       
                        If Err Then
                           If Err.Number = -2147467259 And Err.HelpContext = 5003284 Then
                              goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwSchemaUpdateWork.ProcessSchema, " & Err.Description & " : " & strSQL
                              Resume Next
                           ElseIf Err.Number = -2147467259 And Err.HelpContext = 5003039 Then
                              goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwSchemaUpdateWork.ProcessSchema, " & Err.Description & " : " & strSQL
                              Resume Next
                           ElseIf Err.Number = -2147467259 And Err.HelpContext = 1000430 Then
                              goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwSchemaUpdateWork.ProcessSchema, " & Err.Description & " : " & strSQL
                              Resume Next
                           ElseIf Err.Number = -2147467259 And Err.HelpContext = 1000440 Then
                              goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwSchemaUpdateWork.ProcessSchema, " & Err.Description & " : " & strSQL
                              Resume Next
      
                           ElseIf Err.Number = -2147217900 Then
                              If Left(Err.Description, 9) = "ORA-01408" Or Left(Err.Description, 9) = "ORA-02261" Then
                                 goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwSchemaUpdateWork.ProcessSchema, " & Err.Description & " : " & strSQL
                                 Resume Next
                              Else
                                 strMsg = "Error in " & strConnectionID & " Schema Update ! Script: " & _
                                   strChangeID & ", Error: " & Err.Number & "  " & Err.Description & " SQL: " & strSQL
                                 If IsReplicatedSchema = False Then
                                    MsgBox strMsg, vbCritical
                                 End If
                                 goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
                                 IsScriptError = True
                                 If IsReplicatedSchema = False Then
                                    loCon.RollbackTrans
                                 End If
                                 loRs!Status = "Error"
                                 loRs!Message = Left(strMsg, 255)
                                 loRs.Update
                                 Exit For
                              End If
                           
                           
                           Else
                              strMsg = "Error in " & strConnectionID & " Schema Update ! Script: " & _
                                strChangeID & ", Error: " & Err.Number & "  " & Err.Description & " SQL: " & strSQL
                              If IsReplicatedSchema = False Then
                                 MsgBox strMsg, vbCritical
                              End If
                              goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
                              IsScriptError = True
                              loRs!Status = "Error"
                              loRs!Message = Left(strMsg, 255)
                              loRs.Update
                              Exit For
                           End If
                        Else
                           strMsg = "Executed " & strConnectionID & " Schema Update ! Change ID: " & _
                             strChangeID & " SQL: " & strSQL
                           goSession.Logger.LogIt mwl_Workstation, mwl_Warning, strMsg
                        End If
                     End If
                     On Error GoTo FunctionError
                     j = j + 1
                  Wend
                  If IsScriptError Then
                     '
                     ' SQL Server/Access - Rollback Transaction
                     '
                     If IsReplicatedSchema = False Then
                        loCon.RollbackTrans
                     End If
                     
                     loRs!Status = "Error"
                     loRs!Message = Left(strMsg, 255)
                     loRs.Update
                  Else
                     '
                     ' Commit Transaction - SQL Server/Access
                     '
                     If IsReplicatedSchema = False Then
                        loCon.CommitTrans
                     End If
                     '
                     ' Update mwcSchemaUpdate...
                     '
                     loRsSU.Filter = adFilterNone
                     loRsSU.AddNew
                     loRsSU!ConnectionID = Left(Trim(strConnectionID), 50)
                     If strSourceID <> "" Then
                        loRsSU!SourceID = UCase(Left(Trim(strSourceID), 20))
                     End If
                     loRsSU!ChangeID = Left(Trim(strChangeID), 50)
                     loRsSU!ChangeDate = Now()
                     loRsSU!BuildReference = Left(Trim(strBuildReference), 50)
                     loRsSU!Reference = Left(Trim(strReference), 50)
                     loRsSU!Description = Left(Trim(strChangeDescription), 255)
                     loRsSU.Update
                     goSession.Logger.LogIt mwl_Workstation, mwl_Warning, "Successfully installed Schema Script: " & strChangeID
                     '
                     ' Create RSU Receipt...
                     '
                     If strSendNotification = "1" Or strSendNotification = "TRUE" Then
                        CreateRemoteSchemaReceipt strSourceID, strChangeID
                     End If
                     
                     loRs!Status = "Complete"
                     loRs!Message = "Schema complete"
                     loRs.Update

                     ' ***********************************************************************************************************
                     '
                     ' Insert a record of this schema in the ChangeLog table so it will get replicated to the remote sites.
                     '
                     ' ***********************************************************************************************************
   
                     If IsReplicateThisSchema = True Then
                        CreateSchemaReplicationRecord loRs!Path, loRs!Name
                     End If
                  End If
                  IsScriptError = False
               End If
               loRsSU.Filter = adFilterNone
               
            Next i
         End If
         '
         ' Archive schema update...
         '
         strMoveFile = loRs!Path & "." & Format(Now(), "yyyymmddhhmmss")
         fso.MoveFile loRs!Path, strMoveFile
         If IsScriptError Then
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
            
            If IsReplicatedSchema = False Then
               MsgBox " Errors occurred updating Database Schema through Revision: " & strChangeID & _
                 vbCrLf & vbCrLf & "ShipNet Fleet exit when all updates are installed." & vbCrLf & _
                 "Restart system for successful updates to take effect. Thank you.", vbCritical
            End If
            goSession.SendNotification strConnectionID & " Error in mwSession.mwSchemaUpdateWork. Remote Schema Update Error Encountered for: " & goSession.Site.SiteName, "Script File: " & _
              loRs!Name & vbCrLf & _
              "ConnectionID= " & strConnectionID & vbCrLf & _
              "Last ChangeID= " & strChangeID & vbCrLf & _
              "Description= " & strDescription & vbCrLf & _
              "Error= " & strMsg
            On Error Resume Next
            
            If IsReplicatedSchema = False Then
               loCon.RollbackTrans
            End If
            
            If IsMwConnection Then
               Set loCon = Nothing
            Else
               goSession.CloseConnection loCon
            End If
            goSession.KillObject fso
            CloseRecordset loRsDB
            CloseRecordset loRsSU
            ProcessSchema = True
            Exit Function
         Else
            goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchivedOK, True
'            If IsReplicatedSchema = False Then
'               MsgBox " Database Schema has been updated through Revision: " & strChangeID & _
'                 vbCrLf & vbCrLf & "ShipNet Fleet will exit when all updates are installed." & vbCrLf & _
'                 "Restart system for updates to take effect. Thank you.", vbInformation
'            End If
         End If
         goSession.KillObject loSchema
      End If

   End If

NextFile:
   If IsMwConnection Then
      Set loCon = Nothing
   Else
      goSession.CloseConnection loCon
   End If
   goSession.KillObject fso
   
   CloseRecordset loRsDB
   CloseRecordset loRsSU
   
   ProcessSchema = False
   Exit Function
   
FunctionError:
'Resume Next
   goSession.RaiseError "General Error in mwSession.mwSchemaUpdateWork.ProcessSchema. ", Err.Number, Err.Description
   If strError = "" Then
      strError = "General Error in mwSession.mwSchemaUpdateWork.ProcessSchema: " & Err.Number & " - " & Err.Description
   End If
   strMsg = "Error in Schema Processing on " & goSession.Site.SiteName
   
   If strChangeID <> "" Then
      strMsg = strMsg & ", Schema " & strChangeID
   End If
   
   goSession.SendNotification strMsg, strError
   If Not loFile Is Nothing Then
      strMoveFile = loRs!Path & "." & Format(Now(), "yyyymmddhhmmss")
      fso.MoveFile loRs!Path, strMoveFile
      goSession.Workflow.ArchiveWorkflowItem strMoveFile, ArchiveError, True
      '
   End If
   On Error Resume Next
   If IsReplicatedSchema = False Then
      loCon.RollbackTrans
   End If
   ProcessSchema = True
   loRs!Status = "Error"
   loRs.Update
   
End Function

Private Function CreateRemoteSchemaReceipt(SourceID As String, ChangeID As String) As Boolean
   On Error GoTo FunctionError
   Dim loform As mwForm
   Dim oRec As ConfigKeys
   On Error GoTo FunctionError
   '
   ' Create Remote record
   '
   Set loform = goSession.OpenNewForm()
   If Not loform.MakeNewForm("RSU_RECEIPT") Then
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loform
      Exit Function
   End If
   Set oRec = loform.DataForm.AddRecord
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
   If Not loform.DataForm.WriteDataForm() Then
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   '
   ' Transmit the message
   '
   loform.DeleteAfterSend = True
   If Not loform.Send() Then
      ' Error sending email
      CreateRemoteSchemaReceipt = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   goSession.KillObject loform
   goSession.KillObject oRec
   CreateRemoteSchemaReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.CreateRemoteSchemaReceipt. ", Err.Number, Err.Description
   CreateRemoteSchemaReceipt = False
End Function


Private Function IsValidateSource(SourceID As String) As Boolean
   Dim loRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError:
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwcSchemaUpdateSourceType where SourceID='" & SourceID & "'"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount > 0 Then
      IsValidateSource = True
   Else
      IsValidateSource = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.CreateRemoteSchemaReceipt. ", Err.Number, Err.Description
   IsValidateSource = False
   CloseRecordset loRs
End Function


Public Function MakeNewSchemaUpdate(SourceID As String, SchemaDescription As String, ConnectionID As String, SendNotification As Boolean, SchemaReference As String) As ConfigGroups
   Dim loCfg As ConfigGroups
   Dim loRec As ConfigGroup
   Dim loRs As Recordset
   Dim strSQL As String
   Dim strChangeID As String
   On Error GoTo FunctionError
   '
   ' Fetch max value from log...
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select max(ChangeNumber) from mwcSchemaUpdateMakeLog where SourceID='" & SourceID & "'"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      ' what to do ?
   End If
   strChangeID = Format(loRs.Fields(1).value + 1, "00000")
   loRs.Close
   '
   ' Create new record
   '
   loRs.Open "mwcSchemaUpdateMakeLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   loRs.AddNew
   loRs!SourceID = SourceID
   loRs!ChangeNumber = CLng(strChangeID)
   loRs!SchemaDescription = SchemaDescription
   loRs!CreateDate = Now()
   loRs!mwcUsersKey = goSession.User.UserKey
   loRs!IsSent = 0
   loRs.Update
   '
   ' Create SchemaUpdate Shell
   '
   Set loCfg = New ConfigGroups
   Set loRec = loCfg.Add("HEADER", "HEADER")
   With loRec.ConfigKeys
      .Add "IssueDate", Format(Now(), "Short Date")
      .Add "Description", SchemaDescription
      .Add "ConnectionID", ConnectionID
      If SendNotification Then
         .Add "SendNotification", "True"
      Else
         .Add "SendNotification", "False"
      End If
      .Add "SourceID", SourceID
   End With
   '
   
   Set loRec = loCfg.Add("[" & strChangeID & "]", "[" & strChangeID & "]")
   With loRec.ConfigKeys
      .Add "ChangeID", strChangeID
      .Add "Reference", SchemaReference
      .Add "Description", SchemaDescription
   End With
   ' Return ConfigGroups, wrap up shop
   Set MakeNewSchemaUpdate = loCfg
   Set loCfg = Nothing
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.MakeNewSchemaUpdate. ", Err.Number, Err.Description
End Function


Public Function DeleteFormTemplateLines(ByRef loRec As ConfigGroup, TemplateID As String) As Boolean
   On Error GoTo FunctionError
   Dim loRs As Recordset
   Dim strSQL As String
   Dim iCounter As Integer
   ' Working Recordset
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.CursorType = adOpenForwardOnly
   loRs.LockType = adLockReadOnly
   '
   ' mwwfFormPopulate
   ' mwwfFleetFormTemplate
   ' mwwfFormTemplateEmailList
   ' mwcMenuItemTemplates
   '
   With loRec.ConfigKeys
      .Add "100", "delete from mwwfFormPopulate where TemplateID='" & TemplateID & "'"
      .Add "101", "delete from mwwfFleetFormTemplate where TemplateID='" & TemplateID & "'"
      .Add "102", "delete from mwwfFormTemplateEmailList where TemplateID='" & TemplateID & "'"
      .Add "103", "delete from mwcMenuItemTemplates where TemplateID='" & TemplateID & "'"
   End With
   '
   ' mwwfFormPopulateTable
   ' mwwfFormPopulateColumn
   '
   strSQL = "select * from mwwfFormPopulateTable where TemplateID='" & TemplateID & "'"
   loRs.Open strSQL
   iCounter = 200
   Do While Not loRs.EOF
      '
      With loRec.ConfigKeys
         .Add str(iCounter) & "A", "delete from mwwfFormPopulateColumn where FptKey=" & loRs.Fields("ID").value
         .Add str(iCounter) & "B", "delete from mwwfFormPopulateTable where ID=" & loRs.Fields("ID").value
      End With
      iCounter = iCounter + 1
      loRs.MoveNext
   Loop
   loRs.Close
   '
   ' mwwfEventMapDgTable
   ' mwwfEventMapDgColumn
   '
   strSQL = "select * from mwwfEventMapDgTable where TemplateID='" & TemplateID & "'"
   loRs.Open strSQL
   iCounter = 300
   Do While Not loRs.EOF
      '
      With loRec.ConfigKeys
         .Add str(iCounter) & "A", "delete from mwwfEventMapDgColumn where mwEventMapDgTableKey=" & loRs.Fields("ID").value
         .Add str(iCounter) & "B", "delete from mwwfEventMapDgTable where ID=" & loRs.Fields("ID").value
      End With
      iCounter = iCounter + 1
      loRs.MoveNext
   Loop
   loRs.Close
   '
   ' mwwfEventFormType
   ' mwwfEventFormLog
   '
   strSQL = "select * from mwwfEventFormType where TemplateID='" & TemplateID & "'"
   loRs.Open strSQL
   iCounter = 400
   Do While Not loRs.EOF
      '
      With loRec.ConfigKeys
         .Add str(iCounter) & "A", "delete from mwwfEventFormLog where mwEventFormTypeKey=" & loRs.Fields("ID").value
         .Add str(iCounter) & "B", "delete from mwwfEventFormType where ID=" & loRs.Fields("ID").value
      End With
      iCounter = iCounter + 1
      loRs.MoveNext
   Loop
   loRs.Close
   '
   ' and... mwwfFormTemplate
   '
   With loRec.ConfigKeys
      .Add "500", "delete from mwwfFormTemplate where TemplateID='" & TemplateID & "'"
   End With
   '
   ' That's all folks...
   '
   CloseRecordset loRs
   DeleteFormTemplateLines = True

   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.DeleteFormTemplateLines. ", Err.Number, Err.Description
   DeleteFormTemplateLines = False
End Function


Public Function AddFormTemplateLines(ByRef loRec As ConfigGroup, TemplateID As String) As Boolean
   On Error GoTo FunctionError
   Dim loRs As Recordset
   Dim strSQL As String
   Dim strTemp As String
   ' Working Recordset
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.CursorType = adOpenForwardOnly
   loRs.LockType = adLockReadOnly
   
   '
   ' mwwfFormPopulate
   ' mwwfFleetFormTemplate
   ' mwwfFormTemplateEmailList
   ' mwcMenuItemTemplates
   '
   
   
   
   '
   ' mwwfFormPopulateTable
   ' mwwfFormPopulateColumn
   '
   
   '
   ' mwwfEventFormType
   ' mwwfEventFormLog
   '
   
   '
   ' mwwfEventMapDgTable
   ' mwwfEventMapDgColumn
   '


   '
   ' That's all folks...
   '


   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.AddFormTemplateLines. ", Err.Number, Err.Description
End Function

Private Function CreateSchemaReplicationRecord(filepath As String, SchemaName As String) As Boolean
   Dim loStream As ADODB.Stream
   Dim fso As FileSystemObject
   Dim loRs As Recordset
   Dim sStreamContents As String
   On Error GoTo FunctionError

   Set fso = New FileSystemObject
   If Not fso.FileExists(filepath) Then
      goSession.RaisePublicError "Error in mwSession.mwSchemaUpdateWork.CreateSchemaReplicationRecord, File Not Found: " & filepath
      KillObject fso
      CreateSchemaReplicationRecord = False
      Exit Function
   End If

   Set loStream = New ADODB.Stream
   
   loStream.Type = adTypeText
   loStream.Charset = "ascii"
   loStream.Open
   loStream.LoadFromFile filepath
   sStreamContents = loStream.ReadText
   sStreamContents = Replace(sStreamContents, Chr(0), "")

   If loStream.Size > 30000 Then
      goSession.RaisePublicError "Schema " & SchemaName & " is too large to replicate. Maximum size is 30,000 bytes."
      KillObject fso
      KillObject loStream
      CreateSchemaReplicationRecord = False
      Exit Function
   End If

   Set loRs = New Recordset
   loRs.Fields.Append "SchemaName", adVarWChar, 200
'   loRs.Fields.Append "SchemaContents", adLongVarWChar, 50000
'   loRs.Fields.Append "SchemaContents", adLongVarWChar, (loStream.Size + 10) * 2
   loRs.Fields.Append "SchemaContents", adLongVarWChar, loStream.Size
   loRs.Open
      
   loRs.AddNew
   loRs!SchemaName = SchemaName
   loRs!SchemaContents = sStreamContents
   loRs.Update
   
   CreateSchemaReplicationRecord = goSession.ReplicateWork.LogSchemaChange(loRs.Fields)
   
   CloseRecordset loRs

   KillObject fso
   KillObject loStream

   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwSchemaUpdateWork.AddRecord. ", Err.Number, Err.Description
   CloseRecordset loRs
   KillObject fso
   KillObject loStream
   CreateSchemaReplicationRecord = False
End Function

Public Function ProcessReplicatedSchema(SchemaName As String, SchemaContents As String) As Boolean
   Dim loStream As ADODB.Stream
   Dim strFilename As String
   
   On Error GoTo FunctionError
   
   strFilename = goSession.Workflow.PollingFolder & "\" & SchemaName
   
   Set loStream = New ADODB.Stream
   
   loStream.Type = adTypeText
   loStream.Charset = "ascii"
   loStream.Open
      
   loStream.WriteText SchemaContents
   loStream.SaveToFile strFilename, adSaveCreateOverWrite
      
   KillObject loStream
   
   ProcessReplicatedSchema = RemoteSchemaUpdater(True)
   
   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwSchemaUpdateWork.ProcessReplicatedSchema. ", Err.Number, Err.Description
   KillObject loStream
   ProcessReplicatedSchema = False
End Function
