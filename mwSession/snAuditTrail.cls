VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "snAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const MAX_OLD_VALUE_TEST = 31

Private Function IsLogAuditTrialEnabledForParentEvent(RecordKey As Long, ChangeTableKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset

   On Error GoTo FunctionError
   
   IsLogAuditTrialEnabledForParentEvent = False
   
   Select Case ChangeTableKey
      Case Is = MWRT_mwEventLog       'EventLog
         sSQL = "SELECT 1 FROM mwEventLog" & _
                " INNER JOIN (mwEventType INNER JOIN mwrChangeTable " & _
                " ON mwEventType.TableName = mwrChangeTable.Tablename )" & _
                " ON mwEventLog.mwEventTypeKey = mwEventType.ID" & _
                " WHERE (mwrChangeTable.IsAuditTrailEnabled = 1 OR mwrChangeTable.IsAuditTrailEnabled = -1)" & _
                " AND mwEventLog.ID = " & RecordKey
      Case Is = MWRT_mwEventLinkLog       'Link
         sSQL = "SELECT 1 FROM mwEventLinkLog" & _
                " INNER JOIN (mwEventType INNER JOIN mwrChangeTable " & _
                " ON mwEventType.TableName = mwrChangeTable.Tablename )" & _
                " ON mwEventLinkLog.mwEventTypeKey = mwEventType.ID" & _
                " WHERE (mwrChangeTable.IsAuditTrailEnabled = 1 OR mwrChangeTable.IsAuditTrailEnabled = -1)" & _
                " AND mwEventLinkLog.ID = " & RecordKey
      Case Is = MWRT_mwEventFactLogSN     'Fact
         sSQL = "SELECT 1 FROM mwEventFactLogSN" & _
                " INNER JOIN (mwEventType INNER JOIN mwrChangeTable " & _
                " ON mwEventType.TableName = mwrChangeTable.Tablename )" & _
                " ON mwEventFactLogSN.mwEventTypeKey = mwEventType.ID" & _
                " WHERE (mwrChangeTable.IsAuditTrailEnabled = 1 OR mwrChangeTable.IsAuditTrailEnabled = -1)" & _
                " AND mwEventFactLogSN.ID = " & RecordKey
      Case Is = MWRT_mwEventFormLog       'Form
         sSQL = "SELECT 1 FROM mwEventFormLog" & _
                " INNER JOIN (mwEventType INNER JOIN mwrChangeTable " & _
                " ON mwEventType.TableName = mwrChangeTable.Tablename )" & _
                " ON mwEventFormLog.mwEventTypeKey = mwEventType.ID" & _
                " WHERE (mwrChangeTable.IsAuditTrailEnabled = 1 OR mwrChangeTable.IsAuditTrailEnabled = -1)" & _
                " AND mwEventFormLog.ID = " & RecordKey
      Case Else
         IsLogAuditTrialEnabledForParentEvent = True
         Exit Function
   End Select

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      IsLogAuditTrialEnabledForParentEvent = True
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.IsLogAuditTrialEnabledForParentEvent. ", Err.Number, Err.Description
End Function

Public Function WriteAuditTrailFromChangeLog(ChangeLogKey As Long) As Boolean
   Dim sSQL As String
   Dim sSQL1 As String
   Dim nAuditTrailHeaderKey As Long
   Dim nParentKey As Long
   Dim sRemarks As String
   Dim loRsAT As Recordset
   Dim loRs As Recordset
   Dim sOperation As String
   On Error GoTo FunctionError
   
   sSQL = "SELECT *  from mwrChangeLog where ID = " & ChangeLogKey & " ORDER BY ID"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If Not IsLogAuditTrialEnabledForParentEvent(loRs.Fields("RecordKey").value, loRs.Fields("mwrChangeTableKey").value) Then
         CloseRecordset loRs
         WriteAuditTrailFromChangeLog = True
         Exit Function
      End If
   End If

   Set loRsAT = New Recordset
   loRsAT.CursorLocation = adUseClient
   sSQL1 = "select * from snAuditTrail where ID=-1"
   loRsAT.Open sSQL1, goCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(loRs) Then

      Do While Not loRs.EOF
         loRsAT.AddNew
         loRsAT.Fields("ID").value = goSession.MakePK("snAuditTrail")
         nAuditTrailHeaderKey = loRsAT.Fields("ID").value
         sOperation = loRs.Fields("ChangeType").value
         If sOperation = "D" Then
            nParentKey = GetParentKeyForAuditTrail(loRs.Fields("RecordKey").value, loRs.Fields("mwrChangeTableKey").value, sRemarks)
            loRsAT.Fields("mwrChangeTableKey").value = MWRT_smOccurrence
            loRsAT.Fields("RecordKey").value = nParentKey
            loRsAT.Fields("Remarks").value = sRemarks
         Else
            loRsAT.Fields("mwrChangeTableKey").value = loRs.Fields("mwrChangeTableKey").value
            loRsAT.Fields("RecordKey").value = loRs.Fields("RecordKey").value
         End If
         loRsAT.Fields("ChangeType").value = loRs.Fields("ChangeType").value
         loRsAT.Fields("mwcUsersKey").value = goSession.User.UserKey
         If goSession.User.RoleTypeKey > 0 Then
            loRsAT.Fields("mwcRoleTypeKey").value = goSession.User.RoleTypeKey
         End If
         loRsAT.Fields("ChangeDateTime").value = loRs.Fields("ChangeDateTime").value
         loRsAT.Fields("mwcSitesKey").value = loRs.Fields("mwcSitesKey").value
         loRsAT.Fields("OriginatingSiteKey").value = goSession.Site.SiteKey
         loRsAT.Fields("mwrChangeLogKey").value = ChangeLogKey
         loRsAT.Fields("IsDeleted").value = False

         loRsAT.Update
           
         If sOperation = "D" Then
            MarkAuditTrailForDelete loRs.Fields("RecordKey").value, loRs.Fields("mwrChangeTableKey").value
         Else
            WriteAuditTrailDetailFromChangeLog ChangeLogKey, nAuditTrailHeaderKey
         End If
         loRs.MoveNext
      Loop
   
   End If
   CloseRecordset loRsAT
   CloseRecordset loRs
   WriteAuditTrailFromChangeLog = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.WriteAuditTrailFromChangeLog. ", Err.Number, Err.Description
   CloseRecordset loRsAT
   CloseRecordset loRs
End Function

Private Function WriteAuditTrailDetailFromChangeLog(ChangeLogKey As Long, AuditTrailHeaderKey As Long) As Boolean

   Dim sSQL As String
   Dim sSQL1 As String
   Dim loRs As Recordset
   Dim moRsATDetail As Recordset

   On Error GoTo FunctionError
   
   sSQL = "SELECT *  from mwrChangeLogColumn where mwrChangeLogKey = " & ChangeLogKey & " ORDER BY ID"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   CloseRecordset moRsATDetail
   Set moRsATDetail = New Recordset
   moRsATDetail.CursorLocation = adUseClient
   sSQL1 = "select * from snAuditTrailDetail where ID=-1"
   moRsATDetail.Open sSQL1, goCon, adOpenDynamic, adLockOptimistic

   If IsRecordLoaded(loRs) Then
   
      Do While Not loRs.EOF
         
         moRsATDetail.AddNew
         moRsATDetail.Fields("ID").value = goSession.MakePK("snAuditTrailDetail")
         moRsATDetail.Fields("snAuditTrailKey").value = AuditTrailHeaderKey
         moRsATDetail.Fields("ColumnDataType").value = loRs.Fields("ColumnDataType").value
         moRsATDetail.Fields("ColumnName").value = loRs.Fields("ColumnName").value
         moRsATDetail.Fields("OldValue").value = loRs.Fields("OldValue").value
         moRsATDetail.Fields("OldValueLarge").value = loRs.Fields("OldValueLarge").value
         moRsATDetail.Fields("NewValue").value = loRs.Fields("NewValue").value
         moRsATDetail.Fields("NewValueLarge").value = loRs.Fields("NewValueLarge").value
         
         moRsATDetail.Update

         loRs.MoveNext
      Loop
   End If
   
   CloseRecordset moRsATDetail
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in snAuditTrail.WriteAuditTrailDetailFromChangeLog ", Err.Number, Err.Description
   CloseRecordset moRsATDetail
   CloseRecordset loRs

End Function

Private Function GetParentKeyForAuditTrail(RecordKey As Long, ChangeTableKey As Long, Remarks As String) As Long

   Dim sSQL As String
   Dim loRs As Recordset

   On Error GoTo FunctionError
   
   Remarks = ""

   Select Case ChangeTableKey
      
      Case Is = MWRT_smOccurrenceFinding  'Finding
         sSQL = "SELECT smOccurrenceKey,reference from smOccurrenceFinding where ID = " & RecordKey
      Case Is = MWRT_smOccurrenceInjury   'Injury
         sSQL = "SELECT smOccurrenceKey,InjuredCrewName from smOccurrenceInjury where ID = " & RecordKey
      Case Is = MWRT_mwEventLinkLog       'Link
         sSQL = "SELECT mwEventLinkLog.mwEventDetailKey, mwEventLinkLog.BriefDescription,mweventlinktype.LinkTitle from mwEventLinkLog LEFT OUTER JOIN mweventlinktype ON (mwEventLinkLog.mwEventLinkTypeKey = mweventlinktype.ID) where mwEventLinkLog.ID = " & RecordKey
      Case Is = MWRT_mwEventFactLogSN     'Fact
         sSQL = "SELECT mwEventDetailKey,FactTitle from mwEventFactLogSN where ID = " & RecordKey
      Case Is = MWRT_mwEventFormLog       'Form
         sSQL = "SELECT mwEventFormLog.mwEventDetailKey, mwEventFormType.TemplateID from mwEventFormLog LEFT OUTER JOIN mwEventFormType ON (mwEventFormLog.mwEventFormTypeKey = mwEventFormType.ID) where mwEventFormLog.ID = " & RecordKey
      Case Else
         Exit Function
   End Select

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      GetParentKeyForAuditTrail = loRs.Fields(0).value
   
      Select Case ChangeTableKey
         
         Case Is = MWRT_smOccurrenceFinding  'Finding
            Remarks = "Finding Record Deleted"
            If Not IsNull(loRs.Fields(1)) Then
               Remarks = Remarks & " Reference : " & loRs.Fields(1)
            End If
         Case Is = MWRT_smOccurrenceInjury   'Injury
            Remarks = "Injury Record Deleted"
            If Not IsNull(loRs.Fields(1)) Then
               Remarks = Remarks & " Crew Name : " & loRs.Fields(1)
            End If
         Case Is = MWRT_mwEventLinkLog       'Link
            Remarks = "Link Record Deleted : "
            If Not IsNull(loRs.Fields(2)) Then
               Remarks = Remarks & " Link Type : " & loRs.Fields(2) & " ; "
            End If
            If Not IsNull(loRs.Fields(1)) Then
               Remarks = Remarks & " Link Name : " & loRs.Fields(1)
            End If
         Case Is = MWRT_mwEventFactLogSN     'Fact
            Remarks = "Fact Record Deleted : "
            If Not IsNull(loRs.Fields(1)) Then
               Remarks = Remarks & " Fact Title : " & loRs.Fields(1)
            End If
         Case Is = MWRT_mwEventFormLog       'Form
            Remarks = "Form Record Deleted : "
            If Not IsNull(loRs.Fields(1)) Then
               Remarks = Remarks & " Form Name : " & loRs.Fields(1)
            End If
         Case Else
      End Select
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.GetParentKeyForAuditTrail. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function MarkAuditTrailForDelete(RecordKey As Long, ChangeTableKey As Long) As Boolean

   Dim sSQL As String
   Dim loRs As Recordset

   On Error GoTo FunctionError
   
   sSQL = "SELECT * from snAuditTrail where RecordKey = " & RecordKey & " And mwrChangeTableKey = " & ChangeTableKey & " And IsDeleted = 0 "

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(loRs) Then
      Do While Not loRs.EOF
         loRs.Fields("IsDeleted").value = True
         loRs.Update
         loRs.MoveNext
      Loop
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.MarkAuditTrailForDelete. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function DeleteAuditTrailTransaction(ChangeLogKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError

      sSQL = "delete from snAuditTrailDetail where snAuditTrailKey in (Select ID from snAuditTrail where mwrChangeLogKey =" & ChangeLogKey & ")"
      goCon.Execute sSQL
      sSQL = "delete from snAuditTrail where mwrChangeLogKey=" & ChangeLogKey
      goCon.Execute sSQL
      DeleteAuditTrailTransaction = True

   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.DeleteAuditTrailTransaction. ", Err.Number, Err.Description
End Function

Public Function CheckIsAuditTrailEnabledTable(TargetTable As String) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   strSQL = "select IsAuditTrailEnabled from mwrChangeTable where TableName='" & TargetTable & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      CheckIsAuditTrailEnabledTable = False
   Else
      CheckIsAuditTrailEnabledTable = BoolNull(loRs!IsAuditTrailEnabled)
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.CheckIsAuditTrailEnabledTable. ", Err.Number, Err.Description
   CheckIsAuditTrailEnabledTable = False
   CloseRecordset loRs
End Function

Public Function WriteAuditTrailHeader(ChangeType As String, ChangeTableKey As Long, RecordKey As Long, OriginatingSiteKey As Long, ChangeDateTime As String, UserKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   Dim nRoleTypeKey As Long
   Dim nParentKey As Long
   Dim sRemarks As String
   On Error GoTo FunctionError
   If ChangeTableKey > 0 Then
      Set loRs = New Recordset
      With loRs
         If goSession.IsSqlServer Then
            .CursorLocation = adUseClient
         Else
            .CursorLocation = adUseServer
         End If
         sSQL = "select * from snAuditTrail where ID=-1"
         .Open sSQL, goCon, adOpenDynamic, adLockOptimistic
         .AddNew
         .Fields("ID").value = goSession.MakePK("snAuditTrail")
         
         If ChangeType = "D" Then
            nParentKey = GetParentKeyForAuditTrail(RecordKey, ChangeTableKey, sRemarks)
            .Fields("mwrChangeTableKey").value = MWRT_smOccurrence
            .Fields("RecordKey").value = nParentKey
            .Fields("Remarks").value = sRemarks
         Else
            .Fields("mwrChangeTableKey").value = ChangeTableKey
            .Fields("RecordKey").value = RecordKey
         End If
         
         .Fields("ChangeType").value = ChangeType
         If goSession.Site.SiteType = SITE_TYPE_SHORE Then
            .Fields("mwcSitesKey").value = OriginatingSiteKey
         Else
            .Fields("mwcSitesKey").value = goSession.Site.SiteKey
         End If
         .Fields("OriginatingSiteKey").value = OriginatingSiteKey
         .Fields("mwcUsersKey").value = UserKey
         nRoleTypeKey = GetUserRoleTypeKey(UserKey)
         If nRoleTypeKey > 0 Then
            .Fields("mwcRoleTypeKey").value = nRoleTypeKey
         End If
         .Fields("ChangeDateTime").value = ChangeDateTime
         .Fields("IsDeleted").value = False
         .Update
         If goSession.IsOracle() Then
            .Resync adAffectCurrent, adResyncAllValues
         End If
         If ChangeType = "D" Then
            MarkAuditTrailForDelete RecordKey, ChangeTableKey
         End If
         WriteAuditTrailHeader = .Fields("ID").value
      End With
      CloseRecordset loRs
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.WriteAuditTrailHeader. ", Err.Number, Err.Description
'   Resume Next
   WriteAuditTrailHeader = -1
   CloseRecordset loRs
End Function

Public Function WriteAuditTrailDetail(AuditTrailHeaderKey As Long, RecField As field) As Boolean
   Dim sSQL1 As String
   Dim sSQL2 As String
   Dim sDate As String
   Dim sField As String
   Dim i As Integer
   Static loCmd As Command
   
   On Error GoTo FunctionError
   
   If AuditTrailHeaderKey > 0 Then
      '
      ' Check for compound column names (based on tables constructed using joins)
      '
      If loCmd Is Nothing Then
         Set loCmd = New Command
         Set loCmd.ActiveConnection = goCon
         loCmd.CommandType = adCmdText
         loCmd.CommandText = "insert into snAuditTrailDetail (ID, snAuditTrailKey, ColumnName, OldValue, " & _
          " NewValue, OldValueLarge, NewValueLarge, ColumnDataType) values (?, ?, ?, ?, ?, ?, ?, ?)"
         loCmd.Parameters.Append loCmd.CreateParameter("KEY", adInteger, adParamInput)
         loCmd.Parameters.Append loCmd.CreateParameter("ATK", adInteger, adParamInput)
         loCmd.Parameters.Append loCmd.CreateParameter("CN", adVarChar, adParamInput, 50)
         loCmd.Parameters.Append loCmd.CreateParameter("OV", adVarChar, adParamInput, 30)
         loCmd.Parameters.Append loCmd.CreateParameter("NV", adVarChar, adParamInput, 30)
         loCmd.Parameters.Append loCmd.CreateParameter("OVL", adLongVarWChar, adParamInput, 30000)
         loCmd.Parameters.Append loCmd.CreateParameter("NVL", adLongVarWChar, adParamInput, 30000)
         loCmd.Parameters.Append loCmd.CreateParameter("CDT", adInteger, adParamInput)
         loCmd.Parameters("NV").value = Null
         loCmd.Parameters("NVL").value = Null
         loCmd.Parameters("OVL").value = Null
         loCmd.Parameters("CN").value = Null
         loCmd.Parameters("OV").value = Null
      End If
      loCmd.Parameters("KEY").value = goSession.MakePK("snAuditTrailDetail")
      loCmd.Parameters("ATK").value = AuditTrailHeaderKey
      loCmd.Parameters("CDT").value = RecField.Type
      sField = RecField.Name
      i = InStr(1, RecField.Name, ".")
      If i > 0 Then
         sField = mID$(RecField.Name, i + 1)
      Else
         sField = RecField.Name
      End If
      loCmd.Parameters("CN").value = sField
      '
      If Not IsNull(RecField.OriginalValue) Then
         If RecField.OriginalValue <> "" Then
            If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
               If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
                  sDate = Format(RecField.OriginalValue, DATAGRAM_DATE_FORMAT)
                  loCmd.Parameters("OV").value = MassageField(RecField.OriginalValue)
               ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
                  loCmd.Parameters("OV").value = TranslateBoolean(RecField.OriginalValue)
               Else
                  loCmd.Parameters("OV").value = MassageField(RecField.OriginalValue)
               End If
               loCmd.Parameters("OVL").value = Null
            Else
               loCmd.Parameters("OVL").value = MassageField(RecField.OriginalValue)
               loCmd.Parameters("OV").value = Null
            End If
         End If
      Else
         loCmd.Parameters("OVL").value = Null
         loCmd.Parameters("OV").value = Null
      End If
      If IsNull(RecField.value) Or (RecField.value = "" And RecField.Type = adLongVarWChar) Then
         If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
            loCmd.Parameters("NV").value = "[NULL]"
         Else
            loCmd.Parameters("NVL").value = "[NULL]"
         End If
      Else
         If RecField.value <> "" Then
            If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
               If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
                  sDate = Format(RecField.value, DATAGRAM_DATE_FORMAT)
                  loCmd.Parameters("NV").value = sDate
               ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
                  loCmd.Parameters("NV").value = TranslateBoolean(RecField.value)
               Else
                  loCmd.Parameters("NV").value = MassageField(RecField.value)
               End If
               loCmd.Parameters("NVL").value = Null
            Else
               loCmd.Parameters("NVL").value = MassageField(RecField.value)
               loCmd.Parameters("NV").value = Null
            End If
         End If
      End If
      loCmd.Execute
      WriteAuditTrailDetail = True
   Else
      WriteAuditTrailDetail = False
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.WriteAuditTrailDetail. ", Err.Number, Err.Description
   WriteAuditTrailDetail = False
End Function

Private Function GetUserRoleTypeKey(UserKey As Long) As Long
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   strSQL = "select mwcRoleTypeKey from mwcUsers where ID =" & UserKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetUserRoleTypeKey = -1
   Else
      GetUserRoleTypeKey = ZeroNull(loRs!mwcRoleTypekey)
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.GetUserRoleTypeKey. ", Err.Number, Err.Description
   GetUserRoleTypeKey = -1
   CloseRecordset loRs
End Function

Private Function MassageField(str As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   strTemp = Replace(str, "'", " ")
   strTemp = FixString(strTemp)
   
   If Len(Trim(strTemp)) = 0 Then
      strTemp = str
   End If
   MassageField = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.MassageField: ", Err.Number, Err.Description
   MassageField = str
End Function
Private Function FixString(strVal As String) As String
   Dim strTemp As String
   Dim ByteArray() As Byte
   Dim xx As Long
   
   ' Look for and fix UniCode characters.
   
   On Error GoTo FunctionError
   
   strTemp = Replace(strVal, vbCrLf, "|")
   strTemp = Replace(strTemp, vbLf, "|")
   
   ByteArray = strTemp
   
   For xx = 1 To UBound(ByteArray) Step 2
      If ByteArray(xx) <> 0 Then
         ByteArray(xx - 1) = 32     ' Replace uniCode character with a blank
         ByteArray(xx) = 0
      End If
   Next
   
   FixString = ByteArray
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.FixString.", Err.Number, Err.Description
   FixString = strVal
End Function

Public Function GetAuditEnabledStatusForSite(SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   If goSession.IsOracle Then
      sSQL = "select * from snAuditTrail where ROWNUM <= 1 AND OriginatingSiteKey =" & SiteKey
   Else
      sSQL = "select TOP 1 * from snAuditTrail where OriginatingSiteKey =" & SiteKey
   End If
  
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRs.RecordCount < 1 Then
      GetAuditEnabledStatusForSite = False
   Else
      GetAuditEnabledStatusForSite = True
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in snAuditTrail.GetAuditEnabledStatusForSite. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function
