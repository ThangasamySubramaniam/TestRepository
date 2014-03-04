VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwReplicateWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwReplicateWork
' Aug 2003 ms
'

Option Explicit
Const MAX_OLD_VALUE_TEST = 31



Public Function LogAddChange(ChangeTableKey As Long, RecordKey As Long, Optional SiteKey As Long, Optional RecFields As Fields) As Boolean
   Dim nChangeKey As Long
   Dim sSQL As String
   Dim loField As field
   Dim loCmd As Command
   Dim loPar As Parameter
   On Error GoTo FunctionError
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goSession.DbShapeConnection
   On Error GoTo FunctionError
   If RecFields Is Nothing Then
      '
      ' Track PK of ADD (will send all columns in datagram)
      '
      If SiteKey < 1 Then
         sSQL = "insert into mwrChangeLog (mwrChangeTableKey, ChangeType, RecordKey, mwcUsersKey, mwcRoleTypeKey, ChangeDateTime, OriginatingSiteKey) " & _
           "values(" & ChangeTableKey & ",'A'," & RecordKey & "," & goSession.User.UserKey & "," _
           & goSession.User.RoleTypeKey & ",? ," & goSession.Site.SiteKey & ")"
      Else
         sSQL = "insert into mwrChangeLog (mwrChangeTableKey, ChangeType, RecordKey, mwcSitesKey, mwcUsersKey, mwcRoleTypeKey, ChangeDateTime, OriginatingSiteKey) " & _
           "values(" & ChangeTableKey & ",'A'," & RecordKey & "," & SiteKey & "," & goSession.User.UserKey & _
           "," & goSession.User.RoleTypeKey & ",? ," & goSession.Site.SiteKey & ")"
      End If
      loCmd.CommandText = sSQL
      Set loPar = loCmd.CreateParameter("NowDate", adDate, adParamInput, , Now())
      loCmd.Parameters.Append loPar
      loCmd.Execute
      LogAddChange = True
   Else
      '
      ' Track Column changes on add transaction
      '
      'nChangeKey = WriteChangeHeader("A", ChangeTableKey, RecFields(0).value, goSession.Site.SiteKey, SiteKey)
      nChangeKey = WriteChangeHeader("A", ChangeTableKey, RecordKey, goSession.Site.SiteKey, SiteKey)
      For Each loField In RecFields
         If Not (loField.Type = adChapter Or loField.Type = adLongVarBinary) Then
            If Not IsNull(loField.value) Then
               If loField.value <> "" Then
                  WriteChangeColumn nChangeKey, loField
               End If
            End If
         End If
      Next loField
      LogAddChange = True
   End If
   'ValidateSiteKey ChangeTableKey, SiteKey
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.LogAddChange. ", Err.Number, Err.Description
   LogAddChange = False
End Function

Public Function LogDeleteChange(ChangeTableKey As Long, RecordKey As Long, Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loCmd As Command
   Dim loPar As Parameter
   On Error GoTo FunctionError
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goSession.DbShapeConnection
   
   If SiteKey = 0 Then
      sSQL = "insert into mwrChangeLog (mwrChangeTableKey, ChangeType, RecordKey, mwcUsersKey, mwcRoleTypeKey, ChangeDateTime, OriginatingSiteKey) " & _
        "values(" & ChangeTableKey & ",'D'," & RecordKey & "," & goSession.User.UserKey & "," & _
        goSession.User.RoleTypeKey & ",?" & "," & goSession.Site.SiteKey & ")"
   Else
      sSQL = "insert into mwrChangeLog (mwrChangeTableKey, ChangeType, RecordKey, mwcSitesKey, mwcUsersKey, mwcRoleTypeKey, ChangeDateTime, originatingSiteKey) " & _
        "values(" & ChangeTableKey & ",'D'," & RecordKey & "," & SiteKey & "," & goSession.User.UserKey & "," & _
        goSession.User.RoleTypeKey & ",?" & "," & goSession.Site.SiteKey & ")"
   End If
   loCmd.CommandText = sSQL
   Set loPar = loCmd.CreateParameter("NowDate", adDate, adParamInput, , Now())
   loCmd.Parameters.Append loPar
   loCmd.Execute
   'goCon.Execute sSQL
   Set loCmd = Nothing
   LogDeleteChange = True
   'ValidateSiteKey ChangeTableKey, SiteKey
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.LogDeleteChange. ", Err.Number, Err.Description
   LogDeleteChange = False
End Function


Public Function LogModifyChange(ChangeTableKey As Long, RecFields As Fields, Optional SiteKey As Long) As Boolean
   Dim IsChangeCreated As Boolean
   Dim nChangeKey As Long
   Dim sSQL As String
   Dim sSummary As String
   Dim loField As field
   Dim bWriteColWorked As Boolean
   
   On Error GoTo FunctionError
   For Each loField In RecFields
      '
      ' mwDT dates have "12:00:00 AM" as a null value...
      '
      'If loField.Type = adDate Then
      '   If loField.OriginalValue = "12:00:00 AM" Then
      '      loField.OriginalValue = Null
      '   End If
      '
      '   If loField.value = "12:00:00 AM" Then
      '      loField.value = Null
      '   End If
      'End If
      ' adChapter is RS pointer in a shaped RS
      '
      
      If Not loField.Type = adChapter And Not loField.Type = adLongVarBinary Then
         If loField.value = "" Then
            loField.value = Null
         End If
         If (loField.OriginalValue <> loField.value) Or (IsNull(loField.OriginalValue) And _
           Not IsNull(loField.value)) Or (IsNull(loField.value) And Not IsNull(loField.OriginalValue)) Then
            
            If Not IsChangeCreated Then
               nChangeKey = WriteChangeHeader("M", ChangeTableKey, RecFields(0).value, goSession.Site.SiteKey, SiteKey)
               If nChangeKey < 1 Then
                  LogModifyChange = False
                  Exit Function
               End If
               IsChangeCreated = True
            End If
            bWriteColWorked = WriteChangeColumn(nChangeKey, loField)
            
            If bWriteColWorked = False Then
               goSession.RaiseWorkflowStatus "mwReplicateWork.LogModifyChange mwrChangeLogKey = " & CStr(nChangeKey) & ", mwrChangeTableKey = " & CStr(ChangeTableKey)
            End If
         End If
      End If
   Next loField
   'MsgBox sSummary
   LogModifyChange = True
   'ValidateSiteKey ChangeTableKey, SiteKey
   Exit Function
FunctionError:
   ' Add Operation: Either BOF or EOF is True, or the current record has been deleted.
   ' Requested operation requires a current record.
   If Err.Number <> 3021 Then
      goSession.RaiseError "General Error in mwReplicateWork.LogModifyChange. ", Err.Number, Err.Description
   End If
   LogModifyChange = False
   'Resume Next
End Function

Public Function LogSchemaChange(RecFields As Fields) As Boolean
   Dim nChangeKey As Long
   Dim sSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   On Error GoTo FunctionError
   
   Dim RecordKey As Long
   Dim OriginatingSiteKey As Long
   Dim SiteKey As Long
      
   RecordKey = 0
   OriginatingSiteKey = goSession.Site.SiteKey
   SiteKey = 0
   
   nChangeKey = WriteChangeHeader("S", MWRT_mwcSchReplicate, RecordKey, OriginatingSiteKey, SiteKey)
   For Each loField In RecFields
      If Not (loField.Type = adChapter Or loField.Type = adLongVarBinary) Then
         If Not IsNull(loField.value) Then
            If loField.value <> "" Then
               WriteChangeColumn nChangeKey, loField
            End If
         End If
      End If
   Next loField
   LogSchemaChange = True
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.LogSchemaChange. ", Err.Number, Err.Description
   LogSchemaChange = False
End Function

Private Function WriteChangeColumnSQL(ChangeKey, RecField As field) As Boolean
   Dim sSQL1 As String
   Dim sSQL2 As String
   Dim sField As String
   Dim i As Integer
   On Error GoTo FunctionError
   '
   ' Check for compound column names (based on tables constructed using joins)
   '
   sField = RecField.Name
   i = InStr(1, RecField.Name, ".")
   If i > 0 Then
      sField = mID$(RecField.Name, i + 1)
   Else
      sField = RecField.Name
   End If
   sSQL1 = "insert into mwrChangeLogColumn (mwrChangeLogKey, ColumnName, ColumnDataType"
   sSQL2 = "values(" & ChangeKey & ",'" & sField & "'," & RecField.Type
   If Not IsNull(RecField.OriginalValue) Then
      If RecField.OriginalValue <> "" Then
         If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
            sSQL1 = sSQL1 & ", OldValue"
            sSQL2 = sSQL2 & ",'" & MassageField(RecField.OriginalValue) & "'"
         Else
            sSQL1 = sSQL1 & ", OldValueLarge"
            sSQL2 = sSQL2 & ",'" & MassageField(RecField.OriginalValue) & "'"
         End If
      End If
   End If
   If IsNull(RecField.value) Or (RecField.value = "" And RecField.Type = adLongVarWChar) Then
      If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
         sSQL1 = sSQL1 & ", NewValue"
         sSQL2 = sSQL2 & ",'[NULL]'"
      Else
         sSQL1 = sSQL1 & ", NewValueLarge"
         sSQL2 = sSQL2 & ",'[NULL]'"
      End If
   Else
      If RecField.value <> "" Then
         If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
            sSQL1 = sSQL1 & ", NewValue"
            sSQL2 = sSQL2 & ",'" & MassageField(RecField.value) & "'"
         Else
            sSQL1 = sSQL1 & ", NewValueLarge"
            sSQL2 = sSQL2 & ",'" & MassageField(RecField.value) & "'"
         End If
      End If
   End If
   sSQL1 = sSQL1 & ") " & sSQL2 & ")"
   goCon.Execute sSQL1
   WriteChangeColumnSQL = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.WriteChangeColumnSQL. ", Err.Number, Err.Description
   WriteChangeColumnSQL = False
End Function


Private Function WriteChangeHeader(ChangeType As String, ChangeTableKey As Long, RecordKey As Long, OriginatingSiteKey As Long, Optional SiteKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   With loRs
      .CursorLocation = adUseClient
      sSQL = "select * from mwrChangeLog where ID=-1"
      .Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      .AddNew
      .Fields("mwrChangeTableKey").value = ChangeTableKey
      .Fields("ChangeType").value = ChangeType
      .Fields("RecordKey").value = RecordKey
      If SiteKey > 0 Then
         .Fields("mwcSitesKey").value = SiteKey
      End If
      
      .Fields("OriginatingSiteKey").value = OriginatingSiteKey
      
      .Fields("mwcUsersKey").value = goSession.User.UserKey
      If goSession.User.RoleTypeKey > 0 Then
         .Fields("mwcRoleTypeKey").value = goSession.User.RoleTypeKey
      End If
      .Fields("ChangeDateTime").value = Now()
      .Update
      If goSession.IsOracle() Then
         .Resync adAffectCurrent, adResyncAllValues
      End If
      WriteChangeHeader = .Fields("ID").value
   End With
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.WriteChangeHeader. ", Err.Number, Err.Description
   WriteChangeHeader = -1
   CloseRecordset loRs
End Function


Private Function MassageField(str As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
'   strTemp = Replace(str, "'", " ")
   strTemp = Replace(str, vbCrLf, "|")
   
   If Len(Trim(strTemp)) = 0 Then
      strTemp = str
   End If
   MassageField = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.MassageField: ", Err.Number, Err.Description
   MassageField = str
End Function


Public Function ManualModifyStart(ChangeTableKey As Long, ChangeRecordKey As Long, Optional SiteKey As Long) As Long
   On Error GoTo FunctionError
   ManualModifyStart = WriteChangeHeader("M", ChangeTableKey, ChangeRecordKey, goSession.Site.SiteKey, SiteKey)
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.ManualModifyStart. ", Err.Number, Err.Description
   ManualModifyStart = -1
End Function

Public Function ManualModifyColumn(ChangeKey As Long, oField As field) As Boolean
   On Error GoTo FunctionError
   '
   ' mwDT dates have "12:00:00 AM" as a null value...
   '
   If oField.Type = adDate Then
      If oField.OriginalValue = "12:00:00 AM" Then
         oField.OriginalValue = Null
      End If
      
      If oField.value = "12:00:00 AM" Then
         oField.value = Null
      End If
   End If
   ' adChapter is RS pointer in a shaped RS
   If Not oField.Type = adChapter Then
      If (oField.OriginalValue <> oField.value) Or (IsNull(oField.OriginalValue) And _
       Not IsNull(oField.value)) Or (IsNull(oField.value) And Not IsNull(oField.OriginalValue)) Then
         WriteChangeColumn ChangeKey, oField
      End If
   End If
   ManualModifyColumn = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.ManualModifyColumn. ", Err.Number, Err.Description
   ManualModifyColumn = False
End Function


Public Function ManageWillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal TableKey As Long, _
 ByRef oRs As Recordset, SiteKey As Long) As Boolean
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      LogAddChange TableKey, oRs!ID, SiteKey, oRs.Fields
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      LogDeleteChange TableKey, oRs!ID, SiteKey
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      LogModifyChange TableKey, oRs.Fields, SiteKey
   End If
   Exit Function
SubError:
   goSession.RaisePublicError "General Error in mwReplicateWork.ManageWillChangeRecord. ", Err.Number, Err.Description
End Function


Private Function WriteChangeColumn(ChangeKey, RecField As field) As Boolean
   Dim sSQL1 As String
   Dim sSQL2 As String
   Dim sField As String
   Dim sDate As String
   Dim i As Integer
   Static loCmd As Command
   
   On Error GoTo FunctionError
   If ChangeKey < 1 Then
      Exit Function
   End If
   '
   ' Check for compound column names (based on tables constructed using joins)
   '
   If loCmd Is Nothing Then
      Set loCmd = New Command
      Set loCmd.ActiveConnection = goCon
      loCmd.CommandType = adCmdText
      loCmd.CommandText = "insert into mwrChangeLogColumn (mwrChangeLogKey, ColumnName, OldValue, " & _
       " NewValue, OldValueLarge, NewValueLarge, ColumnDataType) values (?, ?, ?, ?, ?, ?, ?)"
      loCmd.Parameters.Append loCmd.CreateParameter("CLK", adInteger, adParamInput)
      loCmd.Parameters.Append loCmd.CreateParameter("CN", adVarChar, adParamInput, 50)
      loCmd.Parameters.Append loCmd.CreateParameter("OV", adVarChar, adParamInput, 30)
      loCmd.Parameters.Append loCmd.CreateParameter("NV", adVarChar, adParamInput, 30)
      loCmd.Parameters.Append loCmd.CreateParameter("OVL", adLongVarWChar, adParamInput, 30000)
      loCmd.Parameters.Append loCmd.CreateParameter("NVL", adLongVarWChar, adParamInput, 30000)
      loCmd.Parameters.Append loCmd.CreateParameter("CDT", adInteger, adParamInput)
   End If
   loCmd.Parameters("NV").value = Null
   loCmd.Parameters("NVL").value = Null
   loCmd.Parameters("OVL").value = Null
   loCmd.Parameters("CN").value = Null
   loCmd.Parameters("OV").value = Null
   loCmd.Parameters("CLK").value = ChangeKey
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
   WriteChangeColumn = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.WriteChangeColumn. ", Err.Number, Err.Description
   goSession.RaiseWorkflowStatus "mwrChangeLogKey = " & CStr(ChangeKey) & ", Field Name = " & sField
   WriteChangeColumn = False
End Function

Private Function ValidateSiteKey(ChangeTableKey As Long, SiteKey As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   
   On Error GoTo FunctionError
   
   If SiteKey = 0 Then
   
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      
      sSQL = "select mwrBatchTypeKey, TableName from mwrChangeTable where ID=" & ChangeTableKey
      
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      
      If loRs.RecordCount > 0 Then
         If loRs!mwrBatchTypeKey = 100 Then
            goSession.RaiseError "Error in mwReplicateWork.ValidateSiteKey: Site Specific table " & loRs!TableName & " change logged but no SiteKey supplied.", Err.Number, Err.Description
            ValidateSiteKey = False
         End If
      End If
      
      CloseRecordset loRs
   End If
   
   ValidateSiteKey = True
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.ValidateSiteKey. ", Err.Number, Err.Description
   ValidateSiteKey = False
End Function

Public Function LogLogRequest(RecFields As Fields) As Long
   Dim nChangeKey As Long
   Dim sSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   
   Dim RecordKey As Long
   Dim OriginatingSiteKey As Long
   Dim SiteKey As Long
      
   RecordKey = 0
   OriginatingSiteKey = goSession.Site.SiteKey
   SiteKey = 0
   
   
   If RecFields(0).Name = "RequestedSiteKey" Then
      SiteKey = ZeroNull(RecFields(0).value)  ' add sitekey for Fleet dgm to use only RequestedSite body to return log
   End If
   
   '
   'nChangeKey = WriteChangeHeader("L", MWRT_mwcSchReplicate, RecordKey, OriginatingSiteKey, SiteKey)
   nChangeKey = WriteChangeHeader("L", MWRT_mwcSchReplicateShip, RecordKey, OriginatingSiteKey, SiteKey)
   
   For Each loField In RecFields
      If Not (loField.Type = adChapter Or loField.Type = adLongVarBinary) Then
         If Not IsNull(loField.value) Then
            If loField.value <> "" Then
               WriteChangeColumn nChangeKey, loField
            End If
         End If
      End If
   Next loField
   LogLogRequest = nChangeKey
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWork.LogLogRequest. ", Err.Number, Err.Description
   LogLogRequest = False
End Function

