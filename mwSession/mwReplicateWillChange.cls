VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwReplicateWillChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' mwWillChange - Service Class to manage replication.
'
Option Explicit
Private mIsSiteTable As Boolean
Private mIsValidateSiteKey As Boolean
Private mTableKey As Long
Private mTableName As String
Private mIsCancelled As Boolean
Private mIsNonReplicationTable As Boolean
Private mIsAuditTrail As Boolean
Dim moAuditTrail As snAuditTrail

Const BATCH_TYPE_NO_REPLICATION = 0
Const BATCH_TYPE_SITE_SPECIFIC = 100
Const MAX_OLD_VALUE_TEST = 31

Public Function Initialize(TableName As String) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   Dim moKeys As ConfigKeys
   On Error GoTo FunctionError
   '
   ' Validate Table
   '
   If goSession.IsOracle Then
      sSQL = "select ID, mwrBatchTypeKey,IsAuditTrailEnabled from mwrChangeTable where upper(TableName)='" & UCase(TableName) & "'"
   Else
      sSQL = "select ID, mwrBatchTypeKey,IsAuditTrailEnabled from mwrChangeTable where TableName='" & TableName & "'"
   End If
   Set loRs = New Recordset
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.Initialize, table missing: " & TableName
      mIsCancelled = False
      Initialize = False
      mTableKey = -1
      mIsNonReplicationTable = False
   Else
      mTableKey = loRs!ID
      mTableName = TableName
      '
      ' Site Specific Table ?
      '
      If mTableKey = MWRT_mwEventLog Or _
       mTableKey = MWRT_mwEventFactLogSN Or _
       mTableKey = MWRT_mwEventLinkLog Or _
       mTableKey = MWRT_mwEventFormLog Then
         ' Old VRS tables
         mIsSiteTable = True
         mIsValidateSiteKey = False
      ElseIf ZeroNull(loRs!mwrBatchTypeKey) = BATCH_TYPE_SITE_SPECIFIC Then
         mIsSiteTable = True
         mIsValidateSiteKey = True
      Else
         mIsSiteTable = False
         mIsValidateSiteKey = False
      End If
      
      If ZeroNull(loRs!mwrBatchTypeKey) = BATCH_TYPE_NO_REPLICATION Then
         mIsNonReplicationTable = True
         mIsSiteTable = False
         mIsValidateSiteKey = False
      Else
         mIsNonReplicationTable = False
      End If

      mIsAuditTrail = BoolNull(loRs!IsAuditTrailEnabled)
      
      If mIsAuditTrail = True Then
         Set moKeys = New ConfigKeys
         Set moKeys = goSession.GetEventSecurityKeys(MW_EVENTTYPE_ADMIN_CFG_SWITCHES)
         mIsAuditTrail = moKeys.GetBoolKeyValue("IsLogSafetyAuditTrailDetails")
         Set moKeys = Nothing
      End If
      
      Initialize = True
   End If
   
   Set moAuditTrail = New snAuditTrail
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwReplicateWillChange.Initialize. ", Err.Number, Err.Description
End Function


Public Function WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, _
 adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset) As Boolean
   Static IsValidated As Boolean
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   Static nCurrentChangeKey As Long

   On Error GoTo FunctionError
   If mIsCancelled Then
      adStatus = adStatusCancel
      WillChangeRecord = False
      Exit Function
   End If
   If Not IsValidated Then
      '
      '
      If Len(pRecordset.Source) > 0 Then
         If Not InStr(1, pRecordset.Source, mTableName, vbTextCompare) > 0 Then
            goSession.LogIt mwl_Workstation, mwl_Warning, "Warning in mwSession.mwReplicateWillChange.WillChangeRecord, table name missing from SQL Source: " & mTableName & " - " & pRecordset.Source
         End If
      End If
      
      ' pkgPackageSite is not replicated, let it generate errors if someone tries to log chages to it.
      
'      If mTableName = "pkgPackageSite" Then
'         If InStr(1, pRecordset.Fields(0).Name, "pkgPackageKey", vbTextCompare) <= 0 Then
'            goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.WillChangeRecord, Field(0) is not named pkgPackageKey! " & mTableName & " Field(0) Name = " & pRecordset.Fields(0).Name
'         End If
'      Else
      If mIsNonReplicationTable = False Then
      
         If InStr(1, pRecordset.Fields(0).Name, "ID", vbTextCompare) <= 0 Then
            goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.WillChangeRecord, Field(0) is not named ID! " & mTableName & " Field(0) Name = " & pRecordset.Fields(0).Name
'         adStatus = adStatusCancel
'         WillChangeRecord = False
'         mIsCancelled = True
'         Exit Function
         End If
      End If
      
      IsValidated = True
   End If
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
      nCurrentChangeKey = -1
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      
      If mIsValidateSiteKey Then
         If ZeroNull(pRecordset!mwcSitesKey) < 1 Then
            goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.WillChangeRecord, SiteKey is missing in query for Table: " & mTableName
            'adStatus = adStatusCancel
            'WillChangeRecord = False
            'mIsCancelled = True
            'Exit Function
         End If
      End If
      
      nCurrentChangeKey = LogAddChange(mTableKey, pRecordset.Fields(0).value, pRecordset.Fields)
      IsBeginAdd = False
      
      If nCurrentChangeKey > 1 Then
         If mIsAuditTrail Then
            moAuditTrail.WriteAuditTrailFromChangeLog nCurrentChangeKey
         End If
      End If
      
   ElseIf adReason = adRsnDelete Then
      If mIsSiteTable Then
            
         
         If mIsValidateSiteKey Then
            If ZeroNull(pRecordset!mwcSitesKey) < 1 Then
               goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.WillChangeRecord, SiteKey is missing in query for Table: " & mTableName
               'adStatus = adStatusCancel
               'WillChangeRecord = False
               'mIsCancelled = True
               'Exit Function
            End If
         End If
         nCurrentChangeKey = LogDeleteChange(mTableKey, pRecordset.Fields(0).value, ZeroNull(pRecordset!mwcSitesKey))
      Else
         nCurrentChangeKey = LogDeleteChange(mTableKey, pRecordset.Fields(0).value)
      End If
      IsBeginDelete = True
      
      If nCurrentChangeKey > 1 Then
         If mIsAuditTrail Then
            moAuditTrail.WriteAuditTrailFromChangeLog nCurrentChangeKey
         End If
      End If
      
   ElseIf adReason = adRsnUndoAddNew Then
      DeleteTransaction nCurrentChangeKey
   ElseIf adReason = adRsnUndoDelete Then
      DeleteTransaction nCurrentChangeKey
   ElseIf adReason = adRsnUndoUpdate Then
      DeleteTransaction nCurrentChangeKey
   ElseIf adReason <> adRsnFirstChange Then
      
      If mIsValidateSiteKey Then
         If ZeroNull(pRecordset!mwcSitesKey) < 1 Then
            goSession.RaisePublicError "Error in mwSession.mwReplicateWillChange.WillChangeRecord, SiteKey is missing in query for Table: " & mTableName
            'adStatus = adStatusCancel
            'WillChangeRecord = False
            'mIsCancelled = True
            'Exit Function
         End If
      End If
      
      nCurrentChangeKey = LogModifyChange(mTableKey, pRecordset.Fields)
      
      If nCurrentChangeKey > 1 Then
         If mIsAuditTrail Then
            moAuditTrail.WriteAuditTrailFromChangeLog nCurrentChangeKey
         End If
      End If
      
   Else
      nCurrentChangeKey = -1
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwReplicateWillChange.WillChangeRecord. Table = " & mTableName & " ", Err.Number, Err.Description
   adStatus = adStatusCancel
   WillChangeRecord = False
   mIsCancelled = True
End Function


Private Function LogAddChange(ChangeTableKey As Long, RecordKey As Long, RecFields As Fields) As Long
   Dim loField As field
   On Error GoTo FunctionError
   '
   ' Track Column changes on add transaction
   '
   If mIsNonReplicationTable = True Then
      LogAddChange = 1
   Else
      If mIsSiteTable Then
         LogAddChange = WriteChangeHeader("A", ChangeTableKey, RecordKey, ZeroNull(RecFields!mwcSitesKey))
      Else
         LogAddChange = WriteChangeHeader("A", ChangeTableKey, RecordKey)
      End If
      If LogAddChange > 0 Then
         For Each loField In RecFields
            If Not (loField.Type = adChapter Or loField.Type = adLongVarBinary) Then
               If Not IsNull(loField.value) Then
                  If loField.value <> "" Then
                     WriteChangeColumn LogAddChange, loField
                  End If
               End If
            End If
         Next loField
      End If
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWillChange.LogAddChange. ", Err.Number, Err.Description
   LogAddChange = -1
End Function

Private Function LogDeleteChange(ChangeTableKey As Long, RecordKey As Long, Optional SiteKey As Long) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   If mIsNonReplicationTable = True Then
      LogDeleteChange = 1
   Else
   
      Set loRs = New Recordset
      sSQL = "select * from mwrChangeLog where ID=-1"
      
      With loRs
         .CursorLocation = adUseClient
         .Open sSQL, goCon, adOpenDynamic, adLockOptimistic
         .AddNew
         If SiteKey > 0 Then
            .Fields!mwcSitesKey = SiteKey
         End If
         .Fields!mwrChangeTableKey = ChangeTableKey
         .Fields!ChangeType = "D"
         .Fields!RecordKey = RecordKey
         .Fields!mwcUsersKey = goSession.User.UserKey
         .Fields!mwcRoleTypekey = goSession.User.RoleTypeKey
         .Fields!ChangeDateTime = Now()
         .Fields!OriginatingSiteKey = goSession.Site.SiteKey
         .Update
         LogDeleteChange = .Fields!ID
         .Close
      End With
      Set loRs = Nothing
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWillChange.LogDeleteChange. ", Err.Number, Err.Description
   LogDeleteChange = -1
End Function


Private Function LogModifyChange(ChangeTableKey As Long, RecFields As Fields) As Long
   Dim IsChangeCreated As Boolean
   Dim loField As field
   On Error GoTo FunctionError
   
   If mIsNonReplicationTable = True Then
      LogModifyChange = 1
   Else
      For Each loField In RecFields
         
         If Not loField.Type = adChapter And Not loField.Type = adLongVarBinary Then
            If loField.value = "" Then
               loField.value = Null
            End If
            If (loField.OriginalValue <> loField.value) Or (IsNull(loField.OriginalValue) And _
              Not IsNull(loField.value)) Or (IsNull(loField.value) And Not IsNull(loField.OriginalValue)) Then
               
               If Not IsChangeCreated Then
                  If mIsSiteTable = True Then
                     LogModifyChange = WriteChangeHeader("M", ChangeTableKey, RecFields(0).value, ZeroNull(RecFields!mwcSitesKey))
                  Else
                     LogModifyChange = WriteChangeHeader("M", ChangeTableKey, RecFields(0).value)
                  End If
               
                  If LogModifyChange < 1 Then
   '                  LogModifyChange = False
                     Exit Function
                  End If
                  IsChangeCreated = True
               End If
               WriteChangeColumn LogModifyChange, loField
            End If
         End If
      Next loField
   End If
'   LogModifyChange = True
   Exit Function
FunctionError:
' Add Operation: Either BOF or EOF is True, or the current record has been deleted.
' Requested operation requires a current record.
'   If Err.Number <> 3021 Then
      goSession.RaiseError "General Error in mwReplicateWillChange.LogModifyChange. ", Err.Number, Err.Description
'   End If
   LogModifyChange = -1
End Function

Private Function WriteChangeHeader(ChangeType As String, ChangeTableKey As Long, RecordKey As Long, _
 Optional SiteKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   Dim NumTries As Integer
   On Error GoTo FunctionError
   
   If mIsNonReplicationTable = True Then
      WriteChangeHeader = 1
      Exit Function
   Else
   
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      sSQL = "select * from mwrChangeLog where ID=-1"
      loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      NumTries = 10
      On Error Resume Next
      Do While NumTries > 0
         loRs.AddNew
         If Err.Number = 0 Then
            loRs.Fields("mwrChangeTableKey").value = ChangeTableKey
            loRs.Fields("ChangeType").value = ChangeType
            loRs.Fields("RecordKey").value = RecordKey
            If SiteKey > 0 Then
               loRs.Fields("mwcSitesKey").value = SiteKey
            End If
            
            loRs.Fields("OriginatingSiteKey").value = goSession.Site.SiteKey
            
            loRs.Fields("mwcUsersKey").value = goSession.User.UserKey
            If goSession.User.RoleTypeKey > 0 Then
               loRs.Fields("mwcRoleTypeKey").value = goSession.User.RoleTypeKey
            End If
            loRs.Fields("ChangeDateTime").value = Now()
            loRs.Update
            If Err.Number = 0 Then
               '
               ' OKEY DOKEY
               '
               If goSession.IsOracle() Then
                  loRs.Resync adAffectCurrent, adResyncAllValues
               End If
               WriteChangeHeader = loRs.Fields("ID").value
               CloseRecordset loRs
               Exit Function
            Else
               goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwReplicateWillChange.WriteChangeHeader, .Update Failed on Try: " & NumTries & "-" & Err.Number & "-" & Err.Description
               loRs.CancelUpdate
               NumTries = NumTries - 1
               goSession.API.SleepApp 1000
            End If
         Else
            goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwReplicateWillChange.WriteChangeHeader, .AddNew Failed on Try: " & NumTries & "-" & Err.Number & "-" & Err.Description
            loRs.CancelUpdate
            NumTries = NumTries - 1
            goSession.API.SleepApp 1000
         End If
      Loop
   End If
   '
   ' Falls Trough to Error Handler
   '
FunctionError:
   goSession.RaiseError "Error in mwReplicateWillChange.WriteChangeHeader, Failed to Write Record after 10 tries. ChangeType = " & ChangeType & " TableName = " & mTableName & " RecordKey = " & RecordKey
   WriteChangeHeader = -1
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
   goSession.RaiseError "General Error in mwReplicateWillChange.MassageField: ", Err.Number, Err.Description
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
   goSession.RaiseError "General Error in mwReplicateWillChange.FixString.", Err.Number, Err.Description
   FixString = strVal
End Function

Private Function WriteChangeColumn(ChangeKey, RecField As field) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   Dim sField As String
   Dim sBoolean As String
   Dim sTemp As String
   Dim i As Integer
   Dim NumTries As Integer
   Dim sDate As String
   
   On Error GoTo FunctionError
   If ChangeKey < 1 Then
      Exit Function
   End If
   
   sField = RecField.Name
   i = InStr(1, RecField.Name, ".")
   If i > 0 Then
      sField = mID$(RecField.Name, i + 1)
   Else
      sField = RecField.Name
   End If
            
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwrChangeLogColumn where ID=-1"
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   NumTries = 10
   On Error Resume Next
   Do While NumTries > 0
      loRs.AddNew
      If Err.Number = 0 Then
      
         loRs!ColumnName = sField
         loRs!mwrChangeLogKey = ChangeKey
         loRs!ColumnDataType = RecField.Type
         
         If Not IsNull(RecField.OriginalValue) Then
            If RecField.OriginalValue <> "" Then
               If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
                  If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
                     sDate = Format(RecField.OriginalValue, DATAGRAM_DATE_FORMAT)
            
                     loRs!OldValue = MassageField(RecField.OriginalValue)
            
                  ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
                     loRs!OldValue = TranslateBoolean(RecField.OriginalValue)
                  Else
                     loRs!OldValue = MassageField(RecField.OriginalValue)
                  End If
                  loRs!OldValueLarge = Null
               Else
                  loRs!OldValueLarge = MassageField(RecField.OriginalValue)
                  loRs!OldValue = Null
               End If
            End If
         Else
            loRs!OldValueLarge = Null
            loRs!OldValue = Null
         End If
         
         If IsNull(RecField.value) Or (RecField.value = "" And RecField.Type = adLongVarWChar) Then
            If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
               loRs!NewValue = "[NULL]"
            Else
               loRs!NewValueLarge = "[NULL]"
            End If
         Else
            If RecField.value <> "" Then
               If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
                  If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
                     sDate = Format(RecField.value, DATAGRAM_DATE_FORMAT)
                     loRs!NewValue = sDate
                  ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
                     loRs!NewValue = TranslateBoolean(RecField.value)
                  ElseIf RecField.Type = adDouble Or RecField.Type = adNumeric Then       ' (5)
                     sTemp = Format(RecField.value, "#########0.0##")
                     ' Euro Symbol crap...
                     sTemp = Replace(sTemp, ",", ".")
                     loRs!NewValue = sTemp
                  Else
                     loRs!NewValue = MassageField(RecField.value)
                  End If
                  loRs!NewValueLarge = Null
               Else
                  loRs!NewValueLarge = MassageField(RecField.value)
                  loRs!NewValue = Null
               End If
            End If
         End If
         
         
         loRs.Update
         If Err.Number = 0 Then
            '
            ' OKEY DOKEY
            '
            If goSession.IsOracle() Then
               loRs.Resync adAffectCurrent, adResyncAllValues
            End If
            WriteChangeColumn = True
            CloseRecordset loRs
            Exit Function
         Else
            goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwReplicateWillChange.WriteChangeColumn, .Update Failed on Try: " & NumTries & "-" & Err.Number & "-" & Err.Description
            loRs.CancelUpdate
            NumTries = NumTries - 1
            goSession.API.SleepApp 1000
         End If
      Else
         goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwReplicateWillChange.WriteChangeColumn, .Update Failed on Try: " & NumTries & "-" & Err.Number & "-" & Err.Description
         loRs.CancelUpdate
         NumTries = NumTries - 1
         goSession.API.SleepApp 1000
      End If
   Loop
   
   goSession.RaiseError "Error in mwReplicateWillChange.WriteChangeColumn, Failed to Write Record after 10 tries. Field = " & RecField.Name & " ", Err.Number, Err.Description
   WriteChangeColumn = False
   
   Exit Function

FunctionError:
   
   goSession.RaiseError "General Error in mwReplicateWillChange.WriteChangeColumn. Failed to insert ChangeLogColumn for Field - " & RecField.Name & " ", Err.Number, Err.Description
   
   WriteChangeColumn = False
End Function


'Private Function WriteChangeColumn(ChangeKey, RecField As field) As Boolean
'   Dim sSQL1 As String
'   Dim sSQL2 As String
'   Dim sField As String
'   Dim sBoolean As String
'   Dim sTemp As String
'   Dim i As Integer
'   Dim NumTries As Integer
'   Dim sDate As String
'   Dim nSizeOfLongFields As Long
'
'   Static loCmd As Command
'   On Error GoTo FunctionError
'   If ChangeKey < 1 Then
'      Exit Function
'   End If
'   '
'   ' Check for compound column names (based on tables constructed using joins)
'   '
'
'   nSizeOfLongFields = MAX_OLD_VALUE_TEST
'
'   If RecField.ActualSize > MAX_OLD_VALUE_TEST Then
'      nSizeOfLongFields = RecField.ActualSize + 32
'      goSession.LogIt mwl_User_Defined, mwl_Warning, " "
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Field = " & RecField.Name
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Type = " & RecField.Type
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Long ActualSize+32 =  " & nSizeOfLongFields
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Long DefinedSize   =  " & RecField.DefinedSize
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Len(OriginalValue) =  " & Len(BlankNull(RecField.OriginalValue))
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Len(value)         =  " & Len(BlankNull(RecField.value))
'   End If
'
'
'   If loCmd Is Nothing Then
'      Set loCmd = New Command
'      Set loCmd.ActiveConnection = goCon
'      loCmd.CommandType = adCmdText
'      loCmd.CommandText = "insert into mwrChangeLogColumn (mwrChangeLogKey, ColumnName, OldValue, " & _
'       " NewValue, OldValueLarge, NewValueLarge, ColumnDataType) values (?, ?, ?, ?, ?, ?, ?)"
'      loCmd.Parameters.Append loCmd.CreateParameter("CLK", adInteger, adParamInput)
'      loCmd.Parameters.Append loCmd.CreateParameter("CN", adVarChar, adParamInput, 50)
'      loCmd.Parameters.Append loCmd.CreateParameter("OV", adVarChar, adParamInput, 30)
'      loCmd.Parameters.Append loCmd.CreateParameter("NV", adVarChar, adParamInput, 30)
'
'      loCmd.Parameters.Append loCmd.CreateParameter("OVL", adLongVarWChar, adParamInput, nSizeOfLongFields)
'      loCmd.Parameters.Append loCmd.CreateParameter("NVL", adLongVarWChar, adParamInput, nSizeOfLongFields)
'
'      loCmd.Parameters.Append loCmd.CreateParameter("CDT", adInteger, adParamInput)
'   End If
'   loCmd.Parameters("NV").value = Null
'   loCmd.Parameters("NVL").value = Null
'   loCmd.Parameters("OVL").value = Null
'   loCmd.Parameters("CN").value = Null
'   loCmd.Parameters("OV").value = Null
'   loCmd.Parameters("CLK").value = ChangeKey
'   loCmd.Parameters("CDT").value = RecField.Type
'   sField = RecField.Name
'   i = InStr(1, RecField.Name, ".")
'   If i > 0 Then
'      sField = mID$(RecField.Name, i + 1)
'   Else
'      sField = RecField.Name
'   End If
'   loCmd.Parameters("CN").value = sField
'   '
'   If Not IsNull(RecField.OriginalValue) Then
'      If RecField.OriginalValue <> "" Then
'         If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
'            If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
'               sDate = Format(RecField.OriginalValue, DATAGRAM_DATE_FORMAT)
'               loCmd.Parameters("OV").value = MassageField(RecField.OriginalValue)
'            ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
'               loCmd.Parameters("OV").value = TranslateBoolean(RecField.OriginalValue)
'            Else
'               loCmd.Parameters("OV").value = MassageField(RecField.OriginalValue)
'            End If
'            loCmd.Parameters("OVL").value = Null
'         Else
'            loCmd.Parameters("OVL").value = MassageField(RecField.OriginalValue)
'            loCmd.Parameters("OV").value = Null
'         End If
'      End If
'   Else
'      loCmd.Parameters("OVL").value = Null
'      loCmd.Parameters("OV").value = Null
'   End If
'   If IsNull(RecField.value) Or (RecField.value = "" And RecField.Type = adLongVarWChar) Then
'      If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
'         loCmd.Parameters("NV").value = "[NULL]"
'      Else
'         loCmd.Parameters("NVL").value = "[NULL]"
'      End If
'   Else
'      If RecField.value <> "" Then
'         If RecField.DefinedSize < MAX_OLD_VALUE_TEST Then
'            If RecField.Type = adDate Or RecField.Type = adDBDate Or RecField.Type = adDBTime Or RecField.Type = adDBTimeStamp Then
'               sDate = Format(RecField.value, DATAGRAM_DATE_FORMAT)
'               loCmd.Parameters("NV").value = sDate
'            ElseIf RecField.Type = adBoolean Or (RecField.Type = adNumeric And RecField.Precision = 1 And goSession.IsOracle) Then
'               loCmd.Parameters("NV").value = TranslateBoolean(RecField.value)
'            ElseIf RecField.Type = adDouble Or RecField.Type = adNumeric Then       ' (5)
'               sTemp = Format(RecField.value, "#########0.0##")
'               ' Euro Symbol crap...
'               sTemp = Replace(sTemp, ",", ".")
'               loCmd.Parameters("NV").value = sTemp
'            Else
'               loCmd.Parameters("NV").value = MassageField(RecField.value)
'            End If
'            loCmd.Parameters("NVL").value = Null
'         Else
'            loCmd.Parameters("NVL").value = MassageField(RecField.value)
'            loCmd.Parameters("NV").value = Null
'         End If
'      End If
'   End If
'   NumTries = 10
'   On Error Resume Next
'   Do While NumTries > 0
'      loCmd.Execute
'      If Err.Number = 0 Then
'         WriteChangeColumn = True
'         Exit Function
'      End If
'      goSession.LogIt mwl_User_Defined, mwl_Warning, "Warning in mwReplicateWillChange.WriteChangeColumn, .Execute Failed on Try: " & NumTries & "-" & Err.Number & "-" & Err.Description
'      NumTries = NumTries - 1
'      goSession.API.SleepApp 1000
'   Loop
'   '
'   ' Fall through to error handler...
'   '
'FunctionError:
'
'   goSession.RaiseError "General Error in mwReplicateWillChange.WriteChangeColumn. Field - " & RecField.Name & " ", Err.Number, Err.Description
'
'   WriteChangeColumn = False
'End Function

Private Function DeleteTransaction(ChangeLogKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   If mIsNonReplicationTable = True Then
      DeleteTransaction = True
   Else
      sSQL = "delete from mwrChangeLogColumn where mwrChangeLogKey=" & ChangeLogKey
      goCon.Execute sSQL
      sSQL = "delete from mwrChangeLog where ID=" & ChangeLogKey
      goCon.Execute sSQL
      DeleteTransaction = True
      
      moAuditTrail.DeleteAuditTrailTransaction ChangeLogKey
      
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWillChange.DeleteTransaction. ", Err.Number, Err.Description
End Function

Private Sub Class_Terminate()
   On Error GoTo SubError

   KillObject moAuditTrail
   Exit Sub
SubError:
   goSession.RaiseError "Error in mwReplicateWillChange.Class_Terminate", Err.Number, Err.Description
End Sub
