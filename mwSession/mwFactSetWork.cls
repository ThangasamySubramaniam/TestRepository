VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwFactSetWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' Fetch RS offsets
'
Const RS_FSC_ID = 0
Const RS_FSC_mwFactSetKey = 1
Const RS_FSC_ReferenceTitle = 2
Const RS_FSC_MainTitle = 3
Const RS_FSC_AdoDataType = 4
Const RS_FSC_AdoDataTypeWidth = 5
Const RS_FSC_RecordsetColumnName = 6
Const RS_FSC_IsAddToRemarks = 7
Const RS_FSC_IsMandatory = 8
Const RS_FSC_MaxTextWidth = 9
Const RS_FSC_NumericMinValue = 10
Const RS_FSC_NumericMaxValue = 11
Const RS_FSC_IsActive = 12
Const RS_FSC_CustomEditCode = 13

Const RS_FLC_ID = 14
Const RS_FLC_mwFactSetColumnKey = 15
Const RS_FLC_mwFactLogKey = 16
Const RS_FLC_LogValue = 17




Public Function FetchRs(FactLogKey As Long) As Recordset
   Dim loRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   strSQL = "SELECT mwFactSetColumn.*, mwFactLogColumn.*" & _
'     " FROM mwFactSetColumn INNER JOIN mwFactLogColumn ON mwFactSetColumn.ID = " & _
'     " mwFactLogColumn.mwFactSetColumnKey where mwFactLogKey=" & FactLogKey & _
'     " order by mwFactSetColumn.ReferenceTitle"
   
   strSQL = "SELECT mwFactSetColumn.*, mwFactLogColumn.*" & _
      " FROM mwFactSetColumn , mwFactLogColumn " & _
      " WHERE mwFactSetColumn.ID = mwFactLogColumn.mwFactSetColumnKey AND " & _
      " mwFactLogKey = " & FactLogKey & _
      " ORDER BY mwFactSetColumn.ReferenceTitle"
   
   loRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = ""
   Set FetchRs = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwFactSetWork.FetchRs. ", err.Number, err.Description
   CloseRecordset loRs
End Function

' loRsFactLogColumn contains a bunch of records to role into one
Public Function FetchAsMergedRs(FactLogKey As Long, Optional RemarksField As String) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loField As field
   Dim loRsFactLog As Recordset
   Dim loRsFactLogColumn As Recordset
   Dim sTemp As String
   Dim sDecimalSeparator As String
   On Error GoTo FunctionError
   '
   sDecimalSeparator = goSession.API.GetLocaleString(LOCALE_DECIMAL_SEPARATOR)
   Set loRsFactLog = New Recordset
   loRsFactLog.CursorLocation = adUseClient
   ' Start with mwFactLog record...
   sSQL = "select * from mwFactLog where ID=" & FactLogKey
   loRsFactLog.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsFactLog.RecordCount < 1 Then
      goSession.RaiseError "Error in mwFactSetWork.a, No matching records found."
      CloseRecordset loRsFactLog
      Exit Function
   End If
   ' Target recordset...
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   '
   ' Append columns from the "header"
   '
   For Each loField In loRsFactLog.Fields
      If goSession.IsOracle And loField.Type = adNumeric Then
         loRs.Fields.Append loField.Name, adInteger, 4
      Else
         loRs.Fields.Append loField.Name, loField.Type, loField.DefinedSize
      End If
   Next loField
   '
   ' Maybe append Remarks...
   '
   If RemarksField <> "" Then
      loRs.Fields.Append "REMARKS", adLongVarWChar, 536870910
   End If
   '
   ' Get the children
   '
   Set loRsFactLogColumn = FetchRs(FactLogKey)
   
   If loRsFactLogColumn.RecordCount < 1 Then
      goSession.RaiseError "Error in mwFactSetWork.a, No matching Column records found."
      CloseRecordset loRs
      CloseRecordset loRsFactLog
      CloseRecordset loRsFactLogColumn
      Exit Function
   End If
   ' Add columns to fabricated recordset
   With loRsFactLogColumn
      .MoveFirst
      Do While Not .EOF
         If Not IsNull(.Fields(RS_FSC_RecordsetColumnName)) Then
            If IsNull(.Fields(RS_FSC_RecordsetColumnName)) Then
               loRs.Fields.Append .Fields(RS_FSC_RecordsetColumnName), _
                 .Fields(RS_FSC_AdoDataType)
            Else
               loRs.Fields.Append .Fields(RS_FSC_RecordsetColumnName), _
                 .Fields(RS_FSC_AdoDataType), .Fields(RS_FSC_AdoDataTypeWidth)
            End If
         End If
         .MoveNext
      Loop
      '
      loRs.Open
      loRs.AddNew
      '
      ' Add Header information...
      '
      For Each loField In loRsFactLog.Fields
         If Not IsNull(loField.value) Then
            loRs.Fields(loField.Name) = loField.value
         End If
      Next loField
      '
      ' Maybe Add Remarks...
      '
      If RemarksField <> "" Then
         loRs.Fields("REMARKS") = RemarksField
      End If
      
      '
      ' Add column records...
      '
      .MoveFirst
      Do While Not .EOF
         If Not IsNull(.Fields(RS_FSC_RecordsetColumnName)) Then
            sTemp = .Fields(RS_FSC_RecordsetColumnName)
            If Not IsNull(.Fields(RS_FLC_LogValue)) And .Fields(RS_FLC_LogValue) <> "" Then
               '
               ' Euro Conversion required for numbers !!!!
               '
               If loRs.Fields(sTemp).Type = adDouble Then
                  .Fields(RS_FLC_LogValue) = Replace(.Fields(RS_FLC_LogValue), ",", sDecimalSeparator)
                  .Fields(RS_FLC_LogValue) = Replace(.Fields(RS_FLC_LogValue), ".", sDecimalSeparator)
                  loRs.Fields(sTemp).value = CDbl(.Fields(RS_FLC_LogValue))
               Else
                  loRs.Fields(sTemp).value = .Fields(RS_FLC_LogValue)
               End If
            End If
         End If
         .MoveNext
      Loop
      loRs.Update
   End With
   Set FetchAsMergedRs = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwFactSetWork.FetchAsMergedRs. ", err.Number, err.Description
   CloseRecordset loRs
End Function

Public Function AddFactSet(FactSetKey As Long, EventTypeKey As Long, EventDetailKey As Long) As Long
   Dim sSQL As String
   Dim nNewSetKey As Long
   Dim nNewColumnKey As Long
   ' Source
   Dim loRs As Recordset
   Dim loRsSource As Recordset
   Dim loRsTarget As Recordset
   On Error GoTo FunctionError
   '
   ' Create Header record
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwFactlog where mwFactSetKey= -1"
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   With loRs
      .AddNew
      .Fields("ID") = goSession.MakePK("mwFactLog")
      nNewSetKey = .Fields("ID")
      .Fields("mwFactSetKey") = FactSetKey
      .Fields("mwEventTypeKey") = EventTypeKey
      If EventDetailKey > 0 Then
         .Fields("mwEventDetailKey") = EventDetailKey
      End If
      .Fields("Vessel_ID") = goSession.Site.SiteID
      .Fields("Voyage_No") = goSession.Site.CurrentVoyage
      .Fields("Report_DateTime") = Now()
      .Fields("Sent") = False
      .Update
   End With
   'CloseRecordset loRs
   '
   ' Copy template records into log...
   '
   sSQL = "select * from mwFactSetColumn where mwFactSetKey=" & FactSetKey & _
     " and IsActive<>0"
   Set loRsSource = New Recordset
   loRsSource.CursorLocation = adUseClient
   loRsSource.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If loRsSource.RecordCount < 1 Then
      goSession.RaiseError "Error in mwFactSetWork.AddFactSet, no columns found for Fact Set: " & FactSetKey
      loRs.Delete
      CloseRecordset loRs
      CloseRecordset loRsSource
      AddFactSet = -1
      Exit Function
   End If
   Set loRsTarget = New Recordset
   loRsTarget.CursorLocation = adUseClient
   sSQL = "select * from mwFactLogColumn where ID = -1"
   loRsTarget.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   nNewColumnKey = goSession.MakePK("mwFactLogColumn")
   '
   ' Create target records...
   '
   Do While Not loRsSource.EOF
      With loRsTarget
         .AddNew
         .Fields("ID") = nNewColumnKey
         .Fields("mwFactSetColumnKey") = loRsSource("ID")
         .Fields("mwFactLogKey") = nNewSetKey
         .Update
      End With
      nNewColumnKey = nNewColumnKey + 1
      loRsSource.MoveNext
   Loop
   If nNewColumnKey > 0 Then
      goSession.UpdatePrimaryKeySequence "mwFactLogColumn", nNewColumnKey
   End If
   CloseRecordset loRsSource
   CloseRecordset loRs
   CloseRecordset loRsSource
   AddFactSet = nNewSetKey
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwFactSetWork.AddFactSet. ", err.Number, err.Description
   CloseRecordset loRsSource
   CloseRecordset loRs
   CloseRecordset loRsSource
   AddFactSet = -1
End Function

Public Function SetNewEventKey(FactLogKey As Long, NewEventKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwFactlog where ID=" & FactLogKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwFactSetWork.SetNewEventKey, mwFactLog record not found. ", err.Number, err.Description
      CloseRecordset loRs
      SetNewEventKey = False
      Exit Function
   End If
   loRs!mwEventLogKey = NewEventKey
   loRs.Update
   CloseRecordset loRs
   SetNewEventKey = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwFactSetWork.SetNewEventKey. ", err.Number, err.Description
   CloseRecordset loRs
   SetNewEventKey = False
End Function

Public Function DeleteFactSet(FactLogKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   sSQL = "delete from mwFactLogColumn where mwFactLogKey=" & FactLogKey
   goCon.Execute sSQL
   sSQL = "delete from mwFactLog where ID=" & FactLogKey
   goCon.Execute sSQL
   DeleteFactSet = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwFactSetWork.DeleteFactSet. ", err.Number, err.Description
End Function
