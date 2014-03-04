VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventFactsWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwEventFactsWork - Support functions for managing Voyage Activities
' 10/28/02 ms
'

Option Explicit
Private mVoyNo As String

Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventFactTypeKey = 3
Const RS_ActivityTitle = 4
Const RS_DateOrder = 10
Const RS_DatagramCode = 13
Const RS_IsMandatory = 22
Const RS_IsStartDateTimeAllowed = 24
Const RS_StartDateTime = 5
Const RS_IsEndDateTimeRequired = 26
Const RS_EndDateTime = 6
Const RS_IsRemarkRequired = 28
Const RS_Remarks = 7

Public Property Let VoyNo(ByVal Data As String)
    mVoyNo = Data
End Property

Public Property Get VoyNo() As String
    VoyNo = mVoyNo
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

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


Public Function AutoCreate(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loRsType As Recordset
   Dim loRsFact As Recordset
   On Error GoTo FunctionError
   Set loRsType = New Recordset
   loRsType.CursorLocation = adUseClient
   'sSQL = "select * from mwEventFactType where mwEventTypeKey=" & EventType _
   '  & "and IsSuggested=1 order by DisplaySequence"
   'Dim loCmd As Command
   'Dim loPar As Parameter
   sSQL = "SELECT * from mwEventFactType " & _
     " Where mwEventTypeKey=" & EventType & " And IsSuggested <> 0 and IsActive <> 0" & _
     " ORDER BY DisplaySequence"
   'sSQL = "SELECT * from mwEventFactType Where mwEventTypeKey= ? And IsSuggested= ? ORDER BY DisplaySequence;"
   'Set loCmd = New Command
   
   
   
   loRsType.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsType.RecordCount < 1 Then
      CloseRecordset loRsType
      AutoCreate = True
      Exit Function
   End If
   Set loRsFact = New Recordset
   loRsFact.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLog where ID=-1"
   loRsFact.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   '
   Do While Not loRsType.EOF
      '
      '
      loRsFact.AddNew
      loRsFact!ID = goSession.MakePK("mwEventFactLog")
      loRsFact!mwEventTypeKey = EventType
      loRsFact!mwEventDetailKey = EventDetailKey
      loRsFact!mwEventFactTypeKey = loRsType!ID
      loRsFact!ActivityTitle = loRsType!ActivityTitle
      '
      ' Set dummay date for Remarks...
      '
      If loRsType!IsStartDateTimeAllowed = False And loRsType!IsEndDateTimeAllowed = False Then
         loRsFact!DateSort = "01/01/2099"
      End If
      loRsFact.Update
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   CloseRecordset loRsFact
   AutoCreate = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.AutoCreate. ", err.Number, err.Description
   AutoCreate = False
End Function

Public Function AddRecord(EventType As Long, EventDetailKey As Long, ByRef loRs As Recordset) As Long
   Dim loform As frmEventChooser
   Dim strFactTitle As String
   Dim sSQL As String
   Dim iPK As Long
   Dim iFactType As Long
   On Error GoTo FunctionError
   '
   ' Display list of possible activities in form...
   '
   Set loform = New frmEventChooser
   If Not loform.EventFactInitChooser(EventType, EventDetailKey) Then
      AddRecord = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   loform.Show vbModal
   If loform.IsCancelled Then
      AddRecord = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   '
   ' ok - Use Insert Statement - Cleaner on add's
   '
   ' May 2003 - Allow Multi Row select
   '
   iPK = goSession.MakePK("mwEventFactLog")
   Do While loform.IsSelectedRows
      iFactType = loform.FetchNextSelected(strFactTitle)
      If iFactType < 1 Then
         Exit Do
      End If
      loRs.AddNew
      loRs.Fields(RS_ID).value = iPK
      loRs.Fields(RS_mwEventTypeKey).value = EventType
      loRs.Fields(RS_mwEventDetailKey).value = EventDetailKey
      loRs.Fields(RS_mwEventFactTypeKey).value = iFactType
      loRs.Fields(RS_ActivityTitle).value = strFactTitle
      If IsRemarkOnly(iFactType) Then
         loRs.Fields(RS_DateOrder).value = "01/01/2099"
      End If
      loRs.Update
      iPK = iPK + 1
   Loop
   If iPK > 0 Then
      goSession.UpdatePrimaryKeySequence "mwEventFactLog", iPK
   End If
   AddRecord = iPK
   Unload loform
   Set loform = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.AddRecords. ", err.Number, err.Description
   AddRecord = -1
End Function

Public Function FetchFactsRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " Where mwEventFactLog.mwEventTypeKey=" & EventType & " And mwEventFactLog.mwEventDetailKey=" & EventDetailKey & _
'     " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID;"

   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
     " FROM mwEventFactType , mwEventFactLog " & _
     " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
     " mwEventFactLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLog.mwEventDetailKey = " & EventDetailKey & _
     " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID"

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Set FetchFactsRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchFactsRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function

Public Function FetchDelayFactsRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " Where mwEventFactLog.mwEventTypeKey=" & EventType & " And mwEventFactLog.mwEventDetailKey=" & EventDetailKey & _
'     " and mwEventFactType.IsDelayActivity <> 0 " & _
'     " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID;"
     
   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
      " FROM mwEventFactType , mwEventFactLog " & _
      " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
      " mwEventFactLog.mwEventTypeKey=" & EventType & " AND " & _
      " mwEventFactLog.mwEventDetailKey = " & EventDetailKey & _
      " AND mwEventFactType.IsDelayActivity <> 0 " & _
      " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID"

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Set FetchDelayFactsRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchDelayFactsRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function

Public Function FetchFieldTagRS(EventType As Long, EventDetailKey As Long, FieldTag As String) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " WHERE (mwEventFactLog.mwEventTypeKey=" & EventType & ") AND mwEventFactLog.mwEventDetailKey=" & EventDetailKey & _
'     " AND (mwEventFactType.StartFieldTag='" & FieldTag & "'  OR mwEventFactType.EndFieldTag='" & FieldTag & _
'     "' OR mwEventFactType.RemarksFieldTag='" & FieldTag & "')"

   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
     " FROM mwEventFactType , mwEventFactLog " & _
     " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
     " mwEventFactLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLog.mwEventDetailKey = " & EventDetailKey & " AND " & _
     " (mwEventFactType.StartFieldTag = '" & FieldTag & "'  OR " & _
     "  mwEventFactType.EndFieldTag = '" & FieldTag & _
     "' OR mwEventFactType.RemarksFieldTag='" & FieldTag & "')"

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Set FetchFieldTagRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchFieldTagRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function

Public Function FetchExpandedFieldList(EventType As Long, EventDetailKey As Long) As Collection
   Dim sSQL As String
   Dim loRs As Recordset
   Dim s As String
   Dim loCol As Collection
   On Error GoTo FunctionError
   Set loCol = New Collection
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT StartFieldTag, EndFieldTag, RemarksFieldTag, IsDelayActivity " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " WHERE mwEventFactLog.mwEventTypeKey=" & EventType & " AND mwEventFactLog.mwEventDetailKey=" & EventDetailKey
   
   sSQL = "SELECT StartFieldTag, EndFieldTag, RemarksFieldTag, IsDelayActivity " & _
     " FROM mwEventFactType , mwEventFactLog " & _
     " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
     " mwEventFactLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLog.mwEventDetailKey = " & EventDetailKey
   
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      On Error Resume Next
      If Not loRs!IsDelayActivity Then
         If Not IsNull(loRs!StartFieldTag) Then
            If Not IsInCollection(loCol, loRs!StartFieldTag) Then
               s = loRs!StartFieldTag
               loCol.Add Item:=s, Key:=s
            End If
         End If
         If Not IsNull(loRs!EndFieldTag) Then
            If Not IsInCollection(loCol, loRs!EndFieldTag) Then
               s = loRs!EndFieldTag
               loCol.Add Item:=s, Key:=s
            End If
         End If
         If Not IsNull(loRs!RemarksFieldTag) Then
            If Not IsInCollection(loCol, loRs!RemarksFieldTag) Then
               s = loRs!RemarksFieldTag
               loCol.Add Item:=s, Key:=s
            End If
         End If
      End If
      On Error GoTo FunctionError
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   Set FetchExpandedFieldList = loCol
   Set loCol = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchExpandedFieldList. ", err.Number, err.Description
   CloseRecordset loRs
End Function




Public Function Submit(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loform As mwForm
   Dim strFile As String
   On Error GoTo FunctionError
   MsgBox "stump code"
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.Submit. ", err.Number, err.Description
   CloseRecordset loRs
End Function



Public Function DeleteUnusedFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If loRs.RecordCount < 1 Then
      DeleteUnusedFacts = True
      goSession.CloseRecordset loRs
      Exit Function
   End If
   ' delete facts not used...
   Do While Not loRs.EOF
      If IsNull(loRs.Fields(RS_StartDateTime).value) And IsNull(loRs.Fields(RS_EndDateTime).value) _
        And IsNull(loRs.Fields(RS_Remarks).value) Then
         ' Nothing filled in, delete it...
         loRs.Delete
      End If
      loRs.MoveNext
   Loop
   goSession.CloseRecordset loRs
   DeleteUnusedFacts = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteUnusedFacts. ", err.Number, err.Description
   DeleteUnusedFacts = False
End Function

Public Function DeleteFact(FactKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLog where ID=" & FactKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If loRs.RecordCount < 1 Then
      DeleteFact = True
      goSession.CloseRecordset loRs
      Exit Function
   End If
   If Not loRs!IsSent Then
      loRs.Delete
      DeleteFact = True
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventFactsWork, skip delete submitted form: " & loRs!FullFilename
      DeleteFact = False
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteFact. ", err.Number, err.Description
   DeleteFact = False
End Function



Public Function IsSubmittedFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "select count(*) from mwEventFactLog where mwEventTypeKey=" & EventType & _
   '  " and mwEventDetailKey=" & EventKey & " and IsSubmitted=True"
   sSQL = "select count(*) from mwEventFactLog WHERE mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey & " AND (IsSubmitted is not null AND IsSubmitted <> 0)"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   ' Got Forms ?
   If loRs.Fields(0).value > 0 Then
      IsSubmittedFacts = True
   Else
      IsSubmittedFacts = False
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.IsSubmittedFacts. ", err.Number, err.Description
   IsSubmittedFacts = False
End Function



'
' Expand a recordset with all EventFacts fields for a particular event detail record
'
Public Function ExpandRecordset(ByRef oRsEventDetail As Recordset, EventTypeKey As Long, _
  EventDetailKey As Long) As Recordset
   Dim loRsFact As Recordset
   Dim loRs As Recordset
   Dim loSkipList As Collection
   Dim loField As field
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Staging/Validation
   If oRsEventDetail.State = adStateClosed Then
      goSession.RaiseError "Error in mwEventFactsWork.ExpandRecordset, Recordset is closed."
      Exit Function
   End If
   ' Any records in there ?
   If oRsEventDetail.RecordCount < 1 Then
      goSession.RaiseError "Error in mwEventFactsWork.ExpandRecordset, no records in Recordset."
      Exit Function
   End If
   '
   ' Create new recordset...
   '
   Set loSkipList = New Collection
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ' Add columns to fabricated recordset
   For Each loField In oRsEventDetail.Fields
      ' Skip shape RS last column...
      If loField.Type <> adChapter Then
         If goSession.IsOracle And loField.Type = adNumeric Then
            loRs.Fields.Append loField.Name, adInteger, 4
         Else
            loRs.Fields.Append loField.Name, loField.Type, loField.DefinedSize
         End If
      End If
   Next loField
   ' Get the activities
   Set loRsFact = FetchFactsRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   
   Do While Not loRsFact.EOF
      On Error Resume Next
      If Not loRsFact!IsDelayActivity Then
         ' Start Date/Time
         If Not IsNull(loRsFact!StartFieldTag) Then
            loRs.Fields.Append loRsFact.Fields("StartFieldTag").value, loRsFact.Fields("StartDateTime").Type, loRsFact.Fields("StartDateTime").DefinedSize
            If err Then
               If err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("StartFieldTag").value, Key:=loRsFact.Fields("StartFieldTag").value
               Else
                  goSession.RaiseError "General Error in mwEventFactsWork.ExpandRecordset. ", err.Number, err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
         'End Date/Time
         If Not IsNull(loRsFact!EndFieldTag) Then
            loRs.Fields.Append loRsFact.Fields("EndFieldTag").value, loRsFact.Fields("EndDateTime").Type, loRsFact.Fields("EndDateTime").DefinedSize
            If err Then
               If err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("EndFieldTag").value, Key:=loRsFact.Fields("EndFieldTag").value
               Else
                  goSession.RaiseError "General Error in mwEventFactsWork.ExpandRecordset. ", err.Number, err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
         ' Remarks/Other Information
         If Not IsNull(loRsFact!RemarksFieldTag) Then
            loRs.Fields.Append loRsFact.Fields("RemarksFieldTag").value, loRsFact.Fields("Remarks").Type, loRsFact.Fields("Remarks").DefinedSize
            If err Then
               If err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("RemarksFieldTag").value, Key:=loRsFact.Fields("RemarksFieldTag").value
               Else
                  goSession.RaiseError "General Error in mwEventFactsWork.ExpandRecordset. ", err.Number, err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
      End If
      On Error GoTo FunctionError
      loRsFact.MoveNext
   Loop
   On Error GoTo FunctionError
   '
   ' Append data to fabricated recordset...
   '
   loRs.Open
   loRs.AddNew
   ' Event Detail...
   For Each loField In oRsEventDetail.Fields
      With loField
         ' Skip shape RS last column...
         If loField.Type <> adChapter Then
            If Not IsNull(.value) And Not IsEmpty(.value) Then
               loRs(.Name).value = .value
            End If
         End If
      End With
   Next loField
   '
   ' Add Activity record fields to flat fabricated recordset
   '
   If loRsFact.RecordCount > 0 Then
      loRsFact.MoveFirst
      Do While Not loRsFact.EOF
         If Not loRsFact!IsDelayActivity Then
            ' Start Date/Time
            If Not IsNull(loRsFact!StartFieldTag) And Not IsNull(loRsFact!StartDateTime) _
              And Not IsInCollection(loSkipList, loRsFact!StartFieldTag) Then
               loRs.Fields(loRsFact.Fields("StartFieldTag").value).value = loRsFact!StartDateTime
            End If
            ' End Date/Time
            If Not IsNull(loRsFact!EndFieldTag) And Not IsNull(loRsFact!EndDateTime) _
              And Not IsInCollection(loSkipList, loRsFact!EndFieldTag) Then
               loRs.Fields(loRsFact.Fields("EndFieldTag").value).value = loRsFact!EndDateTime
            End If
            ' Remarks/Other Information
            If Not IsNull(loRsFact!RemarksFieldTag) And Not IsNull(loRsFact!Remarks) _
              And Not IsInCollection(loSkipList, loRsFact!RemarksFieldTag) Then
               loRs.Fields(loRsFact.Fields("RemarksFieldTag").value).value = loRsFact!Remarks
            End If
         End If
         loRsFact.MoveNext
      Loop
   End If
   CloseRecordset loRsFact
   ' That's all folks...
   Set ExpandRecordset = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactsWork.ExpandRecordset. ", err.Number, err.Description
   CloseRecordset loRsFact
   CloseRecordset loRs
   ' revert to original recordset
   Set ExpandRecordset = oRsEventDetail
End Function



Public Function LockEventFacts(EventTypeKey As Long, _
  EventDetailKey As Long, Optional DatagramLockCode As String) As Boolean
   Dim loRsFact As Recordset
   Dim sSQL As String
   Dim strLockCode
   On Error GoTo FunctionError
   ' Get the activities
   Set loRsFact = FetchFactsRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   Do While Not loRsFact.EOF
      If DatagramLockCode = "" Or DatagramLockCode = loRsFact!DatagramLockCode Then
         loRsFact!sent = True
         loRsFact.Update
      End If
      loRsFact.MoveNext
   Loop
   CloseRecordset loRsFact
   LockEventFacts = True
   ' That's all folks...
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactsWork.LockEventFacts. ", err.Number, err.Description
   LockEventFacts = False
End Function

Public Function UnLockEventFacts(EventTypeKey As Long, _
  EventDetailKey As Long) As Boolean
   Dim loRsFact As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Get the activities
   Set loRsFact = FetchFactsRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   Do While Not loRsFact.EOF
      loRsFact!sent = False
      loRsFact.Update
      loRsFact.MoveNext
   Loop
   CloseRecordset loRsFact
   UnLockEventFacts = True
   ' That's all folks...
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactsWork.UnLockEventFacts. ", err.Number, err.Description
   UnLockEventFacts = False
End Function


Public Function IsEventFactTypesExist(EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select count(*) from mwEventFactType where mwEventTypeKey=" & EventType & " and IsActive<>0"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.Fields(0).value > 0 Then
      IsEventFactTypesExist = True
   Else
      IsEventFactTypesExist = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactsWork.IsEventFactTypesExist. ", err.Number, err.Description
End Function



Public Function FetchShipnetActivityRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " Where mwEventFactLog.mwEventTypeKey=" & EventType & " And mwEventFactLog.mwEventDetailKey=" & EventDetailKey & _
'     " and mwEventFactType.SN_ACTIVITY_NAME <> Null " & _
'     " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID;"

   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
     " FROM mwEventFactType , mwEventFactLog " & _
     " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
     " mwEventFactLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLog.mwEventDetailKey = " & EventDetailKey & " AND " & _
     " mwEventFactType.SN_ACTIVITY_NAME <> Null " & _
     " ORDER BY mwEventFactLog.StartDateTime, mwEventFactLog.ID"

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Set FetchShipnetActivityRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchShipnetActivityRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function


Private Function IsRemarkOnly(iFactType As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * from mwEventFactType where ID=" & str(iFactType)
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs!IsStartDateTimeAllowed = False And loRs!IsEndDateTimeAllowed = False Then
      IsRemarkOnly = True
   Else
      IsRemarkOnly = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.IsRemarkOnly. ", err.Number, err.Description
   CloseRecordset loRs
   IsRemarkOnly = False
End Function

Public Function IsValidFacts(EventType As Long, EventDetailKey As Long, _
  ByRef ErrorString As String, Optional DatagramCode As String) As Boolean
   Dim lIsError As Boolean
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ' Loop through all records, performing validation...
   lIsError = False
   Set loRs = FetchFactsRS(EventType, EventDetailKey)
   If loRs.RecordCount < 1 Then
      IsValidFacts = True
      CloseRecordset loRs
      Exit Function
   End If
   Do While Not loRs.EOF
      
      ' Arrival Messages require DatagramCode to decide which records to check...
      If DatagramCode = "" Or DatagramCode = loRs.Fields(RS_DatagramCode).value Then
         If loRs.Fields(RS_IsMandatory).value Then
            ' Start Date...
            If loRs.Fields(RS_IsStartDateTimeAllowed).value And _
              IsNull(loRs.Fields(RS_StartDateTime).value) Then
                ErrorString = ErrorString & _
                  "Mandatory Start Date/Time Missing: " & loRs.Fields(RS_ActivityTitle).value & vbCrLf
                lIsError = True
            End If
            '
         
         End If
         ' Is end date/time required ?
         If loRs.Fields(RS_IsEndDateTimeRequired).value Then
            '
            If IsNull(loRs.Fields(RS_EndDateTime).value) Then
               ErrorString = ErrorString & _
                 "Required End Date/Time Missing: " & loRs.Fields(RS_ActivityTitle).value & vbCrLf
               lIsError = True
            End If
         End If
         
         
         ' Is remark required ?
         If loRs.Fields(RS_IsRemarkRequired).value = True Then
            '
            If IsNull(loRs.Fields(RS_Remarks).value) Then
               ErrorString = ErrorString & _
                 "Required Entry Missing: " & loRs.Fields(RS_ActivityTitle).value & vbCrLf
               lIsError = True
            End If
         End If
      End If
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   IsValidFacts = Not lIsError
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFacts.IsValidFacts. ", err.Number, err.Description
   IsValidFacts = False
End Function


Public Function FetchFactLogRecordRS(EventTypeKey As Long, EventDetailKey As Long, FactTypeKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
'     " FROM mwEventFactType INNER JOIN mwEventFactLog ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " Where mwEventFactLog.mwEventTypeKey=" & EventTypeKey & " And mwEventFactLog.mwEventDetailKey=" & EventDetailKey & _
'     " and  mwEventFactLog.mwEventFacttypeKey=" & FactTypeKey

   sSQL = "SELECT mwEventFactLog.*, mwEventFactType.* " & _
     " FROM mwEventFactType , mwEventFactLog " & _
     " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
     " mwEventFactLog.mwEventTypeKey = " & EventTypeKey & " AND " & _
     " mwEventFactLog.mwEventDetailKey = " & EventDetailKey & " AND " & _
     " mwEventFactLog.mwEventFacttypeKey = " & FactTypeKey

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Set FetchFactLogRecordRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchFactLogRecordRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function

Public Function DeleteByFilter(EventType As Long, EventDetail As Long, FilterNo As Integer, FilterCode As String) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
'   sSQL = "SELECT mwEventFactLog.* FROM mwEventFactType INNER JOIN mwEventFactLog " & _
'     " ON mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey " & _
'     " WHERE mwEventFactLog.mwEventTypeKey=" & EventType & " AND mwEventFactLog.mwEventDetailKey=" & EventDetail
   
   sSQL = "SELECT mwEventFactLog.* " & _
      " FROM mwEventFactType , mwEventFactLog " & _
      " WHERE mwEventFactType.ID = mwEventFactLog.mwEventFactTypeKey AND " & _
      " mwEventFactLog.mwEventTypeKey = " & EventType & " AND " & _
      " mwEventFactLog.mwEventDetailKey = " & EventDetail
   
   Select Case FilterNo
      Case 1
         sSQL = sSQL & " AND Filter1='" & FilterCode & "'"
      Case 2
         sSQL = sSQL & " AND Filter2='" & FilterCode & "'"
      Case 3
         sSQL = sSQL & " AND Filter3='" & FilterCode & "'"
      Case Else
         goSession.RaiseError "General Error in mwEventFactsWork.ReduceFactLogList, invalid filter number. "
         DeleteByFilter = False
         Exit Function
   End Select
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLog"
   Do While Not loRs.EOF
      loRs.Delete
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   DeleteByFilter = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteByFilter. ", err.Number, err.Description
End Function

Public Function FetchFactLogRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * from mwEventFactLog Where mwEventTypeKey=" & EventType & _
    " And mwEventDetailKey=" & EventDetailKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Set FetchFactLogRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.FetchFactLogRS. ", err.Number, err.Description
   CloseRecordset loRs
End Function



Public Function DeleteAllFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If loRs.RecordCount < 1 Then
      DeleteAllFacts = True
      goSession.CloseRecordset loRs
      Exit Function
   End If
   ' delete facts not used...
   Do While Not loRs.EOF
      loRs.Delete
      loRs.MoveNext
   Loop
   goSession.CloseRecordset loRs
   DeleteAllFacts = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteAllFacts. ", err.Number, err.Description
   DeleteAllFacts = False
End Function


