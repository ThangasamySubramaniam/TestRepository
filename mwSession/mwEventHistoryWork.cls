VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventHistoryWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MW Utility Class for Event History
'
' 2 April 2003
'


Option Explicit
Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventHistoryTypeKey = 3
Const RS_HistoryTitle = 4
Const RS_DateTimeCreated = 5
Const RS_BriefDescription = 6
Const RS_IsApproved = 7
Const RS_Remarks = 8
Const RS_ReferenceFullFilename = 9
Const RS_mwcUsersKey = 10
Const RS_mwcRoleTypeKey = 11
Const RS_scPersonnelKey = 12
Const RS_HistoryStatus = 16   'MT-48 Occurrence History needs change protection and different save option

Const MWRT_mwEventHistoryLog = 959

'MT-48 Occurrence History needs change protection and different save option
Const SW_HISTORYSTATUS_PENDING = 1
Const SW_HISTORYSTATUS_COMPLETED = 2
Const SW_HISTORYSTATUS_REPORTED = 3

Const EVENT_HISTORY_TITLE = "CREATED"
Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1
Private mSiteKey As Long

Private mFollowupCreateDate As Date
Public Property Let FollowupCreateDate(dCreatedDate As Date)
   mFollowupCreateDate = dCreatedDate
End Property


Public Function FetchHistoryRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "SELECT mwEventHistoryLog.*, mwEventHistoryType.* " & _
     " FROM mwEventHistoryType INNER JOIN mwEventHistoryLog ON mwEventHistoryType.ID = mwEventHistoryLog.mwEventHistoryTypeKey" & _
     " WHERE mwEventHistoryLog.mwEventTypeKey=" & EventType & " AND mwEventHistoryLog.mwEventDetailKey=" & EventDetailKey & _
     " order by DateTimeCreated"
   If goSession.IsOracle Then
      sSQL = "SELECT mwEventHistoryLog.*, mwEventHistoryType.* " & _
         " FROM mwEventHistoryType , mwEventHistoryLog " & _
         " WHERE mwEventHistoryType.ID(+) = mwEventHistoryLog.mwEventHistoryTypeKey AND " & _
         " mwEventHistoryLog.mwEventTypeKey = " & EventType & " AND " & _
         " mwEventHistoryLog.mwEventDetailKey = " & EventDetailKey & _
         " ORDER BY DateTimeCreated DESC"
   Else
      sSQL = "SELECT mwEventHistoryLog.*, mwEventHistoryType.* " & _
         " FROM mwEventHistoryType RIGHT JOIN mwEventHistoryLog ON mwEventHistoryType.ID = mwEventHistoryLog.mwEventHistoryTypeKey" & _
         " WHERE mwEventHistoryLog.mwEventTypeKey=" & EventType & " AND mwEventHistoryLog.mwEventDetailKey=" & EventDetailKey & _
         " ORDER BY DateTimeCreated DESC"
   End If
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventHistoryLog"
   Set FetchHistoryRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventHistoryWork.FetchHistoryRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function AddRecord(EventType As Long, EventDetailKey As Long, _
 ByRef loRs As Recordset, Optional SiteKey As Long) As Long
   Dim loform As frmEventChooser
   Dim sSQL As String
   Dim iPK As Long
   On Error GoTo FunctionError
   '
   ' Display list of possible activities in form...
   '
   Set loform = New frmEventChooser
   If Not loform.EventHistoryInitChooser(EventType) Then
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
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   Else
      mSiteKey = 0
   End If
   '
   ' Note: mwEventHistory OCX handles WillChangeRecord Event...
   '
   iPK = goSession.MakePK("mwEventHistoryLog")
   loRs.AddNew
   loRs!mwcSitesKey = mSiteKey
   loRs.Fields(RS_ID).value = iPK
   loRs.Fields(RS_mwEventTypeKey).value = EventType
   loRs.Fields(RS_mwEventDetailKey).value = EventDetailKey
   loRs.Fields(RS_mwEventHistoryTypeKey).value = loform.FetchMwEventFactTypeKey
   loRs.Fields(RS_HistoryTitle).value = loform.FetchHistoryTitle
   loRs.Fields(RS_DateTimeCreated).value = Now()
   loRs.Fields(RS_mwcUsersKey).value = goSession.User.UserKey
   If goSession.User.RoleTypeKey > 0 Then
      loRs.Fields(RS_mwcRoleTypeKey).value = goSession.User.RoleTypeKey
   End If
   'If goSession.User.EmployeeKey > 0 Then
   '   loRS.Fields(RS_scPersonnelKey).value = goSession.User.EmployeeKey
   'End If
'MT-48 Occurrence History needs change protection and different save option
   If IsShowEventHistoryStatus(EventType) = True Then
      loRs.Fields(RS_HistoryStatus).value = SW_HISTORYSTATUS_PENDING
   End If
   loRs.Update
   AddRecord = iPK
   Unload loform
   Set loform = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventHistoryWork.AddRecord. ", Err.Number, Err.Description
   AddRecord = -1
End Function


Public Function AutoCreate(EventType As Long, EventDetailKey As Long, _
 Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loRsType As Recordset
   On Error GoTo FunctionError
   Set loRsType = New Recordset
   loRsType.CursorLocation = adUseClient
   'sSQL = "select * from mwEventHistoryType where mwEventTypeKey=" & EventType _
   '  & "and IsSuggested=1 order by DisplaySequence"
   'Dim loCmd As Command
   'Dim loPar As Parameter
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   Else
      mSiteKey = 0
   End If
   sSQL = "SELECT * from mwEventHistoryType " & _
     " Where mwEventTypeKey=" & EventType & " And IsSuggested <> 0 " & _
     " ORDER BY DisplaySequence"
   'sSQL = "SELECT * from mwEventHistoryType Where mwEventTypeKey= ? And IsSuggested= ? ORDER BY DisplaySequence;"
   'Set loCmd = New Command
   loRsType.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsType.RecordCount < 1 Then
      CloseRecordset loRsType
      AutoCreate = True
      Exit Function
   End If
   sSQL = "select * from mwEventHistoryLog where ID=-1"
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   '
   Do While Not loRsType.EOF
      '
      '
      moRS.AddNew
      moRS.Fields(RS_ID).value = goSession.MakePK("mwEventHistoryLog")
      moRS.Fields(RS_mwEventTypeKey).value = EventType
      moRS.Fields(RS_mwEventDetailKey).value = EventDetailKey
      moRS.Fields(RS_mwEventHistoryTypeKey).value = loRsType!ID
      moRS.Fields(RS_HistoryTitle).value = loRsType!HistoryTitle
      moRS.Fields(RS_DateTimeCreated).value = Now()
      moRS.Fields(RS_mwcUsersKey).value = goSession.User.UserKey
      If goSession.User.RoleTypeKey > 0 Then
         moRS.Fields(RS_mwcRoleTypeKey).value = goSession.User.RoleTypeKey
      End If
      If goSession.User.EmployeeKey > 0 Then
         moRS.Fields(RS_scPersonnelKey).value = goSession.User.EmployeeKey
      End If
      If mSiteKey > 0 Then
         moRS!mwcSitesKey = mSiteKey
      End If
'MT-48 Occurrence History needs change protection and different save option
      If IsShowEventHistoryStatus(EventType) = True Then
         moRS.Fields(RS_HistoryStatus).value = SW_HISTORYSTATUS_COMPLETED
      End If
      moRS.Update
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   CloseRecordset moRS
   AutoCreate = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventHistoryWork.AutoCreate. ", Err.Number, Err.Description
   AutoCreate = False
End Function



Public Function IsEventHistoryTypesExist(EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select count(*) from mwEventHistoryType where mwEventTypeKey=" & EventType
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.Fields(0).value > 0 Then
      IsEventHistoryTypesExist = True
   Else
      IsEventHistoryTypesExist = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventHistoryWork.IsEventHistoryTypesExist. ", Err.Number, Err.Description
End Function

Public Function AddSystemRecord(EventType As Long, EventDetailKey As Long, _
 HistoryTitle As String, BriefDescription As String, _
 Optional Remarks As String, Optional FromWfStage As String, _
 Optional ToWfStage As String, Optional SiteKey As Long) As Long
   Dim sSQL As String
   Dim iPK As Long
   On Error GoTo FunctionError
   '
   ' Display list of possible activities in form...
   '
   ' ok - Use Insert Statement - Cleaner on add's
'   If SiteKey > 0 Then
'      mSiteKey = SiteKey
'   Else
'      mSiteKey = goSession.Site.SiteKey
'   End If
   sSQL = "select * from mwEventHistoryLog where ID=-1"
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   iPK = goSession.MakePK("mwEventHistoryLog")
   moRS.AddNew
   moRS.Fields(RS_ID).value = iPK
   moRS.Fields(RS_mwEventTypeKey).value = EventType
   moRS.Fields(RS_mwEventDetailKey).value = EventDetailKey
   moRS.Fields(RS_HistoryTitle).value = Left(HistoryTitle, 50)
   If Remarks <> "" Then
      moRS.Fields(RS_Remarks).value = Remarks
   End If
   If BriefDescription <> "" Then
      moRS.Fields(RS_BriefDescription).value = Left(BriefDescription, 99)
   End If
   
   'moRS.Fields(RS_DateTimeCreated).value = Now()  'DEV-1661
   If IsDate(mFollowupCreateDate) Then
      If mFollowupCreateDate <> "12:00:00 AM" Then
         moRS.Fields(RS_DateTimeCreated).value = mFollowupCreateDate
      Else
         moRS.Fields(RS_DateTimeCreated).value = Now()
      End If
   Else
      moRS.Fields(RS_DateTimeCreated).value = Now()
   End If
   
   moRS.Fields(RS_mwcUsersKey).value = goSession.User.UserKey
   If goSession.User.RoleTypeKey > 0 Then
      moRS.Fields(RS_mwcRoleTypeKey).value = goSession.User.RoleTypeKey
   End If
   If goSession.User.EmployeeKey > 0 Then
      moRS.Fields(RS_scPersonnelKey).value = goSession.User.EmployeeKey
   End If
   If SiteKey > 0 Then
      moRS!mwcSitesKey = SiteKey
   Else
      moRS!mwcSitesKey = goSession.Site.SiteKey
   End If
'MT-48 Occurrence History needs change protection and different save option
   If IsShowEventHistoryStatus(EventType) = True Then
      moRS.Fields(RS_HistoryStatus).value = SW_HISTORYSTATUS_COMPLETED
   End If
   
   moRS.Update
   AddSystemRecord = iPK
   CloseRecordset moRS
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwEventHistoryWork.AddSystemRecord. ", Err.Number, Err.Description
   AddSystemRecord = -1
End Function


Public Function SendSystemDatagram(TargetSiteID As String, EventType As Long, EventDetailKey As Long, HistoryTitle As String, _
  BriefDescription As String, Optional Remarks As String, _
  Optional IsLocalUpdate As Boolean) As Boolean
   Dim loform As mwForm
   Dim fso As FileSystemObject
   Dim sSQL As String
   Dim loRec As ConfigKeys
   Dim iPK As Long
   On Error GoTo FunctionError
   Set loform = goSession.OpenNewForm
   If Not loform.MakeNewForm("MWS_EH", , , , TargetSiteID) Then
      SendSystemDatagram = False
      goSession.TearDownForm loform
      Exit Function
   End If
   Set loRec = loform.DataForm.AddRecord
   With loRec
      .Add "Table", "mwEventHistoryLog"
      .Add "Operation", "A"
      ' ID column is MakePK On Process...
      .Add "mwEventTypeKey", goSession.Site.SiteKey
      .Add "mwEventDetailKey", goSession.User.UserKey
      .Add "DateTimeCreated", goSession.User.EmployeeID
      .Add "BriefDescription", goSession.User.RoleType
   End With
   loform.DataForm.WriteDataForm
   If loform.Send(, mw_SendToShip, TargetSiteID) Then
      SendSystemDatagram = True
   Else
      SendSystemDatagram = False
   End If
   Set loRec = Nothing
   goSession.TearDownForm loform
   '
   ' Local Update Too ?
   '
   If IsLocalUpdate Then
      'Loop the history record back for processing...
      loform.FormProperties.TargetSiteID = ""
      loform.DataForm.WriteDataForm
      loform.Send , mw_SendToShore
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventHistoryWork.AddRecords. ", Err.Number, Err.Description
   SendSystemDatagram = False
   goSession.TearDownForm loform
End Function



Public Function DeleteAllHistory(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   On Error GoTo FunctionError
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventHistoryLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteAllHistory = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   ' delete facts not used...
   Do While Not moRS.EOF
      moRS.Delete
      moRS.MoveNext
   Loop
   goSession.CloseRecordset moRS
   DeleteAllHistory = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteAllHistory. ", Err.Number, Err.Description
   DeleteAllHistory = False
End Function


Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwEventHistoryLog") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmSmOccurrence.moRS_WillChangeRecord. ", Err.Number, Err.Description
End Sub

'MT-48 Occurrence History needs change protection and different save option
Public Function IsShowEventHistoryStatus(EventTypeKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   Set loRs = New Recordset
   sSQL = "select IsShowEventHistoryStatus from mwEventType where ID=" & EventTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount = 0 Then
      IsShowEventHistoryStatus = False
   ElseIf IsNull(loRs!IsShowEventHistoryStatus) Then
      IsShowEventHistoryStatus = False
   ElseIf loRs!IsShowEventHistoryStatus = False Then
      IsShowEventHistoryStatus = False
   Else
      IsShowEventHistoryStatus = True
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventHistoryWork.IsShowEventHistoryStatus. ", Err.Number, Err.Description
End Function


Public Function IsEventCreatedBySessionUser(UserKey As Long, EventTypeKey As Long, EventDetailKey As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
     
   Set loRs = New Recordset
   sSQL = "SELECT ID FROM mwEventHistoryLog WHERE  mwEventTypeKey =  " & EventTypeKey & "  AND mwEventDetailKey= " & EventDetailKey & _
   " AND mwcUsersKey = " & UserKey & "  AND  HistoryTitle = '" & EVENT_HISTORY_TITLE & "'"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      IsEventCreatedBySessionUser = True
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventHistoryWork.IsEventCreatedBySessionUser", Err.Number, Err.Description
   CloseRecordset loRs
   IsEventCreatedBySessionUser = False
End Function


