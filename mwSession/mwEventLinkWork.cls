VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventLinkWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MW Utility Class for Event Links
'
' 4 April 2003
'


Option Explicit
Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventLinkTypeKey = 3
Const RS_IsCreated = 4
Const RS_DateTimeCreated = 5
Const RS_BriefDescription = 6
Const RS_mwcSitesKey = 7
Const RS_IsFastTrack = 8
Const RS_IsComplete = 9
Const RS_IsNotComplete = 10
Const RS_smFileCabinetKey = 11
Const RS_mwcRoleTypeKey = 12
Const RS_mwcUsersKey = 13
Const RS_Remarks = 14


Const RST_ID = 15
Const RST_mwEventTypeKey = 16
Const RST_LinkTitle = 17
Const RST_DisplaySequence = 18
Const RST_DefaultDescription = 19
Const RST_DisplayIcon = 20
Const RST_IsMandatory = 21
Const RST_IsSuggested = 22
Const RST_IsReadOnlyLink = 23
Const RST_smFileCabFleetIndexKey = 24
Const RST_smStdTaskKey = 25
Const RST_IsReplicate = 26
Const RST_IsCompleteCheckboxAllowed = 27
Const RST_IsNotCompleteCheckboxAllowed = 28
Const RST_IsFastTrackAllowed = 29
Const RST_FullFileName = 30
Const RST_ContextID = 31
Const RST_MaxImportSizeBytes = 32
Const RST_IsActive = 33
Const RST_LinkType = 34




Private mIsDeleteLink As Boolean
Dim mSiteKey As Long
Public WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1

Public Property Let SiteKey(ByVal Data As Long)
    mSiteKey = Data
End Property

Public Property Get SiteKey() As Long
    SiteKey = mSiteKey
End Property

Public Property Let IsDeleteLink(ByVal Data As Boolean)
    mIsDeleteLink = Data
End Property

Public Property Get IsDeleteLink() As Boolean
    IsDeleteLink = mIsDeleteLink
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



Public Function FetchLinkRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   strSQL = "SELECT mwEventLinkLog.*, mwEventLinkType.* " & _
'     " FROM mwEventLinkType INNER JOIN mwEventLinkLog ON mwEventLinkType.ID = mwEventLinkLog.mwEventLinkTypeKey" & _
'     " WHERE mwEventLinkLog.mwEventTypeKey=" & EventType & " AND mwEventLinkLog.mwEventDetailKey=" & EventDetailKey & _
'     " order by DateTimeCreated"
   
   strSQL = "SELECT mwEventLinkLog.*, mwEventLinkType.* " & _
     " FROM mwEventLinkType , mwEventLinkLog " & _
     " WHERE mwEventLinkType.ID = mwEventLinkLog.mwEventLinkTypeKey AND " & _
     " mwEventLinkLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventLinkLog.mwEventDetailKey = " & EventDetailKey & _
     " ORDER BY DateTimeCreated"
   
   loRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventLinkLog"
   Set FetchLinkRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventLinkWork.FetchLinkRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function AddRecord(EventType As Long, EventDetailKey As Long, _
 ByRef moRS As Recordset, Optional SiteKey As Long) As Long
   Dim loform As frmEventChooser
   Dim strSQL As String
   Dim fso As FileSystemObject
   Dim loRsType As Recordset
   Dim sSQL As String
   Dim ELKey As Long
   On Error GoTo FunctionError
   '
   ' Display list of possible activities in form...
   '
   Set fso = New FileSystemObject
   Set loform = New frmEventChooser
   If Not loform.EventLinkInitChooser(EventType) Then
      AddRecord = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   ELKey = loform.FetchLinkKey
   If ELKey < 1 Then
      loform.Show vbModal
      If loform.IsCancelled Then
         AddRecord = -1
         Unload loform
         Set loform = Nothing
         Exit Function
      End If
      ELKey = loform.FetchLinkKey
   End If
      
   Set loRsType = New Recordset
   loRsType.CursorLocation = adUseClient
   sSQL = "select ID, mwEventTypeKey, LinkTitle, DefaultDescription, IsFastTrackAllowed, smFileCabFleetIndexKey, FullFileName From mwEventLinkType where ID=" & ELKey
   loRsType.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsType.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwEventFormWork.SubmitForm. Unable to store record."
      CloseRecordset loRsType
      CloseRecordset moRS
      goCon.RollbackTrans
      goSession.TearDownForm loform
      Exit Function
   End If
           
   'sSQL = "select * from mwEventLinkLog where ID=-1"
   'moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   With moRS
      .AddNew
      .Fields(RS_ID).value = goSession.MakePK("mwEventLinkLog")
      .Fields(RS_mwEventTypeKey).value = EventType
      .Fields(RS_mwEventDetailKey).value = EventDetailKey
      .Fields(RS_mwEventLinkTypeKey).value = loRsType!ID
      .Fields(RS_IsCreated).value = False
      .Fields(RS_DateTimeCreated).value = Now()
      If Not IsNull(loRsType!DefaultDescription) Then
         .Fields(RS_BriefDescription).value = loRsType!DefaultDescription
      End If
      If SiteKey > 0 Then
         .Fields(RS_mwcSitesKey).value = SiteKey
      End If
      .Fields(RS_IsFastTrack).value = False
      .Fields(RS_IsComplete).value = False
      .Fields(RS_IsNotComplete).value = False
      .Fields(RS_mwcRoleTypeKey).value = goSession.User.RoleTypeKey
      .Fields(RS_mwcUsersKey).value = goSession.User.UserKey
      .Update
   End With
   
   AddRecord = moRS.Fields(RS_ID).value
   CloseRecordset loRsType
   CloseRecordset moRS
   Unload loform
   Set fso = Nothing
   Set loform = Nothing
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwEventLinkWork.AddRecords. ", Err.Number, Err.Description
   AddRecord = -1
End Function

Public Function StoreDocument(mSiteKey As Long, FullFilename As String, FileCabinetFleetIndexKey As Long, LinkTitle As String, EventTypeKey As Long, EventDetailKey As Long) As Boolean
   'Dim nFileCabKey As Long
   'Dim loWork As Object
   'Dim fso As FileSystemObject
   'Dim sSQL As String
   'On Error GoTo FunctionError
   '
   'Set moRS = New Recordset
   'moRS.CursorLocation = adUseClient
   'sSQL = "select * From mwEventLinkLog where ID=" & EventDetailKey
   'moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   '
   'goCon.BeginTrans
   '   Set loWork = CreateObject("mwManuals.smFileCabWork")
   '   loWork.initsession goSession
   '   ' Store record in smFileCabinet
   '   nFileCabKey = loWork.StoreDocument(mSiteKey, FullFilename, FileCabinetFleetIndexKey, LinkTitle, EventTypeKey, EventDetailKey)
   '   If nFileCabKey = -1 Then
   '      goCon.RollbackTrans
   '      goSession.RaisePublicError "Error in mwEventLinkWork.StoreDocument. Unable to store record."
   '      CloseRecordset moRS
   '      KillObject loWork
   '      StoreDocument = False
   '      Exit Function
   '   End If
   '   ' store file cabinet key in Link Log
   '
   '   moRS!smFileCabinetKey = nFileCabKey
   '   moRS.Update
   'goCon.CommitTrans
   '
   'StoreDocument = True
   'KillObject loWork
   'CloseRecordset moRS
   'Exit Function
'FunctionError:
 '  goSession.RaiseError "General Error in mwEventLinkWork.StoreDocument. ", err.Number, err.Description
 '  On Error Resume Next
 '  goCon.RollbackTrans
 '  KillObject loWork
 '  'CloseRecordset moRS
End Function

Public Function AutoCreate(EventType As Long, EventDetailKey As Long, Optional SiteKey As Long) As Boolean
   Dim strSQL As String
   Dim loRsType As Recordset
   On Error GoTo FunctionError
   
   Set loRsType = New Recordset
   '
   ' Get list of link types
   '
   loRsType.CursorLocation = adUseClient
   strSQL = "SELECT * from mwEventLinkType " & _
     " Where mwEventTypeKey=" & EventType & " And IsSuggested <> 0 " & _
     " ORDER BY DisplaySequence"
   loRsType.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsType.RecordCount < 1 Then
      CloseRecordset loRsType
      AutoCreate = True
      Exit Function
   End If
   
   CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open "mwEventLinkLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   '
   Do While Not loRsType.EOF
      '
      '
      moRS.AddNew
      moRS.Fields(RS_ID).value = goSession.MakePK("mwEventLinkLog")
      moRS.Fields(RS_mwEventTypeKey).value = EventType
      moRS.Fields(RS_mwEventDetailKey).value = EventDetailKey
      moRS.Fields(RS_mwEventLinkTypeKey).value = loRsType!ID
      moRS.Fields(RS_IsCreated).value = False
      moRS.Fields(RS_DateTimeCreated).value = Now()
      If Not IsNull(loRsType!DefaultDescription) Then
         moRS.Fields(RS_BriefDescription).value = loRsType!DefaultDescription
      End If
      
      If SiteKey > 0 Then
         moRS.Fields(RS_mwcSitesKey).value = SiteKey
      End If
      
      moRS.Fields(RS_IsFastTrack).value = False
      moRS.Fields(RS_IsComplete).value = False
      moRS.Fields(RS_IsNotComplete).value = False
      moRS.Fields(RS_mwcRoleTypeKey).value = goSession.User.RoleTypeKey
      moRS.Fields(RS_mwcUsersKey).value = goSession.User.UserKey
      moRS.Update
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   CloseRecordset moRS
   AutoCreate = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventLinkWork.AutoCreate. ", Err.Number, Err.Description
   AutoCreate = False
End Function


Public Function EditLink(ByRef moRS As Recordset) As Boolean
   Dim loform As frmEventLink
   Dim strContext As String
   On Error GoTo FunctionError
   If moRS.RecordCount < 1 Or moRS.EOF Or moRS.BOF Then
      goSession.RaiseError "Error in mwEventLinkWork.EditLink, invalid Recordset passed to function."
      EditLink = False
      Exit Function
   End If
   Set loform = New frmEventLink
   If Not loform.InitForm(moRS.Fields(RS_mwEventTypeKey).value, moRS) Then
      EditLink = False
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   '
   ' Edit Link Form
   '
   loform.Show vbModal
   If loform.IsCancelled Then
      EditLink = False
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   ' Is the link deleted ?
   If loform.IsDeleteLink Then
      'moRS.Fields(RS_FullFilename).value = Null
      'moRS.Fields(RS_ContextID).value = Null
      'moRS.Fields(RS_IsCreated).value = False
      mIsDeleteLink = True
   End If
   moRS.Update
   EditLink = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventLinkWork.EditLink. ", Err.Number, Err.Description
   EditLink = False
End Function

Public Function OpenLink(ByRef moRS As Recordset) As Boolean
   On Error GoTo FunctionError
   If moRS.RecordCount < 1 Or moRS.EOF Or moRS.BOF Then
      goSession.RaiseError "Error in mwEventLinkWork.OpenLink, invalid Recordset passed to function."
      OpenLink = False
      Exit Function
   End If
   'If IsNull(moRS.Fields(RS_FullFilename).value) Then
   '   goSession.RaiseError "Error in mwEventLinkWork.OpenLink, No link Filename in record: " & moRS.Fields(RS_ID).value
   '   OpenLink = False
   '   Exit Function
   'End If
   '
   ' Launch External Link...
   '
   'If IsNull(moRS.Fields(RS_ContextID).value) Or moRS.Fields(RS_ContextID).value = -1 Then
   '   goSession.API.LaunchExternalLink moRS.Fields(RS_FullFilename).value
   'Else
   '   goSession.API.LaunchExternalLink moRS.Fields(RS_FullFilename).value, moRS.Fields(RS_ContextID).value
   'End If
   OpenLink = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventLinkWork.OpenLink. ", Err.Number, Err.Description
   OpenLink = False
End Function


Public Function IsEventLinkTypesExist(EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select count(*) from mwEventLinkType where mwEventTypeKey=" & EventType
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.Fields(0).value > 0 Then
      IsEventLinkTypesExist = True
   Else
      IsEventLinkTypesExist = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventLinkWork.IsEventLinkTypesExist. ", Err.Number, Err.Description
End Function


' Return name of new filename...
Public Function CopyLinkFileToEventFolder(FullFilename As String) As String

   MsgBox "mwEventLinkWork.CopyLinkFileToEventFolder. Function Not implemented."

End Function


Public Function DeleteLinkRecords(EventType As Long, EventKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   
   CloseRecordset moRS
   sSQL = "SELECT * FROM mwEventLinkLog WHERE mwEventTypeKey=" & EventType & _
     " AND mwEventDetailKey=" & EventKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   ' Got Links ?
   If moRS.RecordCount < 1 Then
      DeleteLinkRecords = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   
   ' delete Links
   Do While Not moRS.EOF
      moRS.Delete
      moRS.MoveNext
   Loop
   
   goSession.CloseRecordset moRS
   DeleteLinkRecords = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventLinkWork.DeleteLinkRecords. ", Err.Number, Err.Description
   DeleteLinkRecords = False
End Function


Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwEventLinkLog") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwEventLinkWork.moRs_WillChangeRecord. ", Err.Number, Err.Description
End Sub



