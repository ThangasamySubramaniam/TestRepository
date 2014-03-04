VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwAlertWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwAlertWork - ShipNet Fleet Alert Subsystem
' May 2004 tjm/ms

Option Explicit

Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1
'mod variables
   Dim mEventTypeKey As Long
   Dim mEventDetailKey As Long
   
   Dim mAlertEmailAddress  As String
   
 
Public Property Let EventTypeKey(nEventTypeKey As Long)
   mEventTypeKey = nEventTypeKey
End Property

Public Property Let EventDetailKey(nEventDetailKey As Long)
   mEventDetailKey = nEventDetailKey
End Property

  

Public Function CreateUserAlert(ByRef AlertTitle As String, ByRef AlertDetails As String, ByRef AlertTarget As String) As Long
   Dim loform As frmAlertCreate
   
   On Error GoTo FunctionError
   
   Set loform = New frmAlertCreate
   loform.AlertTitle = AlertTitle
   loform.AlertDescription = AlertDetails
   loform.AlertTarget = AlertTarget
   
   ' pass in additional values mwEventTypeKey, mwEventDetailKey to view alert
   loform.EventTypeKey = mEventTypeKey
   loform.EventDetailKey = mEventDetailKey
   
   loform.Show vbModal
   
   If Not loform.IsCancelled Then
      CreateUserAlert = loform.NewID
      AlertDetails = loform.GetAlertDetails
      AlertTitle = loform.GetAlertTitle
      AlertTarget = loform.GetAlertTarget
   Else
      CreateUserAlert = 0
      AlertDetails = ""
      AlertTitle = ""
      AlertTarget = ""
   End If
   
   Unload loform
   KillObject loform
   
   DoEvents
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwAlertWork.CreateUserAlert: ", Err.Number, Err.Description
   KillObject loform
   CreateUserAlert = 0
   AlertDetails = ""
   AlertTitle = ""
   AlertTarget = ""
   
End Function

Public Function DisplayAlertLog(mwRoleTypeKey As Long, Optional ShowNew As Boolean) As Boolean
   Dim loform As frmAlertLog
   
   On Error GoTo FunctionError
   
   Set loform = New frmAlertLog
   loform.SetmwRoleTypeKey mwRoleTypeKey
   loform.SetShowNewAlerts ShowNew
   
   loform.Show vbModal
   
   Unload loform
   KillObject loform
   
   DoEvents
   DisplayAlertLog = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwAlertWork.DisplayAlertLog: ", Err.Number, Err.Description
   KillObject loform
   DisplayAlertLog = False
End Function

Public Function ChangeAlertNotifyInterval() As Boolean
   Dim loform As frmAlertNotifyInterval
   
   On Error GoTo FunctionError
   
   Set loform = New frmAlertNotifyInterval
   
   loform.Show vbModal
   
   If Not loform.IsCancelled Then
      ChangeAlertNotifyInterval = True
   End If
   Unload loform
   KillObject loform
   
   DoEvents
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwAlertWork.ChangeAlertNotifyInterval: ", Err.Number, Err.Description
   KillObject loform
   ChangeAlertNotifyInterval = False
End Function

Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   '
   ' mwAlertLog must use this special WillChangeRecord handler because it is a
   ' Site Specific table that does NOT have an mwcSitesKey.
   
   ' DO NOT change this to the new-style mwReplicateWillChange class!
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRS!ID, moRS!mwcSitesKeyTarget, moRS.Fields
      Else
         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRS!ID, goSession.Site.SiteKey, moRS.Fields
      End If
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRS!ID, moRS!mwcSitesKeyTarget
      Else
         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRS!ID, goSession.Site.SiteKey
      End If
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRS.Fields, moRS!mwcSitesKeyTarget
      Else
         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRS.Fields, goSession.Site.SiteKey
      End If
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwAlertWork.moRs_WillChangeRecord. ", Err.Number, Err.Description
End Sub

Public Function GetUnreadCount(Optional RoleTypeKey As Long) As Long
   Dim sSQL As String
   Dim nRoleTypeKey As Long
   Dim loRs As Recordset
   
   On Error GoTo FunctionError
   
   If RoleTypeKey < 1 Then
      nRoleTypeKey = goSession.User.RoleTypeKey
   Else
      nRoleTypeKey = RoleTypeKey
   End If

   'By N.Angelakis On 28th May 2009
   'DEV-1154-Voyage Instruction Export - Informing user of newly arrived instructions
   'sSQL = "Select count(*) from mwAlertlog WHERE mwcSitesKeyTarget=" & goSession.Site.SiteKey & _
          " and mwcRoleTypeKeyTarget=" & nKey & " and mwAlertLogStatusKey=1"
          
   sSQL = "Select count(*) from mwAlertlog WHERE mwcSitesKeyTarget=" & goSession.Site.SiteKey & " and mwAlertLogStatusKey=1" & _
          " AND ((mwcRoleTypeKeyTarget = " & nRoleTypeKey & " AND mwcUsersKeyTarget Is Null) " & _
          "  OR  (mwcRoleTypeKeyTarget = " & nRoleTypeKey & " AND mwcUsersKeyTarget = " & goSession.User.UserKey & ")) "
          
   Set loRs = New Recordset
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   GetUnreadCount = loRs.Fields(0).value
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.GetUnreadCount: ", Err.Number, Err.Description
   CloseRecordset loRs
End Function


Public Function CreateSysAdminAlert(TargetSiteKey As Long, Title As String, ExternalData As String) As Boolean
   
   'By N.Angelakis On 21st May 2009
   CreateSysAdminAlert = CreateAlert(TargetSiteKey, _
                              Title, _
                              0, _
                              "", _
                              MW_ALERT_STATUS_SENT, _
                              MW_ALERT_TYPE_ADMIN, _
                              goSession.User.RoleTypeKey, _
                              ExternalData)
   
'   Dim sSQL As String
'   On Error GoTo FunctionError
'   Set moRS = New Recordset
'   moRS.CursorLocation = adUseClient
'   sSQL = "select * from mwAlertlog where ID=-1"
'   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
'   With moRS
'      moRS.AddNew
'      .Fields!ID = goSession.MakePK("mwAlertLog")
'      .Fields!Title = Title
'      .Fields!AlertDateTime = Now()
'      .Fields!mwcSitesKeySource = goSession.Site.SiteKey
'      .Fields!mwcRoleTypeKeySource = goSession.User.RoleTypeKey
'      .Fields!mwcSitesKeyTarget = TargetSiteKey
'      .Fields!ExternalData = ExternalData
'      .Fields!mwAlertLogStatusKey = MW_ALERT_STATUS_SENT
'      .Fields!mwAlertTypeKey = MW_ALERT_TYPE_ADMIN
'      .Update
'   End With
'   CloseRecordset moRS

   CreateSysAdminAlert = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.CreateSysAdminAlert: ", Err.Number, Err.Description
   CloseRecordset moRS
End Function

Public Function CreateApplicationAlert(TargetSiteKey As Long, _
                                       TargetRoleTypeKey As Long, _
                                       Title As String, _
                                       AlertDetails As String) As Boolean
   
   'By N.Angelakis On 21st May 2009
   CreateApplicationAlert = CreateAlert(TargetSiteKey, _
                                          Title, TargetRoleTypeKey, _
                                          AlertDetails, _
                                          MW_ALERT_STATUS_SENT, _
                                          MW_ALERT_TYPE_ADMIN, _
                                          goSession.User.RoleTypeKey, "")
   
   CreateApplicationAlert = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.CreateApplicationAlert: ", Err.Number, Err.Description
   CloseRecordset moRS
   CreateApplicationAlert = False
End Function
Public Function CreateApplicationEventAlert(TargetSiteKey As Long, _
                                            TargetRoleTypeKey As Long, _
                                            Title As String, _
                                            AlertDetails As String, _
                                            UsersKeySource As Long, _
                                            UsersKeyTarget As Long, _
                                            mwEventTypeKey As Long, _
                                            mwEventDetailKey As Long _
                                            ) As Boolean
   
   CreateApplicationEventAlert = CreateAlert(TargetSiteKey, _
                                          Title, TargetRoleTypeKey, _
                                          AlertDetails, _
                                          MW_ALERT_STATUS_SENT, _
                                          MW_ALERT_TYPE_ADMIN, _
                                          goSession.User.RoleTypeKey, "", _
                                          UsersKeySource, UsersKeyTarget, _
                                          mwEventTypeKey, mwEventDetailKey)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.CreateApplicationAlert: ", Err.Number, Err.Description
   CreateApplicationEventAlert = False
End Function

Friend Function ProcessSysAdminAlerts() As Boolean
   Dim sSQL As String
   Dim s() As String
   Dim loRepWork As mwReplicateBatchWork
   On Error GoTo FunctionError
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwAlertlog where mwcSitesKeyTarget=" & goSession.Site.SiteKey & _
          " and mwcRoleTypeKeyTarget is null and mwAlertLogStatusKey=" & MW_ALERT_STATUS_SENT
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If moRS.RecordCount < 1 Then
      CloseRecordset moRS
      ProcessSysAdminAlerts = True
      Exit Function
   End If
   Do While Not moRS.EOF
      If Not IsNull(moRS!Title) Then
         Select Case UCase(moRS!Title)
            Case Is = "TRANSMIT_ACKBACK"
               'Format: SiteKey|BatchTypeKey|BatchKey
               If IsNull(moRS!ExternalData) Then
                  goSession.RaiseWorkflowStatus "Error in ProcessSysAdminAlerts, TRANSMIT_ACKBACK - ExternalData is Null" & moRS!ID
                  ProcessSysAdminAlerts = True
               Else
                  '
                  ' TRANSMIT_ACKBACK
                  '
                  s = Split(moRS!ExternalData, "|")
                  If UBound(s) < 2 Then
                     goSession.RaiseWorkflowStatus "Error in ProcessSysAdminAlerts, TRANSMIT_ACKBACK - ExternalData is Null: " & moRS!ID
                     ProcessSysAdminAlerts = False
                  ElseIf Not (IsNumeric(s(0)) And IsNumeric(s(1)) And IsNumeric(s(2))) Then
                     goSession.RaiseWorkflowStatus "Error in ProcessSysAdminAlerts, TRANSMIT_ACKBACK - ExternalData is Invalid: " & moRS!ID
                     ProcessSysAdminAlerts = False
                  Else
                     Set loRepWork = New mwReplicateBatchWork
                     If Not loRepWork.UpdateAckBack(CLng(s(0)), CLng(s(1)), CLng(s(2))) Then
                        goSession.RaiseWorkflowStatus "Error in ProcessSysAdminAlerts, TRANSMIT_ACKBACK - Rejected: " & moRS!ID
                        ProcessSysAdminAlerts = False
                     Else
                        moRS!ReceiverNotes = "TRANSMIT_ACKBACK OK."
                        ProcessSysAdminAlerts = True
                     End If
                  End If
               End If
            Case Is = "RESEND_REQUEST"
               'Format: SiteKey|BatchTypeKey|BatchKey
               goSession.RaiseWorkflowStatus "Stump Code in mwAlertWork.ProcessSysAdminAlerts. RESEND_REQUEST."
               moRS!ReceiverNotes = "REQUEST_RESEND OK."
               ProcessSysAdminAlerts = True
            Case Else
               goSession.RaiseWorkflowStatus "Error in mwAlertWork.ProcessSysAdminAlerts, Invalid Title in record: " & moRS!ID
               moRS!ReceiverNotes = "Error, Invalid Title"
               ProcessSysAdminAlerts = False
         End Select
      Else
         goSession.RaiseWorkflowStatus "Error in mwAlertWork.ProcessSysAdminAlerts, Invalid Title in record: " & moRS!ID
         moRS!ReceiverNotes = "Error, Invalid Title"
         ProcessSysAdminAlerts = False
      End If
      '
      ' Close Out Alert
      '
      moRS!mwAlertLogStatusKey = MW_ALERT_STATUS_CLOSED
      moRS!ReceivedDateTime = Now()
      moRS.Update
      moRS.MoveNext
   Loop
   CloseRecordset moRS
   ProcessSysAdminAlerts = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.CreateApplicationAlert: ", Err.Number, Err.Description
End Function


Public Function CreateAlert(ByVal TargetSiteKey As Long, _
                            ByVal AlertTitle As String, _
                              Optional ByVal TargetRoleTypeKey As Long, _
                              Optional ByVal AlertDetails As String, _
                              Optional ByVal AlertStatus As Long, _
                              Optional ByVal AlertType As Long, _
                              Optional ByVal SourceRoleTypeKey As Long, _
                              Optional ByVal ExternalData As String, _
                              Optional ByVal UsersKeySource As Long, _
                              Optional ByVal UsersKeyTarget As Long, _
                              Optional ByVal mwEventTypeKey As Long, _
                              Optional ByVal mwEventDetailKey As Long _
                              ) As Boolean

   'By N.Angelakis On 21st May 2009
   Dim sSQL As String

   On Error GoTo FunctionError

   sSQL = "SELECT * FROM mwAlertLog WHERE ID = -1"
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   
   moRS.AddNew
      
      'common entries
      moRS!ID = goSession.MakePK("mwAlertLog")
      moRS!Title = AlertTitle
      moRS!AlertDateTime = Now()
      moRS!mwcSitesKeySource = goSession.Site.SiteKey
      moRS!mwcSitesKeyTarget = TargetSiteKey
         
      If TargetRoleTypeKey > 0 Then
         moRS!mwcRoleTypeKeyTarget = TargetRoleTypeKey
      Else
         moRS!mwcRoleTypeKeyTarget = Null
      End If
      If AlertDetails <> "" Then
         moRS!AlertDetails = AlertDetails
      Else
         moRS!AlertDetails = Null
      End If
      If AlertStatus > 0 Then
         moRS!mwAlertLogStatusKey = AlertStatus
      Else
         moRS!mwAlertLogStatusKey = Null
      End If
      If AlertType > 0 Then
         moRS!mwAlertTypeKey = AlertType
      Else
         moRS!mwAlertTypeKey = Null
      End If
      If SourceRoleTypeKey > 0 Then
         moRS!mwcRoleTypeKeySource = SourceRoleTypeKey
      Else
         moRS!mwcRoleTypeKeySource = Null
      End If
      If ExternalData <> "" Then
         moRS!ExternalData = ExternalData
      Else
         moRS!ExternalData = Null
      End If

      If UsersKeySource > 0 Then
         moRS!mwcUsersKeySource = UsersKeySource
      Else
         moRS!mwcUsersKeySource = Null
      End If
      
      If UsersKeyTarget > 0 Then
         moRS!mwcUsersKeyTarget = UsersKeyTarget
      Else
         moRS!mwcUsersKeyTarget = Null
      End If
      
      If mwEventTypeKey > 0 Then
         moRS!mwEventTypeKey = mwEventTypeKey
      Else
         moRS!mwEventTypeKey = Null
      End If
      
      If mwEventDetailKey > 0 Then
         moRS!mwEventDetailKey = mwEventDetailKey
      Else
         moRS!mwEventDetailKey = Null
      End If
         
   moRS.Update
   
   If TargetRoleTypeKey > 0 Then
      CheckRoleEmail TargetRoleTypeKey
   End If
    
   If UsersKeyTarget > 0 Then
      CheckUserEmail UsersKeyTarget
   End If
    

   
   CloseRecordset moRS
   CreateAlert = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwAlertWork.CreateAlert: ", Err.Number, Err.Description
   CloseRecordset moRS
   CreateAlert = False
End Function


Public Function CascadeDeleteAlert(EventTypeKey As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   
   If EventTypeKey < 1 Or EventDetailKey < 1 Then
      Exit Function
   End If
   
   CloseRecordset moRS
   
   sSQL = "SELECT * FROM mwAlertLog WHERE mwEventTypeKey = " & EventTypeKey _
      & " And mwEventDetailKey = " & EventDetailKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(moRS) Then
      Do While moRS.EOF = False
         moRS.Delete
         moRS.MoveNext
      Loop
   End If
   
   CloseRecordset moRS
   CascadeDeleteAlert = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwAlertWork.CascadeDeleteAlert: ", Err.Number, Err.Description
   CloseRecordset moRS
   CascadeDeleteAlert = False
End Function


Private Function CheckRoleEmail(nRoleTypeKey As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   If nRoleTypeKey > 0 Then
      
      ' lookup RoleTypeEmailAddress
      sSQL = "SELECT IsSendAlertsToEmail, RoleTypeEmailAddress FROM mwcRoleType WHERE ID = " & nRoleTypeKey
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If IsRecordLoaded(loRs) Then
         If ZeroNull(loRs!IsSendAlertsToEmail) <> 0 And Not IsNull(loRs!RoleTypeEmailAddress) Then
            mAlertEmailAddress = loRs!RoleTypeEmailAddress
            CheckRoleEmail = SendEmailAlert(nRoleTypeKey, 0)
         End If
      End If
   
      CloseRecordset loRs
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwAlertWork.CheckRoleEmail: ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function CheckUserEmail(nUserKey As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   If nUserKey > 0 Then
   
      ' lookup FromEmailAddress
      sSQL = "SELECT IsSendAlertsToEmail, FromEmailAddress FROM mwcUsers WHERE ID = " & nUserKey
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If IsRecordLoaded(loRs) Then
         If ZeroNull(loRs!IsSendAlertsToEmail) <> 0 And Not IsNull(loRs!FromEmailAddress) Then
            mAlertEmailAddress = loRs!FromEmailAddress
            CheckUserEmail = SendEmailAlert(0, nUserKey)
         End If
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwAlertWork.CheckUserEmail: ", Err.Number, Err.Description
   CloseRecordset loRs
End Function
 
Private Function SendEmailAlert(Optional nRoleTypeKey As Long, Optional nUserKey As Long) As Boolean
   Dim sNotifyBody As String
   Dim sEmailTextBody As String
   Dim sFileName As String
   On Error GoTo FunctionError
   
   If nRoleTypeKey < 1 And nUserKey < 1 Then
      Exit Function
   End If
   If IsRecordLoaded(moRS) Then
      With goSession.Message
         If IsNull(moRS!Title) Then
            .subject = "Alert from " & goSession.Site.GetSiteName(ZeroNull(moRS!mwcSitesKeyTarget))
         Else
            .subject = moRS!Title
         End If
         .Body = BlankNull(moRS!AlertDetails)
         .AttachmentFile = ""
         .SendToAddress = mAlertEmailAddress
         SendEmailAlert = .Send()
      End With
   Else
      goSession.RaisePublicError "Error in mwSession.mwAlertWork.SendEmailAlert, no record loaded: "
   End If
      
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwAlertWork.SendEmailAlert: ", Err.Number, Err.Description
End Function

Public Function SetAlertNotification(oRs As Recordset, Optional RTkey As Long, Optional UserKey As Long) As Boolean
   On Error GoTo FunctionError
   
   'Pass in Alert to SendNotification
   If oRs Is Nothing Then Exit Function
   
   CloseRecordset moRS
   Set moRS = oRs.Clone ' need to be clone or will mess up ors from form
   
   'check key then sendnotification
   If RTkey > 0 Then
      SetAlertNotification = CheckRoleEmail(RTkey)
   ElseIf UserKey > 0 Then
      SetAlertNotification = CheckUserEmail(UserKey)
   End If
   
   Set moRS = Nothing
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwAlertWork.SetAlertNotification: ", Err.Number, Err.Description
   Set moRS = Nothing
End Function
