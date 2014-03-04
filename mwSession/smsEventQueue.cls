VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "smsEventQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' SMS EventQueue Class -
' 2/17/2002 ms
'

Option Explicit

Private moSession As Session
Private moCon As ADODB.Connection
Private moRS As Recordset
Private moNewDirectives As Collection
Private mIsEvents As Boolean
Private mIsDirectiveEmailErrors As Boolean
Private mIsWorkflowTransitionErrors As Boolean

Public Property Set Session(ByVal Data As Session)
   If moSession Is Nothing Then
      Set moSession = Data
   End If
End Property

Public Property Get Session() As Session
   If moSession Is Nothing Then
      Set Session = moSession
   End If
End Property

Public Property Set NewDirectives(ByVal Data As Collection)
   Set moNewDirectives = Data
End Property

Public Property Get NewDirectives() As Collection
   Set NewDirectives = moNewDirectives
End Property

Public Property Get IsEvents() As Boolean
   IsEvents = mIsEvents
End Property

Public Property Let IsDirectiveEmailErrors(ByVal Data As Boolean)
   mIsDirectiveEmailErrors = Data
End Property

Public Property Get IsDirectiveEmailErrors() As Boolean
   IsDirectiveEmailErrors = mIsDirectiveEmailErrors
End Property

Public Property Let IsWorkflowTransitionErrors(ByVal Data As Boolean)
   mIsWorkflowTransitionErrors = Data
End Property

Public Property Get IsWorkflowTransitionErrors() As Boolean
   IsWorkflowTransitionErrors = mIsWorkflowTransitionErrors
End Property

'Public Property Let A(ByVal Data As String)
'   mA = Data
'End Property

'Public Property Get A() As String
'   A = mA
'End Property

' Open SQL Connection, create record set with all events to process
Public Function OpenEventQueue() As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   '
   ' Open connection
   '
   Set moCon = New Connection
   moCon.Open moSession.Workflow.DbEmailRouterString
   '
   ' Filter out records with not used wf_stepTID...
   '
   strSQL = "Update WorkflowStepHistory set IsProcessed = 1 where (Wf_StepTID<>4 and Wf_StepTID<>7)"
   moCon.Execute strSQL
   
   '
   ' Recordset of event queue items
   '
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   strSQL = "select * from WorkflowStepHistory where IsProcessed=0"
   moRS.Open strSQL, moCon, adOpenStatic, adLockOptimistic
   '
   ' Trap no events to populate
   '
   If moRS.RecordCount < 1 Then
      mIsEvents = False
   Else
      mIsEvents = True
   End If
   '
   ' Extract collection of Directives
   '
   ' StepTID 7 are Directives, as per WorkflowStepType table...
   moRS.Filter = "WF_StepTID=7"
   If moRS.RecordCount > 0 Then
      '
      ' Build the collection of Directives
      '
      Do While Not moRS.EOF
         ProcessDirectiveBody moRS!Body
         moRS.MoveNext
      Loop
   End If
   moRS.Filter = adFilterNone
   OpenEventQueue = True
   Exit Function
FunctionError:
   moSession.RaiseError "General error in smsEventQueue.OpenEventQueue.  ", err.Number, err.Description
   OpenEventQueue = False
End Function

Private Function ProcessDirectiveBody(Body As String) As Boolean
   Dim strDirectiveID As String
   Dim iStart As Integer
   Dim iLength As Integer
   On Error GoTo FunctionError
   iStart = InStr(1, Body, "Directive ID =", vbBinaryCompare) + 15
   iLength = (InStr(1, Body, "Role Type", vbBinaryCompare) - 2) - iStart
   strDirectiveID = mID(Body, iStart, iLength)
   moNewDirectives.Add strDirectiveID, strDirectiveID
   ProcessDirectiveBody = True
   Exit Function
FunctionError:
   ' Duplicate - cool, only want one entry per Directive
   If err.Number = 457 Then
      Resume Next
   Else
      moSession.RaiseError "General Error in smsEventQueue.ProcessDirectiveBody ", err.Number, err.Description
   End If
   ProcessDirectiveBody = False
End Function

'
' Return recordset pointer, with filter turned on...
'
Public Function GetWorkflowTransitions() As Recordset
   moRS.Filter = "WF_StepTID<>7"
   Set GetWorkflowTransitions = moRS
End Function

Public Function MaybeCloseWorkflowItems() As Boolean
   On Error GoTo FunctionError
   moRS.Filter = adFilterNone
   moRS.MoveFirst
   Do While Not moRS.EOF
      If (Not mIsDirectiveEmailErrors) And moRS!WF_StepTID = 7 Then
         moRS!IsProcessed = "1"
      ElseIf (Not mIsWorkflowTransitionErrors) And moRS!WF_StepTID = 4 Then
         moRS!IsProcessed = "1"
      End If
      moRS.Update
      moRS.MoveNext
   Loop
   MaybeCloseWorkflowItems = True
   Exit Function
FunctionError:
   moSession.RaiseError "General error in smsEventQueue.CloseWorkflowItems.  ", err.Number, err.Description
   MaybeCloseWorkflowItems = False
   Exit Function
End Function


Public Function GetOccurrenceCode(Body) As String
   Dim strDirectiveID As String
   Dim iStart As Integer
   Dim iLength As Integer
   On Error GoTo FunctionError
   iStart = InStr(1, Body, "Internal Coding Number:", vbBinaryCompare) + 24
   iLength = (InStr(1, Body, "Originating Site Name", vbBinaryCompare) - 2) - iStart
   If iStart <> 0 And iLength <> 0 Then
      GetOccurrenceCode = mID(Body, iStart, iLength)
   Else
      GetOccurrenceCode = ""
   End If
   Exit Function
FunctionError:
   moSession.RaiseError "General error in smsEventQueue.GetOccurrenceCode.  ", err.Number, err.Description
   GetOccurrenceCode = ""
   Exit Function
End Function


Private Sub Class_Initialize()
   Set moNewDirectives = New Collection
End Sub

Private Sub Class_Terminate()
   On Error GoTo SubError
   If Not moRS Is Nothing Then
      If moRS.State = adStateOpen Then
         If moRS.LockType <> adLockReadOnly Then
            If Not (moRS.EOF Or moRS.BOF) Then moRS.Update
         End If
         moRS.Close
      End If
      Set moRS = Nothing
   End If
   If Not moCon Is Nothing Then
      If moCon.State = adStateOpen Then
         moCon.Close
      End If
      Set moCon = Nothing
   End If
   Set moNewDirectives = Nothing
   Exit Sub
SubError:
   Session.RaiseError "General error in mwWorkflow.smsEventQueue.class_terminate. ", err.Number, err.Description
End Sub


Public Function SendDirectiveDatagrams() As Boolean

   '
   '  Creates Directive Datagrams for transmission to vessels
   '
   Dim loform As mwForm
   Dim fso As FileSystemObject
   Dim loRS As Recordset
   Dim strDirectiveID As String
   Dim i As Integer
   Dim strSQL As String
   Dim strEmail As String
   On Error GoTo FunctionError
   '
   ' Any to send ?
   '
   If moNewDirectives.Count < 1 Then
      SendDirectiveDatagrams = True
      Exit Function
   End If
   '
   ' Must be configured for email (not loopback or SendByMedia...
   '
   'If moSession.User.DefaultTransport <> mwt_ElectronicMail Then
   '   If moSession.Workflow.IsSilentErrors Then
   '      moSession.RaiseWorkflowStatus "Error in smsEventQueue.SendDirectiveDatagrams, Default Transport for User is not Electronic Mail."
   '   Else
   '      moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, Default Transport for User is not Electronic Mail."
   '   End If
   '   SendDirectiveDatagrams = False
   '   Exit Function
   'End If
   Set fso = New FileSystemObject
   '
   ' Loop through each Directive...
   '
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   For i = 1 To moNewDirectives.Count
      strDirectiveID = moNewDirectives.Item(i)
      '
      ' open up form template...
      '
      Set loform = moSession.OpenNewForm
      If Not loform.MakeNewForm("SMS_DIRECTIVE_DGM") Then
         moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, unable to open Form Template: SMS_DIRECTIVE_DGM"
         SendDirectiveDatagrams = False
         moSession.TearDownForm loform
         moSession.KillObject fso
         Exit Function
      End If
      '
      ' make sure delete after send is off - we will manage that manually...
      '
      loform.TemplateProperties.DeleteAfterTransmit = False
      '
      ' Fetch Directive, and lookup names...
      '
'      strSQL = "SELECT Directive.*, DirectiveStatus.Name as DirStatusName, DirectiveType.Name as DirTypeName, " & _
'        "DepartmentType.DepartmentShortName, AcknowledgementType.Name as AckName " & _
'        "FROM ((Directive INNER JOIN DirectiveStatus ON Directive.DirectiveStatusID = DirectiveStatus.DirectiveStatusID) " & _
'        " INNER JOIN DirectiveType ON Directive.DirectiveTID = DirectiveType.DirectiveTID) " & _
'        " INNER JOIN DepartmentType ON Directive.DepartmentTID = DepartmentType.DepartmentTID " & _
'        " INNER JOIN AcknowledgementType ON Directive.AcknowledgementTID = AcknowledgementType.AcknowledgementTID " & _
'        " where DirectiveID=" & strDirectiveID

      strSQL = "SELECT Directive.*, DirectiveStatus.Name as DirStatusName, DirectiveType.Name as DirTypeName, " & _
        " DepartmentType.DepartmentShortName, AcknowledgementType.Name as AckName " & _
        " FROM Directive , DirectiveStatus, DirectiveType, DepartmentType, AcknowledgementType " & _
        " WHERE Directive.DirectiveStatusID = DirectiveStatus.DirectiveStatusID AND " & _
        " Directive.DirectiveTID = DirectiveType.DirectiveTID AND " & _
        " Directive.DepartmentTID = DepartmentType.DepartmentTID AND " & _
        " Directive.AcknowledgementTID = AcknowledgementType.AcknowledgementTID AND " & _
        " DirectiveID = " & strDirectiveID

      'strSQL = "select * from Directive where DirectiveID=" & strDirectiveID
      loRS.Open strSQL, moCon, adOpenForwardOnly, adLockReadOnly
      If loRS.RecordCount < 1 Then
         moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, Directive record missing: Directive ID: " & strDirectiveID
         SendDirectiveDatagrams = False
         moSession.CloseRecordset loRS
         moSession.TearDownForm loform
         moSession.KillObject fso
         Exit Function
      End If
      
      '
      ' Convert Directive to Record in datagram
      '
      If Not loform.DataForm.CreateDatagramFromRecordset(loRS, "smsDirective") Then
         moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, error creating datagram for : " & strDirectiveID
         SendDirectiveDatagrams = False
         moSession.CloseRecordset loRS
         moSession.TearDownForm loform
         moSession.KillObject fso
         Exit Function
      End If
      '
      ' Now go after the Directive Occurrences
      '
      loRS.Close
      
'      strSQL = "SELECT DISTINCT DirectiveOccurrence.DirectiveID, DirectiveOccurrence.OccurrenceID, " & _
'        "Site.Name as OccurrenceSiteName, Occurrence.BriefDescription, Occurrence.OccurrenceCode  FROM (DirectiveOccurrence INNER JOIN Directive ON " & _
'        "DirectiveOccurrence.DirectiveID = Directive.DirectiveID) INNER JOIN (Site INNER JOIN Occurrence " & _
'        "ON Site.SiteID = Occurrence.RefSiteID) ON (DirectiveOccurrence.OccurrenceSiteID = Occurrence.SiteID) " & _
'        "AND (DirectiveOccurrence.OccurrenceID = Occurrence.OccurrenceID)" & _
'        "WHERE (((DirectiveOccurrence.DirectiveID)=" & strDirectiveID & "));"
      
      strSQL = "SELECT DISTINCT DirectiveOccurrence.DirectiveID, DirectiveOccurrence.OccurrenceID, " & _
        " Site.Name as OccurrenceSiteName, Occurrence.BriefDescription, Occurrence.OccurrenceCode  " & _
        " FROM DirectiveOccurrence , Directive, Site, Occurrence " & _
        " WHERE DirectiveOccurrence.DirectiveID = Directive.DirectiveID AND " & _
        " Site.SiteID = Occurrence.RefSiteID AND " & _
        " DirectiveOccurrence.OccurrenceSiteID = Occurrence.SiteID AND " & _
        " DirectiveOccurrence.OccurrenceID = Occurrence.OccurrenceID AND " & _
        " DirectiveOccurrence.DirectiveID = " & strDirectiveID
      
      loRS.Open strSQL, moCon, adOpenForwardOnly, adLockReadOnly
      If loRS.RecordCount > 0 Then
         If Not loform.DataForm.CreateDatagramFromRecordset(loRS, "smsDirectiveOccurrence") Then
            moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, error creating DirectiveOccurrence datagram for : " & strDirectiveID
            SendDirectiveDatagrams = False
            moSession.CloseRecordset loRS
            moSession.TearDownForm loform
            moSession.KillObject fso
            Exit Function
         End If
      End If
      loRS.Close
      '
      ' Documents and Notes...
      '
'      strSQL = "SELECT Directive.DirectiveID, DirectiveNote.NoteText, DocumentStatusType.Name as StatusName " & _
'        "FROM (DirectiveNote INNER JOIN Directive ON DirectiveNote.DirectiveID = Directive.DirectiveID) " & _
'        " INNER JOIN DocumentStatusType ON DirectiveNote.StatusTID = DocumentStatusType.DocumentStatusTID " & _
'        " WHERE (((Directive.DirectiveID)=" & strDirectiveID & "));"
      
      strSQL = "SELECT Directive.DirectiveID, DirectiveNote.NoteText, DocumentStatusType.Name as StatusName " & _
        " FROM DirectiveNote , Directive, DocumentStatusType " & _
        " WHERE DirectiveNote.DirectiveID = Directive.DirectiveID AND " & _
        " DirectiveNote.StatusTID = DocumentStatusType.DocumentStatusTID AND " & _
        " Directive.DirectiveID = " & strDirectiveID
      
      loRS.Open strSQL, moCon, adOpenForwardOnly, adLockReadOnly
      If loRS.RecordCount > 0 Then
         If Not loform.DataForm.CreateDatagramFromRecordset(loRS, "smsDirectiveNote") Then
            moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, error creating DirectiveNote datagram for : " & strDirectiveID
            SendDirectiveDatagrams = False
            moSession.CloseRecordset loRS
            moSession.TearDownForm loform
            moSession.KillObject fso
            Exit Function
         End If
      End If
      loRS.Close
      '
      ' Directive Site Roles
      '
'   strSQL = "SELECT distinct Directive.DirectiveID, DirectiveSiteRole.DestRoleTID as RoleTypeID " & _
'     "FROM Directive INNER JOIN DirectiveSiteRole ON Directive.DirectiveID = DirectiveSiteRole.DirectiveID " & _
'     "WHERE (((Directive.DirectiveID)=" & strDirectiveID & "));"
      
      strSQL = "SELECT distinct Directive.DirectiveID, DirectiveSiteRole.DestRoleTID as RoleTypeID " & _
         " FROM Directive , DirectiveSiteRole " & _
         " WHERE Directive.DirectiveID = DirectiveSiteRole.DirectiveID AND " & _
         " Directive.DirectiveID = " & strDirectiveID
      
      loRS.Open strSQL, moCon, adOpenForwardOnly, adLockReadOnly
      If loRS.RecordCount > 0 Then
         If Not loform.DataForm.CreateDatagramFromRecordset(loRS, "smsDirectiveRole") Then
            moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, error creating DirectiveNote datagram for : " & strDirectiveID
            SendDirectiveDatagrams = False
            moSession.CloseRecordset loRS
            moSession.TearDownForm loform
            moSession.KillObject fso
            Exit Function
         End If
      End If
      loRS.Close
      '
      ' Create the physical file to send..
      '
      If Not loform.DataForm.WriteDataForm() Then
         moSession.RaiseError "Error in smsEventQueue.SendDirectiveDatagrams, error creating DirectiveOccurrence datagram for : " & strDirectiveID
         SendDirectiveDatagrams = False
         moSession.CloseRecordset loRS
         moSession.TearDownForm loform
         moSession.KillObject fso
         Exit Function
      End If
      
      '
      ' Test Send Datagram - This will be a loop to all destinations (or pass an array of SiteID to a new function)
      '
'      strSQL = "SELECT distinct  SiteRole.SiteID " & _
'        "FROM SiteRole INNER JOIN (DirectiveSiteRole " & _
'        "  INNER JOIN Directive ON DirectiveSiteRole.DirectiveID = Directive.DirectiveID) " & _
'        "     ON (SiteRole.RoleTID = DirectiveSiteRole.DestRoleTID) " & _
'        "     AND (SiteRole.SiteID = DirectiveSiteRole.DestSiteID) " & _
'        "     WHERE (((Directive.DirectiveID)=" & strDirectiveID & ")) " & _
'        "     ORDER BY SiteRole.SiteID;"

      strSQL = "SELECT distinct  SiteRole.SiteID " & _
        " FROM SiteRole , DirectiveSiteRole, Directive " & _
        " WHERE DirectiveSiteRole.DirectiveID = Directive.DirectiveID AND " & _
        " SiteRole.RoleTID = DirectiveSiteRole.DestRoleTID AND " & _
        " SiteRole.SiteID = DirectiveSiteRole.DestSiteID AND " & _
        " Directive.DirectiveID = " & strDirectiveID & _
        " ORDER BY SiteRole.SiteID"

      loRS.Open strSQL, moCon, adOpenStatic, adLockReadOnly
      If loRS.RecordCount < 1 Then
         moSession.RaiseWorkflowStatus "Directive sent with no destinations: " & strDirectiveID
      Else
         '
         ' Loop through DirectiveSites to send to...
         '
         Do While Not loRS.EOF
            strEmail = GetEmailFromYetAnotherSiteID(loRS!SiteID)
            If strEmail <> "" Then
               If Not loform.Send(strEmail) Then
                  moSession.RaiseWorkflowStatus "Error sending in smsEventQueue.SendDirectiveDatagrams, sending email to Site: " & _
                    loRS!SiteID & ", Email: " & strEmail
                  IsDirectiveEmailErrors = True
               Else
                  moSession.RaiseWorkflowStatus "Directive Datagram sent to Site: " & _
                    loRS!SiteID & ", Email: " & strEmail
               End If
            End If
            loRS.MoveNext
            '
            '
         Loop
      End If
      '
      ' Cleanup objects
      '
      ' Manually delete Directive Datagram file...
      moSession.Message.AttachmentFile = ""
      fso.DeleteFile loform.FormProperties.FullFilename
      moSession.TearDownForm loform
      loRS.Close
   Next i ' Next Directive
   '
   ' We Be Done...
   '
   If IsDirectiveEmailErrors Then
      SendDirectiveDatagrams = False
   Else
      SendDirectiveDatagrams = True
   End If
   ' Final cleanup
   moSession.CloseRecordset loRS
   Exit Function
FunctionError:
   moSession.RaiseError "General error in smsEventQueue.SendDirectiveDatagrams.  ", err.Number, err.Description
   IsDirectiveEmailErrors = True
   SendDirectiveDatagrams = False
   Exit Function
End Function

Private Function GetEmailFromYetAnotherSiteID(yaSiteID As String) As String
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select ShipWorkflowSendToAddress from mwcSites where YetAnotherSiteID = '" & yaSiteID & "'"
   loRS.Open strSQL, moSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.RecordCount < 1 Then
      If moSession.Workflow.IsSilentErrors Then
         moSession.RaiseWorkflowStatus "Error in smsEventQueue.GetEmailFromYetAnotherSiteID, Yet Another SiteID missing for: " & yaSiteID
      Else
         moSession.RaiseError "Error in smsEventQueue.GetEmailFromYetAnotherSiteID, Yet Another SiteID missing for: " & yaSiteID
      End If
      GetEmailFromYetAnotherSiteID = ""
   Else
      GetEmailFromYetAnotherSiteID = loRS!ShipWorkflowSendToAddress
   End If
   moSession.CloseRecordset loRS
   Exit Function
FunctionError:
   moSession.RaiseError "General error in smsEventQueue.GetEmailFromYetAnotherSiteID.  ", err.Number, err.Description
   GetEmailFromYetAnotherSiteID = ""
   Exit Function
End Function

