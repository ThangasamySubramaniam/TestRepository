VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
Option Explicit
Private mIsUpdateable As Boolean
Private moRSEvent As Recordset
Private moVoyInterface As mwEventInterface
Private mIsShipSelectAllowed As Boolean
Private mIsVoyageSelectAllowed As Boolean
Private mIsDateSelectAllowed As Boolean
Private mIsEventGroupOpen As Boolean
Private mIsAutoChangeStatusAllowed As Boolean
Private mGroupName As String
Private mEventGroup As Long
Private mSiteKey As Long
Private mIsEnableMultipleWFStage As Boolean

' parent child
   Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1
   Dim WithEvents moRsAudit As Recordset
Attribute moRsAudit.VB_VarHelpID = -1
   Dim WithEvents moRsWorkQ As Recordset
Attribute moRsWorkQ.VB_VarHelpID = -1

' SendAlert Sitekey
   Dim mRecordSiteKey As Long
   

Const IS_UPDATEABLE_STAGES_ONLY = True

Const WQSRAT_ACTION_OWNER = 1
Const WQSRAT_APPROVER = 2
Const WQSRAT_WATCHER = 3
Const WQSRAT_GOVERNANCE_OWNER = 4
Const WQSRAT_ACTION_CONTRIBUTOR = 5

Public Property Let IsUpdateable(ByVal Data As Boolean)
    mIsUpdateable = Data
End Property

Public Property Get IsUpdateable() As Boolean
    IsUpdateable = mIsUpdateable
End Property

Public Property Get VoyInterface() As mwEventInterface
   If moVoyInterface Is Nothing Then
      Set moVoyInterface = New mwEventInterface
   End If
   Set VoyInterface = moVoyInterface
End Property

Public Property Let VoyInterface(ByVal vData As mwEventInterface)
   Set moVoyInterface = vData
End Property

Public Property Get IsShipSelectAllowed() As Boolean
    IsShipSelectAllowed = mIsShipSelectAllowed
End Property

Public Property Let IsShipSelectAllowed(ByVal vData As Boolean)
    mIsShipSelectAllowed = vData
End Property

Public Property Get IsVoyageSelectAllowed() As Boolean
    IsVoyageSelectAllowed = mIsVoyageSelectAllowed
End Property

Public Property Let IsVoyageSelectAllowed(ByVal vData As Boolean)
    mIsVoyageSelectAllowed = vData
End Property

Public Property Get IsDateSelectAllowed() As Boolean
    IsDateSelectAllowed = mIsDateSelectAllowed
End Property

Public Property Let IsDateSelectAllowed(ByVal vData As Boolean)
    mIsDateSelectAllowed = vData
End Property
Public Property Get IsAutoChangeStatusAllowed() As Boolean
    IsAutoChangeStatusAllowed = mIsAutoChangeStatusAllowed
End Property

Public Property Let IsAutoChangeStatusAllowed(ByVal vData As Boolean)
    mIsAutoChangeStatusAllowed = vData
End Property

Public Property Get IsEnableMultipleWFStage() As Boolean
    IsEnableMultipleWFStage = mIsEnableMultipleWFStage
End Property

Public Property Let IsEnableMultipleWFStage(ByVal vData As Boolean)
    mIsEnableMultipleWFStage = vData
End Property

Public Property Get EventGroup() As Long
    EventGroup = mEventGroup
End Property

Public Property Let EventGroup(ByVal vData As Long)
    mEventGroup = vData
End Property

Public Property Get GroupName() As String
    GroupName = mGroupName
End Property

Public Property Let GroupName(ByVal vData As String)
    mGroupName = vData
End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

Public Function OpenEventGroup(EventGroup As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   If mIsEventGroupOpen Then
      goSession.RaiseError "Error in mwSession.mwEventWork.OpenEventGroup, EventGroup already open. "
      OpenEventGroup = False
      Exit Function
   End If
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwEventGroup where ID=" & EventGroup
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwSession.mwEventWork.OpenEventGroup, EventGroup not found: " & EventGroup
      OpenEventGroup = False
      CloseRecordset loRs
      Exit Function
   End If
   '
   mGroupName = loRs!GroupName
   mIsShipSelectAllowed = loRs!IsShipSelectAllowed
   mIsVoyageSelectAllowed = loRs!IsVoyageSelectAllowed
   mIsDateSelectAllowed = loRs!IsDateSelectAllowed
   If Not IsNull(loRs!IsAutoChangeStatusAllowed) Then
      mIsAutoChangeStatusAllowed = loRs!IsAutoChangeStatusAllowed
   End If
   If Not IsNull(loRs!IsEnableMultipleWFStage) Then
      mIsEnableMultipleWFStage = loRs!IsEnableMultipleWFStage
   End If
   CloseRecordset loRs
   OpenEventGroup = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.OpenEventGroup. ", Err.Number, Err.Description
   OpenEventGroup = False
End Function

Public Function FetchEventGroupTypes(EventGroup As Long) As Recordset
   Dim sSQL As String
   Dim loCmd As Command
   Dim loPar As Parameter
   Dim loRs As Recordset
   On Error GoTo FunctionError
   '
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goCon
'   sSQL = "SELECT mwEventType.* FROM mwEventGroup INNER JOIN" & _
'     " (mwEventType INNER JOIN mwEventGroupType ON mwEventType.ID = mwEventGroupType.mwEventTypeKey) " & _
'     " ON mwEventGroup.ID = mwEventGroupType.mwEventGroupKey WHERE mwEventGroup.ID=? and IsCreateEventAllowed<>0;"
   sSQL = "SELECT mwEventType.* " & _
          " FROM mwEventGroup , mwEventType , mwEventGroupType " & _
          " WHERE mwEventType.ID = mwEventGroupType.mwEventTypeKey AND " & _
          " mwEventGroup.ID = mwEventGroupType.mwEventGroupKey AND " & _
          " mwEventGroup.ID=? and IsCreateEventAllowed<>0"
   
   Set loPar = New Parameter
   loPar.Type = adInteger
   loPar.value = EventGroup
   loCmd.Parameters.Append loPar
   loCmd.CommandText = sSQL
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open loCmd, , adOpenDynamic, adLockOptimistic, adCmdUnknown
   'loRS.Properties("Unique Table").value = "mwEventType"
   Set FetchEventGroupTypes = loRs
   Set loRs = Nothing
   Set loCmd = Nothing
   Set loPar = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchEventGroupTypes. ", Err.Number, Err.Description
   Set loRs = Nothing
   Set loCmd = Nothing
   Set loPar = Nothing
End Function




Public Function AddEvent(EventType As Long, ByRef EventDetailRS As Recordset, Optional Description As String, _
 Optional EventGroupKey As Long, Optional IsAddFactsFormsLinksHistory As Boolean, Optional SiteKey As Long) As Long
   'Dim loRs As Recordset
   Dim loEhWork As mwEventHistoryWork
   Dim iDefaultStage As Long
   On Error GoTo FunctionError
   
   goSession.CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open "mwEventLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   moRS.AddNew
   moRS!ID = goSession.MakePK("mwEventLog")
   moRS!mwEventTypeKey = EventType
   moRS!mwEventDetailKey = EventDetailRS.Fields(0).value
   If Not IsNull(EventDetailRS!Vessel_ID) Then
      moRS!Vessel_ID = EventDetailRS!Vessel_ID
   Else
      moRS!Vessel_ID = goSession.Site.SiteID
   End If
   If Not IsNull(EventDetailRS!Voyage_NO) Then
      moRS!Voyage_NO = EventDetailRS!Voyage_NO
   Else
      moRS!Voyage_NO = goSession.Site.CurrentVoyage
   End If
   If Not IsNull(EventDetailRS!Report_DateTime) Then
      moRS!ReportDateTime = EventDetailRS!Report_DateTime
   Else
      moRS!ReportDateTime = Now()
   End If
   
   If SiteKey > 0 Then
      moRS!mwcSitesKey = SiteKey
   Else
      moRS!mwcSitesKey = goSession.Site.SiteKey
   End If
   moRS!IsSubmitted = EventDetailRS!sent
   '
   ' For Workflow Events, set default stage...
   '  - If EventGroupKey > 0, then its going to mwEventGroup.DefaultWfStage...
   '
   'If EventGroupKey > 0 Then
   iDefaultStage = FetchDefaultWfStageKey(EventType, EventGroupKey)
   If iDefaultStage > 0 Then
      moRS!mwEventWfStageKey = iDefaultStage
   End If
   'End If
   If Description <> "" Then
      moRS!AbstractDescription = Description
   End If
   
   moRS.Update
   AddEvent = moRS!ID
   Set loEhWork = New mwEventHistoryWork
   loEhWork.AddSystemRecord EventType, EventDetailRS.Fields(0).value, "CREATED", "This event was created.", , , , SiteKey
   KillObject loEhWork
   '
   ' Add the other stuff
   '
   If IsAddFactsFormsLinksHistory Then
      AddEventFactsFormsLinksHistory EventType, moRS!mwEventDetailKey, SiteKey
   End If
   '
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.AddEvent. ", Err.Number, Err.Description
   goSession.CloseRecordset moRS
   AddEvent = -1
End Function


Public Function FetchEventRS(EventGroup As Long, Optional Vessel_ID As String, Optional Voyage_NO As String, _
  Optional FromDate As Date, Optional ToDate As Date) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loCmd As Command
   Dim loPar As Parameter
   On Error GoTo FunctionError
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goCon
   '
   ' Step 1 Build basic SQL statement
   '
'   sSQL = "SELECT mwEventLog.ID, mwEventLog.mwEventTypeKey, mwEventType.Description, mwEventLog.ReportDateTime, " & _
'     " mwEventLog.AbstractDescription, mwEventLog.IsSubmitted, mwEventLog.mwEventDetailKey, " & _
'     " mwEventLog.Vessel_ID, mwEventLog.Voyage_NO FROM (mwEventType INNER JOIN mwEventLog " & _
'     " ON mwEventType.ID = mwEventLog.mwEventTypeKey) INNER JOIN mwEventGroupType ON mwEventType.ID" & _
'     " = mwEventGroupType.mwEventTypeKey" & _
'     " WHERE mwEventGroupType.mwEventGroupKey=? "

   sSQL = "SELECT mwEventLog.ID, mwEventLog.mwEventTypeKey, mwEventType.Description, " & _
          " mwEventLog.ReportDateTime, mwEventLog.AbstractDescription, mwEventLog.IsSubmitted, " & _
          " mwEventLog.mwEventDetailKey, mwEventLog.Vessel_ID, mwEventLog.Voyage_NO, mwEventLog.mwcSitesKey " & _
          " FROM mwEventType , mwEventLog, mwEventGroupType " & _
          " WHERE mwEventType.ID = mwEventLog.mwEventTypeKey AND " & _
          " mwEventType.ID = mwEventGroupType.mwEventTypeKey AND " & _
          " mwEventGroupType.mwEventGroupKey = ? "

   ' Event Group
   Set loPar = New Parameter
   loPar.Type = adInteger
   loPar.value = EventGroup
   loCmd.Parameters.Append loPar
   
   If Trim(Voyage_NO) <> "" And Voyage_NO <> "-1" Then
      '
      ' Ship + Voyage Number Fetch
      '
      sSQL = sSQL & " AND mwEventLog.Vessel_ID = ? AND mwEventLog.Voyage_NO = ?"
      ' Vessel...
      Set loPar = New Parameter
      loPar.Type = adVarChar
      loPar.Size = 20
      loPar.value = Vessel_ID
      loCmd.Parameters.Append loPar
      ' Voyage
      Set loPar = New Parameter
      loPar.Type = adVarChar
      loPar.Size = 20
      loPar.value = Voyage_NO
      loCmd.Parameters.Append loPar
      'locmd.
   ElseIf Trim(Vessel_ID) <> "" Then
      sSQL = sSQL & " and mwEventLog.Vessel_ID = ?"
      ' Vessel...
      Set loPar = New Parameter
      loPar.Type = adVarChar
      loPar.Size = 20
      loPar.value = Vessel_ID
      loCmd.Parameters.Append loPar
   
   End If
   If Trim(FromDate) <> "12:00:00 AM" Or Trim(ToDate) <> "12:00:00 AM" Then
      '
      ' Date Fetch
      '
      If FromDate = "12:00:00 AM" Then
         FromDate = "1/1/1901"
      End If
      If ToDate = "12:00:00 AM" Then
         ToDate = Now()
      End If
      sSQL = sSQL & " and  mwEventLog.ReportDateTime >= ? and mwEventLog.ReportDateTime <= ? "
      Set loPar = New Parameter
      loPar.Type = adDate
      loPar.value = FromDate
      loCmd.Parameters.Append loPar
      Set loPar = New Parameter
      loPar.Type = adDate
      loPar.value = ToDate
      loCmd.Parameters.Append loPar
   End If
   'sSQL = sSQL & " ORDER BY mwEventLog.ReportDateTime DESC"
   sSQL = sSQL & " ORDER BY mwEventLog.ReportDateTime DESC"
   
   '
   ' open recordset
   '
   loCmd.CommandText = sSQL
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If mIsUpdateable Then
      loRs.Open loCmd, , adOpenDynamic, adLockOptimistic, adCmdUnknown
      loRs.Properties("Unique Table").value = "mwEventLog"
      Set FetchEventRS = loRs
   Else
      loRs.Open loCmd, , adOpenDynamic, adLockOptimistic, adCmdUnknown
      loRs.Properties("Unique Table").value = "mwEventLog"
      If loRs.RecordCount < 1 Then
         goSession.RaiseError "Error in mwEventWork.FetchEventRS, no records in non-updatable table fetch."
         CloseRecordset loRs
      Else
         Set FetchEventRS = loRs
      End If
   End If
   Set loRs = Nothing
   Set loCmd = Nothing
   Set loPar = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchEventRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function RefreshEventLog(moInterface As mwEventInterface, oRsEventRecord As Recordset) As Boolean
   Dim sSQL As String
'   Dim loCmd As Command
'   Dim loPar As Parameter
'   Dim loRs2 As Recordset
'   Dim loEaWork As mwEventFactsWork
   On Error GoTo FunctionError
   '
   ' Some events don't have event detail (i.e. Shipyard Spec)
   '
   If oRsEventRecord Is Nothing Then
      RefreshEventLog = False
      Exit Function
   End If
   
   ' 03/29/05 ljh    revised for replication (EventLog Updates from detail form) [orseventrecord]
   CloseRecordset moRS
   sSQL = "SELECT * FROM mwEventLog WHERE ID = " & moInterface.EventKey
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   moRS!IsSubmitted = oRsEventRecord!sent
   moRS!ReportDateTime = oRsEventRecord!Report_DateTime
   moRS!AbstractDescription = BuildAbstractDescription(oRsEventRecord, goSession.EventTypes.Item(moInterface.EventType).AbstractTemplate)
   moRS!Voyage_NO = oRsEventRecord!Voyage_NO
   moRS.Update
   
   CloseRecordset moRS
   
   '
'   Set loCmd = New Command
'   Set loCmd.ActiveConnection = goCon
'   sSQL = "Update mwEventLog set IsSubmitted=? , ReportDateTime= ?, AbstractDescription= ?, Voyage_no=? where ID=?"
'   loCmd.CommandText = sSQL
   '
   ' Query Parameters
   '
'   Set loPar = New Parameter
'   If goSession.IsOracle Then
'      loPar.Type = adInteger
'   Else
'      loPar.Type = adBoolean
'   End If
'   loPar.value = oRsEventRecord!sent
'   loCmd.Parameters.Append loPar
'
'   Set loPar = New Parameter
'   loPar.Type = adDate
'   loPar.value = oRsEventRecord.Fields("report_Datetime").value
'   loCmd.Parameters.Append loPar
'
'   Set loPar = New Parameter
'   loPar.Type = adVarChar
'   loPar.Size = 200
'   Set loEaWork = New mwEventFactsWork
'   Set loRs2 = loEaWork.ExpandRecordset(oRsEventRecord, moInterface.EventType, oRsEventRecord!ID)
'   loPar.value = BuildAbstractDescription(loRs2, goSession.EventTypes.Item(moInterface.EventType).AbstractTemplate)
'   loCmd.Parameters.Append loPar
'
'   Set loPar = New Parameter
'   loPar.Type = adVarChar
'   loPar.value = oRsEventRecord!Voyage_NO
'   loPar.Size = 20
'   loCmd.Parameters.Append loPar
'
'   Set loPar = New Parameter
'   If goSession.IsOracle Then
'      loPar.Type = adNumeric
'   Else
'      loPar.Type = adInteger
'   End If
'   loPar.value = moInterface.EventKey
'   loCmd.Parameters.Append loPar
'
'   loCmd.Execute
'   Set loPar = Nothing
'   Set loCmd = Nothing
'   goSession.CloseRecordset loRs2
'   Set loEaWork = Nothing
   'goCon.Execute sSQL
   'goSession.LogIt mwl_Workstation, mwl_Information, "Logging mwVoyEvent.RefreshEventLog: " & sSQL
   RefreshEventLog = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.RefreshEventLog. ", Err.Number, Err.Description
   CloseRecordset moRS
'   Set loPar = Nothing
'   Set loCmd = Nothing
   RefreshEventLog = False
End Function

Public Function BuildVoyEventTable(EventGroup As Long) As Boolean
   Dim sSQL As String
   Dim loType As mwEventType
   Dim i As Integer
   On Error GoTo FunctionError
   '
   ' Purge Table for Event Group
   '
   MsgBox "Feature under construction"
   BuildVoyEventTable = True
   Exit Function
   '
   'Update for Event Groups
   '
   sSQL = "delete from mwEventType where "
   goCon.Execute sSQL
   '
   ' Rebuild table...
   '
   CloseRecordset moRSEvent
   Set moRSEvent = New Recordset
   moRSEvent.CursorLocation = adUseClient
   moRSEvent.Open "mwEventType", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   '
   goSession.LogIt mwl_Workstation, mwl_Information, "Begin Rebuild Voyage Event table, deleting all records."
   '
   ' Loop through event types...
   '
   For i = 1 To goSession.EventTypes.Count
      Set loType = goSession.EventTypes.Item(i)
      If loType.IsEnabled Then
         If Not AddEventTable(loType.TableName, loType.ID, loType.AbstractTemplate) Then
            BuildVoyEventTable = False
            Exit Function
         End If
      End If
   Next i
   ' Close up shop
   moRSEvent.Close
   BuildVoyEventTable = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.BuildVoyEventTable. ", Err.Number, Err.Description
   BuildVoyEventTable = False
   moRSEvent.Close
End Function


Private Function AddEventTable(TableName As String, WhichEvent As Integer, AbstractTemplate As String) As Boolean
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   goSession.LogIt mwl_Workstation, mwl_Information, "Rebuilding Voyage Event table, loading: " & TableName
   loRs.Open TableName, goCon, adOpenForwardOnly, adLockReadOnly, adCmdTable
   Do While Not loRs.EOF
      moRSEvent.AddNew
      moRSEvent!ID = goSession.MakePK("mwEventType")
      moRSEvent!mwEventTypeKey = WhichEvent
      moRSEvent!mwEventDetailKey = loRs!ID
      moRSEvent!Vessel_ID = loRs!Vessel_ID
      moRSEvent!Voyage_NO = loRs!Voyage_NO
      moRSEvent!ReportDateTime = loRs!Report_DateTime
      moRSEvent!IsSubmitted = loRs!sent
      moRSEvent!AbstractDescription = Left(BuildAbstractDescription(loRs, AbstractTemplate), 99)
      moRSEvent.Update
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   AddEventTable = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.AddEventTable. ", Err.Number, Err.Description
   AddEventTable = False
End Function

Public Function BuildAbstractDescription(ByRef loRs As Recordset, AbstractTemplate As String) As String
   Dim strTemp As String
   Dim strDateTimeFormat As String
   Dim loField As field
   On Error GoTo FunctionError
   strTemp = AbstractTemplate
   'strDateTimeFormat = goSession.GetAppKey("VRS", "AbstractDateTimeFormat")
   'If Trim(strDateTimeFormat) = "" Then
      strDateTimeFormat = "General Date"
   'End If
   For Each loField In loRs.Fields
      If Not IsNull(loField.value) Then
         If loField.Type = adDate Then
            strTemp = Replace(strTemp, "[" & UCase(loField.Name) & "]", Format(loField.value, strDateTimeFormat), , , vbTextCompare)
         Else
            strTemp = Replace(strTemp, "[" & UCase(loField.Name) & "]", loField.value)
         End If
      Else
         strTemp = Replace(strTemp, "[" & UCase(loField.Name) & "]", " ")
      End If
   Next loField
   strTemp = Replace(strTemp, "|", vbCrLf)
   BuildAbstractDescription = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.BuildAbstractDescription. ", Err.Number, Err.Description
   BuildAbstractDescription = "*** Error Occurred in Build Abstract Description ***"
End Function

Public Function FetchVoyListRS(Vessel_ID As String, EventGroup As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "select DISTINCT Voyage_NO from mwEventLog where Vessel_ID='" & Vessel_ID & "' order by Voyage_NO"
   
'   sSQL = "SELECT distinct mwEventLog.Voyage_NO " & _
'     " FROM mwEventGroup INNER JOIN (mwEventLog INNER JOIN mwEventGroupType ON mwEventLog.mwEventTypeKey = mwEventGroupType.mwEventTypeKey) ON mwEventGroup.ID = mwEventGroupType.mwEventGroupKey " & _
'     " WHERE mwEventLog.Vessel_ID='" & Vessel_ID & "' and mwEventGroup.ID=" & EventGroup

   sSQL = "SELECT distinct mwEventLog.Voyage_NO " & _
     " FROM mwEventGroup , mwEventLog , mwEventGroupType " & _
     " WHERE mwEventLog.mwEventTypeKey = mwEventGroupType.mwEventTypeKey AND " & _
     " mwEventGroup.ID = mwEventGroupType.mwEventGroupKey AND " & _
     " mwEventLog.Vessel_ID = '" & Vessel_ID & "' AND mwEventGroup.ID=" & EventGroup
   
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchVoyListRS = loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchVoyListRS. ", Err.Number, Err.Description
End Function

Public Function IsCreateableEvents(EventGroup As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventType.IsCreateEventAllowed " & _
'     "FROM mwEventType INNER JOIN (mwEventGroup INNER JOIN " & _
'     " mwEventGroupType ON mwEventGroup.ID = mwEventGroupType.mwEventGroupKey) " & _
'     " ON mwEventType.ID = mwEventGroupType.mwEventTypeKey " & _
'     " WHERE mwEventGroup.ID=" & EventGroup
   
   sSQL = "SELECT mwEventType.IsCreateEventAllowed " & _
     "FROM mwEventType , mwEventGroup , mwEventGroupType " & _
     " WHERE mwEventGroup.ID = mwEventGroupType.mwEventGroupKey AND " & _
     " mwEventType.ID = mwEventGroupType.mwEventTypeKey AND " & _
     " mwEventGroup.ID = " & EventGroup

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   loRs.Filter = "IsCreateEventAllowed <> 0"
   If loRs.RecordCount > 0 Then
      IsCreateableEvents = True
   Else
      IsCreateableEvents = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.IsCreateableEvents. ", Err.Number, Err.Description
   IsCreateableEvents = False
End Function

Public Function AutoCreateEventFromDg(EventType As Long, ByRef loRec As ConfigKeys) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim iDefaultStage As Long
   Dim loEhWork As mwEventHistoryWork
   On Error GoTo FunctionError
   Set loRs = New Recordset
   With loRs
      .CursorLocation = adUseClient
      .Open "mwEventLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
      .AddNew
      .Fields("ID") = goSession.MakePK("mwEventLog")
      .Fields("mwEventTypeKey") = EventType
      .Fields("mwEventDetailKey") = loRec.GetKeyValue("ID")
      If loRec.GetKeyValue("Vessel_ID") <> "" Then
         .Fields("Vessel_ID") = loRec.GetKeyValue("Vessel_ID")
      End If
      If loRec.GetKeyValue("Voyage_NO") <> "" Then
         .Fields("Voyage_NO") = loRec.GetKeyValue("Voyage_NO")
      End If
      If loRec.GetKeyValue("report_Datetime") <> "" Then
         .Fields("ReportDateTime") = loRec.GetKeyValue("report_Datetime")
      Else
         .Fields("ReportDateTime") = Now()
      End If
      .Fields("IsSubmitted") = False
      .Fields("AbstractDescription") = "Open Event to Create Description"
      iDefaultStage = FetchDefaultWfStageKey(EventType)
      If iDefaultStage > 0 Then
         .Fields("mwEventWfStageKey") = iDefaultStage
      End If
      'Left(BuildAbstractDescription(EventDetailRS, goSession.EventTypes.Item(EventType).AbstractTemplate), 99)
      .Update
   End With
   '
   ' Add the other stuff
   '
   AddEventFactsFormsLinksHistory EventType, loRs.Fields("mwEventDetailKey")
   '
   ' System Event History
   '
   Set loEhWork = New mwEventHistoryWork
   loEhWork.AddSystemRecord EventType, loRs!mwEventDetailKey, "EVENT CREATED", "This event has been automatically created by Workflow Agent"
   Set loEhWork = Nothing
   goSession.RaiseWorkflowStatus "Successfully Auto-Created New MW Event: " & EventType & "/" & loRs!ID
   goSession.CloseRecordset loRs
   AutoCreateEventFromDg = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.AutoCreateEventFromDg. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   AutoCreateEventFromDg = False
End Function

Private Sub Class_Initialize()
   Set moRSEvent = New Recordset
   moRSEvent.CursorLocation = adUseClient
End Sub

Private Sub Class_Terminate()
   CloseRecordset moRSEvent
End Sub

Public Function FetchDefaultWfStageKey(EventType As Long, Optional EventGroupKey As Long) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If EventGroupKey > 0 Then
      If goSession.Site.SiteType = SITE_TYPE_SHIP Then
         sSQL = "select DefaultWfStageKey from mwEventGroup where ID=" & EventGroupKey
      Else
         sSQL = "select DefaultWfStageShoreKey from mwEventGroup where ID=" & EventGroupKey
      End If
   Else
      sSQL = "select DefaultWfStageKey from mwEventType where ID=" & EventType
   End If
   
   
   
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      FetchDefaultWfStageKey = -1
   Else
      If IsNull(loRs.Fields(0).value) Then
         FetchDefaultWfStageKey = -1
      Else
         FetchDefaultWfStageKey = loRs.Fields(0).value
      End If
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchDefaultWfStageKey. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   FetchDefaultWfStageKey = -1
End Function

Public Function FetchEventWfShapeRs(EventGroup As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
      
'   sSQL = "SHAPE {SELECT mwEventWfStage.ID as WfStageID, mwEventWfStage.mwEventGroupKey, " & _
'    " mwEventWfStage.WfStageName, mwEventWfStage.IsClosedWfStage, mwEventWfStage.IsLockedWfStage, " & _
'    " mwEventWfStage.IsApprovalOnTransitionRequired, mwEventWfStage.IsApprovePasswordRequired, mwEventWfStage.RequiredNextStage, " & _
'    " mwEventWfStagePermissions.*, mwcRoleType.* " & _
'    " FROM mwEventWfStage , mwcRoleType , mwEventWfStagePermissions " & _
'    " WHERE mwcRoleType.ID = mwEventWfStagePermissions.mwcRoleTypeKey " & _
'      " AND mwEventWfStage.ID = mwEventWfStagePermissions.mwEventWfStageKey " & _
'      " AND mwcRoleType.ID = " & goSession.User.RoleTypeKey & _
'      " AND mwEventWfStage.mwEventGroupKey = " & EventGroup & "order by mwEventWfStage.DisplaySequence" & _
'    " } AS EventGroupStages " & _
'    "APPEND ({SELECT mwEventLog.ID, mwEventLog.mwEventTypeKey, mwEventType.Description, " & _
'     " mwEventLog.ReportDateTime,  mwEventLog.AbstractDescription, mwEventLog.IsSubmitted, " & _
'     " mwEventLog.mwEventDetailKey,  mwEventLog.Vessel_ID, mwEventLog.Voyage_NO, " & _
'     " mwEventLog.mwEventWfStageKey, mwEventGroupType.mwEventGroupKey, mwEventLog.mwcSitesKey " & _
'     " FROM mwEventType , mwEventLog, mwEventGroupType  " & _
'     " WHERE mwEventType.ID = mwEventLog.mwEventTypeKey " & _
'       " AND mwEventType.ID = mwEventGroupType.mwEventTypeKey " & _
'       " AND mwEventGroupType.mwEventGroupKey = " & EventGroup & _
'     "}  AS cmdLog RELATE mwEventGroupKey TO mwEventGroupKey, WfStageID TO mwEventWfStageKey) AS cmdLog "
  
   sSQL = "SHAPE {SELECT mwEventWfStage.ID as WfStageID, mwEventWfStage.mwEventGroupKey, " & _
    " mwEventWfStage.WfStageName, mwEventWfStage.IsClosedWfStage, mwEventWfStage.IsLockedWfStage, " & _
    " mwEventWfStage.IsApprovalOnTransitionRequired, mwEventWfStage.IsApprovePasswordRequired, mwEventWfStage.RequiredNextStage," & _
    " mwEventWfStagePermissions.*, mwcRoleType.* " & _
    " FROM mwEventWfStage , mwcRoleType , mwEventWfStagePermissions " & _
    " WHERE mwcRoleType.ID = mwEventWfStagePermissions.mwcRoleTypeKey " & _
      " AND mwEventWfStage.ID = mwEventWfStagePermissions.mwEventWfStageKey " & _
      " AND mwcRoleType.ID = " & goSession.User.RoleTypeKey & _
      " AND mwEventWfStage.mwEventGroupKey = " & EventGroup & " order by mwEventWfStage.DisplaySequence" & _
    " } AS EventGroupStages " & _
    "APPEND ({SELECT mwEventLog.ID, mwEventLog.mwEventTypeKey, mwEventType.Description, " & _
     " mwEventLog.ReportDateTime,  mwEventLog.AbstractDescription, mwEventLog.IsSubmitted, " & _
     " mwEventLog.mwEventDetailKey,  mwEventLog.Vessel_ID, mwEventLog.Voyage_NO, " & _
     " mwEventLog.mwEventWfStageKey, mwEventGroupType.mwEventGroupKey, mwEventLog.mwcSitesKey " & _
     " FROM mwEventType , mwEventLog, mwEventGroupType  " & _
     " WHERE mwEventType.ID = mwEventLog.mwEventTypeKey " & _
       " AND mwEventType.ID = mwEventGroupType.mwEventTypeKey " & _
       " AND mwEventGroupType.mwEventGroupKey = " & EventGroup & " ORDER BY mwEventLog.mwcSitesKey, mwEventLog.ID DESC " & _
     "}  AS cmdLog RELATE mwEventGroupKey TO mwEventGroupKey, WfStageID TO mwEventWfStageKey) AS cmdLog "
  
  
   loRs.Open sSQL, goSession.DbShapeConnection, adOpenDynamic, adLockOptimistic
   'loRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Set FetchEventWfShapeRs = loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchEventWfShapeRs. ", Err.Number, Err.Description
End Function


Public Function GetVoyageFolder(EventType As Long, Optional EventKey As Long) As String
   Dim fso As FileSystemObject
   Dim strPath As String
   On Error GoTo FunctionError
   '
   ' Create Baseline Voyage folder, if necessary...
   '
   Set fso = New FileSystemObject
   strPath = goSession.GetPlace(goSession.EventTypes(EventType).FormsPlaceID, mw_DriveLetterPath)
   If Not fso.FolderExists(strPath) Then
      fso.CreateFolder (strPath)
   End If
   '
   ' Create Ship based sub-folder if necessary...
   '
   strPath = strPath & "\" & goSession.Site.SiteID
   If Not fso.FolderExists(strPath) Then
      fso.CreateFolder (strPath)
   End If
   '
   ' Create Year-based sub-folder if necessary...
   '
   strPath = strPath & "\" & Format(Now(), "yyyy")
   If Not fso.FolderExists(strPath) Then
      fso.CreateFolder (strPath)
   End If
   '
   ' Create Month-based Sub Folder based on month...
   '
   strPath = strPath & "\" & Format(Now(), "mmmm")
   If Not fso.FolderExists(strPath) Then
      fso.CreateFolder (strPath)
   End If
   '
   ' Create optional EventKey folder underneath...
   '
   If EventKey > 0 Then
      strPath = strPath & "\" & LTrim(str(EventKey))
      If Not fso.FolderExists(strPath) Then
         fso.CreateFolder (strPath)
      End If
   End If
   Set fso = Nothing
   GetVoyageFolder = strPath
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.GetVoyageFolder. ", Err.Number, Err.Description
End Function


Public Function SetAbstractDescription(EventKey As Long, NewDescription As String, _
   Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   sSQL = "select * from mwEventLog where ID=" & EventKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If moRS.RecordCount > 0 Then
      moRS!AbstractDescription = NewDescription
      If mSiteKey > 0 Then
         moRS!mwcSitesKey = mSiteKey
      End If
      
      moRS.Update
      SetAbstractDescription = True
   Else
      SetAbstractDescription = False
   End If
   CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.SetAbstractDescription. ", Err.Number, Err.Description
   SetAbstractDescription = False
End Function

Public Function SetAbstractDescription2(EventTypeKey As Long, EventDetailKey As Long, NewDescription As String, _
   Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   sSQL = "select * from mwEventLog where mwEventTypeKey=" & EventTypeKey & _
    " and mwEventDetailKey=" & EventDetailKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If moRS.RecordCount > 0 Then
      moRS!AbstractDescription = NewDescription
      If mSiteKey > 0 Then
         moRS!mwcSitesKey = mSiteKey
      End If
      
      moRS.Update
      SetAbstractDescription2 = True
   Else
      SetAbstractDescription2 = False
   End If
   CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.SetAbstractDescription2. ", Err.Number, Err.Description
   SetAbstractDescription2 = False
End Function



Public Function MoveWorkflowStage(EventKey As Long, NewWfStageKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   
   sSQL = "select * from mwEventLog where ID=" & EventKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If Not IsNull(moRS!mwcSitesKey) Then
      mSiteKey = moRS!mwcSitesKey
   End If
   If moRS.RecordCount > 0 Then
      moRS!mwEventWfStageKey = NewWfStageKey
      moRS.Update
      MoveWorkflowStage = True
   Else
      goSession.RaiseError "Error in mwSession.mwEventWork.MoveWorkflowStage, Event record is missing: " & EventKey
      MoveWorkflowStage = False
   End If
   
   CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.MoveWorkflowStage. ", Err.Number, Err.Description
   MoveWorkflowStage = False
End Function

Public Function AddEventFactsFormsLinksHistory(EventTypeKey As Long, EventDetailKey As Long, Optional SiteKey As Long) As Boolean
   Dim loErWork As mwEventFormWork
   Dim loEaWork As mwEventFactsWork
   Dim loEsWork As mwEventFactWorkSN
   Dim loElWork As mwEventLinkWork
   Dim loEhWork As mwEventHistoryWork
   On Error GoTo FunctionError
   '
   ' Add Forms
   '
   Set loErWork = New mwEventFormWork
   loErWork.AutoCreate EventTypeKey, EventDetailKey, SiteKey
   '
   Set loErWork = Nothing
   '
   ' Add Event Fact...
   '
   Set loEaWork = New mwEventFactsWork
   loEaWork.AutoCreate EventTypeKey, EventDetailKey
   Set loEaWork = Nothing
   '
   ' Add Event Fact SN...
   '
   Set loEsWork = New mwEventFactWorkSN
   loEsWork.AutoCreate EventTypeKey, EventDetailKey, SiteKey
   Set loEsWork = Nothing
   '
   ' Add Event Links...
   '
   Set loElWork = New mwEventLinkWork
   loElWork.AutoCreate EventTypeKey, EventDetailKey, SiteKey
   Set loElWork = Nothing
   '
   ' Add Event History...
   '
   Set loEhWork = New mwEventHistoryWork
   loEhWork.AutoCreate EventTypeKey, EventDetailKey, SiteKey
   Set loEhWork = Nothing
   AddEventFactsFormsLinksHistory = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.AddEventFactsFormsLinksHistory. ", Err.Number, Err.Description
End Function


Public Function DeleteEvent(EventTypeKey As Long, EventDetailKey As Long) As Boolean
   Dim loErWork As mwEventFormWork
   Dim loEaWork As mwEventFactsWork
   Dim loElWork As mwEventLinkWork
   Dim loEhWork As mwEventHistoryWork
   Dim sSQL As String
   On Error GoTo FunctionError
   '
   ' Delete Forms
   '
   Set loErWork = New mwEventFormWork
   loErWork.DeleteAllForms EventTypeKey, EventDetailKey
   '
   Set loErWork = Nothing
   '
   ' Delete Facts
   '
   Set loEaWork = New mwEventFactsWork
   loEaWork.DeleteAllFacts EventTypeKey, EventDetailKey
   Set loEaWork = Nothing
   '
   ' Delete Links...
   '
   Set loElWork = New mwEventLinkWork
   loElWork.DeleteLinkRecords EventTypeKey, EventDetailKey
   Set loElWork = Nothing
   '
   ' Delete History...
   '
   Set loEhWork = New mwEventHistoryWork
   loEhWork.DeleteAllHistory EventTypeKey, EventDetailKey
   Set loEhWork = Nothing

   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventLog where mwEventTypeKey=" & EventTypeKey & _
     " and mwEventDetailKey=" & EventDetailKey
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If moRS.RecordCount > 0 Then
   
      DeleteAudit moRS!ID
   
      moRS.Delete
   End If
   CloseRecordset moRS
   DeleteEvent = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.DeleteEvent. ", Err.Number, Err.Description
   DeleteEvent = False
   CloseRecordset moRS
End Function


Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwEventLog") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmSmOccurrence.moRS_WillChangeRecord. ", Err.Number, Err.Description
End Sub


Public Function AddEventExplorer(EventType As Long, ByRef EventDetailRS As Recordset, Optional Description As String, _
 Optional EventGroupKey As Long, Optional IsAddFactsFormsLinksHistory As Boolean, _
 Optional SiteKey As Long, Optional AllowChooseStartingWfStage As Boolean) As Long
   Dim loEhWork As mwEventHistoryWork
   Dim iDefaultStage As Long
   On Error GoTo FunctionError
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open "mwEventLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   moRS.AddNew
   moRS!ID = goSession.MakePK("mwEventLog")
   moRS!mwEventTypeKey = EventType
   moRS!mwEventDetailKey = EventDetailRS.Fields(0).value
   moRS!ReportDateTime = GetReportDateTime(EventDetailRS)
   
   ' For Workflow Events, set default stage...
   '  - If EventGroupKey > 0, then its going to mwEventGroup.DefaultWfStage...
   '
   'If EventGroupKey > 0 Then
   If Not AllowChooseStartingWfStage Then
      iDefaultStage = FetchDefaultWfStageKey(EventType, EventGroupKey)
      If iDefaultStage < 1 Then
         goSession.RaiseError "Error in mwSession.mwEventWork.AddEventExplorer, no Starting Workflow Stage Defined. "
      End If
   Else
      iDefaultStage = FetchWfStageKey(EventType, EventGroupKey)
   End If
   If iDefaultStage > 0 Then
      moRS!mwEventWfStageKey = iDefaultStage
   Else
      goSession.LogIt mwl_Workstation, mwl_Information, "AddEventExplorer User cancelled NewEvent."
      AddEventExplorer = -1
      CloseRecordset moRS
      Exit Function
   End If
   'End If
   If Description <> "" Then
      moRS!AbstractDescription = Description
   End If
   If mSiteKey > 0 Then
      moRS!mwcSitesKey = mSiteKey
   End If
   
   moRS.Update
   AddEventExplorer = moRS!ID
   Set loEhWork = New mwEventHistoryWork
   loEhWork.AddSystemRecord EventType, EventDetailRS.Fields(0).value, "CREATED", "This event was created.", , , , mSiteKey
   KillObject loEhWork
   '
   ' Add the other stuff
   '
   If IsAddFactsFormsLinksHistory Then
      AddEventFactsFormsLinksHistory EventType, moRS!mwEventDetailKey, SiteKey
   End If

   '
   'JIRA: VBCOMP-9 Workflow and Work Queue Compatibility Layer
   '
   CreateWorkQueueItem moRS!ID
   
   '
   ' DEV-2120 ljh Create Alert/Email according to wf.IsNotifyOnEntry
   CreateAlertFromNewEvent moRS!ID
   '
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.AddEventExplorer. ", Err.Number, Err.Description
   AddEventExplorer = -1
End Function

Public Function CreateWorkQueueItem(EvLogKey As Long) As Boolean
   Dim loRsEvLog As Recordset
   Dim sSQL As String
   Dim sEventDetailCode As String
   Dim sEventDescription As String
   Dim nEvTypeKey As Long
   Dim nEvWfSiteKey As Long
   Dim nEvDetailKey  As Long
   Dim nEvWfStageKey As Long
   Dim nWorkQueueKey As Long
   Dim nDefaultAssignedUser As Long
   Dim dReportDate   As Date
   On Error GoTo FunctionError
      
   If EvLogKey < 1 Then Exit Function

   '1. Get EventLog Details
   Set loRsEvLog = FetchEventLogRS(EvLogKey)
   If Not IsRecordLoaded(loRsEvLog) Then
      CloseRecordset loRsEvLog
      Exit Function
   End If
   
   nEvTypeKey = loRsEvLog!mwEventTypeKey
   nEvDetailKey = loRsEvLog!mwEventDetailKey
   nEvWfStageKey = loRsEvLog!mwEventWfStageKey
   nEvWfSiteKey = ZeroNull(loRsEvLog!mwcSitesKey)
   dReportDate = loRsEvLog!ReportDateTime
   sEventDescription = loRsEvLog!AbstractDescription
   
   '2. Fetch Event Detail Code
   sEventDetailCode = FetchEventDetailCode(nEvTypeKey, nEvDetailKey)
   
   '3. Fetch WorkQueueKey & Get Default Assigned User
   If nEvWfStageKey > 0 Then
      nWorkQueueKey = FetchmwWorkQueueKey(nEvWfStageKey)
      If nWorkQueueKey > 0 Then
         nDefaultAssignedUser = GetDefaultAssignedToUser(nWorkQueueKey, nEvWfSiteKey)
      End If
   End If
   
   '4. Create Work Queue Item
   Set moRsWorkQ = New Recordset
   moRsWorkQ.CursorLocation = adUseClient
   sSQL = "select * from MwWorkQueueItem where ID=-1"
   moRsWorkQ.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   moRsWorkQ.AddNew
   moRsWorkQ!ID = goSession.MakePK("MwWorkQueueItem")
   moRsWorkQ!mwEventTypeKey = nEvTypeKey
   moRsWorkQ!EventDetailKey = nEvDetailKey
   moRsWorkQ!mwcSitesKey = nEvWfSiteKey
   moRsWorkQ!CreatedByUserKey = goSession.User.UserKey
   moRsWorkQ!CreatedDateTime = dReportDate
   moRsWorkQ!mwEventLogKey = EvLogKey
   
   If nWorkQueueKey > 0 Then
      moRsWorkQ!MwWorkQueueStateKey = nWorkQueueKey
   End If
   
   If sEventDetailCode <> "" Then
      moRsWorkQ!EventDetailCode = sEventDetailCode
   End If
   
   If GetDefaultDueDate(nEvTypeKey, dReportDate) Then
      moRsWorkQ!AssignedToDateDue = dReportDate
   End If
   
   If nDefaultAssignedUser > 0 Then
      moRsWorkQ!AssignedToUserKey = nDefaultAssignedUser
   End If
   
   moRsWorkQ!EventDetailBriefDescription = sEventDescription
   moRsWorkQ.Update
   
   CreateWorkQueueItem = True
   CloseRecordset moRsWorkQ
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwEventWork.CreateWorkQueueItem. ", Err.Number, Err.Description
   CloseRecordset moRsWorkQ
End Function

Private Function FetchEventLogRS(EvLogKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * FROM mwEventLog WHERE ID = " & EvLogKey
      
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchEventLogRS = loRs
   
   'CloseRecordset loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwEventWork.FetchEventLogRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function FetchEventDetailCode(nEvTypeKey As Long, EvDetailKey As Long) As String
   Dim loEvWork As Object
   Dim sSQL As String
   Dim loRs As Recordset
   Dim sTableName As String
   On Error GoTo FunctionError
   
   Set loEvWork = CreateObject("mwEventWorkflow.mwEventsWork")
   loEvWork.InitSession goSession
   sTableName = loEvWork.GetTableName(nEvTypeKey)
   
   Select Case UCase(sTableName)
      Case "SMOCCURRENCE"
         sSQL = "SELECT OccurrenceCode FROM smOccurrence WHERE ID = " & EvDetailKey
      Case Else
         Set loEvWork = Nothing
         Exit Function
   End Select
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If Not IsNull(loRs.Fields(0).value) Then
         FetchEventDetailCode = loRs.Fields(0).value
      End If
   End If
   
   Set loEvWork = Nothing
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwEventWork.FetchEventDetailCode. ", Err.Number, Err.Description
   CloseRecordset loRs
   Set loEvWork = Nothing
End Function

Public Function FetchmwWorkQueueKey(nEvWfStageKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   FetchmwWorkQueueKey = -1
   
   sSQL = "Select mwWorkQueueKey FROM MwWorkQueueState WHERE mwEventWfStage = " & nEvWfStageKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If Not IsNull(loRs.Fields(0).value) Then
         FetchmwWorkQueueKey = loRs.Fields("mwWorkQueueKey").value
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchmwWorkQueueKey. ", Err.Number, Err.Description
End Function

Public Function GetDefaultDueDate(nEvTypeKey As Long, ByRef dEvDate As Date) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   Dim sDateDueDefaultIncrement As String
   Dim sDateDueDefaultInterval As String
   Dim nDateDueDefaultIncrement As Long
   On Error GoTo FunctionError
   
   sSQL = "Select * FROM mwEventType WHERE ID = " & nEvTypeKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      sDateDueDefaultIncrement = Trim$(BlankNull(loRs!DateDueDefaultIncrement))
      If BlankNull(loRs!DateDueDefaultIncrement) <> "" Then
         nDateDueDefaultIncrement = Val(loRs!DateDueDefaultIncrement)
         If sDateDueDefaultIncrement <> "" And nDateDueDefaultIncrement > 0 Then
            If InStr(UCase(sDateDueDefaultIncrement), "D") > 0 Then
               sDateDueDefaultInterval = "d" 'd Day interval
            ElseIf InStr(UCase(sDateDueDefaultIncrement), "W") > 0 Then
               sDateDueDefaultInterval = "ww" 'ww Week interval
            ElseIf InStr(UCase(sDateDueDefaultIncrement), "M") > 0 Then
               sDateDueDefaultInterval = "m" 'm Month interval
            ElseIf InStr(UCase(sDateDueDefaultIncrement), "Y") > 0 Then
               sDateDueDefaultInterval = "yyyy" 'yyyy Year interval
            Else
               sDateDueDefaultInterval = "d" 'Else 'd' Day interval
            End If
         End If
         GetDefaultDueDate = True
         dEvDate = DateAdd(sDateDueDefaultInterval, nDateDueDefaultIncrement, dEvDate)
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.GetDefaultDueDate. ", Err.Number, Err.Description
End Function

Public Function GetDefaultAssignedToUser(workQueueStateKey As Long, SiteKey As Long) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   GetDefaultAssignedToUser = -1
   'Get Default Assigned User
   sSQL = "SELECT DISTINCT mwcFleetPositionLog.MwcUsersKey, mwWorkQueueStateRole.WorkQueueStateRoleActionTypeEnum " & _
       "FROM MwcFleetPositionLog mwcFleetPositionLog " & _
       "LEFT JOIN MwcFleetPositionTRole mwcFleetPositionTRole ON mwcFleetPositionTRole.ID = mwcFleetPositionLog.MwcFleetPositionTRoleKey  " & _
       "LEFT JOIN MwWorkQueueStateRole mwWorkQueueStateRole ON mwWorkQueueStateRole.MwcPositionRoleKey = mwcFleetPositionTRole.MwcPositionRoleKey " & _
       "LEFT JOIN MwWorkQueueState mwWorkQueueState ON mwWorkQueueState.ID = mwWorkQueueStateRole.MwWorkQueueStateKey  " & _
       "LEFT JOIN MwcFleetPositionLogSite mwcFleetPositionRoleSite ON mwcFleetPositionRoleSite.MwcFleetPositionLogKey = mwcFleetPositionLog.ID " & _
       "WHERE mwWorkQueueStateRole.WorkQueueStateRoleActionTypeEnum <> " & WQSRAT_WATCHER & _
       " AND mwWorkQueueStateRole.MwWorkQueueStateKey = " & workQueueStateKey & _
       " AND mwcFleetPositionRoleSite.MwcSitesKey = " & SiteKey
   
   Set loRs = New Recordset
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If IsRecordLoaded(loRs) Then
      loRs.Filter = "WorkQueueStateRoleActionTypeEnum = " & WQSRAT_ACTION_OWNER
      If loRs.RecordCount > 0 Then
         GetDefaultAssignedToUser = loRs.Fields(0).value
      Else
         loRs.Filter = adFilterNone
         loRs.MoveFirst
         loRs.Filter = "WorkQueueStateRoleActionTypeEnum = " & WQSRAT_ACTION_CONTRIBUTOR
         If loRs.RecordCount > 0 Then
            GetDefaultAssignedToUser = loRs.Fields(0).value
         End If
      End If
   End If

CloseRecordset loRs
Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.GetDefaultAssignedToUser. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Sub moRsWorkQ_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("MwWorkQueueItem") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.mwEventWork.moRsWorkQ_WillChangeRecord. ", Err.Number, Err.Description
End Sub

Public Function FetchWfStageKey(EventType As Long, Optional EventGroupKey As Long) As Long
   Dim loform As frmEventWfChooser
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   ' get EvType.DefaultWfStageKey or lookup Workflow Stage by chooser
   If EventGroupKey < 1 Then
      sSQL = "select DefaultWfStageKey from mwEventType where ID=" & EventType
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If loRs.RecordCount < 1 Then
         FetchWfStageKey = -1
      Else
         If IsNull(loRs.Fields(0).value) Then
            FetchWfStageKey = -1
         Else
            FetchWfStageKey = loRs.Fields(0).value
         End If
      End If
      CloseRecordset loRs
      Exit Function
   End If
   
   Set loform = New frmEventWfChooser
   loform.InitForm EventGroupKey, IS_UPDATEABLE_STAGES_ONLY
   loform.Show vbModal
   If loform.IsCancelled Then
      FetchWfStageKey = -1
      KillObject loform
      Exit Function
   Else
      FetchWfStageKey = loform.GetEventWfStageKey
   End If
'
   Unload loform
   KillObject loform
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwEventWork.FetchWfStageKey. ", Err.Number, Err.Description
   FetchWfStageKey = -1
   KillObject loform
End Function

Private Function GetReportDateTime(ByRef EventDetailRS As Recordset) As Date
   On Error GoTo FunctionError

   If Not IsNull(EventDetailRS!ReportDate) Then
      GetReportDateTime = EventDetailRS!ReportDate
   Else
      GetReportDateTime = Now()
   End If
   Exit Function

FunctionError:
   GetReportDateTime = Now()
End Function

Public Function DeleteEventRALink(RAHeaderKey As Long, EventTypeKey As Long) As Long
'DEV-1743 RA Checklist, Participants, WP Linking RA
'By N.Angelakis On 22th Aug 2009

   Dim sSQL As String
   Dim loEhWork As mwEventHistoryWork

   On Error GoTo FunctionError

   sSQL = "SELECT * From mwEventLog WHERE mwEventDetailKey = " & RAHeaderKey & " AND mwEventTypeKey= " & EventTypeKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Do While moRS.EOF = False
      moRS.Delete
      moRS.MoveNext
   Loop
   CloseRecordset moRS
   
   '
   ' Add Event History...
   '
   Set loEhWork = New mwEventHistoryWork
   'Modified By N.Angelakis On 9th September 2009
   'loEhWork.AddSystemRecord EventTypeKey, RAHeaderKey, "DELETED", "This event was deleted.", , , , goSession.Site.SiteID
   loEhWork.AddSystemRecord EventTypeKey, RAHeaderKey, "DELETED", "This event was deleted.", , , , goSession.Site.SiteKey
   Set loEhWork = Nothing
   
Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwEventWork.DeleteEventExplorer: ", Err.Number, Err.Description
   CloseRecordset moRS
End Function

Private Sub moRsAudit_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwEventLogAudit") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.mwEventWork.moRsAudit_WillChangeRecord. ", Err.Number, Err.Description
End Sub

Private Function DeleteAudit(EventLogKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   
   ' delete child records (called from parent)
   
   CloseRecordset moRsAudit
   sSQL = "SELECT ID, mwcSitesKey FROM mwEventLogAudit WHERE mwEventLogKey = " & EventLogKey
   Set moRsAudit = New Recordset
   moRsAudit.CursorLocation = adUseClient
   moRsAudit.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(moRsAudit) Then
      Do While moRsAudit.EOF = False
         moRsAudit.Delete
         moRsAudit.MoveNext
      Loop
   End If
   
   CloseRecordset moRsAudit
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwEventWork.DeleteAudit. ", Err.Number, Err.Description
   CloseRecordset moRsAudit
End Function


Public Function CreateAlertFromNewEvent(EvLogKey As Long) As Boolean
   Dim strBriefDescription As String
   Dim sAlertBody As String
   Dim strComments As String
   Dim nEvLogKey As Long
   Dim nEvDetailKey As Long
   Dim nEvTypeKey As Long
   Dim nEvWfStageKey As Long
   Dim loRsET As Recordset
   Dim loRsEvLog As Recordset
   Dim loRsEvWfStage As Recordset
   Dim loRsEvWfStagePerm As Recordset
   Dim loRs As Recordset
   Dim sSQL As String
   Dim sTableName As String
   On Error GoTo FunctionError          ' On New, check IsNotifOnEventEntry & automatically send alert
   
   If EvLogKey < 1 Then Exit Function

   
   '1. get mwEventLog record (could be internal/external call)
   Set loRsEvLog = goCon.Execute("Select ID, mwEventTypeKey, mwEventDetailKey, AbstractDescription, mwEventWfStageKey, mwcSitesKey " _
      & " From mwEventLog WHERE ID = " & EvLogKey)
   If Not IsRecordLoaded(loRsEvLog) Then
      CloseRecordset loRsEvLog                     ' could not find EventLog ref
      Exit Function
   ElseIf ZeroNull(loRsEvLog!mwEventWfStageKey) < 1 Then
      CloseRecordset loRsEvLog                     ' mwEventWfStageKey missing ref
      Exit Function
   End If
   nEvTypeKey = loRsEvLog!mwEventTypeKey
   nEvDetailKey = loRsEvLog!mwEventDetailKey
   nEvWfStageKey = loRsEvLog!mwEventWfStageKey
   mRecordSiteKey = ZeroNull(loRsEvLog!mwcSitesKey)   ' send Alert option
   
   '2. EventType is qualified for NewEvent
   Set loRsET = goCon.Execute("SELECT TableName FROM mwEventType WHERE ID = " & nEvTypeKey)
   If Not IsRecordLoaded(loRsET) Then
      CloseRecordset loRsEvLog
      CloseRecordset loRsET
      Exit Function
   ElseIf UCase(BlankNull(loRsET!TableName)) <> "SMOCCURRENCE" Then  ' Allowance for Safety at this time
      CloseRecordset loRsEvLog
      CloseRecordset loRsET
      Exit Function
   End If
   CloseRecordset loRsET      ' done with this...close it
   
   
   '3. get mwEventWfStage record
   Set loRsEvWfStage = goCon.Execute("Select * FROM mwEventWfStage WHERE ID = " & nEvWfStageKey)
   If Not IsRecordLoaded(loRsEvWfStage) Then
      CloseRecordset loRsEvWfStage                 ' no stages...exit
      CloseRecordset loRsEvLog
      Exit Function
   End If
   
   '4. getmwEventWfStagePermissions
   Set loRsEvWfStagePerm = goCon.Execute("SELECT * FROM mwEventWfStagePermissions " _
      & " WHERE mwEventWfStageKey = " & nEvWfStageKey & " And NOT (IsAlertOnCreation Is Null Or IsAlertOnCreation = 0) ")
   If Not IsRecordLoaded(loRsEvWfStagePerm) Then
      CloseRecordset loRsEvWfStage                 ' IsNotifyOnEventEntry=false...no alerts to send
      CloseRecordset loRsEvLog
      CloseRecordset loRsEvWfStagePerm
      
      Exit Function
   End If
   
   
   '
   '5. Initial workflow stage assignment check
   '
   ' Send notification email ?
   '
   strComments = "Created Event initiated by: " & goSession.User.UserID & " Automatic Alert Notification" & vbCrLf
   strBriefDescription = "New Event Stage: " & loRsEvWfStage!WfStageName & "  User Initiated: " & goSession.User.UserID
   sAlertBody = "This alert has been sent to you because the " & BlankNull(loRsEvLog!AbstractDescription) & " Workflow Event: " & vbCrLf & vbCrLf & _
     "  " & BlankNull(loRsEvLog!AbstractDescription) & _
     vbCrLf & vbCrLf & "has been assigned to you." & _
     vbCrLf & "-------------------------------------------------------------" & vbCrLf & strComments
   
   'SendWfNotifications "Creation Event Alert Notification", sAlertBody, nEvWfStageKey, nEvTypeKey, nEvDetailKey
   ' DEV-2120 qualified IsAlertOnCreate & RoleTypeEmailAddress<>'' & mwcSiteType=Shore & IsSendAlertsToEmail
   sSQL = "SELECT mwEventWfStagePermissions.*, RoleTypeEmailAddress, mwcSiteType, IsSendAlertsToEmail " _
      & " From mwcRoleType, mwEventWfStagePermissions WHERE mwcRoleType.ID = mwEventWfStagePermissions.mwcRoleTypekey " _
      & " And mwEventWfStageKey = " & nEvWfStageKey _
      & " And NOT (IsAlertOnCreation Is Null Or IsAlertOnCreation = 0) " _
      & " And NOT (IsSendAlertsToEmail Is Null Or IsSendAlertsToEmail = 0) " _
      & " And RoleTypeEmailAddress is not null And mwcSiteType = " & SITE_TYPE_SHORE
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF               ' RT Permission
      If UCase(goSession.ThisSite.GetExtendedProperty("IsAlertSubSystemActive")) = "TRUE" _
       Or UCase(goSession.ThisSite.GetExtendedProperty("IsAlertSubSystemActive")) = "1" _
       Or goSession.ThisSite.GetExtendedProperty("IsAlertSubSystemActive") = "-1" Then
          '
          '6a. Send alert  / RS_0_P_mwcRoleTypeKey
          '
         If Not IsNull(loRs!mwcSiteType) Then
            SendRoleTypeAlertEvent loRs!mwcRoleTypekey, loRs!mwcSiteType, sAlertBody, nEvTypeKey, nEvDetailKey
         Else
            goSession.RaisePublicError "Error in mwSession.mwEventWork.CreateAlertFromNewEvent, Site Type is null ! No notification sent."
         End If
      Else
          '
          '6b. Attempt Send email
          '
          goSession.SendNotification strBriefDescription, sAlertBody, loRs.Fields("RoleTypeEmailAddress")
      End If
      loRs.MoveNext
   Loop
      
         
   CreateAlertFromNewEvent = True
   CloseRecordset loRs
   CloseRecordset loRsEvLog
   CloseRecordset loRsEvWfStage
   CloseRecordset loRsEvWfStagePerm
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwEventWorkCreateAlertFromNewEvent. ", Err.Number, Err.Description
   CloseRecordset loRsEvLog
   CloseRecordset loRsEvWfStage
   CloseRecordset loRsEvWfStagePerm
End Function

Private Function SendRoleTypeAlertEvent(RoleTypeKey As Long, RtSiteType As Long, AlertMsg As String, mwEventType As Long, mwEventDetailKey As Long) As Boolean
   Dim loWork As mwAlertWork
   Dim IsAlertCreated As Boolean
   Dim nTargetSiteKey As Long
   Dim UsersKeySource As Long
   Dim UsersKeyTarget As Long
   On Error GoTo FunctionError
   
   Set loWork = New mwAlertWork
   
   UsersKeySource = 0
   UsersKeyTarget = 0
   
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      If RtSiteType = SITE_TYPE_SHORE Then          ' send to targetsiteid address
         nTargetSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
      Else                                          ' send to gosession.site.sitekey
         nTargetSiteKey = goSession.Site.SiteKey
      End If
   Else
      If RtSiteType = SITE_TYPE_SHORE Then          ' send to goesssion.site.sitekey
         nTargetSiteKey = goSession.Site.SiteKey
      Else                                          ' send to  mRecordSiteKey
         nTargetSiteKey = mRecordSiteKey
      End If
   End If
   
   IsAlertCreated = loWork.CreateApplicationEventAlert(nTargetSiteKey, RoleTypeKey, "Automatic Workflow Event Alert", AlertMsg, UsersKeySource, UsersKeyTarget, mwEventType, mwEventDetailKey)
   
   SendRoleTypeAlertEvent = True
   KillObject loWork
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwEventWork.SendRoleTypeAlertEvent. ", Err.Number, Err.Description
   KillObject loWork
End Function



