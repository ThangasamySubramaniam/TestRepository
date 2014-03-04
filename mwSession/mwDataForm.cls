VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwDataForm - Class supports manipulation of incoming "data forms"
' built from Maritime Systems Custom Form.
' Started: 3/6/2001 ms
' Overhauled: 8/2001 ms
' 5/27/10 ljh vb custom formats to .textbody send email
'
Option Explicit

Const IS_MODIFY = True
Const MAX_OLD_VALUE_TEST = 31

Private moParent As mwForm
Private moRS As Recordset
Private moRecords As ConfigGroups
Private mIsWritten As Boolean
Private moCon As Connection
Private mMtmlReqBody As String
Private mTextBody As String
Private mCurrentParentRecordID As Long
Private mEventType As Long
Private mDgGeneralDateFormat As String
Private mDgShortDateFormat As String
Private mDgDecimalSeparator As String
Private mIsEventDatagram As Boolean
Private mIsBatchDatagram As Boolean
Private mCurrentRecord As Long
Private mTotalErrors As Long
Private mMaxErrorsBeforeQuit As Long
Private mIsShowAllDatagramErrors As Boolean
Private msRecordID As String
Private mSatelliteDBName As String

' 1/2003 - Event Datagram support
Private mEventTypeKey As Long
Private mDatagramCode As String
Private moRsEventColumnMap As Recordset
Private mLocalDecimalSeparator As String
Private mEmailAddress As String

Private moNGradeRS As Recordset

Private mStrBerthKeys As String
Private mStrCargoKeys As String
Private mPortCallKey As Long

Dim moAuditTrail As snAuditTrail

'nGradeBunkers enum
   Const DT_ATSEA = 1
   Const DT_PORTCALL_ARRIVAL = 2
   Const DT_PORTCALL_DEPART = 3
   Const DT_PORTPASSAGE = 4
   
      
'nGradeBunkers Parameters   ---PBT-203
Const BUNKER_PARAM_FMT = " FMT"
Const BUNKER_PARAM_NUMBER = " NUMBER_FMT"
Const BUNKER_PARAM_GRADE = " HIDE_EMPTY"
Const BUNKER_PARAM_SEPARATOR = " SEPARATOR"
Const BUNKER_PARAM_HEADER = " HIDE_HEADER"

Const BUNKER_PURPOSE_ROB = 1
Const BUNKER_PURPOSE_TAKEN = 6

Public Property Let sBerthKeys(ByVal sData As String)
    mStrBerthKeys = Trim(sData)
End Property

Public Property Let sCargoKeys(ByVal sData As String)
    mStrCargoKeys = Trim(sData)
End Property

Public Property Let sPortCallKey(ByVal sData As String)
    mPortCallKey = CLng(sData)
End Property

Public Property Set Parent(oValue As mwForm)
   'If moParent Is Nothing Then
      Set moParent = oValue
   'End If
   
End Property

Public Property Get Parent() As mwForm
   Set Parent = moParent
End Property


Public Property Set Records(ByVal Data As ConfigGroups)
    Set moRecords = Data
End Property

Public Property Get Records() As ConfigGroups
    Set Records = moRecords
End Property

Public Property Get IsWritten() As Boolean
    IsWritten = mIsWritten
End Property

Public Property Let MtmlReqBody(ByVal Data As String)
    mMtmlReqBody = Data
End Property

Public Property Get MtmlReqBody() As String
    MtmlReqBody = mMtmlReqBody
End Property

Public Property Let TextBody(ByVal Data As String)
    mTextBody = Data
End Property

Public Property Get TextBody() As String
    TextBody = mTextBody
End Property

Public Property Let EventType(ByVal Data As Long)
    mEventType = Data
End Property

Public Property Get EventType() As Long
    EventType = mEventType
End Property

Public Property Let DgGeneralDateFormat(ByVal Data As String)
    mDgGeneralDateFormat = Data
End Property

Public Property Get DgGeneralDateFormat() As String
    DgGeneralDateFormat = mDgGeneralDateFormat
End Property

Public Property Let DgShortDateFormat(ByVal Data As String)
    mDgShortDateFormat = Data
End Property

Public Property Get DgShortDateFormat() As String
    DgShortDateFormat = mDgShortDateFormat
End Property

Public Property Let DgDecimalSeparator(ByVal Data As String)
    mDgDecimalSeparator = Data
End Property

Public Property Get DgDecimalSeparator() As String
    DgDecimalSeparator = mDgDecimalSeparator
End Property

Public Property Let IsEventDatagram(ByVal Data As Boolean)
    mIsEventDatagram = Data
End Property

Public Property Get IsEventDatagram() As Boolean
    IsEventDatagram = mIsEventDatagram
End Property

Public Property Let EventTypeKey(ByVal Data As Long)
    mEventTypeKey = Data
End Property

Public Property Get EventTypeKey() As Long
    EventTypeKey = mEventTypeKey
End Property

Public Property Let DatagramCode(ByVal Data As String)
    mDatagramCode = Data
End Property

Public Property Get DatagramCode() As String
    DatagramCode = mDatagramCode
End Property

Public Property Let IsBatchDatagram(ByVal Data As Boolean)
    mIsBatchDatagram = Data
End Property

Public Property Get IsBatchDatagram() As Boolean
    IsBatchDatagram = mIsBatchDatagram
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'
'

Public Function OpenForm(Optional FullFilename As Variant) As Boolean
   Dim sTempVal As String
   
   On Error GoTo FunctionError
   '
   If moParent.IsOpen Then
      goSession.RaiseError "Error in msSession.mwDataForm.OpenForm, Form is already Open."
      OpenForm = False
      Exit Function
   End If
   If IsMissing(FullFilename) Then
      FullFilename = moParent.FormProperties.FullFilename
   End If
   '
   If Not moParent.IsFileExists(FullFilename) Then
      goSession.RaiseError "Error in mwSession.mwDataForm.OpenForm, File not found: " & FullFilename
      OpenForm = False
      Exit Function
   End If
   '
   ' Load File into ConfigGroups...
   '
   Set moRecords = goSession.LoadConfigGroupsFile(FullFilename)
   If moRecords Is Nothing Or moRecords.Count < 1 Then
      OpenForm = False
      Exit Function
   End If
   
   mMaxErrorsBeforeQuit = goSession.Workflow.MaxErrorsBeforeQuit
   mIsShowAllDatagramErrors = goSession.Workflow.IsShowAllDatagramErrors
   '
   ' Validate Confirm right stuff is in file
   '
   With moRecords("header").ConfigKeys
   ' Gotta have TemplateID
   
      If Not moParent.OpenTemplate(.Item("TemplateID").KeyValue) Then
         goSession.KillObject moRecords
         OpenForm = False
         Exit Function
      End If
      moParent.FormProperties.FormID = .GetKeyValue("FormID")
      moParent.FormProperties.RoleType = .GetKeyValue("RoleTypeKey")
      moParent.FormProperties.CreatedDate = .GetKeyValue("CreatedDate")
      moParent.FormProperties.subject = .GetKeyValue("Subject")
      If IsNumeric(.GetKeyValue("TargetSiteKey")) Then
         moParent.FormProperties.TargetSiteKey = .GetKeyValue("TargetSiteKey")
         moParent.FormProperties.TargetSiteID = goSession.Site.GetSiteID(.GetKeyValue("TargetSiteKey"))
      Else
         moParent.FormProperties.TargetSiteID = .GetKeyValue("TargetSiteID")
      End If
      '
      ' 11/2005 - If SiteKey in Header, use that for all validation...
      '
      If IsNumeric(.GetKeyValue("SiteKey")) Then
         moParent.FormProperties.SiteKey = .GetKeyValue("SiteKey")
         moParent.FormProperties.SiteID = goSession.Site.GetSiteID(.GetKeyValue("SiteKey"))
      Else
         moParent.FormProperties.SiteKey = 0
         moParent.FormProperties.SiteID = .GetKeyValue("SiteID")
      End If
      moParent.WorkflowProperties.WFStage = .GetKeyValue("WFStage")
      moParent.FormProperties.IsFleetTransmission = .GetBoolKeyValue("IsFleetTransmission")
      
'      If .GetKeyValue("IsFleetTransmission") = "True" Then
'         moParent.FormProperties.IsFleetTransmission = True
'      Else
'         moParent.FormProperties.IsFleetTransmission = False
'      End If
      '
      ' 1/2003 ms - Event Datagrams have new properties in header...
      '
      If IsNumeric(.GetKeyValue("EventTypeKey")) Then
         mIsEventDatagram = True
         mEventTypeKey = .GetKeyValue("EventTypeKey")
         mDatagramCode = .GetKeyValue("DatagramCode")
      Else
         mIsEventDatagram = False
      End If
      '
      ' 10/2003 ms - Batch Datagrams have new properties in header...
      '
      If IsNumeric(.GetKeyValue("BATCHTYPE")) Then
         IsBatchDatagram = True
      Else
         IsBatchDatagram = False
      End If
      
      If IsNumeric(.GetKeyValue("MAXERRORSBEFOREQUIT")) Then
         mMaxErrorsBeforeQuit = .GetKeyValue("MAXERRORSBEFOREQUIT")
      End If
      
      sTempVal = .GetKeyValue("SHOWALLDATAGRAMERRORS")
      If Len(Trim(sTempVal)) <= 0 Then
         sTempVal = .GetKeyValue("ISSHOWALLDATAGRAMERRORS")
      End If
      
      If Len(Trim(sTempVal)) > 0 Then
         sTempVal = UCase(sTempVal)
         If sTempVal = "TRUE" Then
            mIsShowAllDatagramErrors = True
         End If
         If sTempVal = "1" Then
            mIsShowAllDatagramErrors = True
         End If
         If sTempVal = "FALSE" Then
            mIsShowAllDatagramErrors = False
         End If
         If sTempVal = "0" Then
            mIsShowAllDatagramErrors = False
         End If
         
      End If
      
   End With
   '
   ' Check required Properties
   '
   'If Trim(moParent.FormProperties.FormID) = "" Then
   '   goSession.RaiseError "Error in mwSession.mwDataForm.OpenForm, FormID missing from FataForm."
   '   KillObject moRecords
   '   OpenForm = False
   '   Exit Function
   'End If
   OpenForm = True
   moParent.IsOpen = True
   ' Existing Form, so it doth persist...
   mIsWritten = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General error in mwSession.mwDataForm.OpenForm:" & FullFilename, Err.Number, Err.Description
   OpenForm = False
   moParent.IsOpen = False
   goSession.KillObject moRecords
End Function

Friend Function MakeDataForm() As Boolean
   On Error GoTo FunctionError
   '
   ' Setup Header Information
   '
   moRecords.Add "HEADER", "HEADER"
   With moRecords("HEADER")
      .ConfigKeys.Add "FormID", moParent.FormProperties.FormID, "FormID"
      .ConfigKeys.Add "UserID", goSession.User.UserID, "UserID"
      '.ConfigKeys.Add "EmployeeID", goSession.User.EmployeeID, "EmployeeID"
      .ConfigKeys.Add "RoleType", goSession.User.RoleTypeKey
      .ConfigKeys.Add "SiteID", moParent.FormProperties.SiteID, "SiteID"
      .ConfigKeys.Add "SiteKey", moParent.FormProperties.SiteKey, "SiteKey"
      '.ConfigKeys.Add "CreatedDate", moParent.FormProperties.CreatedDate, "CreatedDate"
      .ConfigKeys.Add "CreatedDate", Format(goSession.API.GetGmtNow, DATAGRAM_DATE_FORMAT) & " GMT"
      '.ConfigKeys.Add "Modified", goSession.API.GetGmtNow()
      ' In a DataForm, startup
      moParent.FormProperties.subject = moParent.TemplateProperties.Description
      .ConfigKeys.Add "Subject", moParent.FormProperties.subject, "Subject"
      .ConfigKeys.Add "TemplateID", moParent.TemplateProperties.TemplateID, "TemplateID"
      .ConfigKeys.Add "WFStage", moParent.WorkflowProperties.WFStage, "WFStage"
      If moParent.FormProperties.TargetSiteID <> "" Then
         .ConfigKeys.Add "TargetSiteID", moParent.FormProperties.TargetSiteID, "TargetSiteID"
         ' Compatibility...write the site key just in case...
         .ConfigKeys.Add "TargetSiteKey", goSession.Site.GetSiteKey(moParent.FormProperties.TargetSiteID), "TargetSiteKey"
      End If
      If moParent.FormProperties.TargetSiteKey > 0 Then
         .ConfigKeys.Add "TargetSiteKey", moParent.FormProperties.TargetSiteKey, "TargetSiteKey"
      End If
   End With
   MakeDataForm = True
   mIsWritten = False
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwDataForm.MakeDataForm.", Err.Number, Err.Description
   MakeDataForm = False

End Function

Public Function AddRecord() As ConfigKeys
   On Error GoTo FunctionError
   Dim strRecord As String
   strRecord = "RECORD_" & Format(str(moRecords.Count), "000")
   moRecords.Add strRecord, strRecord
   Set AddRecord = moRecords(strRecord).ConfigKeys
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwDataForm.AddRecord.", Err.Number, Err.Description
   Set AddRecord = Nothing
End Function

Public Function WriteDataForm(Optional TargetFolder As Variant, Optional ArchiveCopy As Boolean) As Boolean
   Dim strTarget As String
   Dim strArchive As String
   Dim strXmlFile As String
   Dim strType As String
   Dim fso As FileSystemObject
   Dim ts As TextStream
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   '
   ' Possible DLL Execution:
   If Trim(moParent.WorkflowProperties.VbFormEventDLL) <> "" Then
      ExecuteDllOnWriteDatagram
   End If
   '
   If IsMissing(TargetFolder) Then
      TargetFolder = goSession.Workflow.WorkingSessionFolder
   End If
   '
   ' 30 March 2002 - Eliminate use of TargetFolder - using Working Session Folder...
   '
   '
   ' Check if target folder is there...
   If Not fso.FolderExists(TargetFolder) Then
      goSession.RaiseError "Error in mwSession.mwDataForm.WriteDataForm, Folder Doesn't exist: " & TargetFolder
      WriteDataForm = False
      goSession.KillObject fso
      Exit Function
   End If
   ' Assign FullFilename property
   
   moParent.FormProperties.FullFilename = TargetFolder & _
     "\" & moParent.FormProperties.FileName
   strType = ""
   If ZeroNull(moParent.GetExtendedProperty("ConvertToDatagram")) > 0 Then
      strType = moParent.GetExtendedProperty("ConvertToDataGramType")
      If strType = "0" Then strType = ""
   End If
   If strType = "" Or strType = "1" Then
      If Not goSession.WriteConfigGroupsFile(moRecords, moParent.FormProperties.FullFilename) Then
         WriteDataForm = False
         goSession.KillObject fso
         Exit Function
      End If
   Else
      '
      ' Write Mtml String....
      '
      ' Change filename extension
      moParent.FormProperties.FullFilename = Left(moParent.FormProperties.FullFilename, Len(moParent.FormProperties.FullFilename) - 4) & ".xml"
      
      
      Set ts = fso.OpenTextFile(moParent.FormProperties.FullFilename, ForWriting, True)
      ts.Write mMtmlReqBody
      ts.Close
      Set ts = Nothing
   End If
   '
   ' Optional Archive File after write
   '
   If ArchiveCopy Then
      goSession.Workflow.ArchiveWorkflowItem moParent.FormProperties.FullFilename, ArchivedOK, False
   End If
   goSession.KillObject fso
   WriteDataForm = True
   mIsWritten = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwDataForm.WriteDataForm.", Err.Number, Err.Description
   WriteDataForm = False
End Function


Private Sub Class_Initialize()
   On Error GoTo SubError
   Set moRecords = New ConfigGroups
   Set moAuditTrail = New snAuditTrail
   mIsWritten = False
   ' Default regional settings for target...
   mDgGeneralDateFormat = "General Date"
   mDgShortDateFormat = "Short Date"
   mDgDecimalSeparator = "."
   mLocalDecimalSeparator = goSession.API.GetLocaleString(LOCALE_DECIMAL_SEPARATOR)
   Exit Sub
SubError:
   goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataform.Class_Initialize. ", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
   On Error GoTo SubError
   If Not moRS Is Nothing Then
      If moRS.State = adStateOpen Then
         If moRS.LockType <> adLockReadOnly Then
            If Not (moRS.EOF Or moRS.BOF) Then moRS.Update
         End If
         If Trim(moParent.WorkflowProperties.ConnectionID) <> "" Then
            moRS.Close
         End If
      End If
      Set moRS = Nothing
      CloseRecordset moRsEventColumnMap
   End If
   
      
   If Not moCon Is Nothing Then
      If moCon.State = adStateOpen Then
         ' Don't close if form is using the main connection...
         If Trim(moParent.WorkflowProperties.ConnectionID) <> "" Then
            moCon.Close
         End If
      End If
      Set moCon = Nothing
   End If
   If Not moRecords Is Nothing Then Set moRecords = Nothing
   If Not moParent Is Nothing Then Set moParent = Nothing
   KillObject moAuditTrail
SubError:

End Sub

Public Function UpdateAsRecordset() As Boolean
   Dim strOperation As String
   Dim sGet_SqlServerDATEFORMAT As String
   Dim loDataWork As mwDataWork
   Dim loRec As ConfigKeys
   Dim i As Long
   'Dim i As Integer
   Dim IsMappedEvent As Boolean
   Dim loEventWork As mwEventWork
   Dim loBatchWork As mwReplicateBatchWork
   Dim nRecStart As Long
   Dim nBatchType As Long
   Dim nBatchNo As Long
   Dim IsPrimaryRecord As Boolean
   Dim IsCancelUpdate As Boolean
   Dim bIsAPK As Boolean
   Dim bIsAMPK As Boolean
   Dim strTableName As String
   Dim IsRedistributableTable As Boolean
   Dim mwrChangeLogKey As Long
   Dim OriginatingSiteID As String
   Dim OriginatingSiteKey As Long
   Dim SchemaName As String
   Dim SchemaContents As String
   Dim AuditTrailHeaderKey As Long
   Dim sChangeDate As String
   Dim nUserKey As Long
   On Error GoTo FunctionError
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaiseError "Error in mwSession.mwDataForm.UpdateAsRecordset, DataForm Not Open."
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' and it must be a dataform...
   '
   If moParent.FormType <> mwft_DATA_FORM Then
      goSession.RaiseError "Error in mwSession.mwDataForm.UpdateAsRecordset, Wrong Form Type."
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' Get Connection - either from FormTemplate, or use Global Connection...
   '
   Set moCon = moParent.GetDataConnection
   If moCon Is Nothing Then
      UpdateAsRecordset = False
      Exit Function
   End If
      
   '
   ' For each record (except first moRecord, which is the Header)
   ' 1-based collection
   '
   '
   ' If BatchType: 1.Validate Sequence, 2.Validate Starting Record, let it rip....
   '
   If mIsBatchDatagram Then
      Set loBatchWork = New mwReplicateBatchWork
      '
      ' Is valid Batch ?
      '
       If Not loBatchWork.ValidateIncomingDgmBatch(moRecords(1).ConfigKeys, moRecords(moRecords.Count).ConfigKeys, moRecords.Count) Then
         UpdateAsRecordset = False
         Exit Function
      End If
      nBatchType = CLng(moRecords(1).ConfigKeys.GetKeyValue("BatchType"))
      nBatchNo = CLng(moRecords(1).ConfigKeys.GetKeyValue("BatchNumber"))
      OriginatingSiteID = moRecords(1).ConfigKeys.GetKeyValue("SiteID")
      OriginatingSiteKey = goSession.Site.GetSiteKey(OriginatingSiteID)
   End If
   
   mSatelliteDBName = ""
   mSatelliteDBName = moRecords(1).ConfigKeys.GetKeyValue("SatelliteDBName")
   nRecStart = 2
   
   '
   ' Begin Transaction...
   '
   goSession.RaiseWorkflowStatus "Processing Data Form..........: " & moParent.FormProperties.FormID
   goSession.RaiseWorkflowStatus "Begin Transaction..."
   moCon.BeginTrans
   
   For i = nRecStart To moRecords.Count
'      RcdTimer = StartTimer
      DoEvents
      If i Mod 100 = 0 Then

         goSession.RaiseWorkflowStatus "... " & i & " records processed."
      End If
      
      mCurrentRecord = i
      IsCancelUpdate = False
      Set loRec = moRecords(i).ConfigKeys
      msRecordID = moRecords(i).ConfigGroup
      '
      ' Validate: Need an Operation Code
      '
      strOperation = UCase(Trim(loRec.GetKeyValue("Operation")))
      strTableName = Trim(loRec.GetKeyValue("table"))
      
      If strOperation = "" Then
         ' Must be "Operation" Code.
         If Not LogRsError(msRecordID, 0, "Operation Code missing from Record.") Then
            UpdateAsRecordset = False
            RollbackTransactions
            Exit Function
         End If
      ElseIf strTableName = "" And InStr(1, strOperation, "AMD") Then
                     '
         If Not LogRsError(msRecordID, 0, "Table name missing from Record.") Then
            IsCancelUpdate = True
            UpdateAsRecordset = False
            RollbackTransactions
            Exit Function
         End If
      
      Else
         AuditTrailHeaderKey = 0
         If moAuditTrail.CheckIsAuditTrailEnabledTable(strTableName) Then
            If Len(loRec.GetKeyValue("ID")) < 1 Then
               goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateAsRecordset, no ID value found in Record: " & msRecordID
               strOperation = "R"
            Else
               If Trim(loRec.GetKeyValue("CHANGEDATE")) <> "" And Trim(loRec.GetKeyValue("USERKEY")) <> "" Then
                  sChangeDate = Format(Trim(loRec.GetKeyValue("CHANGEDATE")), DATAGRAM_DATE_FORMAT)
                  nUserKey = loRec.GetKeyValue("USERKEY")
                  AuditTrailHeaderKey = moAuditTrail.WriteAuditTrailHeader(strOperation, ChangeTableKey(strTableName), loRec.GetKeyValue("ID"), OriginatingSiteKey, sChangeDate, nUserKey)
               End If
            End If
         End If
         mwrChangeLogKey = 0
         Select Case strOperation
            Case Is = "A", "APK", "D", "M"
               If CheckIsRedistributableTable(strTableName) Then
                  If Len(loRec.GetKeyValue("ID")) < 1 Then
                     goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateAsRecordset, no ID value found in Record: " & msRecordID
                     strOperation = "R"
                  Else
                     mwrChangeLogKey = WriteChangeHeader(strOperation, ChangeTableKey(strTableName), loRec.GetKeyValue("ID"), OriginatingSiteKey)
                  End If
               ElseIf Len(mSatelliteDBName) > 0 Then
                  If Len(loRec.GetKeyValue("ID")) < 1 Then
                     goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateAsRecordset, no ID value found in Record: " & msRecordID
                     strOperation = "R"
                  Else
                     mwrChangeLogKey = WriteChangeHeader(strOperation, ChangeTableKey(strTableName), loRec.GetKeyValue("ID"), goSession.Site.SiteKey)
                  End If
               End If
               If goSession.Site.SiteType = SITE_TYPE_SHORE Then
                  If goSession.ThisSite.IsLogIncomingTransactions Then
                     If Len(loRec.GetKeyValue("ID")) < 1 Then
                        goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateAsRecordset, no ID value found in Record: " & msRecordID
                        strOperation = "R"
                     Else
                        mwrChangeLogKey = WriteChangeHeader(strOperation, ChangeTableKey(strTableName), loRec.GetKeyValue("ID"), OriginatingSiteKey, goSession.Site.SiteKey, True)
                     End If
                  End If
               End If
         End Select
         
         '
         ' X is Passthrough SQL Command
         '
         Select Case strOperation
            Case Is = "X"
               '
               ' Execute SQL commands for each ConfigKey
               '
               If Not ExecuteRecordQueries(loRec) Then
                  If Not LogRsError(msRecordID, 0, "SQL Execute Command failed in Record.") Then
                     UpdateAsRecordset = False
                     RollbackTransactions
                     Exit Function
                  End If
               End If
            
            Case Is = "A", "APK", "AMPK"
               ' "APK" generates new PK if duplicate error occurs
               '
               ' Validate: Need a table name for rest of commands...
               '
               '
               ' 1/2003 - Convert Datagram to Insert Statement option
               '
               If mIsEventDatagram Then
                  If OpenEventColumnMap(strTableName) Then
                     IsMappedEvent = True
                  Else
                     IsMappedEvent = False
                  End If
               End If
               If moParent.TemplateProperties.MakeInsertStatement = True Then
                  If Not UpdateAsInsertStatement(loRec, mwrChangeLogKey) Then
                     If Not LogRsError(msRecordID, 0, "Error executing SQL Statement.") Then
                        RollbackTransactions
                        UpdateAsRecordset = False
                        Exit Function
                     End If
                  End If
               Else
                  '
                  ' Add record to database table...
                  '
                  ' First update record is the "primary" (all others children)...
                  If mCurrentRecord = 2 Then
                     IsPrimaryRecord = True
                  Else
                     IsPrimaryRecord = False
                  End If
                  '
                  If strOperation = "APK" Then
                     bIsAPK = True
                  Else
                     bIsAPK = False
                  End If
                  If strOperation = "AMPK" Then
                     bIsAMPK = True
                  Else
                     bIsAMPK = False
                  End If
                  
                  If Not UpdateAsRecord(loRec, mwrChangeLogKey, IsPrimaryRecord, bIsAPK, bIsAMPK, AuditTrailHeaderKey) Then
                     IsCancelUpdate = True
                     If Not LogRsError(msRecordID, 0, "Error in Update As record: ") Then
                        UpdateAsRecordset = False
                        RollbackTransactions
                        Exit Function
                     End If
                  End If
               End If
               '
               ' Maybe Create Event Record ?
               '
               If mIsEventDatagram And IsPrimaryRecord Then
                  If goSession.EventTypes(mEventTypeKey).IsAutoCreateEvent Then
                     Set loEventWork = New mwEventWork
                     If Not loEventWork.AutoCreateEventFromDg(mEventTypeKey, loRec) Then
                        If Not LogRsError(msRecordID, 0, "Error in Auto Create Event.") Then
                           KillObject loEventWork
                           UpdateAsRecordset = False
                           RollbackTransactions
                           Exit Function
                        End If
                     End If
                     KillObject loEventWork
                  End If
               End If
            Case Is = "M", "U"
               '
               If Not UpdateAsModifyRecord(loRec, mwrChangeLogKey, strOperation, AuditTrailHeaderKey) Then
                  IsCancelUpdate = True
                  If Not LogRsError(msRecordID, 0, "Error in Update As Modify record.") Then
                     UpdateAsRecordset = False
                     RollbackTransactions
                     Exit Function
                  End If
               End If
               '
               
            Case Is = "D"
               '
               '
               '
               If Len(mSatelliteDBName) > 0 Then
                  FixupChangeLogSites
               End If
               
               If Not UpdateAsDeletedRecord(loRec, mwrChangeLogKey) Then
                  IsCancelUpdate = True
                  If Not LogRsError(msRecordID, 0, "Error in Update As Deleted Record.") Then
                     UpdateAsRecordset = False
                     RollbackTransactions
                     Exit Function
                  End If
               End If
               
               
            Case Is = "S"
            
               ' Remote Schema
               
               
               SchemaName = Trim(loRec.GetKeyValue("SchemaName"))
               SchemaContents = loRec.GetKeyValue("SchemaContents")
               
               SchemaContents = Replace(SchemaContents, "|", vbCrLf)
               
               If Not ProcessReplicatedSchema(SchemaName, SchemaContents) Then
                  IsCancelUpdate = True
                  If Not LogRsError(msRecordID, 0, "Error in Remote Schema record.") Then
                     UpdateAsRecordset = False
                     RollbackTransactions
                     Exit Function
                  End If
               End If
               
            Case Is = "L"
               '
               ' Send Log request
               '
               SchemaName = Trim(loRec.GetKeyValue("REQUSERKEY"))
               SchemaContents = loRec.GetKeyValue("SENDTOEMAIL")
               Const ZIP_ATTACHMENT_LOG = 1

               If Not ProcessLogRequest(CLng(SchemaName), ZIP_ATTACHMENT_LOG, SchemaContents) Then
                  IsCancelUpdate = True
                  If Not LogRsError(msRecordID, 0, "Error in Remote Schema record.") Then
                     UpdateAsRecordset = False
                     RollbackTransactions
                     Exit Function
                  End If
                  goSession.RaiseWorkflowStatus "Request Log FAILED to Send to  " & SchemaContents
               Else
                  goSession.RaiseWorkflowStatus "Request Log Send to  " & SchemaContents
               End If
            
               '
            Case Is = "I"
               'Debug.Print "i"
               '
               ' Batch Integrity Record = Ignore
               '
               'goSession.RaiseWorkflowStatus "Batch Integrity Record..."
            Case Is = "R"
               'Debug.Print "R"
               '
               ' Rejected Record = Ignore
               '
               'goSession.RaiseWorkflowStatus "Rejected Record..."
         End Select
         '
         ' Possible Duplicate Error...
         '
         If Not moRS Is Nothing Then
            If moRS.State = adStateOpen Then
               If moRS.LockType <> adLockReadOnly Then
                  If Not (moRS.EOF Or moRS.BOF) Then
                     If IsCancelUpdate Then
                        moRS.CancelUpdate
                     Else
                        moRS.Update
                     End If
                  End If
               End If
               moRS.Close
            End If
            Set moRS = Nothing
         End If
      End If
      ' That's it for this loRec...
      goSession.KillObject loRec
      On Error GoTo FunctionError
         
   Next i
   
   ' If mTotalErrors is > mMaxErrorsBeforeQuit then we need to get out of here now
   ' we must have had ShowAllDatagramErrors on so we could see the errors but we
   ' have exceeded the max so we have to fail.
   
   If mTotalErrors >= mMaxErrorsBeforeQuit Then
      UpdateAsRecordset = False
      RollbackTransactions
      Exit Function
   End If
   
   '
   ' We be done...
   '
   If mIsBatchDatagram Then
      '
      ' Update Batch Log
      '
      loBatchWork.CreateInboundBatchLog nBatchType, moParent.FormProperties.SiteKey, nBatchNo, _
       moParent.FormProperties.FormID, moRecords.Count, mTotalErrors
   End If
   
   If Len(mSatelliteDBName) > 0 Then
      FixupChangeLogSites
   End If
   
   
   '
  
   moCon.CommitTrans
   goSession.RaiseWorkflowStatus "Commit Transaction Complete..."
   goSession.RaiseWorkflowStatus "Processed Data Form..........: " & moParent.FormProperties.FormID & " at: " & Now()
   UpdateAsRecordset = True
   Exit Function
FunctionError:
   'Resume Next
   If (moParent.WorkflowProperties.TolerateDuplicateRecords And Err.Number = -2147467259) Or _
    moParent.WorkflowProperties.TolerateOtherErrors Then
      'Future: Need to raise an event here to inform that a duplicate encountered...
      If Not LogRsError(msRecordID, Err.Number, Err.Description) Then
         Resume Next
      Else
         On Error Resume Next
         If moCon.State = adStateOpen Then
            RollbackTransactions
         End If
         UpdateAsRecordset = False
      End If
   Else
      goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateAsRecordset.", Err.Number, Err.Description
'      Resume Next
      On Error Resume Next
      If moCon.State = adStateOpen Then
         RollbackTransactions
      End If
      UpdateAsRecordset = False
   End If
End Function
Private Function ProcessReplicatedSchema(SchemaName As String, SchemaContents As String) As Boolean

   Dim loWork As mwSchemaUpdateWork
   On Error GoTo FunctionError
   Set loWork = New mwSchemaUpdateWork
   ProcessReplicatedSchema = loWork.ProcessReplicatedSchema(SchemaName, SchemaContents)
   Set loWork = Nothing
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwSession.ProcessReplicatedSchema. ", Err.Number, Err.Description
   Set loWork = Nothing
   ProcessReplicatedSchema = False
End Function

Private Sub FixupChangeLogSites()
   Dim nSiteKey As Long
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo SubError
   
   sSQL = "SELECT * from mwrChangeLog WHERE mwcSitesKey Is Null AND mwrBatchLogOutboundKey Is Null AND (IsLoggedTransaction Is Null or IsLoggedTransaction = 0) AND " & _
          " mwrChangeTableKey IN (SELECT ID FROM mwrChangeTable WHERE mwrBatchTypeKey = 100) "
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Do While Not loRs.EOF
      
      nSiteKey = GetSiteKey(loRs!mwrChangeTableKey, loRs!RecordKey)
      
      If nSiteKey > 0 Then
         loRs!mwcSitesKey = nSiteKey
         loRs.Update
         
      End If
      
      loRs.MoveNext
   Loop
   QuietCloseRecordset loRs
   
   Exit Sub
SubError:
   'Resume Next
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.FixupChangeLogSites. ", Err.Number, Err.Description
   QuietCloseRecordset loRs
End Sub
Private Function GetSiteKey(mwrChangeTableKey As Long, RecordKey As Long) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   Dim TableName As String
   
   On Error GoTo FunctionError
   
   TableName = GetTableName(mwrChangeTableKey)
   
   If Len(TableName) > 0 And RecordKey > 0 Then
      sSQL = "SELECT mwcSitesKey FROM " & TableName & " WHERE ID = " & RecordKey
      
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly

      If loRs.RecordCount > 0 Then
         GetSiteKey = ZeroNull(loRs!mwcSitesKey)
      Else
         GetSiteKey = -1
      End If
   Else
      GetSiteKey = -1
   End If
   QuietCloseRecordset loRs
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.GetSiteKey. ", Err.Number, Err.Description
   QuietCloseRecordset loRs
   GetSiteKey = -1

End Function
Private Function GetTableName(mwrChangeTableKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo FunctionError
   
   sSQL = "SELECT Tablename from mwrChangeTable WHERE ID = " & mwrChangeTableKey
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly

   If loRs.RecordCount > 0 Then
      GetTableName = loRs!TableName
   Else
      GetTableName = ""
   End If
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.GetTableName. ", Err.Number, Err.Description
   QuietCloseRecordset loRs
   GetTableName = ""

End Function

Private Function LogRsError(RecordID As String, ErrorNumber As Long, ErrorMsg As String) As Boolean
   Static sErrorFile As String
   Static fso As FileSystemObject
   Dim ts As TextStream
   '
   '
   goSession.RaiseWorkflowStatus "Error logged in mwSession.mwDataForm.LogRsError, " & RecordID, ErrorNumber, ErrorMsg
   If mTotalErrors < 1 Then
      '
      ' Create error file
      Set fso = New FileSystemObject
      sErrorFile = goSession.Workflow.ArchiveErrorZipFile & "\" & moParent.FormProperties.FormID & ".ErrorLog"
      Set ts = fso.OpenTextFile(sErrorFile, ForAppending, True)
      ts.WriteLine "Log File generated: " & Now()
      ts.WriteLine "Max Errors Allowed is " & goSession.Workflow.MaxErrorsBeforeQuit
      ts.WriteLine "----------------------------------------------------"
      ts.WriteLine RecordID & " - " & str(ErrorNumber) & " - " & ErrorMsg
      ts.Close
      Set ts = Nothing
      mTotalErrors = mTotalErrors + 1
   Else
      '
      ' Continue logging
      '
      Set ts = fso.OpenTextFile(sErrorFile, ForAppending, False)
      ts.WriteLine RecordID & " - " & str(ErrorNumber) & " - " & ErrorMsg
      ts.Close
      Set ts = Nothing
      mTotalErrors = mTotalErrors + 1
   End If
   If moParent.WorkflowProperties.TolerateOtherErrors Then
      If mTotalErrors >= mMaxErrorsBeforeQuit And mIsShowAllDatagramErrors = False Then
         LogRsError = False
      Else
         LogRsError = True
      End If
   Else
      LogRsError = False
   End If
   
   goSession.Workflow.IsWorkflowError = True
   
   DoEvents
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.LogRsError.", Err.Number, Err.Description
   LogRsError = False
End Function

Private Function UpdateAsInsertStatement(loRec As ConfigKeys, mwrChangeLogKey As Long) As Boolean
   Dim strFields As String
   Dim strData As String
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim i As Integer
   On Error GoTo FunctionError
   strFields = "Insert into " & loRec.GetKeyValue("Table") & " ("
   strData = " values ("
   For Each loKey In loRec
      Select Case loKey.KeyName
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Else
            If (Not IsNull(loKey.KeyValue)) And (Not IsEventColumnDiscard(loKey.KeyName)) Then
               If loKey.KeyName <> "" And loKey.KeyValue <> "" Then
                  strFields = strFields & loKey.KeyName & ", "
                  If IsEventColumnNumeric(loKey.KeyName) Then
                     strData = strData & loKey.KeyValue & ", "
                  Else
                     'strip unwanted characters...
                     strData = strData & "'" & StripInsertChars(loKey.KeyValue) & "', "
                  End If
'                  WriteChangeColumn(mwrChangeLogKey, RecField As field)
               End If
            End If
      End Select
   Next loKey
   ' strip last "," and put it together...
   strFields = Left(strFields, Len(strFields) - 2)
   strData = Left(strData, Len(strData) - 2)
   sSQL = strFields & ") " & strData & ")"
   goSession.RaiseWorkflowStatus "mwDataform.UpdateAsInsertStatement, about to execute: " & sSQL
   moCon.Execute sSQL
   UpdateAsInsertStatement = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateAsInsertStatement: " & sSQL, Err.Number, Err.Description
   UpdateAsInsertStatement = False
End Function

Private Function UpdateAsRecord(loRec As ConfigKeys, mwrChangeLogKey As Long, Optional IsPrimaryRecord As Boolean, Optional IsAPK As Boolean, Optional IsAMPK As Boolean, Optional AuditTrailHeaderKey As Long) As Boolean
   Dim strTable As String
   Dim sSQL As String
   Dim nRecKey As Long
   Dim i As Long
   Dim bIsMakePk As Boolean
   Dim nErrNumber As Long
   Dim sErrDesc As String
   On Error GoTo FunctionError
   Set moRS = New Recordset
  ' If goSession.IsOracle Then
    '  moRS.CursorLocation = adUseServer
  ' Else
    moRS.CursorLocation = adUseClient
  ' End If
   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Or moParent.GetExtendedProperty("IsMakePkOnProcess") = "-1" Or moParent.GetExtendedProperty("IsMakePkOnProcess") = "1" Or IsAMPK = True Then
      bIsMakePk = True
   Else
      bIsMakePk = False
   End If
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   strTable = loRec.GetKeyValue("table")
   If bIsMakePk Then
      sSQL = "select * from " & loRec.GetKeyValue("table") & " where ID=-1 order by ID"
      moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   Else
      If mIsBatchDatagram Or (moParent.GetExtendedProperty("NoIdColumnInTable") <> "1" And moParent.GetExtendedProperty("NoIdColumnInTable") <> "True") Then
      
         sSQL = "select * from " & loRec.GetKeyValue("table") & " where ID=-1 order by ID"
         moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
      
      Else
         sSQL = "select * from " & loRec.GetKeyValue("table") & " where 1=0"
         moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
      End If
   End If
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord, Unable to open Table: " & strTable, Err.Number, Err.Description
      QuietCloseRecordset moRS
      UpdateAsRecord = False
      Exit Function
   End If
   moRS.AddNew
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord, Unable to AddNew, Table: " & strTable, Err.Number, Err.Description
      QuietCloseRecordset moRS
      UpdateAsRecord = False
      Exit Function
   End If
   If bIsMakePk Then
      moRS!ID = goSession.MakePK(strTable, moCon)
      If Err Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord, Unable to assign PK value, Table: " & strTable, Err.Number, Err.Description
         QuietCloseRecordset moRS
         UpdateAsRecord = False
         Exit Function
      End If
      If moRS.Fields.Count < 75 And Not goSession.IsSqlServer Then
         moRS.Update
         If Err Then
            '
            ' APK - Create new key if True and PK Constraint error encountered..
            '
            'If err.Number = 1 Then
            '   moRS!ID = goSession.MakePK(strTable, moCon)
            '   moRS.Update
            '   If err Then
            '      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord: ", err.Number, err.Description
            '      goSession.RaiseWorkflowStatus "Failed to Update Table during Call to MakePK/APK: " & strTable
            '      goSession.CloseRecordset moRS
            '      UpdateAsRecord = False
            '      Exit Function
            '   End If
            'Else
               goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord: ", Err.Number, Err.Description
               goSession.RaiseWorkflowStatus "Failed to Update Table during Call to MakePK: " & strTable
               QuietCloseRecordset moRS
               UpdateAsRecord = False
               Exit Function
            'End If
         End If
      End If
   End If
   If UpdateRSFields(loRec, moRS, mwrChangeLogKey, , AuditTrailHeaderKey) Then
      '
      ' If Child, save key to the parent record...
      '
      If loRec.GetKeyValue("IsChildKeyName") <> "" Then
         moRS.Fields(loRec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
      End If
      '
      ' Update new record...
      '
      moRS.Update
      If Err Then
         '
         ' APK - Create new key if True and PK Constraint error encountered..
         '
         ' -2147467259 The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data,
         ' remove the index, or redefine the index to permit duplicate entries and try again."
         If (Err.Number = -2147467259 Or Err.Number = -2147217873) And (bIsMakePk Or IsAPK) Then
            moRS!ID = goSession.MakePK(strTable, moCon)
            moRS.Update
            If Err Then
               goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord: ", Err.Number, Err.Description
               goSession.RaiseWorkflowStatus "Failed to Update Table during Call to MakePK/APK: " & strTable
               moRS.CancelUpdate
               QuietCloseRecordset moRS
               UpdateAsRecord = False
               'moCon.RollbackTrans
               Exit Function
            End If
         'ElseIf Err.Number = -2147467259 Then
            ' keep going, may be compound key...
          '  goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord: ", Err.Number, Err.Description
           ' moRS.CancelUpdate
         
         ' Unique constraint, duplicate key error. Try doing a Modify instead.
         ElseIf Err.Number = -2147217873 Or Err.Number = -2147467259 Or Err.Number = -2147217900 Then
'''               goSession.RaiseWorkflowStatus "Duplicate/FK Constraint error during Add transaction " & msRecordID, Err.Number, Err.Description
'''               moRS.CancelUpdate
'''               QuietCloseRecordset moRS
'''               UpdateAsRecord = False
            nErrNumber = Err.Number
            sErrDesc = Err.Description
            If CheckIsTempRecordInsertAllowed(strTable, moRS!ID) Then
               UpdateAsRecord = True
            Else
               goSession.RaiseWorkflowStatus "Duplicate/FK Constraint error during Add transaction " & msRecordID, nErrNumber, sErrDesc
               UpdateAsRecord = False
            End If
            moRS.CancelUpdate
            QuietCloseRecordset moRS
            Exit Function
         Else
            goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsRecord: ", Err.Number, Err.Description
            goSession.RaiseWorkflowStatus "Final .Update Failed on table: " & strTable
            moRS.CancelUpdate
            QuietCloseRecordset moRS
            UpdateAsRecord = False
            'moCon.RollbackTrans
            Exit Function
         End If
      End If
      ' If parent record, save parent key for future use
      If loRec.GetKeyValue("IsParentKeyName") <> "" Then
         mCurrentParentRecordID = moRS.Fields(loRec.GetKeyValue("IsParentKeyName")).value
      End If
   Else
      moRS.CancelUpdate
      goSession.RaiseWorkflowStatus "Error encountered in mwDataForm.UpdateAsRecord: " & Err.Number & "-" & Err.Description & ", Form: " & moParent.FormProperties.FormID
      UpdateAsRecord = False
      QuietCloseRecordset moRS
      UpdateAsRecord = False
      'moCon.RollbackTrans
      Exit Function
   End If
   UpdateAsRecord = True
   '
   If IsPrimaryRecord Then
      nRecKey = moRS!ID
      UpdateProcessLog strTable, nRecKey
   End If
   'moCon.CommitTrans
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateAsRecord.", Err.Number, Err.Description
   UpdateAsRecord = False
   On Error Resume Next
   moCon.RollbackTrans
End Function

Private Function UpdateProcessLog(TableName As String, RecordPK As Long) As Boolean
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "mwwfFormProcessLog", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   With loRs
      .AddNew
      .Fields("ID").value = goSession.MakePK("mwwfFormProcessLog")
      .Fields("TemplateID").value = moParent.TemplateProperties.TemplateID
      .Fields("SiteID").value = moParent.FormProperties.SiteID
      .Fields("Tablename").value = TableName
      .Fields("DetailRecordKey").value = RecordPK
      .Fields("ProcessedDate").value = Now()
      .Update
      .Close
   End With
   Set loRs = Nothing
   UpdateProcessLog = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateProcessLog.", Err.Number, Err.Description
   UpdateProcessLog = False
End Function

Private Function UpdateRSFields(ByRef oRec As ConfigKeys, ByRef oRs As Recordset, mwrChangeLogKey As Long, Optional IsModify As Boolean, Optional AuditTrailHeaderKey As Long) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim IsSkipField As Boolean
   Dim strTable As String
   Dim strTemp As String
   Dim IsFoundadLongVarChar As Boolean
   
   On Error GoTo FunctionError
   
   IsFoundadLongVarChar = False
   
   For Each loKey In oRec
      Debug.Print loKey.KeyName
      IsSkipField = False
      Select Case loKey.KeyName
         Case Is = "TABLE"
            strTable = loKey.KeyValue
         Case Is = "OPERATION"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
'         Case Is = "CHANGEDATE"
'         Case Is = "USERKEY"
         Case Else
            If strTable = "" Then
               If loKey.KeyName = "CHANGEDATE" Or loKey.KeyName = "USERKEY" Then
                  IsSkipField = True
               End If
            End If
            
            If IsSkipField = False Then
               If Not IsNull(loKey.KeyValue) Then
                  ' During Modify, cannot allow PK value to updated...
                  If loKey.KeyValue = "[NULL]" Then
                     oRs.Fields(loKey.KeyName) = Null
                     If mwrChangeLogKey > 0 Then
                        WriteChangeColumn mwrChangeLogKey, oRs.Fields(loKey.KeyName)
                     End If
                     '---
                     If AuditTrailHeaderKey > 0 Then
                        moAuditTrail.WriteAuditTrailDetail AuditTrailHeaderKey, oRs.Fields(loKey.KeyName)
                     End If
                     '---
                  ElseIf loKey.KeyName <> "" And loKey.KeyValue <> "" And Not (IsModify = True And loKey.KeyName = "ID") Then
                     ' Multi-step avoidance in numeric data types.
                     Select Case oRs.Fields(loKey.KeyName).Type
                        Case adNumeric
                           '
                           ' Numeric Fields...
                           '
                           If IsNumeric(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator)) Then
                              If oRs.Fields(loKey.KeyName).NumericScale > 0 Then
                                 'oRs.Fields(loKey.KeyName) = CSng(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                                 oRs.Fields(loKey.KeyName) = CDbl(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                              Else
                                 oRs.Fields(loKey.KeyName) = CLng(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                              End If
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in numeric field: " & strTable & "." & loKey.KeyName
                           End If
                        Case adSingle
                           '
                           ' Numeric Fields...
                           '
                           If IsNumeric(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator)) Then
                              oRs.Fields(loKey.KeyName) = CSng(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in numeric field: " & strTable & "." & loKey.KeyName
                           End If
                        Case adDouble, adVarNumeric
                           '
                           ' Numeric Fields...
                           '
                           If IsNumeric(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator)) Then
                              oRs.Fields(loKey.KeyName) = CDbl(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in numeric field: " & strTable & "." & loKey.KeyName
                           End If
                        Case adDate, adDBTimeStamp
                           '
                           ' Date Fields
                           '
                           If IsDate(loKey.KeyValue) Then
                              If Trim(loKey.KeyValue) > "1899-12-30 00:00:00" Then
                                 oRs.Fields(loKey.KeyName) = CDate(Format(loKey.KeyValue, DATAGRAM_DATE_FORMAT))
                              End If
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in Date field: " & strTable & "." & loKey.KeyName
                           End If
   '                     Case adDBTimeStamp
   '                        '
   '                        ' Oracle Date Fields - Funky, pass it through
   '                        '
   '                        oRs.Fields(loKey.KeyName) = loKey.KeyValue
                           
                        Case adInteger
                           '
                           ' Integer Field
                           '
   '                        If IsNumeric(loKey.KeyValue) Then
                           If IsNumeric(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator)) Then
                              
                              oRs.Fields(loKey.KeyName) = CLng(Replace(loKey.KeyValue, ".", mLocalDecimalSeparator))
                              'oRs.Fields(loKey.KeyName) = Replace(loKey.KeyValue, ".", mLocalDecimalSeparator)
                              'oRs.Fields(loKey.KeyName) = loKey.KeyValue
                           ElseIf UCase(loKey.KeyValue) = "TRUE" Then
                              oRs.Fields(loKey.KeyName) = 1
                           ElseIf UCase(loKey.KeyValue) = "FALSE" Then
                              oRs.Fields(loKey.KeyName) = 0
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in Long Integer field: " & strTable & "." & loKey.KeyName
                           End If
                        Case adSmallInt, adBoolean
                           '
                           ' Integer Field
                           '
                           If IsNumeric(loKey.KeyValue) Then
                              oRs.Fields(loKey.KeyName) = loKey.KeyValue
                           ElseIf UCase(loKey.KeyValue) = "TRUE" Then
                              oRs.Fields(loKey.KeyName) = 1
                           ElseIf UCase(loKey.KeyValue) = "FALSE" Then
                              oRs.Fields(loKey.KeyName) = 0
                           Else
                              If Not IsSkipField Then goSession.RaiseWorkflowStatus "Bad value in Small Integer field: " & strTable & "." & loKey.KeyName
                           End If
                        
                        'Case adBoolean
                        '   oRs.Fields(loKey.KeyName) = loKey.KeyValue
                        
                        Case adLongVarChar
                           IsFoundadLongVarChar = True
                           
                        Case Else
                           '
                           ' All other fields...trim length if required
                           '
                           strTemp = Replace(loKey.KeyValue, "|", vbCrLf)
                           
                           If oRs.Fields(loKey.KeyName).DefinedSize < Len(strTemp) Then
                              goSession.RaiseWorkflowStatus "String value exceeds max length: " & strTable & "." & loKey.KeyName & _
                                ". Extra data is truncated."
                           End If
                           oRs.Fields(loKey.KeyName) = Left(strTemp, oRs.Fields(loKey.KeyName).DefinedSize)
   '                        If oRs.Fields(loKey.KeyName).Type <> 201 Then
   '                           oRs.Fields(loKey.KeyName) = Left(strTemp, oRs.Fields(loKey.KeyName).DefinedSize)
   '                        Else
   '                           oRs.Fields(loKey.KeyName) = strTemp
   '                        End If
   '                        ' CLOB is 201...
   '                        If goSession.IsOracle And oRs.Fields(loKey.KeyName).Type = 201 Then
   '                           oRs.Update
   '                        End If
                     End Select
                     'Re-route transactions (Type 2)
                     If mwrChangeLogKey > 0 Then
                        WriteChangeColumn mwrChangeLogKey, oRs.Fields(loKey.KeyName)
                     End If
                     '---
                     If AuditTrailHeaderKey > 0 Then
                        moAuditTrail.WriteAuditTrailDetail AuditTrailHeaderKey, oRs.Fields(loKey.KeyName)
                     End If
                     '---
                  End If
            End If
         End If
      End Select
   Next loKey
   
   '
   ' CLOB LOOP
   '
   If IsFoundadLongVarChar Then
      For Each loKey In oRec

         Select Case loKey.KeyName
            Case Is = "TABLE"

            Case Is = "OPERATION"
            Case Is = "ISPARENTKEYNAME"
            Case Is = "ISCHILDKEYNAME"
            Case Else
               If Not IsNull(loKey.KeyValue) Then
                  ' During Modify, cannot allow PK value to updated...
                  If loKey.KeyValue = "[NULL]" Then
                     ' no action
                  ElseIf loKey.KeyName <> "" And loKey.KeyValue <> "" And Not (IsModify = True And loKey.KeyName = "ID") Then

                     Select Case oRs.Fields(loKey.KeyName).Type
                        Case adLongVarChar



                           strTemp = Replace(loKey.KeyValue, "|", vbCrLf)
                           oRs.Fields(loKey.KeyName) = strTemp

                     End Select

                  End If
               End If
         End Select
      Next loKey

   End If
   
   UpdateRSFields = True
   Exit Function
FunctionError:
'Resume Next
   If Err.Number = 3265 Then
      '
      ' Missing Column handling...
      '
      If mIsEventDatagram And IsEventColumnDiscard(loKey.KeyName) Then
         'Debug.Print "Discarding table/column: " & strTable & "/" & loKey.KeyName
         IsSkipField = True
         Resume Next
      End If
      goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateRSFields, Column Missing: " & strTable & "." & loKey.KeyName
      If moParent.WorkflowProperties.TolerateMissingColumns Then
         IsSkipField = True
         Resume Next
      Else
         UpdateRSFields = False
      End If
   Else
      '
      ' General Error Handling
      '
      iErr = Err.Number
      sErr = Err.Description
      'Assume invalid field, cancel operation...
      goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataForm.UpdateRSFields: ", Err.Number, Err.Description
      goSession.RaiseWorkflowStatus "Failed Recordset Update on: " & strTable & "." & loKey.KeyName & " - " & loKey.KeyValue
      UpdateRSFields = False
   End If
End Function


Private Function ExecuteRecordQueries(ByRef oRec As ConfigKeys) As Boolean
   Dim i As Integer
   Dim IsError As Boolean
   Dim strSQL As String
   Dim loDataWork As mwDataWork
   Dim strOracleContinuationSQL As String
   Dim strOracleContinuationSQL2 As String
   
   On Error Resume Next
   
   Set loDataWork = goSession.OpenNewDataWork
   
   '
   ' Begin Transaction Code Here
   '
   For i = 2 To oRec.Count
      strSQL = oRec(i).KeyValue
      
      If loDataWork.IsSqlServer() Then
         strSQL = loDataWork.ConvertToSQLSyntax(strSQL)
      ElseIf loDataWork.IsOracle() Then
         strSQL = loDataWork.ConvertToOracleSyntax(strSQL, strOracleContinuationSQL, strOracleContinuationSQL2)
      End If
      
      moCon.Execute strSQL
      If Err Then
         goSession.RaiseWorkflowStatus "Error executing SQL, error: " & Err.Number & " - " & Err.Description
         'goSession.RaiseWorkflowStatus "Failed Execution Query: " & strSQL
         If moParent.WorkflowProperties.TolerateOtherErrors Then
            IsError = True
         Else
            ExecuteRecordQueries = False
            goSession.RaiseWorkflowStatus "Tolerate Other Errors set to false, aborting..."
            KillObject loDataWork
            Exit Function
         End If
      Else
         goSession.RaiseWorkflowStatus "Successfully Executed SQL Statement: " & strSQL
      End If
   Next i
   On Error GoTo FunctionError:
   If IsError Then
      ExecuteRecordQueries = False
   Else
      ExecuteRecordQueries = True
   End If
   '
   ' Commit Transaction Code Here...
   '
   KillObject loDataWork
   Exit Function
FunctionError:
   ExecuteRecordQueries = False
   KillObject loDataWork
End Function


'
'
Public Function CreateDatagramFromRecordset(ByRef oRs As Recordset, TableName As String, Optional IsAPKOperation As Boolean) As Boolean
   Dim oField As ADODB.field
   Dim loRec As ConfigKeys
   Dim i As Integer
   
   On Error GoTo FunctionError
   '
   ' Must be open datagram to start
   '
   If moParent.IsOpen = False Or moParent.FormType <> mwft_DATA_FORM Then
      goSession.RaiseError "Error in CreateDatagramFromRecordset, Form not open or wrong form type."
      CreateDatagramFromRecordset = False
      Exit Function
   End If
   '
   ' Any records in there ?
   '
   If oRs.State = adStateClosed Then
      goSession.RaiseError "Error in CreateDatagramFromRecordset, Recordset is closed."
      CreateDatagramFromRecordset = False
      Exit Function
   End If
   
   If oRs.RecordCount < 1 Then
      goSession.RaiseError "Error in CreateDatagramFromRecordset, Recordset has no records."
      CreateDatagramFromRecordset = False
      Exit Function
   End If
   '
   ' OK, Loop through the recordset
   '
   oRs.MoveFirst
   Do While Not oRs.EOF
      '
      '
      Set loRec = AddRecord()
      With loRec
         .Add "Table", TableName
         If IsAPKOperation Then
            ' Special, new PK generated if PK constraint error
            .Add "Operation", "APK"
         Else
            .Add "Operation", "A"
         End If
         For Each oField In oRs.Fields
            ' Add to datagram, massage for data type
            If Not IsNull(oField.value) And oField.value <> "" Then
               '.Add oField.Name, Replace(oField.value, vbCrLf, "|")
               '
               ' Format field value for target database
               '
               AddRecordField loRec, oField.Name, oField
            End If
         Next oField
      End With
      oRs.MoveNext
   Loop
   oRs.MoveFirst
   ' That's all folks
   CreateDatagramFromRecordset = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in CreateDatagramFromRecordset. ", Err.Number, Err.Description
   CreateDatagramFromRecordset = False

End Function

Public Function TearDown(Optional LeaveAppOpen As Boolean)
   On Error GoTo FunctionError
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
         ' Don't close if form is using the main connection...
         If Trim(moParent.WorkflowProperties.ConnectionID) <> "" Then
            moCon.Close
         End If
      End If
      Set moCon = Nothing
   End If
   If Not moRecords Is Nothing Then Set moRecords = Nothing
   If Not moParent Is Nothing Then Set moParent = Nothing
   Exit Function
FunctionError:
   If Err.Number = -2147467259 And moParent.WorkflowProperties.TolerateDuplicateRecords Then
      moRS.CancelUpdate
      Resume Next
   Else
      goSession.RaiseError "General Error in mwDataForm.Teardown. ", Err.Number, Err.Description
   End If
End Function

Public Function GetField(FieldName, Optional RecordNumber As Variant) As String
   Dim strFieldName As String
   On Error GoTo FunctionError:
   strFieldName = UCase(FieldName)
   If moParent.IsOpen Then
      If IsMissing(RecordNumber) Then
         If moRecords(2).ConfigKeys(strFieldName) Is Nothing Then
            GetField = ""
         Else
            GetField = moRecords(2).ConfigKeys(strFieldName).KeyValue
         End If
      Else
         If Not IsNumeric(RecordNumber) Then
            goSession.RaiseError "Error in mwSession.mwDataForm.GetField, RecordNumber must contain a number."
            GetField = ""
         Else
            'if RecordNumber > moDataForm.
            GetField = moRecords(RecordNumber).ConfigKeys(strFieldName).KeyValue
         End If
      End If
   Else
      GetField = ""
   End If
   Exit Function
FunctionError:
   'moParent.RaiseError "General Error in mwSession.GetField.", err.Number, err.Description
   GetField = ""
End Function



'
' 5/2002 - Convert Datagram to Email TextBody
' 8/2005 - Used for shipboard email only
Public Function ConvertToEmailTextBody(Optional Template As String, Optional NoDllCallout As Boolean) As String
   Dim strTemplate As String
   Dim strTextBody As String
   Dim strShipCode As String
   Dim nSiteKey As Long
   Dim sSQL As String
   Dim loRs As Recordset
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim loFormEmailTemplate As frmEmailTemplate
   Dim sEmailSelected As String
   On Error GoTo FunctionError
   
   '  Default
   If Trim(Template) <> "" Then
      strTemplate = goSession.Workflow.ScriptFolder & "\" & Template
   End If
   '
   ' Fetch Template from database...
   '
   nSiteKey = goSession.Site.GetSiteKey(moParent.FormProperties.SiteID)
   sSQL = "SELECT mwwfFormTemplateEmailList.EmailTemplateTitle, mwwfFormTemplateEmailList.EmailTemplatePath, " & _
     " mwwfFormTemplateEmailList.TemplateID, mwwfFormTemplateEmailList.EmailAddress FROM (mwcFleets INNER JOIN mwwfFormTemplateEmailList ON " & _
     " mwcFleets.ID = mwwfFormTemplateEmailList.FleetKey) INNER JOIN mwcFleetSites ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey " & _
     " WHERE mwwfFormTemplateEmailList.IsShipsideEmail<>0 AND mwcFleetSites.mwcSitesKey=" & nSiteKey & _
     " AND mwwfFormTemplateEmailList.TemplateID='" & moParent.TemplateProperties.TemplateID & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   
   If loRs.RecordCount = 1 Then
      If Not IsNull(loRs!EmailTemplatePath) Then
         strTemplate = goSession.Workflow.ScriptFolder & "\" & loRs!EmailTemplatePath
         moParent.SetExtendedProperty "ConvertDgToEmailTemplate", loRs!EmailTemplatePath, moParent.TemplateProperties.TemplateID
         mEmailAddress = BlankNull(loRs!EmailAddress)
      End If
   ElseIf loRs.RecordCount > 1 Then
      '
      ' Display form to choose a template...
      '
      Set loFormEmailTemplate = New frmEmailTemplate
      If Not loFormEmailTemplate.InitForm(loRs) Then
         goSession.LogIt mwl_Transport, mwl_Information, "Email template was not selected for Send"
      Else
         loFormEmailTemplate.Show vbModal
         
         If loFormEmailTemplate.mIsCancelled Then
            goSession.LogIt mwl_Transport, mwl_Information, "Email template was cancelled for Send"
         Else
            ' of many you selected 1
            If Not IsNull(loFormEmailTemplate.mEmailTemplatePath) Then
               strTemplate = goSession.Workflow.ScriptFolder & "\" & loFormEmailTemplate.mEmailTemplatePath
               moParent.SetExtendedProperty "ConvertDgToEmailTemplate", loFormEmailTemplate.mEmailTemplatePath, moParent.TemplateProperties.TemplateID
            End If
            mEmailAddress = BlankNull(loFormEmailTemplate.mEmailAddress)
         End If
      End If
      Unload loFormEmailTemplate
      KillObject loFormEmailTemplate
   Else
      goSession.RaiseWorkflowStatus "Warning in mwSession.mwDataForm.Convert to EmailTextBody, Email Template not Found on query: " & sSQL
   End If
   

   If Trim(strTemplate) = "" Then         ' vrs_Pos???.textbody template
      '- no email
      ConvertToEmailTextBody = ""
      Exit Function
   End If
   
   
   '
   ' Custom DLL option...
   '
   If Trim(moParent.WorkflowProperties.VbFormEventDLL) <> "" And Not NoDllCallout Then
      '
      strTextBody = ExecuteDllOnDatagramEmail()
      
   Else
      '
      ' Read in Template...
      '
      strTextBody = ""
      Set fso = New FileSystemObject
      If fso.FileExists(strTemplate) Then
         Set ts = fso.OpenTextFile(strTemplate, ForReading, False)
         strTextBody = ts.ReadAll
         ts.Close
         goSession.KillObject ts
      End If
   End If
   '
   ' Basic Datagram translation...
   '
   If Trim(strTextBody) <> "" Then
      strTextBody = ConvertBasicDgText(strTextBody)
   
      If UCase(Left(moParent.TemplateProperties.TemplateID, 4)) = "VRS_" Then
         strTextBody = ConvertCustomVrsDgText(strTextBody)
      End If
   End If
   ConvertToEmailTextBody = strTextBody
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwDataForm.ConvertToEmailTextBody. ", Err.Number, Err.Description
   ConvertToEmailTextBody = ""
   KillObject loFormEmailTemplate
End Function

Public Function FetchEmailTemplateRs() As Recordset
   Dim nSiteKey As Long
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loFormEmailTemplate As frmEmailTemplate
   On Error GoTo FunctionError
   
   nSiteKey = goSession.Site.GetSiteKey(moParent.FormProperties.SiteID)
   sSQL = "SELECT mwwfFormTemplateEmailList.EmailTemplateTitle, mwwfFormTemplateEmailList.EmailTemplatePath, " & _
     " mwwfFormTemplateEmailList.TemplateID, mwwfFormTemplateEmailList.EmailAddress FROM (mwcFleets INNER JOIN mwwfFormTemplateEmailList ON " & _
     " mwcFleets.ID = mwwfFormTemplateEmailList.FleetKey) INNER JOIN mwcFleetSites ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey " & _
     " WHERE mwwfFormTemplateEmailList.IsShipsideEmail<>0 AND mwcFleetSites.mwcSitesKey=" & nSiteKey & _
     " AND mwwfFormTemplateEmailList.TemplateID='" & moParent.TemplateProperties.TemplateID & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      Set loFormEmailTemplate = New frmEmailTemplate
      If Not loFormEmailTemplate.InitForm(loRs) Then
         goSession.LogIt mwl_Transport, mwl_Information, "Email was not selected for Send"
      Else
         If loRs.RecordCount > 1 Then
            loFormEmailTemplate.Show vbModal
         End If
         
         If loFormEmailTemplate.mIsCancelled Then
            goSession.LogIt mwl_Transport, mwl_Information, "Email was cancelled for Send"
            Unload loFormEmailTemplate
            KillObject loFormEmailTemplate
            Exit Function
         Else
            Set FetchEmailTemplateRs = loFormEmailTemplate.GetEmailTemplateRs
         End If
      End If
      Unload loFormEmailTemplate
      KillObject loFormEmailTemplate
   Else
      goSession.RaiseWorkflowStatus "Warning in mwSession.mwDataForm.Convert to EmailTextBody, Email not Found on query: " & sSQL
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwsession.mwDataForm.FetchEmailTemplateRs. ", Err.Number, Err.Description
End Function

Public Function ConvertToEmailTextBody1(strTemplate As String, Optional NoDllCallout As Boolean) As String
   Dim strTextBody As String
   Dim fso As FileSystemObject
   Dim ts As TextStream
   On Error GoTo FunctionError
   
   If Trim(strTemplate) = "" Then         ' vrs_Pos???.textbody template
      '- no email
      ConvertToEmailTextBody1 = ""
      Exit Function
   End If
   
   '
   ' Custom DLL option...
   '
   If Trim(moParent.WorkflowProperties.VbFormEventDLL) <> "" And Not NoDllCallout Then
      '
      strTextBody = ExecuteDllOnDatagramEmail()
      
   Else
      '
      ' Read in Template...
      '
      strTextBody = ""
      Set fso = New FileSystemObject
      If fso.FileExists(strTemplate) Then
         Set ts = fso.OpenTextFile(strTemplate, ForReading, False)
         strTextBody = ts.ReadAll
         ts.Close
         goSession.KillObject ts
      End If
   End If
   '
   ' Basic Datagram translation...
   '
   If Trim(strTextBody) <> "" Then
      strTextBody = ConvertBasicDgText(strTextBody)
   
      If UCase(Left(moParent.TemplateProperties.TemplateID, 4)) = "VRS_" Then
         strTextBody = ConvertCustomVrsDgText(strTextBody)
      End If
   End If
   ConvertToEmailTextBody1 = strTextBody
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.ConvertToEmailTextBody1. ", Err.Number, Err.Description
End Function


Private Function ConvertCustomVrsDgText(TextBody As String) As String
   Dim sTextBody As String
   Dim sShipCode As String
   Dim oVrsWork As Object
   Dim sFACTS_TO_DATE_Selected As String
   Dim sFACTS_Spec As String
   Dim sFACTS_FMT As String
   Dim sFACTS_List As String
   Dim IsLoop As Boolean
   Dim PosStart As Long
   Dim PosEnd As Long
   On Error GoTo FunctionError
   
   sTextBody = TextBody
   
   ' alternate of All  '[FACTS_TO_DATE] is [FACTS_TO_DATE_SELECTED]  Keys selected list received as public properties
   If mPortCallKey > 0 And InStr(1, sTextBody, "[FACTS_TO_DATE_SELECTED ", vbTextCompare) Then
      Set oVrsWork = CreateObject("mwVrsMessage.vrsWork")
      oVrsWork.InitSession Me.Parent.Parent
      
      'DEV-2176 Parameterized Facts (maybe multiple tags)
      PosStart = InStr(1, sTextBody, "[Facts_TO_DATE_SELECTED ", vbTextCompare)
      If PosStart > 0 Then
         IsLoop = True
         Do While IsLoop
            
            PosEnd = InStr(PosStart, sTextBody, "]", vbTextCompare)
            If PosEnd <= PosStart Then
               IsLoop = False                                        ' end
               Exit Do
            End If
            
            sFACTS_Spec = mID(sTextBody, PosStart, ((PosEnd + 1) - PosStart))
            oVrsWork.FactsFMT = sFACTS_Spec                          ' pass the spec then get the list
            If oVrsWork.FactsFMTType = "B" Then
               sFACTS_List = oVrsWork.GetSOFtoDateSelectedFMTB(CStr(mPortCallKey), mStrBerthKeys, mStrCargoKeys)
               sTextBody = Replace(sTextBody, sFACTS_Spec, sFACTS_List) ' replace the list
            ElseIf oVrsWork.FactsFMTType = "C" Then
               sFACTS_List = oVrsWork.GetSOFtoDateSelectedFMTC(CStr(mPortCallKey), mStrBerthKeys, mStrCargoKeys)
               sTextBody = Replace(sTextBody, sFACTS_Spec, sFACTS_List) ' replace the list
            Else
               sFACTS_List = ""
               sTextBody = Replace(sTextBody, sFACTS_Spec, sFACTS_List) ' replace the list
            End If
            
            ' set for next tag
            PosStart = InStr(PosStart, sTextBody, "[FACTS_TO_DATE_SELECTED ", vbTextCompare)
            If PosStart = 0 Then IsLoop = False                ' end
         Loop
      End If
   End If
   If mStrBerthKeys <> "" And mStrCargoKeys <> "" And mPortCallKey > 0 And InStr(1, sTextBody, "[FACTS_TO_DATE_SELECTED]", vbTextCompare) > 1 Then
      If Not IsObject(oVrsWork) Then
         Set oVrsWork = CreateObject("mwVrsMessage.vrsWork")
         oVrsWork.InitSession Me.Parent.Parent
      End If
      
      sFACTS_TO_DATE_Selected = oVrsWork.GetSOFtoDateSelected(CStr(mPortCallKey), mStrBerthKeys, mStrCargoKeys)
      sTextBody = Replace(sTextBody, "[FACTS_TO_DATE_SELECTED]", sFACTS_TO_DATE_Selected)
   End If
   If IsObject(oVrsWork) Then
      Set oVrsWork = Nothing
   End If
   
   sTextBody = Replace(sTextBody, "[VRS_BUNKERS1_LABEL]", goSession.Site.GetExtendedProperty("vrsBunkers1Label"))
   sTextBody = Replace(sTextBody, "[VRS_BUNKERS2_LABEL]", goSession.Site.GetExtendedProperty("vrsBunkers2Label"))
   sTextBody = Replace(sTextBody, "[VRS_ROB1_LABEL]", goSession.Site.GetExtendedProperty("vrsROB1Label"))
   sTextBody = Replace(sTextBody, "[VRS_ROB2_LABEL]", goSession.Site.GetExtendedProperty("vrsROB2Label"))
   ' Vessel Code and Vessel name as well...
   If Not moRecords(2).ConfigKeys("vessel_code") Is Nothing Then
      sShipCode = moRecords(2).ConfigKeys("vessel_code").KeyValue
      sTextBody = Replace(sTextBody, "[VESSEL_CODE]", sShipCode)
      sTextBody = Replace(sTextBody, "[VESSEL_NAME]", goSession.Site.GetShipProperty(sShipCode, "SiteName"))
   End If
   
   sTextBody = ConvertCustomVrsNGrades(sTextBody)
   
   ' Strip out unused remarks...
   sTextBody = Replace(sTextBody, "[REMARK1]", "")
   sTextBody = Replace(sTextBody, "[REMARK2]", "")
   sTextBody = Replace(sTextBody, "[REMARK3]", "")
   sTextBody = Replace(sTextBody, "[REMARK4]", "")
   sTextBody = Replace(sTextBody, "[REMARK5]", "")
   ConvertCustomVrsDgText = sTextBody
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.ConvertCustomVrsDgText. ", Err.Number, Err.Description
   ConvertCustomVrsDgText = TextBody
End Function

Public Function ConvertBasicDgText(strTextBody As String) As String
   Dim loKey As ConfigKey
   Dim strReplace As String
   Dim loRec As ConfigGroup
   Dim nRecNum As Integer
   Dim nLeft As Long
   Dim nRight As Long
   Dim nCurrPos As Long
   Dim nStartPos As Long
   Dim nEndPos As Long
   Dim sKeyPlusFormat As String
   
   Dim sFormattedData As String
   Dim sFormat As String              ' string format
   Dim IsCustomValue As Boolean
   Dim IsDashFormat As Boolean
   Dim sFKeyName As String       ' string formatted KeyName (from textbody)
   Dim sSearch As String         ' string search for
   Dim sBadMask As String
   
   Dim nPos As Long
   Dim ovrsExportWork As Object
   Dim IsAltPortCode As Boolean
   Dim sAltPortCode As String
   Dim sWeatherCode As String
   Dim sWeatherDescription As String
   
   On Error GoTo FunctionError
   
   Set ovrsExportWork = CreateObject("mwVrs6.vrsExportWork")
   
   For nRecNum = 2 To moRecords.Count
      '  Skip Header record...
      Set loRec = moRecords(nRecNum)
      
      For Each loKey In loRec.ConfigKeys
                              
         ' First translate all occurrences of [KeyName] with no optional format
                              
         sSearch = "[" & UCase(loKey.KeyName) & "]"
         
         If sSearch = "[VRSWEATHERKEY]" Then
            If IsNumeric(loKey.KeyValue) Then
               sWeatherCode = ovrsExportWork.GetWeatherCode(CLng(loKey.KeyValue), sWeatherDescription)
               strTextBody = Replace(strTextBody, "[WEATHERDESCRIPTION]", sWeatherDescription)
               strTextBody = Replace(strTextBody, "[WEATHERCODE]", sWeatherCode)
            End If
         End If
         
         strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue)
                              
         ' Now look for all remaining occurrences of [KeyName xxx] where xxx is an optional format string
         ' Since we don't want to search for all possible format strings we have to do this in a
         ' two step process.
         ' First, look for the beginning of a possible match by looking for "[KeyName "
         ' Then look at what follows that to see if it is "format]"
         ' We have to consider mistakes and invalid formats too so check for a missing "]"
         ' and allow for "[Keyname BadFormat]"
         
         ' The problem with a missing "]" is we have to make sure we don't find the "]" on a subsequent variable
         ' Like this "[KeyOne [KeyTwo]", the "]" on the end is for KeyTwo, not KeyOne
         ' or "This is the [Keyname1 line [KeyName2]" we don't want to treat the "line [KeyName2]" portion as the format
         ' for the "[Keyname1" portion
         
         ' If we find an invalid format like [KeyName1 ThisIsABadFormatString] but there is a valid [KeyName1 and "]"
         ' then just ignore the bad format and treat it like [KeyName1]
         
         ' Remember we might have multiple occurrences of a KeyName reference so we have to keep searching until we
         ' don't find any more [KeyName occurrences, we can't just do it once and assume we found them all
         
         '  Replace(expression, find, replace[, start[, count[, compare]]])
         
         ' The Search string includes a trailing blank because we want to make sure the user didn't do something silly
         ' like have variables with similar names like KeyName, KeyName2 and KeyName3.
         ' If we just searched for "[KeyName" then we would find matches on "[KeyName2" and "[KeyName3" by mistake
         
         sSearch = "[" & UCase(loKey.KeyName) & " "
         nCurrPos = 1
         
         nCurrPos = InStr(nCurrPos, strTextBody, sSearch, vbTextCompare)
         
         Do While nCurrPos > 0
         
            nStartPos = nCurrPos
            
            ' Found "[KeyName " so see if we can find the trailing "]"
            
            nEndPos = InStr(nStartPos, strTextBody, "]", vbTextCompare)
            
            ' Did we find an ending "]"? If not just replace the beginning portion and move on
            If nEndPos < 1 Then
               
               ' Replace just the [KeyName portion and leave the rest alone
               
'               strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, nStartPos, 1)
               strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, 1, 1)
               nCurrPos = nStartPos + 1
            Else
               ' sKeyPlusFormat is now the whole matching string, everything from the "[" to the "]"
               sKeyPlusFormat = mID(strTextBody, nStartPos, (nEndPos - nStartPos) + 1)
            
               ' Is there a CRLf in there somewhere? If so then this is a bad format. Probably an extra CrLf in the format
               ' or a missing ] and we found the ] on the next keyname
               If InStr(1, sKeyPlusFormat, vbCrLf, vbTextCompare) > 0 Then
               
                  ' Replace just the [KeyName portion and leave the rest alone
                  
'                  strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, nStartPos, 1)
                  strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, 1, 1)
                  nCurrPos = nStartPos + 1
                  
               ' No CrLf found, see if there is a second "[" which might indicate a missing "]"
               ElseIf InStr(2, sKeyPlusFormat, "[", vbTextCompare) > 0 Then
   
                  ' We found two "[" characters, might be "[KeyName1 [KeyName2]
                  ' Just replace the [KeyName1 portion with the value
                  
'                  strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, nStartPos, 1)
                  strTextBody = Replace(strTextBody, sSearch, loKey.KeyValue, 1, 1)
                  nCurrPos = nStartPos + 1
               
               ' Ok, looks like we might have a valid "[KeyName Format]"
               Else
               
                  ' Break it up into pieces and see if we can format the data
               
                  ' Extract just the format string
                  sFormat = mID(sKeyPlusFormat, Len(sSearch), Len(sKeyPlusFormat) - Len(sSearch))
                  
                  sFormattedData = FormatKeyString(loKey.KeyValue, sFormat)
                  
                  strTextBody = Replace(strTextBody, sKeyPlusFormat, sFormattedData)
                  
                  nCurrPos = nEndPos ' Increment our Current Position to the end of this keyword pair
               End If
            End If
            nCurrPos = InStr(nCurrPos, strTextBody, sSearch, vbTextCompare)
         Loop
         
         
      Next loKey
      Set loRec = Nothing
   Next nRecNum
   
   ' Finally, replace all pipe symbols "|" with a CrLf. The pipe symbol is used in text fields sent via datagrams
   strTextBody = Replace(strTextBody, "|", vbCrLf)
   ConvertBasicDgText = strTextBody
   KillObject ovrsExportWork
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.ConvertBasicDgText. ", Err.Number, Err.Description
   ConvertBasicDgText = strTextBody
   KillObject ovrsExportWork
End Function

Private Function FormatKeyString(sInputData As String, sFormat As String) As String
   Dim nLongValue As Long
   Dim fDoubleValue As Double
   Dim dDateValue As Date
   Dim sModInputData As String
   Dim sF As String
   
   On Error GoTo FunctionError
   
   ' Set the return value to the input string incase there are any errors.
   ' If we get an error we'll just return the original input string without formatting
   
   FormatKeyString = sInputData
   
   sF = Trim(sFormat)
   If Len(sF) < 1 Then
      Exit Function
   End If
   
   ' In order to handle regional settings we have to convert the input string into an internal format
   ' like Long or Date so the Format function will be able to re-convert it using the regional settings
   ' If we try to convert the input string directly we'll get errors going from one format to another
   
   Select Case sF
   
      Case "MOL_LAT", "MOL_LONG"
   
         ' Special MOL Latitude/longitude display format. Replace blanks with "-"
         FormatKeyString = Replace(sInputData, " ", "-")
         Exit Function
      
      Case "General Date", "Long Date", "Medium Date", "Short Date"
      
         If IsDate(sInputData) Then
            dDateValue = CDate(sInputData)
            FormatKeyString = Format(dDateValue, sF)
            Exit Function
         End If
         
      Case "Long Time", "Medium Time", "Short Time"
      
         If IsDate(sInputData) Then
            dDateValue = CDate(sInputData)
            FormatKeyString = Format(dDateValue, sF)
            Exit Function
         End If
         
      Case "General Number"
      
         sModInputData = Replace(sInputData, ".", mLocalDecimalSeparator)
         If IsNumeric(sModInputData) Then
            If InStr(1, sInputData, ".") > 0 Then
               fDoubleValue = CDbl(sModInputData)
               FormatKeyString = Format(fDoubleValue, sF)
            Else
               nLongValue = CLng(sInputData)
               FormatKeyString = Format(nLongValue, sF)
            End If
            Exit Function
         End If
         
      Case "Currency"
      
         sModInputData = Replace(sInputData, ".", mLocalDecimalSeparator)
         If IsNumeric(sModInputData) Then
            fDoubleValue = CCur(sModInputData)
            FormatKeyString = Format(fDoubleValue, sF)
            Exit Function
         End If
      
      Case "Fixed", "Standard", "Percent", "Scientific"
      
         sModInputData = Replace(sInputData, ".", mLocalDecimalSeparator)
         If IsNumeric(sModInputData) Then
            fDoubleValue = CDbl(sModInputData)
            FormatKeyString = Format(fDoubleValue, sF)
            Exit Function
         End If
      
      Case "Yes/No", "True/False", "On/Off"
      
         If IsNumeric(sInputData) Then
            nLongValue = CLng(sInputData)
            FormatKeyString = Format(nLongValue, sF)
            Exit Function
         End If
   End Select
   
   ' If we get here then we don't have a "named format" so we don't know yet what type of data we have
   ' could be numeric, date, boolean, string, whatever
   
   ' Check for Numeric first
   sModInputData = Replace(sInputData, ".", mLocalDecimalSeparator)
   If IsNumeric(sModInputData) Then
      If InStr(1, sInputData, ".") > 0 Then
         fDoubleValue = CDbl(sModInputData)
         FormatKeyString = Format(fDoubleValue, sF)
      Else
         nLongValue = CLng(sInputData)
         FormatKeyString = Format(nLongValue, sF)
      End If
      If FormatKeyString = sF Then
         FormatKeyString = sInputData
      End If
      
      Exit Function
   End If
   
   ' Ok, how about Dates and Times?
   If IsDate(sInputData) Then
      dDateValue = CDate(sInputData)
      FormatKeyString = Format(dDateValue, sF)
      If FormatKeyString = sF Then
         FormatKeyString = sInputData
      End If
      Exit Function
   End If
   
   ' If we get here then it's not a number or a date so just try to format the input string
   
   FormatKeyString = Format(sInputData, sF)
   If FormatKeyString = sF Then
      FormatKeyString = sInputData
   End If
   
   Exit Function
FunctionError:

   ' No errors from this function, just return the input string
   FormatKeyString = sInputData
   Exit Function
End Function

'
' Human Readable VRS Message -
'
Private Function UpdateVrsTextBodyHeader(ByRef TextBody As String, Flag As String)
   On Error GoTo FunctionError
   If Flag = "LT" Then
      TextBody = TextBody & vbCrLf & _
        "B/L Lading     Dest Port  Bl# Product  Tmp Vsl Qty B/L Fig Vsl Fig Receiver   BL" & vbCrLf & _
        "--- ---------- ---------- --- -------- --- ------- ------- ------- ---------- -----" & vbCrLf
        '"B/L Lading     Dest Port  Bl# Product  Tmp Vsl Qty B/L Fig Vsl Fig Receiver   BL Load BL" & vbCrLf & _
        '"--- ---------- ---------- --- -------- --- ------- ------- ------- ---------- ------- -----" & vbCrLf
   ElseIf Flag = "SS" Then
      TextBody = TextBody & vbCrLf & _
        "Stop Date  Time  Start Date Time  Total Stop Reason" & vbCrLf & _
        "---------- ----- ---------- ----- ----- -----------" & vbCrLf
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.UpdateVrsTextBodyHeader. ", Err.Number, Err.Description
End Function

Private Function UpdateVrsTextBodyRow(ByRef TextBody As String, Flag As String)
   Dim strVrsShortDateFormat As String
   Dim strVrsGeneralDateFormat As String
   Dim strTemp As String
   Dim strOffset As String
   Dim sDateStart As String
   Dim sDateStop As String
   Dim i As Integer
   Dim iTotal As Long
   Dim loRec As ConfigKeys
   On Error GoTo FunctionError
   strVrsShortDateFormat = goSession.GetAppKey("vrs", "VrsShortDateFormat")
   strVrsGeneralDateFormat = goSession.GetAppKey("vrs", "VrsGeneralDateFormat")
   strTemp = ""
   If Flag = "LT" Then
      '
      ' Loop through 5 possible load time sets in Record_001...
      '
      Set loRec = moRecords(2).ConfigKeys
      For i = 1 To 5
         strOffset = "LOAD_BILL_NO" & Right(str(i), 1)
         If Trim(loRec.GetKeyValue(strOffset)) <> "" Then
            'Process line
            strTemp = strTemp & Format(loRec.GetKeyValue("LOAD_BILL_NO" & Right(str(i), 1)), "@@@") & " " & _
              Format(loRec.GetKeyValue("LOAD_DATE" & Right(str(i), 1)), strVrsShortDateFormat) & " " & _
              Format(loRec.GetKeyValue("LOAD_DEST_PORT" & Right(str(i), 1)), "@@@@@@@@@@") & " " & _
              Format(loRec.GetKeyValue("LOAD_BILL_NO_ORG" & Right(str(i), 1)), "000") & " " & _
              Format(Left(loRec.GetKeyValue("LOAD_PRODUCT" & Right(str(i), 1)), 8), "@@@@@@@@") & " " & _
              Format(loRec.GetKeyValue("LOAD_TEMP" & Right(str(i), 1)), "000") & " " & _
              Format(loRec.GetKeyValue("LOAD_CAR_VSL_NOM" & Right(str(i), 1)), "000,000") & " " & _
              Format(loRec.GetKeyValue("LOAD_CAR_BL_LOAD" & Right(str(i), 1)), "000,000") & " " & _
              Format(loRec.GetKeyValue("LOAD_CAR_VSL_LOAD" & Right(str(i), 1)), "000,000") & " " & _
              Format(Left(loRec.GetKeyValue("LOAD_RECEIVER" & Right(str(i), 1)), 10), "@@@@@@@@@@") & " " & _
              Format(Left(loRec.GetKeyValue("LOAD_CAR_BL_WORD" & Right(str(i), 1)), 5), "@@@@@") & vbCrLf
         End If
      Next i
      Set loRec = Nothing
   ElseIf Flag = "SS" Then
      For i = 3 To moRecords.Count
         Set loRec = moRecords(i).ConfigKeys
         If UCase(loRec.GetKeyValue("table")) = "VRS_POSDEPSTOPS" Then
            '
            ' Calculate Total Minutes
            '
            If IsDate(loRec.GetKeyValue("STOP_DATE")) And IsDate(loRec.GetKeyValue("START_DATE")) Then
               sDateStop = Format(loRec.GetKeyValue("STOP_DATE"), "MM/DD/YYYY hh:mm")
               sDateStart = Format(loRec.GetKeyValue("START_DATE"), "MM/DD/YYYY hh:mm")
               iTotal = DateDiff("n", sDateStop, sDateStart)
               
               strTemp = strTemp & Format(loRec.GetKeyValue("STOP_DATE"), strVrsGeneralDateFormat) & " " & _
                 Format(loRec.GetKeyValue("START_DATE"), strVrsGeneralDateFormat) & " " & _
                 Format(iTotal, "00000") & " " & _
                 loRec.GetKeyValue("STOP_REASON") & vbCrLf
            End If
         End If
         Set loRec = Nothing
      Next i
   End If
   TextBody = TextBody & strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.UpdateVrsTextBodyRow. ", Err.Number, Err.Description
End Function



Public Function CreateDgFromMappedRs(ByRef oRs As Recordset, EventTypeKey As Long, DatagramCode As String) As Boolean
   Dim loRsMap As Recordset
   Dim loRec As ConfigKeys
   Dim strSQL As String
   Dim strValue As String
   Dim strTargetTableName As String
   Dim strTargetConnectionID As String
   Dim iKey As Long
   Dim oField As ADODB.field
   Dim IsPrefixSymbol As Boolean
   Dim i As Integer
   Dim loKeys As Object
   Dim IsShowPlusSignInVRSDGMTimeZone As Boolean
   
   On Error GoTo FunctionError
   '
   ' Preliminary Staging/Validation
   '
   goSession.LogIt mwl_Workstation, mwl_Information, "Starting mwDataForm.CreateDgFromMappedRs, EventTypeKey: " & EventTypeKey & _
     ", TemplateID: " & moParent.TemplateProperties.TemplateID
   mEventType = EventTypeKey
   If moParent.IsOpen = False Or moParent.FormType <> mwft_DATA_FORM Then
      goSession.RaiseError "Error in mwDataForm.CreateDgFromMappedRs, Form not open or wrong form type."
      CreateDgFromMappedRs = False
      Exit Function
   End If
   '
   If oRs.State = adStateClosed Then
      goSession.RaiseError "Error in CreateDgFromMappedRs, Recordset is closed."
      CreateDgFromMappedRs = False
      Exit Function
   End If
   ' Any records in there ?
   If oRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataForm.CreateDgFromMappedRs, Recordset has no records."
      CreateDgFromMappedRs = False
      Exit Function
   End If
   
   Set loKeys = goSession.GetEventSecurityKeys(MW_EVENT_VRS_VOYEVENT)
   If loKeys.GetBoolKeyValue("IsShowPlusSignInVRSDGMTimeZone") Then
      IsShowPlusSignInVRSDGMTimeZone = True
   End If
   KillObject loKeys

   '
   ' Fetch Mapping information...
   '
   Set loRsMap = New Recordset
   loRsMap.CursorLocation = adUseClient
   strSQL = "select * from mwEventMapDgTable where mwEventTypeKey=" & EventTypeKey & _
     " and DatagramCode='" & DatagramCode & "' and not TemplateID is NULL"
   loRsMap.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsMap.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataForm.CreateDgFromMappedRs, No record found in mwEventMapDgTable for EventType: " & EventTypeKey & _
        ", DatagramCode: " & DatagramCode
      CreateDgFromMappedRs = False
      QuietCloseRecordset loRsMap
      Exit Function
   End If
   iKey = loRsMap!ID
   strTargetTableName = loRsMap!TargetTableName
   strTargetConnectionID = BlankNull(loRsMap!TargetConnectionID)
   loRsMap.Close
   strSQL = "select * from mwEventMapDgColumn where mwEventMapDgTableKey=" & iKey
   loRsMap.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsMap.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataForm.CreateDgFromMappedRs, No mapping records found in mwEventMapDgColumn."
      CreateDgFromMappedRs = False
      QuietCloseRecordset loRsMap
      Exit Function
   End If
   '
   ' Add EventTypeKey/DatagramCode to Header of datagram...
   '
   moRecords(1).ConfigKeys.Add "EventTypeKey", str(EventTypeKey), "EventTypeKey"
   moRecords(1).ConfigKeys.Add "DatagramCode", DatagramCode, "DatagramCode"
   
   '
   ' create the datagram records
   '
   oRs.MoveFirst
   Do While Not oRs.EOF
      '
      '
      loRsMap.MoveFirst
      Set loRec = AddRecord()
      loRec.Add "TABLE", strTargetTableName, "TABLE"
      loRec.Add "OPERATION", "A", "OPERATION"
      ' Loop through mapped records...
      On Error Resume Next
      Do While Not loRsMap.EOF
         'strValue = a
         If loRsMap.Fields("RsColumnName").value = "VRS_NGRADEBUNKERS" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_ROB" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_CONS" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_ROBCONS" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_ROB_ARR" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_CONS_ARR" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_ROB_DEP" Or _
            loRsMap.Fields("RsColumnName").value = "VRS_NGRADE_CONS_DEP" Then
         
            ' These are dummy VRS NGRADE variables, ignore them.
            
         Else
            If Not IsNull(oRs.Fields(loRsMap.Fields("RsColumnName").value).value) And Not IsEmpty(oRs.Fields(loRsMap.Fields("RsColumnName").value).value) Then
               If Err Then
                  ' Taking out gosession.logit, and it all goes to shit... 3/2005 ms
                  '
                  'If goSession.DebugMode Then
                  goSession.LogIt mwl_Workstation, mwl_Warning, "Warning in mwDataForm.CreateDgFromMappedRS, Expected Column missing: " & loRsMap.Fields("RsColumnName").value
                  'End If
               Else
                  '
                  ' Format field value for target database
                  '
                  IsPrefixSymbol = False
                  If IsShowPlusSignInVRSDGMTimeZone = True And UCase(loRsMap.Fields("RsColumnName").value) = "TIMEZONE" Then
                     IsPrefixSymbol = True
                  End If
                  If Trim(oRs.Fields(loRsMap.Fields("RsColumnName").value).value) <> "" Then
                     AddRecordField loRec, loRsMap!DgPairName, oRs.Fields(loRsMap.Fields("RsColumnName").value), IsPrefixSymbol
                  End If
               End If
            End If
         End If
         loRsMap.MoveNext
      Loop
      On Error GoTo FunctionError
      oRs.MoveNext
   Loop
   ' That's all folks
   goSession.CloseRecordset loRsMap
   CreateDgFromMappedRs = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.CreateDgFromMappedRs. ", Err.Number, Err.Description
   CreateDgFromMappedRs = False
   QuietCloseRecordset loRsMap
   Exit Function
MappingError:
   If Err.Number = 3265 Then
      goSession.RaiseError "Error in mwDataForm.CreateDgFromMappedRs, invalid column in source file: " & loRsMap.Fields("RsColumnName").value
   Else
      goSession.RaiseError "General Error in mwDataForm.CreateDgFromMappedRs. ", Err.Number, Err.Description
   End If
   CreateDgFromMappedRs = False
   QuietCloseRecordset loRsMap
   Exit Function

End Function

Private Function AddRecordField(ByRef loRec As ConfigKeys, FieldName As String, loField As field, Optional IsPrefixSymbol As Boolean)
   Dim strTemp As String
   On Error GoTo FunctionError
   Select Case loField.Type
      Case Is = adInteger        ' (3)
         loRec.Add FieldName, loField.value, FieldName
      Case Is = adDouble, adVarNumeric        ' (5)
         strTemp = Format(loField.value, "#########0.0##")
         ' Euro Symbol crap...
         ' 6 Feb 2005 change - datagrams use . for decimal...
         If goSession.API.GetLocaleString(LOCALE_DECIMAL_SEPARATOR) = "," Then
            strTemp = Replace(strTemp, ",", ".")
         End If
         If IsPrefixSymbol = True Then
            If Left(strTemp, 1) <> "-" Then
                strTemp = "+" & strTemp
            End If
         End If
         loRec.Add FieldName, strTemp, FieldName
      Case Is = adDate, adDBTimeStamp           ' (7) (135)
         If Not Format(loField.value, "yyyymmdd") < "19000101" Then
            'loRec.Add FieldName, Format(loField.value, mDgGeneralDateFormat), FieldName
            loRec.Add FieldName, MakeDatagramDate(loField.value), FieldName
         End If
      Case Is = adVarWChar       ' (202) Text
         loRec.Add FieldName, Replace(loField.value, vbCrLf, "|"), FieldName
      Case Is = adLongVarWChar   ' (203) - Memo Field
         loRec.Add FieldName, Replace(loField.value, vbCrLf, "|"), FieldName
      Case Is = adBoolean
         If loField.value Then
            loRec.Add FieldName, "1", FieldName
         Else
            loRec.Add FieldName, "0", FieldName
         End If
      Case Else
         loRec.Add FieldName, Replace(loField.value, vbCrLf, "|"), FieldName
   End Select
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.AddRecordField. ", Err.Number, Err.Description
End Function


Public Function SendDatagramEmail(Optional SubjectAbstract As String) As Boolean
   Dim strEmailBody As String
   Dim strSubject As String
   Dim loFormSession As Object
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim loRs As Recordset
   Dim strTemplate As String
   On Error GoTo FunctionError
   
   Set loRs = FetchEmailTemplateRs
   
   If IsRecordLoaded(loRs) Then
      loRs.MoveFirst
      Do While Not loRs.EOF
         If Not IsNull(loRs!EmailTemplatePath) Then
            strTemplate = goSession.Workflow.ScriptFolder & "\" & loRs!EmailTemplatePath
            moParent.SetExtendedProperty "ConvertDgToEmailTemplate", loRs!EmailTemplatePath, moParent.TemplateProperties.TemplateID
         End If
         mEmailAddress = BlankNull(loRs!EmailAddress)
         strEmailBody = ConvertToEmailTextBody1(strTemplate)
    
         If Trim(strEmailBody) = "" Then
            SendDatagramEmail = False
            Exit Function
         End If
    
         Set fso = New FileSystemObject
        
         goSession.LogIt mwl_User_Defined, mwl_Information, "Debug code: Start SendDatagramEmail: "
         strEmailBody = Replace(strEmailBody, "[VESSEL_NAME]", goSession.Site.SiteName)
         If SubjectAbstract = "" Then
            strSubject = moParent.TemplateProperties.TemplateID & " Message from " & goSession.Site.SiteName
         Else
            strSubject = SubjectAbstract
         End If
         '
         strSubject = Replace(SubjectAbstract, "[VESSEL_NAME]", goSession.Site.SiteName)
         strSubject = ConvertBasicDgText(strSubject)
         '
         ' Final Massaging...
         '
         strEmailBody = StripInsertChars(strEmailBody)
         strEmailBody = StripUnusedEmailFields(strEmailBody)
         '
         '
         ' If SendByMedia or SMTP, then
         '
         ' Write the file to .txt in working folder, and launch using goSesiioon.API.LaunchExternalApp...
         '
         ' else
         '
         ' Put in draft mode of email system and send the wee little bugger on its way...
         '
         If goSession.User.DefaultTransport = mwt_Media Or goSession.User.DefaultEmailCarrier = mw_SMTP Then
         
            Set ts = fso.OpenTextFile(goSession.Workflow.WorkingFolder & "\" & moParent.TemplateProperties.TemplateID & ".txt", ForWriting, True)
            ts.Write strEmailBody
            ts.Close
           
            goSession.API.RunShellExecute goSession.Workflow.WorkingFolder & "\" & moParent.TemplateProperties.TemplateID & ".txt"
            SendDatagramEmail = False
            ' you are manually sending by Cut/Past Text message.
         Else
    
       With goSession.Message
            .UserResolveAddress = True
          .Body = strEmailBody
          .AttachmentFile = ""
          .subject = strSubject
          .SendToAddress = mEmailAddress
          .Send
       End With
       goSession.Message.UserResolveAddress = False
       
       SendDatagramEmail = True
       End If
       
       loRs.MoveNext
    Loop
   
   End If
   KillObject ts
   KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.SendDatagramEmail. ", Err.Number, Err.Description
   KillObject ts
   KillObject fso
End Function

Public Function FetchTargetRegionalSettings(ConnectionID As String) As Boolean
   'Dim loRsDB As Recordset
   'Dim strSQL As String
   On Error GoTo FunctionError
   
   mDgGeneralDateFormat = DATAGRAM_DATE_FORMAT
   mDgShortDateFormat = DATAGRAM_SHORT_DATE_FORMAT
   mDgDecimalSeparator = DATAGRAM_DECIMAL_FORMAT
   
   FetchTargetRegionalSettings = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.FetchTargetRegionalSettings. ", Err.Number, Err.Description
   'CloseRecordset loRsDB
   FetchTargetRegionalSettings = False
End Function

Private Function OpenEventColumnMap(TargetTable As String) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   Dim iPK As Long
   On Error GoTo FunctionError
   ' Get PK...
   strSQL = "select ID from mwEventMapDgTable where mwEventTypeKey=" & mEventTypeKey & _
     " and DatagramCode='" & mDatagramCode & "' and TargetTableName='" & TargetTable & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      'goSession.RaiseError "Error in mwDataForm.OpenEventColumnMap, event datagram has no mapping record for target table: " & TargetTable
      OpenEventColumnMap = False
   Else
      iPK = loRs!ID
      If moRsEventColumnMap Is Nothing Then
         Set moRsEventColumnMap = New Recordset
         moRsEventColumnMap.CursorLocation = adUseClient
      Else
         If moRsEventColumnMap.State = adStateOpen Then
            moRsEventColumnMap.Close
         End If
      End If
      strSQL = "select * from mwEventMapDgColumn where mwEventMapDgTableKey=" & iPK
      moRsEventColumnMap.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If moRsEventColumnMap.RecordCount < 1 Then
         goSession.RaiseError "Error in mwDataForm.OpenEventColumnMap, event datagram " & str(mEventTypeKey) & _
           "/" & mDatagramCode & ", has no mapped columns for target table: " & TargetTable
         OpenEventColumnMap = False
      Else
         OpenEventColumnMap = True
      End If
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.OpenEventColumnMap. ", Err.Number, Err.Description
   OpenEventColumnMap = False
End Function

Private Function IsEventColumnDiscard(ColumnName As String) As Boolean
   On Error GoTo FunctionError
   IsEventColumnDiscard = False
   If moRsEventColumnMap Is Nothing Then
      Exit Function
   End If
   moRsEventColumnMap.MoveFirst
   Do While Not moRsEventColumnMap.EOF
      'Debug.Print moRsEventColumnMap!DgPairName
      If Trim(UCase(moRsEventColumnMap!DgPairName)) = Trim(UCase(ColumnName)) Then
         If moRsEventColumnMap!IsDiscardOnUpdateAsRs Then
            IsEventColumnDiscard = True
            Exit Do
         Else
            IsEventColumnDiscard = False
            Exit Do
         End If
      End If
      moRsEventColumnMap.MoveNext
   Loop
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.IsEventColumnDiscard. ", Err.Number, Err.Description
   IsEventColumnDiscard = False
End Function

Private Function IsEventColumnNumeric(ColumnName As String) As Boolean
   On Error GoTo FunctionError
   moRsEventColumnMap.Filter = adFilterNone
   moRsEventColumnMap.Filter = "RsColumnName='" & ColumnName & "'"
   If moRsEventColumnMap.RecordCount < 1 Then
      IsEventColumnNumeric = False
   Else
      If moRsEventColumnMap!IsNumericOnUpdateAsRs Then
         IsEventColumnNumeric = True
      Else
         IsEventColumnNumeric = False
      End If
   End If
   moRsEventColumnMap.Filter = adFilterNone
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.IsEventColumnNumeric. ", Err.Number, Err.Description
   IsEventColumnNumeric = False
End Function

Public Function MakeDatagramDate(RawDate As String, Optional IsShortDate As Boolean) As String
   Static IsErrored As Boolean
   Dim s As String
   On Error GoTo FunctionError
   s = Format(RawDate, "YYYY-MM-DD hh:nn:ss")
   
   MakeDatagramDate = s
   Exit Function
FunctionError:
   If Not IsErrored Then
      goSession.RaiseError "General Error in mwDataForm.MakeDatagramDate. ", Err.Number, Err.Description
      IsErrored = True
   End If
   MakeDatagramDate = RawDate
End Function


Private Function RollbackTransactions() As Boolean
   On Error GoTo FunctionError
   moCon.RollbackTrans
   goSession.RaiseWorkflowStatus "Error Occurred, rolling back transactions..."
   RollbackTransactions = True
   Exit Function
FunctionError:
   'goSession.RaiseError "General Error in mwDataForm.RollbackTransactions. ", err.Number, err.Description
   RollbackTransactions = True
End Function

Private Function StripInsertChars(str As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   strTemp = Replace(str, "'", " ")
   'strTemp = Replace(strTemp, "|", " ")
   strTemp = Replace(strTemp, "|", vbCrLf)
   If Len(Trim(strTemp)) = 0 Then
      strTemp = str
   End If
   StripInsertChars = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in modWFDProcessing.StripInsertChars: ", Err.Number, Err.Description
   StripInsertChars = str
End Function

Private Function StripUnusedEmailFields(str As String) As String
   Dim iStart As Integer
   Dim iEnd As Integer
   Dim str1 As String
   Dim str2 As String
   On Error GoTo FunctionError
   ' infinite loop with these chars below...
   str = Replace(str, "[[", "[")
   str = Replace(str, "]]", "]")
   iStart = InStr(str, "[")
   Do While iStart > 0
      iEnd = InStr(iStart, str, "]")
      If iEnd < 1 Then
         Exit Do
      End If
      str1 = mID(str, 1, iStart - 1)
      str2 = mID(str, iEnd + 1)
      str = str1 & str2
      iStart = InStr(iStart, str, "[")
   Loop
   StripUnusedEmailFields = str
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in modWFDProcessing.StripUnusedEmailFields: ", Err.Number, Err.Description
   StripUnusedEmailFields = str
End Function


Public Function InsertRecord(loRec As ConfigKeys, ByRef oRs As Recordset, IsMakePK As Boolean) As Boolean
   Dim strTable As String
   Dim mwrChangeLogKey As Long
   
   mwrChangeLogKey = 0
   
   On Error GoTo FunctionError
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   strTable = loRec.GetKeyValue("table")
   If Not oRs.State = adStateOpen Then
      If moCon Is Nothing Then
         Set moCon = goCon
      End If
      oRs.Open strTable, moCon, adOpenStatic, adLockOptimistic, adCmdTable
   End If
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.InsertRecord, Unable to open Table: " & strTable, Err.Number, Err.Description
      QuietCloseRecordset oRs
      InsertRecord = False
      Exit Function
   End If
   oRs.AddNew
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.InsertRecord, Unable to AddNew, Table: " & strTable
      QuietCloseRecordset oRs
      InsertRecord = False
      Exit Function
   End If
   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
      oRs!ID = goSession.MakePK(strTable)
      oRs.Update
      If Err Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.InsertRecord: ", Err.Number, Err.Description
         goSession.RaiseWorkflowStatus "Failed to Update Table during Call to MakePK: " & strTable
         QuietCloseRecordset oRs
         InsertRecord = False
         Exit Function
      End If
   End If
   If UpdateRSFields(loRec, oRs, mwrChangeLogKey) Then
      '
      ' If Child, save key to the parent record...
      '
      If loRec.GetKeyValue("IsChildKeyName") <> "" Then
         oRs.Fields(loRec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
      End If
      '
      ' Update new record...
      '
      oRs.Update
      If Err Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.InsertRecord: ", Err.Number, Err.Description
         goSession.RaiseWorkflowStatus "Failed to Update Table: " & strTable
         QuietCloseRecordset oRs
         InsertRecord = False
         Exit Function
      End If
      ' If parent record, save parent key for future use
      If loRec.GetKeyValue("IsParentKeyName") <> "" Then
         mCurrentParentRecordID = oRs.Fields(loRec.GetKeyValue("IsParentKeyName")).value
      End If
   Else
      oRs.CancelUpdate
      goSession.RaiseWorkflowStatus "Error in mwDataForm.InsertRecord. CancelUpdate on Form: " & moParent.FormProperties.FormID
      InsertRecord = False
      QuietCloseRecordset oRs
      InsertRecord = False
      Exit Function
   End If
   InsertRecord = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.InsertRecord.", Err.Number, Err.Description
   InsertRecord = False
End Function


Private Function ExecuteDllOnDatagramEmail() As String
   Dim strVbDLL As String
   Dim loFormSession As Object
   On Error GoTo FunctionError
   '
   ' Run external VB DLL - if it exists...
   '
   strVbDLL = moParent.WorkflowProperties.VbFormEventDLL & ".FormSession"
   Set loFormSession = CreateObject(strVbDLL)
   Set loFormSession.ActiveMwSession = goSession
   Set loFormSession.ActiveMwForm = moParent
   loFormSession.TriggerFormEvent ("OnDatagramEmail")
   If loFormSession.IsCancel Then
      '
      ' Cancel the process
      '
      'goSession.RaiseWorkflowStatus "OnDatagramEmail DLL Code has cancelled workflow processing for: " & moParent.FormProperties.FormID
      ExecuteDllOnDatagramEmail = ""
   Else
      ExecuteDllOnDatagramEmail = loFormSession.GetEmailTextBody
   End If
   '
   Set loFormSession = Nothing

   Exit Function
FunctionError:
   If Err.Number = 429 Or Err.Number = 13 Then
      goSession.RaiseError "Error in mwDataForm.ExecuteDllOnDatagramEmail, unable to launch DLL: " & strVbDLL
   Else
      goSession.RaiseError "General Error in mwDataForm.ExecuteDllOnDatagramEmail, Dll: " & strVbDLL, Err.Number, Err.Description
   End If
   ExecuteDllOnDatagramEmail = ""
End Function

Private Function ExecuteDllOnWriteDatagram() As Boolean
   Dim strVbDLL As String
   Dim loFormSession As Object
   On Error GoTo FunctionError
   '
   ' Run external VB DLL - if it exists...(last minute changes to datagram records...)
   '
   strVbDLL = moParent.WorkflowProperties.VbFormEventDLL & ".FormSession"
   Set loFormSession = CreateObject(strVbDLL)
   Set loFormSession.ActiveMwSession = goSession
   Set loFormSession.ActiveMwForm = moParent
   loFormSession.TriggerFormEvent ("OnWriteDatagram")
   If loFormSession.IsCancel Then
      '
      ' Cancel the process
      '
      goSession.RaiseWorkflowStatus "OnDatagramEmail DLL Code has cancelled workflow processing for: " & moParent.FormProperties.FormID
      ExecuteDllOnWriteDatagram = False
   Else
      ExecuteDllOnWriteDatagram = True
   End If
   '
   Set loFormSession = Nothing

   Exit Function
FunctionError:
   If Err.Number = 429 Or Err.Number = 13 Then
      goSession.RaiseError "Error in mwDataForm.ExecuteDllOnWriteDatagram, unable to launch DLL: " & strVbDLL
   Else
      goSession.RaiseError "General Error in mwDataForm.ExecuteDllOnWriteDatagram, Dll: " & strVbDLL, Err.Number, Err.Description
   End If
   ExecuteDllOnWriteDatagram = ""
End Function


Public Function SendCustomDatagramEmail(EmailTemplate As String, EmailSubject As String) As Boolean
   Dim strEmailBody As String
   Dim strSubject As String
   Dim sSiteName As String
   Dim loFormSession As Object
   On Error GoTo FunctionError
   '
   sSiteName = goSession.Site.GetShipProperty(moParent.FormProperties.SiteID, "SiteName")
   strEmailBody = ConvertToEmailTextBody(EmailTemplate, True)
   strEmailBody = Replace(strEmailBody, "[VESSEL_NAME]", sSiteName)
   EmailSubject = Replace(EmailSubject, "[VESSEL_NAME]", sSiteName)
   '
   ' Final Massaging...
   '
   strEmailBody = StripInsertChars(strEmailBody)
   strEmailBody = StripUnusedEmailFields(strEmailBody)
   '
   ' Put in draft mode of email system and send the wee little bugger on its way...
   '
   With goSession.Message
      .UserResolveAddress = False
      .Body = strEmailBody
      .AttachmentFile = ""
      .subject = EmailSubject
      .SendToAddress = moParent.GetExtendedProperty("ConvertDgToEmailAddress")
      .Send
   End With
   goSession.RaiseWorkflowStatus "Custom datagram sent to: " & moParent.GetExtendedProperty("ConvertDgToEmailAddress")
   SendCustomDatagramEmail = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.SendCustomDatagramEmail. ", Err.Number, Err.Description
End Function



Private Function UpdateAsModifyRecord(loRec As ConfigKeys, mwrChangeLogKey As Long, strOperation As String, Optional AuditTrailHeaderKey As Long) As Boolean
   Dim sTable As String
   Dim sSQL As String
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
'   moRS.CursorLocation = adUseServer
   sTable = loRec.GetKeyValue("table")
   sSQL = "select * from " & sTable & " where ID=" & loRec.GetKeyValue("ID")
   moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsModifyRecord, Unable to execute Select: " & sSQL
      QuietCloseRecordset moRS
      UpdateAsModifyRecord = False
      Exit Function
   End If
   If moRS.RecordCount < 1 Then
      If strOperation = "U" Then
         moRS.AddNew
         moRS!ID = loRec.GetKeyValue("ID")
      Else
         goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsModifyRecord, Record is missing (Table/Key): " & sTable & "/" & loRec.GetKeyValue("ID")
         QuietCloseRecordset moRS
         UpdateAsModifyRecord = False
         Exit Function
      End If
   End If
   If UpdateRSFields(loRec, moRS, mwrChangeLogKey, IS_MODIFY, AuditTrailHeaderKey) Then
      '
      ' Update changed record
      '
      moRS.Update
      If Err Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsModifyRecord: ", Err.Number, Err.Description
         goSession.RaiseWorkflowStatus "Failed to Update Table: " & sTable
         QuietCloseRecordset moRS
         UpdateAsModifyRecord = False
         Exit Function
      End If
   Else
      moRS.CancelUpdate
      goSession.RaiseWorkflowStatus "Error in mwDataForm.UpdateAsModifyRecord. CancelUpdate on Form: " & moParent.FormProperties.FormID
      UpdateAsModifyRecord = False
      QuietCloseRecordset moRS
      UpdateAsModifyRecord = False
      Exit Function
   End If
   UpdateAsModifyRecord = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateAsModifyRecord.", Err.Number, Err.Description
   QuietCloseRecordset moRS
   UpdateAsModifyRecord = False
End Function

Private Function UpdateAsDeletedRecord(loRec As ConfigKeys, mwrChangeLogKey As Long) As Boolean
   Dim sTable As String
   Dim sSQL As String
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
'   moRS.CursorLocation = adUseServer
   sTable = loRec.GetKeyValue("table")
   sSQL = "select * from " & sTable & " where ID=" & loRec.GetKeyValue("ID")
   moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsDeletedRecord, Unable to execute Select: " & sSQL
      moRS.CancelUpdate
      QuietCloseRecordset moRS
      UpdateAsDeletedRecord = False
      Exit Function
   End If
   moRS.Delete
   If Err Then
      goSession.RaiseWorkflowStatus "Error in mwSession.mwDataForm.UpdateAsDeletedRecord: ", Err.Number, Err.Description
      goSession.RaiseWorkflowStatus "Failed to Update Table: " & sTable
      moRS.CancelUpdate
      QuietCloseRecordset moRS
      UpdateAsDeletedRecord = False
      Exit Function
   End If
   UpdateAsDeletedRecord = True
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.UpdateAsDeletedRecord.", Err.Number, Err.Description
   UpdateAsDeletedRecord = False
End Function

Public Function ConvertEventType(NewEventType As Long) As Boolean
   Dim sTemp As String
   Dim i As Integer
   Dim loRec As ConfigGroup
   Dim loKeys As ConfigKeys
   On Error GoTo FunctionError
   If moRecords.Count < 1 Then
      goSession.RaiseError "Error in mwSession.mwDataForm.ConvertEventType, no records to convert."
   End If
   For Each loRec In moRecords
      Set loKeys = loRec.ConfigKeys
      sTemp = loKeys.GetKeyValue("MWEVENTTYPEKEY")
      If sTemp <> "" Then
         loKeys.Remove "mwEventTypeKey"
         loKeys.Add "mwEventTypeKey", str(NewEventType), "mwEventTypeKey"
      End If
   Next loRec
   ConvertEventType = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataForm.ConvertEventType.", Err.Number, Err.Description
   ConvertEventType = False
End Function

Private Function CheckIsRedistributableTable(TargetTable As String) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ' Only shore redistributes...
   If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
      CheckIsRedistributableTable = False
      Exit Function
   End If
   
   strSQL = "select mwrBatchTypeKey from mwrChangeTable where mwrBatchTypeKey = 2 AND TableName='" & TargetTable & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      CheckIsRedistributableTable = False
   Else
      CheckIsRedistributableTable = True
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.CheckIsRedistributableTable. ", Err.Number, Err.Description
   CheckIsRedistributableTable = False
   QuietCloseRecordset loRs
End Function

Private Function ChangeTableKey(TargetTable As String) As Long
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   strSQL = "select ID from mwrChangeTable where TableName='" & TargetTable & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      ChangeTableKey = -1
   Else
      ChangeTableKey = loRs!ID
   End If
   goSession.CloseRecordset loRs
   KillObject loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.ChangeTableKey. ", Err.Number, Err.Description
   ChangeTableKey = -1
   QuietCloseRecordset loRs
   KillObject loRs
End Function

Private Function WriteChangeHeader(ChangeType As String, ChangeTableKey As Long, RecordKey As Long, OriginatingSiteKey As Long, Optional SiteKey As Long, Optional IsLoggedTransaction As Boolean) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   If ChangeTableKey > 0 Then
      Set loRs = New Recordset
      With loRs
         If goSession.IsSqlServer Then
            .CursorLocation = adUseClient
         Else
            .CursorLocation = adUseServer
         End If
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
         .Fields("IsLoggedTransaction").value = IsLoggedTransaction
         .Update
         If goSession.IsOracle() Then
            .Resync adAffectCurrent, adResyncAllValues
         End If
         WriteChangeHeader = .Fields("ID").value
      End With
      goSession.CloseRecordset loRs
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.WriteChangeHeader. ", Err.Number, Err.Description
'   Resume Next
   WriteChangeHeader = -1
   QuietCloseRecordset loRs
End Function

Private Function WriteChangeColumn(ChangeKey, RecField As field) As Boolean
   Dim sSQL1 As String
   Dim sSQL2 As String
   Dim sDate As String
   Dim sField As String
   Dim i As Integer
   Static loCmd As Command
   
   On Error GoTo FunctionError
   
   If ChangeKey > 0 Then
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
         loCmd.Parameters("NV").value = Null
         loCmd.Parameters("NVL").value = Null
         loCmd.Parameters("OVL").value = Null
         loCmd.Parameters("CN").value = Null
         loCmd.Parameters("OV").value = Null
      End If
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
   Else
      WriteChangeColumn = False
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataForm.WriteChangeColumn. ", Err.Number, Err.Description
   WriteChangeColumn = False
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
   goSession.RaiseError "General Error in mwDataForm.MassageField: ", Err.Number, Err.Description
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
   goSession.RaiseError "General Error in mwDataForm.FixString.", Err.Number, Err.Description
   FixString = strVal
End Function

Private Function QuietCloseRecordset(ByRef oRs As Recordset)
   On Error Resume Next
   If Not oRs Is Nothing Then
      If oRs.State = adStateOpen Then
         If oRs.LockType <> adLockReadOnly Then
            If Not (oRs.EOF Or oRs.BOF) Then
               'oRs.Update
               If oRs.RecordCount > 0 Then
                  oRs.Move (0)
               End If
            End If
         End If
         oRs.Close
      End If
      Set oRs = Nothing
   End If
   Exit Function
End Function

Private Function ConvertCustomVrsNGrades(TextBody As String) As String
   Dim sTextBody As String
   Dim sTemplateID As String
   Dim nDetailKey As Long
   Dim nDetailType As Integer
   
   On Error GoTo FunctionError
   
   sTextBody = TextBody
   
   sTemplateID = UCase(Left(moRecords(1).ConfigKeys("TEMPLATEID").KeyValue, 10))
   nDetailKey = 0
   
   If sTemplateID = "VRS_POSDEP" Then

      If Not moRecords(2).ConfigKeys("PORTCALLKEY") Is Nothing Then
         nDetailKey = moRecords(2).ConfigKeys("PORTCALLKEY").KeyValue
         nDetailType = DT_PORTCALL_DEPART
      End If
      
   ElseIf sTemplateID = "VRS_POSARR" Then
   
      If Not moRecords(2).ConfigKeys("PORTCALLKEY") Is Nothing Then
         nDetailKey = moRecords(2).ConfigKeys("PORTCALLKEY").KeyValue
         nDetailType = DT_PORTCALL_ARRIVAL
      End If
   
   ElseIf sTemplateID = "VRS_POSASS" Or sTemplateID = "VRS_POSPAS" Then
   
      If Not moRecords(2).ConfigKeys("POSASSKEY") Is Nothing Then
         nDetailKey = moRecords(2).ConfigKeys("POSASSKEY").KeyValue
         nDetailType = DT_ATSEA
      End If
   
   End If
   
   
   If nDetailKey > 0 Then
   
      FetchNGradeBunkers nDetailKey, nDetailType
      If nDetailType = DT_PORTCALL_DEPART Then
         FetchNGradeBunkersNew nDetailKey, DT_PORTCALL_ARRIVAL
      ElseIf nDetailType = DT_PORTCALL_ARRIVAL Then
         FetchNGradeBunkersNew nDetailKey, DT_PORTCALL_DEPART
      End If
      
      '---------  pbt-203
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADEBUNKERS]", nDetailType)
      sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADEBUNKERS ", nDetailType)
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB]", nDetailType)
      sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_ROB ", nDetailType)
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS]", nDetailType)
      sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_CONS ", nDetailType)
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_TAKEN]", nDetailType)
      sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_TAKEN ", nDetailType)
      
      If nDetailType = DT_PORTCALL_DEPART Then
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB_ARR]", DT_PORTCALL_ARRIVAL)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_ROB_ARR ", DT_PORTCALL_ARRIVAL)
        
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS_ARR]", DT_PORTCALL_ARRIVAL)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_CONS_ARR ", DT_PORTCALL_ARRIVAL)
        
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_TAKEN_ARR]", DT_PORTCALL_ARRIVAL)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_TAKEN_ARR ", DT_PORTCALL_ARRIVAL)
         
      ElseIf nDetailType = DT_PORTCALL_ARRIVAL Then
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB_DEP]", DT_PORTCALL_DEPART)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_ROB_DEP ", DT_PORTCALL_DEPART)
        
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS_DEP]", DT_PORTCALL_DEPART)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_CONS_DEP ", DT_PORTCALL_DEPART)
        
         sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_TAKEN_DEP]", DT_PORTCALL_DEPART)
         sTextBody = ProcessNGradeBunkersNew(sTextBody, "[VRS_NGRADE_TAKEN_DEP ", DT_PORTCALL_DEPART)
      End If
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROBCONS]", nDetailType)
      
   
      CloseNGradeBunkers
   End If
   
   ConvertCustomVrsNGrades = sTextBody
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwDataForm.ConvertCustomVrsNGrades. ", Err.Number, Err.Description
   ConvertCustomVrsNGrades = TextBody
End Function


Private Sub FetchNGradeBunkers(IdKey As Long, DetailTypeEnum As Integer)
   Dim lorsB As Recordset
   Dim sSQL As String
   Dim sNGradeBunkers As String
   Dim nDetailType As Long
   Dim loform As frmEmailBunkers
   Dim IdKey2 As Long
   
   Dim nGrp As Long
   Dim nRcd As Long
   Dim IsAddRecord As Boolean
   Dim nSortId As Long
   Dim Fld As field
   
   On Error GoTo FunctionError
   
   
   'Validate
   If IdKey < 1 Or (DetailTypeEnum < 1 Or DetailTypeEnum > 3) Then Exit Sub
   ' DetailTypeEnum determines which ngradebunkers group to get.
'   Const DT_ATSEA = 1
'   Const DT_PORTCALL_ARRIVAL = 2
'   Const DT_PORTCALL_DEPART = 3
   
   ' replace with inspection of vrs   (GET STRUCTURE FOR FAB RS ONLY)
   sSQL = "SELECT vrsBunker.ID AS IDKey, DetailType, DetailKey, vrsBunkPurposeKey, " _
      & " vrsBunkGrade.Description AS BgDescription, vrsBunkPurpose.Description AS BpDescription, " _
      & " vrsBunker.vrsBunkGradeKey, UnitType, Quantity, SortID " _
      & " FROM vrsBunker, vrsBunkPurpose, vrsBunkGrade " _
      & " Where vrsBunker.DetailKey = " & IdKey & " And vrsBunker.DetailType = " & DetailTypeEnum _
      & " AND vrsBunkGrade.ID = vrsBunker.vrsBunkGradeKey " _
      & " AND vrsBunkPurpose.ID = vrsBunker.vrsBunkPurposeKey "

   sSQL = sSQL & " ORDER BY SortID, vrsBunkGrade.Description, vrsBunker.vrsBunkPurposeKey "
   
   Set lorsB = New Recordset
   lorsB.CursorLocation = adUseClient
   lorsB.Open sSQL, goCon, adOpenKeyset, adLockOptimistic
   
   'copy to fabricated Rs for modification  JUST copy structure
   FabricateRS lorsB
            
   ' get data from vrsXXMasterXXX.dgm (vessel lightweight datagram NOT standard replication tables)
   nGrp = 1            ' reset group counter
   Do While (nGrp) < moRecords.Count                                    ' config group loop

      'Debug.Print moRecords(nGrp).ConfigGroup
        
      If UCase(Left(moRecords.Item(nGrp).ConfigGroup, 6)) = "RECORD" Then   ' avoid HEADER
         nRcd = 1    ' reset record counter
         
         Do While (nRcd) <= moRecords.Item(nGrp).ConfigKeys.Count  '  records loop
            
            ' look for VRS_POSARR_BUNK
            'Debug.Print moRecords.Item(nGrp).ConfigKeys.Item(nRcd).KeyName
                
            If moRecords.Item(nGrp).ConfigKeys.Item(nRcd).KeyName = "TABLE" Then
                        
               'Debug.Print moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue & " nGradeBunker"
                    
               If moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue = "VRS_POSARR_BUNK" Or _
                  moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue = "VRS_POSDEP_BUNK" Or _
                  moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue = "VRS_POSASS_BUNK" Then
                  
                  moNGradeRS.AddNew
                  IsAddRecord = True      ' flag to prep fields
               End If
            End If
            
            If IsAddRecord Then
               ' keyname
               'Debug.Print "Record " & moRecords.Item(nGrp).ConfigKeys(nRcd).KeyName
                    
               If moRecords.Item(nGrp).ConfigKeys(nRcd).KeyName = "GRADE" Then
               
                  moNGradeRS!BgDescription = moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue
                  
                  If Trim(moNGradeRS!BgDescription) <> "" Then
                     
                     moNGradeRS!UnitType = GetUnitType(moNGradeRS!BgDescription, nSortId)
                     moNGradeRS!SortID = nSortId
                  End If
                            
               End If
               
               If moRecords.Item(nGrp).ConfigKeys(nRcd).KeyName = "QUANTITY" Then
               
                  If Not IsNull(moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue) Then
                     If moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue = "[NULL]" Then
                        moNGradeRS!Quantity = Null
                     Else
                        If IsNumeric(Replace(moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue, ".", mLocalDecimalSeparator)) Then
                           moNGradeRS!Quantity = CDbl(Replace(moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue, ".", mLocalDecimalSeparator))
                        End If
                     End If
                  Else
                     moNGradeRS!Quantity = Null
                  End If
               End If
               
               If moRecords.Item(nGrp).ConfigKeys(nRcd).KeyName = "PURPOSE" Then
                  
                  moNGradeRS!vrsBunkPurposeKey = moRecords.Item(nGrp).ConfigKeys(nRcd).KeyValue
                  
                  If moNGradeRS!vrsBunkPurposeKey > 0 Then
                     
                     moNGradeRS!BpDescription = GetBunkerPurposeDescription(moNGradeRS!vrsBunkPurposeKey)
                  End If
               End If
               moNGradeRS!DetailType = DetailTypeEnum
            End If
            
            nRcd = nRcd + 1
         Loop                    ' record group loop
            
         If IsAddRecord Then
            moNGradeRS.Update
            IsAddRecord = False
         End If
      End If

      nGrp = nGrp + 1

   Loop                ' Config group loop
      
   CloseRecordset lorsB
      
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.FetchNGradeBunkers ", Err.Number, Err.Description
   CloseRecordset lorsB
End Sub

Private Sub CloseNGradeBunkers()
   
   On Error GoTo SubError
   
   CloseRecordset moNGradeRS
   KillObject moNGradeRS
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.CloseNGradeBunkers ", Err.Number, Err.Description
   
End Sub

Private Function ProcessNGradeBunkers(TextBody As String, DGVariable As String, nDetailType As Integer) As String
   Dim sTextBody As String
   Dim sShipCode As String
   Dim sNGradeBunkers As String
   Dim sPortCallKey As String
   Dim sAtSeaPerfKey As String

   Const EtPORTCALL = 15
   Const EtAtSea = 3
   On Error GoTo FunctionError
   
   sTextBody = TextBody
   sNGradeBunkers = ""
   
'   sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADEBUNKERS]")
'   sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB]")
'   sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS]")
'   sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROBCONS]")

    ' limit to group type
   If moNGradeRS.RecordCount > 0 And InStr(1, sTextBody, DGVariable) > 0 Then
      moNGradeRS.Sort = "SortID, BgDescription, vrsBunkPurposeKey"
      
      If DGVariable = "[VRS_NGRADEBUNKERS]" Then
          ' Show everything
         moNGradeRS.Filter = "DetailType = " & nDetailType & ""
         sNGradeBunkers = DisplayNGradeBunkers
         
      ElseIf DGVariable = "[VRS_NGRADE_ROB]" Or DGVariable = "[VRS_NGRADE_ROB_ARR]" Or DGVariable = "[VRS_NGRADE_ROB_DEP]" Then
          ' Show just ROB
         moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey = 1"
         sNGradeBunkers = DisplayNGradeBunkers
          
      ElseIf DGVariable = "[VRS_NGRADE_CONS]" Or DGVariable = "[VRS_NGRADE_CONS_ARR]" Or DGVariable = "[VRS_NGRADE_CONS_DEP]" Then
          ' Show everything except ROB
         moNGradeRS.Filter = "DetailType = " & nDetailType & " and  vrsBunkPurposeKey > 1"
         sNGradeBunkers = DisplayNGradeBunkers
          
      ElseIf DGVariable = "[VRS_NGRADE_ROBCONS]" Then
          
         moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey = 1"
         sNGradeBunkers = DisplayNGradeBunkers
              
         moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey > 1"
         sNGradeBunkers = sNGradeBunkers & DisplayNGradeBunkers
      
      ElseIf DGVariable = "[VRS_NGRADE_TAKEN]" Or DGVariable = "[VRS_NGRADE_TAKEN_ARR]" Or DGVariable = "[VRS_NGRADE_TAKEN_DEP]" Then         '---PBT-109
            ' Show just bunker Taken
           moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey = 6"
           sNGradeBunkers = DisplayNGradeBunkers
      End If
      
      moNGradeRS.Filter = adFilterNone
    End If
            
            
   ' replace keyword in variated textbody
   If sNGradeBunkers <> "" Then
      
      sNGradeBunkers = vbCrLf & "Grade    Purpose                  Quantity Units" & vbCrLf & sNGradeBunkers
      
      If goSession.User.DefaultTransport <> mwt_Media And goSession.User.DefaultEmailCarrier <> mw_SMTP Then
         
         sNGradeBunkers = Replace(sNGradeBunkers, " ", "&nbsp;")
         sNGradeBunkers = "<html><body><font face=" & """" & "Courier New" & """" & " size=" & """" & "2" & """" & "> " & sNGradeBunkers
         sNGradeBunkers = sNGradeBunkers & "</font></body></html>"
         
      End If
      
      ProcessNGradeBunkers = Replace(TextBody, DGVariable, sNGradeBunkers)
      
   Else
      ProcessNGradeBunkers = TextBody
   End If
 
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwDataForm.ProcessNGradeBunkers. ", Err.Number, Err.Description
   ProcessNGradeBunkers = TextBody
End Function

Private Function ProcessNGradeBunkersNew(TextBody As String, DGVariable As String, nDetailType As Integer) As String
    Dim sTextBody As String
    Dim sNGradeBunkers As String
    Dim TAG As String
    Dim FMTType As String
    Dim NumFMT As String
    Dim IsHideEmptyGrade As Boolean
    Dim Separator As String
    Dim IsHideHeader As Boolean
    
    On Error GoTo FunctionError
    
    sTextBody = TextBody
    sNGradeBunkers = ""
   
    TAG = ParseBunkerTAG(TextBody, DGVariable, FMTType, NumFMT, IsHideEmptyGrade, Separator, IsHideHeader)
            
            
    If FMTType <> "" Then
        ' limit to group type
        If moNGradeRS.RecordCount > 0 And InStr(1, sTextBody, DGVariable) > 0 Then
            moNGradeRS.Sort = "SortID, BgDescription, vrsBunkPurposeKey"
            
            If DGVariable = "[VRS_NGRADEBUNKERS " Then
                ' Show everything
                moNGradeRS.Filter = "DetailType = " & nDetailType & ""
                sNGradeBunkers = DisplayNGradeBunkersCONS_ALL(FMTType, NumFMT, IsHideEmptyGrade, Separator, IsHideHeader)
            ElseIf DGVariable = "[VRS_NGRADE_ROB " Or DGVariable = "[VRS_NGRADE_ROB_ARR " Or DGVariable = "[VRS_NGRADE_ROB_DEP " Then
                ' Show just ROB
                moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey = " & BUNKER_PURPOSE_ROB
                sNGradeBunkers = DisplayNGradeBunkersROB_TAKEN(FMTType, NumFMT, IsHideEmptyGrade, Separator, IsHideHeader)
            ElseIf DGVariable = "[VRS_NGRADE_TAKEN " Or DGVariable = "[VRS_NGRADE_TAKEN_ARR " Or DGVariable = "[VRS_NGRADE_TAKEN_DEP " Then
                ' Show just bunker Taken
                moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey = " & BUNKER_PURPOSE_TAKEN
                sNGradeBunkers = DisplayNGradeBunkersROB_TAKEN(FMTType, NumFMT, IsHideEmptyGrade, Separator, IsHideHeader)
            ElseIf DGVariable = "[VRS_NGRADE_CONS " Or DGVariable = "[VRS_NGRADE_CONS_ARR " Or DGVariable = "[VRS_NGRADE_CONS_DEP " Then
                ' Show everything except ROB
                moNGradeRS.Filter = "DetailType = " & nDetailType & " and vrsBunkPurposeKey > " & BUNKER_PURPOSE_ROB
                sNGradeBunkers = DisplayNGradeBunkersCONS_ALL(FMTType, NumFMT, IsHideEmptyGrade, Separator, IsHideHeader)
            End If
            
            moNGradeRS.Filter = adFilterNone
        End If
    End If
      
    ' replace keyword in variated textbody
    If sNGradeBunkers <> "" Then
        ProcessNGradeBunkersNew = Replace(TextBody, TAG, sNGradeBunkers)
    Else
        ProcessNGradeBunkersNew = TextBody
    End If
    
    Exit Function
FunctionError:
    goSession.RaiseError "General Error in mwSession.mwDataForm.ProcessNGradeBunkersNew. ", Err.Number, Err.Description
    ProcessNGradeBunkersNew = TextBody
End Function

Private Function DisplayNGradeBunkers() As String
   Dim sNGradeBunkers As String
   Dim tempStr1 As String
   Dim tempStr2 As String
   Dim tempStr3 As String
   Dim tempStr4 As String
   
   On Error GoTo FunctionError
   
   ' format ngradebunkers for email (send)
   
   If moNGradeRS.RecordCount > 0 Then
      moNGradeRS.MoveFirst
   End If
   
   Do While moNGradeRS.EOF = False
         tempStr1 = moNGradeRS!BgDescription & "        " ' 8
         tempStr2 = moNGradeRS!BpDescription & "                    " ' 20
         tempStr3 = "            " & moNGradeRS!Quantity  ' 12
         tempStr4 = moNGradeRS!UnitType & "   "      ' 3
         
         sNGradeBunkers = sNGradeBunkers & Left(tempStr1, 8)
         sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr2, 20)
         sNGradeBunkers = sNGradeBunkers & " " & Right(tempStr3, 12)
         sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr4, 3)
         sNGradeBunkers = sNGradeBunkers & vbCrLf
      moNGradeRS.MoveNext
   Loop
         
   DisplayNGradeBunkers = sNGradeBunkers
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.DisplayNGradeBunkers ", Err.Number, Err.Description
End Function

Private Sub FabricateRS(Rs As Recordset)
   Dim loField As field
   Dim ii As Integer
   Dim IsNewField As Boolean
   Dim NewFName As String
   On Error GoTo FunctionError
      
   Set moNGradeRS = New Recordset
   moNGradeRS.CursorLocation = adUseClient
   
   ' fabricate rs with index # appended for duplicate names
   For Each loField In Rs.Fields
      If loField.Type <> adChapter Then
         IsNewField = False
         ' replace table.fieldname with X
         If InStr(1, loField.Name, ".") > 1 Then
            IsNewField = True
            NewFName = loField.Name
            NewFName = Replace(NewFName, ".", "X")
         End If
      
         If IsNewField Then
            moNGradeRS.Fields.Append NewFName, loField.Type, loField.DefinedSize, loField.Attributes And adFldIsNullable
         Else
            moNGradeRS.Fields.Append loField.Name, loField.Type, loField.DefinedSize, loField.Attributes And adFldIsNullable
         End If
         If IsNewField Then
            moNGradeRS(NewFName).Precision = loField.Precision
            moNGradeRS(NewFName).NumericScale = loField.NumericScale
         Else
            moNGradeRS(loField.Name).Precision = loField.Precision
            moNGradeRS(loField.Name).NumericScale = loField.NumericScale
         End If
         ii = ii + 1
      End If
   Next loField
   
   moNGradeRS.Open
   
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.FabricateRS", Err.Number, Err.Description
End Sub

Private Function GetUnitType(BunkerGradeDescription As String, Optional SortID As Long, Optional IdKey As Long) As String
    Dim sSQL As String
    Dim loRs As Recordset
    On Error GoTo FunctionError
    
    If Trim(BunkerGradeDescription) = "" Then
        Exit Function
    End If
    
    sSQL = "SELECT * FROM vrsBunkGrade WHERE Description = '" & BunkerGradeDescription & "'"
    Set loRs = New Recordset
    loRs.CursorLocation = adUseClient
    loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
    
    If IsRecordLoaded(loRs) Then
        GetUnitType = BlankNull(loRs!UnitType)
        SortID = ZeroNull(loRs!SortID)
        IdKey = loRs!ID
    End If
    CloseRecordset loRs
    
    Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.GetUnitType", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function GetBunkerPurposeDescription(BpKey As Long) As String
    Dim sSQL As String
    Dim loRs As Recordset
    On Error GoTo FunctionError
    
    If BpKey < 1 Then
        Exit Function
    End If
    
    sSQL = "SELECT * FROM vrsBunkPurpose WHERE ID = " & BpKey
    Set loRs = New Recordset
    loRs.CursorLocation = adUseClient
    loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
    
    If IsRecordLoaded(loRs) Then
        GetBunkerPurposeDescription = BlankNull(loRs!Description)
    End If

    CloseRecordset loRs
    
    Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.GetBunkerPurposeDescription", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Function ProcessLogRequest(UserKey As Long, IsZipAttachment As Long, SendToEmail As String) As Boolean
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim sBody As String
   Dim sUserLogFileName As String
   On Error GoTo FunctionError
   
   'fetch User log file
   sUserLogFileName = FetchUserLogName(UserKey)
   If SendToEmail = "" Then
      goSession.LogIt mwl_User_Defined, mwl_Information, "Send to Email missing: " & SendToEmail
      Exit Function
   End If
   'check log file
   sBody = "Requested log file being sent to Original Site"
   
   Set fso = New FileSystemObject
   If Not fso.FileExists(sUserLogFileName) Then
      ' must have a file to respond (even if empty)
      Set ts = fso.OpenTextFile(sUserLogFileName, ForWriting, True)
      ts.WriteLine "Log file Established at this time " & Now()
      ts.Close
      Set fso = Nothing
   End If
   '
   ' ok, let's go... Send away...
   '
   If IsZipAttachment > 0 Then
      goSession.Workflow.IsZipAttachmentFile = True
   End If
   If goSession.Workflow.SendAnyFile(sUserLogFileName, SendToEmail, , sBody) Then
      goSession.LogIt mwl_User_Defined, mwl_Information, "Log File have been sent to: " & SendToEmail
      ProcessLogRequest = True
   End If
   
   KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataForm.ProcessLogRequest. ", Err.Number, Err.Description
   KillObject fso
   ProcessLogRequest = False
End Function

Private Function FetchUserLogName(UserKey) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   If UserKey < 1 Then
      Exit Function
   End If
   
   sSQL = "SELECT LogFile FROM mwcUsers WHERE ID = " & UserKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      FetchUserLogName = BlankNull(loRs!LogFile)
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataForm.FetchUserLogName. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function


Private Function ParseBunkerTAG(TextBody As String, DGVariable As String, FMTType As String, NumFMT As String, IsHideEmptyGrade As Boolean, Separator As String, IsHideHeader As Boolean) As String
    Dim PosStart As Long
    Dim PosEnd As Long
    Dim sTAG As String
    Dim sEmptyGrade As String
    Dim sHideHeader As String

    On Error GoTo FunctionError:
    
    If InStr(1, TextBody, DGVariable, vbTextCompare) Then
        PosStart = InStr(1, TextBody, DGVariable, vbTextCompare)
        PosEnd = InStr(PosStart, TextBody, "]", vbTextCompare)
        sTAG = mID(TextBody, PosStart, PosEnd + 1 - PosStart)
        
        '1. format type
        FMTType = ParseBunkerParameter(sTAG, BUNKER_PARAM_FMT)
        FMTType = UCase(FMTType)
        
        '2. Number format
        NumFMT = ParseBunkerParameter(sTAG, BUNKER_PARAM_NUMBER)
         
        '3. Hide Empty Grade
        sEmptyGrade = ParseBunkerParameter(sTAG, BUNKER_PARAM_GRADE)
        If sEmptyGrade <> "" Then
            If UCase(sEmptyGrade) = "Y" Then
                IsHideEmptyGrade = True
            Else
                IsHideEmptyGrade = False
            End If
        Else
            IsHideEmptyGrade = True
        End If
         
        '4. Separator
        Separator = ParseBunkerParameter(sTAG, BUNKER_PARAM_SEPARATOR)
        
        '5. Hide Header
        sHideHeader = ParseBunkerParameter(sTAG, BUNKER_PARAM_HEADER)
        If sHideHeader <> "" Then
            If UCase(sHideHeader) = "Y" Then
                IsHideHeader = True
            Else
                IsHideHeader = False
            End If
        Else
            IsHideHeader = False
        End If
    End If
    
    ParseBunkerTAG = sTAG
    Exit Function
FunctionError:
    goSession.RaiseError "General Error in mwSession.mwDataForm.ParseBunkerTAG. ", Err.Number, Err.Description
End Function

Private Function GetParameterValue(ParamString As String) As String
    Dim sValue As String
    On Error GoTo FunctionError
    
    sValue = Replace(ParamString, BUNKER_PARAM_FMT, "")
    sValue = Replace(sValue, BUNKER_PARAM_NUMBER, "")
    sValue = Replace(sValue, BUNKER_PARAM_GRADE, "")
    sValue = Replace(sValue, BUNKER_PARAM_SEPARATOR, "")
    sValue = Replace(sValue, BUNKER_PARAM_HEADER, "")
    GetParameterValue = Trim(sValue)
    Exit Function
FunctionError:
    goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataForm.GetParameterValue. ", Err.Number, Err.Description
End Function

Private Function ParseBunkerParameter(TAG As String, ParamString As String) As String
    Dim PosStart As Long
    Dim Pos1 As Long
    Dim Pos2 As Long
    Dim strParam As String
    Dim Parameter As String
    On Error GoTo FunctionError

    PosStart = InStr(1, TAG, ParamString, vbTextCompare)
        
    If PosStart > 0 Then
        Pos1 = InStr(PosStart, TAG, "=", vbTextCompare)  'skip 1st '='
        If Pos1 > 0 Then
            Pos2 = InStr(Pos1 + 1, TAG, "=", vbTextCompare)
            If Pos2 > 0 Then
                strParam = mID(TAG, Pos1 + 1, Pos2 - (Pos1 + 1))
                Parameter = GetParameterValue(strParam)
            Else
                Pos2 = InStr(Pos1 + 1, TAG, "]", vbTextCompare)
                If Pos2 > 0 Then
                    strParam = mID(TAG, Pos1 + 1, Pos2 - (Pos1 + 1))
                    Parameter = GetParameterValue(strParam)
                Else
                    Parameter = ""
                End If
            End If
        Else
            Parameter = ""
        End If
    Else
        Parameter = ""
    End If
    
    ParseBunkerParameter = Parameter
    Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwSession.mwDataForm.ParseBunkerParameter. ", Err.Number, Err.Description
End Function

Private Function DisplayNGradeBunkersROB_TAKEN(FMTType As String, NumFMT As String, IsHideEmptyGrade As Boolean, Separator As String, IsHideHeader As Boolean) As String
    Dim sNGradeBunkers As String
    Dim tempStr1 As String
    Dim tempStr2 As String
    Dim tempStr3 As String
    Dim tempStr4 As String
    Dim strBunker As String
    
    On Error GoTo FunctionError
    
    If moNGradeRS.RecordCount > 0 Then
        moNGradeRS.MoveFirst
    End If

    Select Case FMTType
    
    Case "A":
    
        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription & "        " ' 8
            tempStr2 = moNGradeRS!BpDescription & "                    " ' 20
            tempStr3 = "            " & Format(moNGradeRS!Quantity, NumFMT)  ' 12
            tempStr4 = moNGradeRS!UnitType & "   "      ' 3
            
            sNGradeBunkers = sNGradeBunkers & Left(tempStr1, 8)
            sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr2, 20)
            sNGradeBunkers = sNGradeBunkers & " " & Right(tempStr3, 12)
            sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr4, 3)
            sNGradeBunkers = sNGradeBunkers & vbCrLf
        
            moNGradeRS.MoveNext
        Loop
        
        If sNGradeBunkers <> "" Then
            sNGradeBunkers = vbCrLf & sNGradeBunkers
        End If
        
    Case "B", "C":
        
        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription
            tempStr2 = moNGradeRS!BpDescription
            tempStr3 = Format(moNGradeRS!Quantity, NumFMT)
            tempStr4 = moNGradeRS!UnitType
            
            If strBunker <> "" Then
                strBunker = strBunker & " " & Separator
            End If
            
            If FMTType = "B" Then
                strBunker = strBunker & " " & tempStr3 & " " & tempStr4 & " " & tempStr1
            Else
                strBunker = strBunker & " " & tempStr1 & " " & tempStr3 & " " & tempStr4
            End If
            moNGradeRS.MoveNext
        Loop
        
        If strBunker <> "" Then
            sNGradeBunkers = sNGradeBunkers & strBunker & vbCrLf
        End If

    Case "D", "E":
       
        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription
            tempStr2 = moNGradeRS!BpDescription
            tempStr3 = Format(moNGradeRS!Quantity, NumFMT)
            tempStr4 = moNGradeRS!UnitType
            
            If FMTType = "D" Then
                strBunker = tempStr1 & " " & tempStr3 & " " & tempStr4
            Else
                strBunker = tempStr3 & " " & tempStr4 & " " & tempStr1
            End If
            
            sNGradeBunkers = sNGradeBunkers & strBunker & vbCrLf
            moNGradeRS.MoveNext
        Loop
        
        If sNGradeBunkers <> "" Then
            sNGradeBunkers = vbCrLf & sNGradeBunkers
        End If
    
    End Select
    DisplayNGradeBunkersROB_TAKEN = sNGradeBunkers
    Exit Function
FunctionError:
    goSession.RaisePublicError "General Error in mwSession.mwDataForm.DisplayNGradeBunkersROB_TAKEN ", Err.Number, Err.Description
End Function

Private Function DisplayNGradeBunkersCONS_ALL(FMTType As String, NumFMT As String, IsHideEmptyGrade As Boolean, Separator As String, IsHideHeader As Boolean) As String
    Dim sNGradeBunkers As String
    Dim tempStr1 As String
    Dim tempStr2 As String
    Dim tempStr3 As String
    Dim tempStr4 As String
    Dim strBunker As String
    Dim nPurposeKey As Long
    
    On Error GoTo FunctionError
    
    If moNGradeRS.RecordCount > 0 Then
        moNGradeRS.MoveFirst
    End If
   
    Select Case FMTType
    
    Case "A":
    
        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription & "        " ' 8
            tempStr2 = moNGradeRS!BpDescription & "                    " ' 20
            tempStr3 = "            " & Format(moNGradeRS!Quantity, NumFMT)  ' 12
            tempStr4 = moNGradeRS!UnitType & "   "      ' 3
            
            sNGradeBunkers = sNGradeBunkers & Left(tempStr1, 8)
            sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr2, 20)
            sNGradeBunkers = sNGradeBunkers & " " & Right(tempStr3, 12)
            sNGradeBunkers = sNGradeBunkers & " " & Left(tempStr4, 3)
            sNGradeBunkers = sNGradeBunkers & vbCrLf
        
        moNGradeRS.MoveNext
        Loop
        
        If sNGradeBunkers <> "" Then
            sNGradeBunkers = vbCrLf & sNGradeBunkers
        End If
       
    Case "B", "C":

        moNGradeRS.Sort = "vrsBunkPurposeKey,SortID,BgDescription"
        strBunker = ""
        nPurposeKey = -1
        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription
            tempStr2 = moNGradeRS!BpDescription
            tempStr3 = Format(moNGradeRS!Quantity, NumFMT)
            tempStr4 = moNGradeRS!UnitType
            
            If moNGradeRS!vrsBunkPurposeKey <> nPurposeKey Then
                If strBunker <> "" Then
                    sNGradeBunkers = sNGradeBunkers & strBunker & vbCrLf
                End If
                strBunker = tempStr2 & ":"
            Else
                strBunker = strBunker & " " & Separator
            End If
            
            If FMTType = "B" Then
                strBunker = strBunker & " " & tempStr3 & " " & tempStr4 & " " & tempStr1
            Else
                strBunker = strBunker & " " & tempStr1 & " " & tempStr3 & " " & tempStr4
            End If
    
            nPurposeKey = moNGradeRS!vrsBunkPurposeKey
            moNGradeRS.MoveNext
        Loop
       
        If strBunker <> "" Then
            sNGradeBunkers = sNGradeBunkers & strBunker & vbCrLf
        End If
        
        If sNGradeBunkers <> "" Then
            sNGradeBunkers = vbCrLf & sNGradeBunkers
        End If

    Case "D", "E":
        
        moNGradeRS.Sort = "vrsBunkPurposeKey,SortID,BgDescription"

        Do While moNGradeRS.EOF = False
            tempStr1 = moNGradeRS!BgDescription
            tempStr2 = moNGradeRS!BpDescription
            tempStr3 = Format(moNGradeRS!Quantity, NumFMT)
            tempStr4 = moNGradeRS!UnitType
            
            If FMTType = "D" Then
                strBunker = tempStr2 & ":" & " " & tempStr1 & " " & tempStr3 & " " & tempStr4
            Else
                strBunker = tempStr3 & " " & tempStr4 & " " & tempStr1 & " " & tempStr2
            End If
            
            sNGradeBunkers = sNGradeBunkers & strBunker & vbCrLf
            
            moNGradeRS.MoveNext
        Loop
        
            If sNGradeBunkers <> "" Then
                sNGradeBunkers = vbCrLf & sNGradeBunkers
            End If

    End Select
    
    DisplayNGradeBunkersCONS_ALL = sNGradeBunkers
    Exit Function
FunctionError:
    goSession.RaisePublicError "General Error in mwSession.mwDataForm.DisplayNGradeBunkersCONS_ALL ", Err.Number, Err.Description
End Function

Private Function CheckIsTempRecordInsertAllowed(TargetTable As String, RecordKey As Long) As Boolean
   Dim sSQL As String
   Dim sSQL1 As String
   Dim loRs As Recordset
   Dim loRs1 As Recordset

   On Error GoTo FunctionError
   
   sSQL = "select IsTempRecordInsertAllowed from mwrChangeTable where TableName='" & TargetTable & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If IsRecordLoaded(loRs) Then
      If BoolNull(loRs!IsTempRecordInsertAllowed) Then
      
         sSQL1 = "select IsTempRecordInsert from " & TargetTable & " where ID = " & RecordKey
         Set loRs1 = New Recordset
         loRs1.CursorLocation = adUseClient
         loRs1.Open sSQL1, goCon, adOpenForwardOnly, adLockReadOnly
         If IsRecordLoaded(loRs1) Then
            If BoolNull(loRs1!IsTempRecordInsert) Then
               CheckIsTempRecordInsertAllowed = True
            End If
         End If
         
      End If
   End If
   CloseRecordset loRs
   CloseRecordset loRs1
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwDataForm.CheckIsTempRecordInsertAllowed. ", Err.Number, Err.Description
   CheckIsTempRecordInsertAllowed = False
   CloseRecordset loRs
   CloseRecordset loRs1
End Function

Private Sub FetchNGradeBunkersNew(IdKey As Long, DetailTypeEnum As Integer)
   Dim lorsB As Recordset
   Dim sSQL As String
   Dim nSortId As Long
   
   On Error GoTo FunctionError
   
   
   If IdKey < 1 Or (DetailTypeEnum < 2 Or DetailTypeEnum > 3) Then
      Exit Sub
   End If
   
   sSQL = "SELECT vrsBunker.ID AS IDKey, DetailType, DetailKey, vrsBunkPurposeKey, " _
      & " vrsBunkGrade.Description AS BgDescription, vrsBunkPurpose.Description AS BpDescription, " _
      & " vrsBunker.vrsBunkGradeKey, UnitType, Quantity, SortID " _
      & " FROM vrsBunker, vrsBunkPurpose, vrsBunkGrade " _
      & " Where vrsBunker.DetailKey = " & IdKey & " And vrsBunker.DetailType = " & DetailTypeEnum _
      & " AND vrsBunkGrade.ID = vrsBunker.vrsBunkGradeKey " _
      & " AND vrsBunkPurpose.ID = vrsBunker.vrsBunkPurposeKey "

   sSQL = sSQL & " ORDER BY SortID, vrsBunkGrade.Description, vrsBunker.vrsBunkPurposeKey "
   
   Set lorsB = New Recordset
   lorsB.CursorLocation = adUseClient
   lorsB.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
            
   If IsRecordLoaded(lorsB) Then
      Do While lorsB.EOF = False
      
         moNGradeRS.AddNew
         
         If Not IsNull(lorsB!BgDescription) Then
            moNGradeRS!BgDescription = lorsB!BgDescription
            
            moNGradeRS!UnitType = GetUnitType(moNGradeRS!BgDescription, nSortId)
            moNGradeRS!SortID = nSortId
         End If
         If Not IsNull(lorsB!Quantity) Then
            moNGradeRS!Quantity = CDbl(Replace(lorsB!Quantity, ".", mLocalDecimalSeparator))
         End If
         If Not IsNull(lorsB!vrsBunkPurposeKey) Then
            moNGradeRS!vrsBunkPurposeKey = lorsB!vrsBunkPurposeKey
            moNGradeRS!BpDescription = GetBunkerPurposeDescription(moNGradeRS!vrsBunkPurposeKey)
         End If
         moNGradeRS!DetailType = DetailTypeEnum
         
         moNGradeRS.Update
                     
         lorsB.MoveNext
      Loop
   End If
            
   CloseRecordset lorsB
      
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwDataForm.FetchNGradeBunkersNew ", Err.Number, Err.Description
End Sub






