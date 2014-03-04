VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwWorkflow Class -
' 9/3/2001 ms


Option Explicit

Private mIDKey As Long
Private mWorkflowID As String
Private mDescription As String
Private mPollingFolder As String
Private mSendByMediaFolder As String
Private mScriptFolder As String
Private mCommonScriptFilename As String
Private mPollingIntervalMinutes As Long
Private mIsDebugMode As Boolean
Private mIsDebugMemory As Boolean
Private mIsSilentErrors As Boolean
Private mIsWorkflowError As Boolean
Private mIsWorkflowOperationError As Boolean
Private mIsNotifyPkgReceipt As String
Private mIsRunOnceMode As Boolean
Private mIsShowAllDatagramErrors As Boolean
Private mMaxErrorsBeforeQuit As Long
Private mMaxErrorsBeforeNotify As Long
Private mWorkingFolder As String
Private mWorkingSessionFolder As String
Private mArchiveOkZipFile As String
Private mArchiveErrorZipFile As String
Private mIsPost0392 As Boolean
Private mIsLoadWorkflow As Boolean
'
Private mInboundQueueFolder As String
Private mSendByCdromFolder As String
Private mIsZipWfFiles As Boolean
Private mIsZipEncryptWfFiles As Boolean
Private mIsSendDataChangesAllowed As Boolean
Private mIsSendBlobChangesAllowed As Boolean
Private mIsSendErrorAsAlert As Boolean
Private mIsSchedulerAllowed As Boolean
Private mRoleTypeKeyAlert As Long

Private mStartBlackout As Date
Private mBlackoutMinutes As Long
Private mIsResendRequested As Boolean
Private mIsDeferAckUntilMoreData As Boolean
Private mIsManualArchiveAllowed As Boolean      '---PBT-109
Private mIsStopUpdateDBOnError As Boolean       '---PBT-1567
Private mIsHideSystemTray As Boolean
Private mIsNotifyManualRevisionUpdate As Boolean       '---PBT-1567
Private mIsAutoSendFCChangesAllowed As Boolean     'PBT 1742
Private mIsSendDAWithInThresholdLimit   As Boolean
Private mIsStripOffExtendedAsciiChars As Boolean
'
Private mIsProcessOfficeForms As Boolean

' 2/2002 - New Intrinsic properties...
Private mIsInboxAllowed As Boolean
Private mIsFormsProcessingAllowed As Boolean
Private mIsPackageAllowed As Boolean
Private mIsEmailRouterAllowed As Boolean
Private mIsEventQueueAllowed As Boolean
Private mIsManualMode As Boolean

Private mTemplateWorkFolder As String
Private mTemplatePlacesID As Long
Private mSignedFormsPlaceID As Long

Const FLEET_TYPE_FORM_TEMPLATE_EMAILS = 1

'Working Objects
Private moParent As Session
Private moPackage As mwPackage
Private moExtendedKeys As ConfigKeys

Private moRecords As ConfigGroups
Dim WithEvents moRsFormHeader As Recordset
Attribute moRsFormHeader.VB_VarHelpID = -1

Private mLocalDecimalSeparator As String
   Private mIsZipAttachmentFile As Boolean   ' SendAnyFile to compress?

Enum mw_ArchiveStatus
   ArchivedOK = 1
   ArchiveError = 2
   ArchiveQueue = 3
End Enum

Private moNGradeRS As Recordset

'nGradeBunkers enum
   Const DT_ATSEA = 1
   Const DT_PORTCALL_ARRIVAL = 2
   Const DT_PORTCALL_DEPART = 3

   
Const SEND_CDROM_TRUE = True
Const SEND_CDROM_FALSE = False

Dim moKeys As ConfigKeys
Dim moFso As FileSystemObject


'
' Workflow Status Event
'
Event Status(Msg As String)

Friend Property Set Parent(oParent As mwSession.Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Friend Property Get Parent() As mwSession.Session
   'If moParent Is Nothing Then
   '   Set moParent = oParent
   'End If
   Set Parent = moParent
End Property

Public Property Get ID() As Long
    ID = mIDKey
End Property

Public Property Let ID(ByVal vData As Long)
    mIDKey = vData
End Property

Public Property Get WorkflowID() As String
    WorkflowID = mWorkflowID
End Property

Public Property Let WorkflowID(ByVal vData As String)
    mWorkflowID = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get PollingFolder() As String
   PollingFolder = mPollingFolder
   MakeFolderExist PollingFolder
End Property

Public Property Let PollingFolder(ByVal vData As String)
    mPollingFolder = vData
End Property

Public Property Get ScriptFolder() As String
   ScriptFolder = mScriptFolder
   MakeFolderExist ScriptFolder
End Property

Public Property Let ScriptFolder(ByVal vData As String)
    mScriptFolder = vData
End Property

Public Property Get CommonScriptFilename() As String
    CommonScriptFilename = mCommonScriptFilename
End Property

Public Property Let CommonScriptFilename(ByVal vData As String)
    mCommonScriptFilename = vData
End Property

Public Property Get PollingIntervalMinutes() As String
    PollingIntervalMinutes = mPollingIntervalMinutes
End Property

Public Property Let PollingIntervalMinutes(ByVal vData As String)
    mPollingIntervalMinutes = vData
End Property

Public Property Get IsDebugMode() As Boolean
    IsDebugMode = mIsDebugMode
End Property

Public Property Let IsDebugMode(ByVal vData As Boolean)
    mIsDebugMode = vData
End Property

Public Property Get IsDebugMemory() As Boolean
    IsDebugMemory = mIsDebugMemory
End Property

Public Property Let IsDebugMemory(ByVal vData As Boolean)
    mIsDebugMemory = vData
End Property

Public Property Get IsSilentErrors() As Boolean
    IsSilentErrors = mIsSilentErrors
End Property

Public Property Let IsSilentErrors(ByVal vData As Boolean)
    mIsSilentErrors = vData
End Property

Public Property Get IsNotifyPkgReceipt() As Boolean
    IsNotifyPkgReceipt = mIsNotifyPkgReceipt
End Property

Public Property Let IsNotifyPkgReceipt(ByVal vData As Boolean)
    mIsNotifyPkgReceipt = vData
End Property

Public Property Get IsRunOnceMode() As Boolean
    IsRunOnceMode = mIsRunOnceMode
End Property

Public Property Let IsRunOnceMode(ByVal vData As Boolean)
    mIsRunOnceMode = vData
End Property

Public Property Get IsShowAllDatagramErrors() As Boolean
    IsShowAllDatagramErrors = mIsShowAllDatagramErrors
End Property

Public Property Let IsShowAllDatagramErrors(ByVal vData As Boolean)
    mIsShowAllDatagramErrors = vData
End Property

Public Property Get MaxErrorsBeforeQuit() As Long
    MaxErrorsBeforeQuit = mMaxErrorsBeforeQuit
End Property

Public Property Let MaxErrorsBeforeQuit(ByVal vData As Long)
    mMaxErrorsBeforeQuit = vData
End Property

Public Property Get MaxErrorsBeforeNotify() As String
    MaxErrorsBeforeNotify = mMaxErrorsBeforeNotify
End Property

Public Property Let MaxErrorsBeforeNotify(ByVal vData As String)
    mMaxErrorsBeforeNotify = vData
End Property

Public Property Get WorkingFolder() As String
   WorkingFolder = mWorkingFolder
   MakeFolderExist WorkingFolder
End Property

Public Property Let WorkingFolder(ByVal vData As String)
    mWorkingFolder = vData
End Property

Public Property Get ArchiveOkZipFile() As String
    ArchiveOkZipFile = mArchiveOkZipFile
End Property

Public Property Let ArchiveOkZipFile(ByVal vData As String)
    mArchiveOkZipFile = vData
End Property

Public Property Get ArchiveErrorZipFile() As String
    ArchiveErrorZipFile = mArchiveErrorZipFile
End Property

Public Property Let ArchiveErrorZipFile(ByVal vData As String)
    mArchiveErrorZipFile = vData
End Property

Public Property Get WorkingSessionFolder() As String
   WorkingSessionFolder = mWorkingSessionFolder
   MakeFolderExist WorkingSessionFolder
End Property

Public Property Let WorkingSessionFolder(ByVal vData As String)
    mWorkingSessionFolder = vData
End Property

Public Property Get IsFormsProcessingAllowed() As Boolean
    IsFormsProcessingAllowed = mIsFormsProcessingAllowed
End Property

Public Property Let IsFormsProcessingAllowed(ByVal vData As Boolean)
    mIsFormsProcessingAllowed = vData
End Property

Public Property Get IsWorkflowError() As Boolean
    IsWorkflowError = mIsWorkflowError
End Property

Public Property Let IsWorkflowError(ByVal vData As Boolean)
    mIsWorkflowError = vData
End Property

Public Property Get SendByMediaFolder() As String
   SendByMediaFolder = mSendByMediaFolder
   MakeFolderExist SendByMediaFolder
End Property

Public Property Let SendByMediaFolder(ByVal vData As String)
    mSendByMediaFolder = vData
    UpdateUserData "SendByMediaFolder", mSendByMediaFolder
End Property

Public Property Get IsInboxAllowed() As Boolean
    IsInboxAllowed = mIsInboxAllowed
End Property

Public Property Let IsWorkflowOperationError(ByVal vData As Boolean)
    mIsWorkflowOperationError = vData
End Property

Public Property Get IsWorkflowOperationError() As Boolean
    IsWorkflowOperationError = mIsWorkflowOperationError
End Property


Public Property Get InboundQueueFolder() As String
   InboundQueueFolder = mInboundQueueFolder
   MakeFolderExist InboundQueueFolder
End Property

Public Property Get SendByCdromFolder() As String
   SendByCdromFolder = mSendByCdromFolder
   MakeFolderExist SendByCdromFolder
End Property

Public Property Get IsZipWfFiles() As Boolean
    IsZipWfFiles = mIsZipWfFiles
End Property

Public Property Get IsZipEncryptWfFiles() As Boolean
    IsZipEncryptWfFiles = mIsZipEncryptWfFiles
End Property

Public Property Get IsSendErrorAsAlert() As Boolean
    IsSendErrorAsAlert = mIsSendErrorAsAlert
End Property

Public Property Get RoleTypeKeyAlert() As Long
    RoleTypeKeyAlert = mRoleTypeKeyAlert
End Property

Public Property Get IsSchedulerAllowed() As Boolean
    IsSchedulerAllowed = mIsSchedulerAllowed
End Property

Public Property Get StartBlackout() As Date
    StartBlackout = mStartBlackout
End Property

Public Property Get BlackoutMinutes() As Long
    BlackoutMinutes = mBlackoutMinutes
End Property

'Public Property Get A() As Boolean
'    A = mA
'End Property

'Public Property Get A() As Boolean
'    A = mA
'End Property

'Public Property Get A() As Boolean
'    A = mA
'End Property

Public Property Let IsInboxAllowed(ByVal vData As Boolean)
    mIsInboxAllowed = vData
End Property

Public Property Get IsPackageAllowed() As Boolean
    IsPackageAllowed = mIsPackageAllowed
End Property

Public Property Let IsPackageAllowed(ByVal vData As Boolean)
    mIsPackageAllowed = vData
End Property

Public Property Let IsZipAttachmentFile(ByVal bData As Boolean)
    mIsZipAttachmentFile = bData
End Property

'Public Property Get IsEmailRouterAllowed() As Boolean
'    IsEmailRouterAllowed = False
'End Property

'Public Property Let IsEmailRouterAllowed(ByVal vData As Boolean)
'    mIsEmailRouterAllowed = False
'End Property

'Public Property Get IsEventQueueAllowed() As Boolean
'    IsEventQueueAllowed = False
'End Property

Public Property Get IsSendDataChangesAllowed() As Boolean
    IsSendDataChangesAllowed = mIsSendDataChangesAllowed
End Property

Public Property Let IsSendDataChangesAllowed(ByVal vData As Boolean)
    mIsSendDataChangesAllowed = vData
End Property

Public Property Get IsSendBlobChangesAllowed() As Boolean
    IsSendBlobChangesAllowed = mIsSendBlobChangesAllowed
End Property

Public Property Let IsSendBlobChangesAllowed(ByVal vData As Boolean)
    mIsSendBlobChangesAllowed = vData
End Property

Public Property Get IsManualArchiveAllowed() As Boolean
    IsManualArchiveAllowed = mIsManualArchiveAllowed
End Property

Public Property Let IsManualArchiveAllowed(ByVal vData As Boolean)
    mIsManualArchiveAllowed = vData
End Property

Public Property Get IsStopUpdateDBOnError() As Boolean
    IsStopUpdateDBOnError = mIsStopUpdateDBOnError
End Property

Public Property Let IsStopUpdateDBOnError(ByVal vData As Boolean)
    mIsStopUpdateDBOnError = vData
End Property

Public Property Get IsNotifyManualRevisionUpdate() As Boolean
    IsNotifyManualRevisionUpdate = mIsNotifyManualRevisionUpdate
End Property

Public Property Let IsNotifyManualRevisionUpdate(ByVal vData As Boolean)
    mIsNotifyManualRevisionUpdate = vData
End Property
Public Property Get IsAutoSendFCChangesAllowed() As Boolean
    IsAutoSendFCChangesAllowed = mIsAutoSendFCChangesAllowed
End Property

Public Property Let IsAutoSendFCChangesAllowed(ByVal vData As Boolean)
    mIsAutoSendFCChangesAllowed = vData
End Property
Public Property Get IsSendDAWithInThresholdLimit() As Boolean
    IsSendDAWithInThresholdLimit = mIsSendDAWithInThresholdLimit
End Property

Public Property Let IsSendDAWithInThresholdLimit(ByVal vData As Boolean)
    mIsSendDAWithInThresholdLimit = vData
End Property

Public Property Get IsStripOffExtendedAsciiChars() As Boolean
    IsStripOffExtendedAsciiChars = mIsStripOffExtendedAsciiChars
End Property

Public Property Let IsStripOffExtendedAsciiChars(ByVal vData As Boolean)
    mIsStripOffExtendedAsciiChars = vData
End Property

Public Property Get IsHideSystemTray() As Boolean
    IsHideSystemTray = mIsHideSystemTray
End Property

Public Property Let IsHideSystemTray(ByVal vData As Boolean)
    mIsHideSystemTray = vData
End Property

'Public Property Get A() As Boolean
'    A = mA
'End Property

'Public Property Let A(ByVal vData As Boolean)
'    mA = vData
'End Property



'Public Property Get A() As Boolean
'    A = mA
'End Property

Public Property Get IsManualMode() As Boolean
    IsManualMode = mIsManualMode
End Property

Public Property Let IsManualMode(ByVal vData As Boolean)
    mIsManualMode = vData
End Property

'Public Property Let DgShortDateFormat(ByVal vData As String)
'    mDgShortDateFormat = vData
'End Property

'Public Property Get DgShortDateFormat() As String
'    DgShortDateFormat = mDgShortDateFormat
'End Property

'Public Property Let DgGeneralDateFormat(ByVal vData As String)
'    mDgGeneralDateFormat = vData
'End Property

'Public Property Get DgGeneralDateFormat() As String
'    DgGeneralDateFormat = mDgGeneralDateFormat
'End Property

'Public Property Let DbEmailRouterString(ByVal vData As String)
'    mDbEmailRouterString = vData
'End Property

'Public Property Get DbEmailRouterString() As String
'    DbEmailRouterString = mDbEmailRouterString
'End Property


Public Property Get Package() As mwPackage
   If moPackage Is Nothing Then
      Set moPackage = New mwPackage
      Set moPackage.Parent = Me
      Set moPackage.Session = moParent
   End If
   Set Package = moPackage
End Property

Public Property Set Package(ByVal vData As mwPackage)
   If moPackage Is Nothing Then
      Set moPackage = New mwPackage
      Set moPackage.Parent = Me
      Set moPackage.Session = moParent
   End If
   Set moPackage = vData
End Property

Public Property Get TemplateWorkFolder() As String
   TemplateWorkFolder = mTemplateWorkFolder
   MakeFolderExist TemplateWorkFolder
End Property

Public Property Let TemplateWorkFolder(ByVal vData As String)
    mTemplateWorkFolder = vData
End Property

Public Property Get TemplatePlacesID() As Long
    TemplatePlacesID = mTemplatePlacesID
End Property

Public Property Let TemplatePlacesID(ByVal vData As Long)
    mTemplatePlacesID = vData
End Property

Public Property Get SignedFormsPlaceID() As Long
    SignedFormsPlaceID = mSignedFormsPlaceID
End Property

Public Property Let SignedFormsPlaceID(ByVal vData As Long)
    mSignedFormsPlaceID = vData
End Property

Public Property Get IsResendRequested() As Boolean
    IsResendRequested = mIsResendRequested
End Property

Public Property Let IsResendRequested(ByVal bData As Boolean)
    IsResendRequested = bData
End Property
Public Property Get IsDeferAckUntilMoreData() As Boolean
   IsDeferAckUntilMoreData = mIsDeferAckUntilMoreData
End Property

Public Property Let IsDeferAckUntilMoreData(ByVal bData As Boolean)
    IsDeferAckUntilMoreData = bData
End Property

' Object Template
'Public Property Get A() As Object
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set A = moA
'End Property

'Public Property Let A(ByVal vData As Object)
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set moA = vData
'End Property



Friend Function LoadConfiguration() As Boolean
   Dim loRs As Recordset
   Dim strID As String
   Dim sSQL As String
   
   Dim loField As field
   Dim IsSessionFolder As Boolean
   On Error GoTo FunctionError
   '
   ' Get WF record to use from mwcThisSite...
   '
   mIsLoadWorkflow = True
   
   If Trim(goSession.User.GetExtendedProperty("WorkflowCfgOverride")) <> "" Then
      strID = goSession.User.GetExtendedProperty("WorkflowCfgOverride")
   ElseIf Trim(goSession.ThisSite.WorkflowCfg) <> "" Then
      strID = goSession.ThisSite.WorkflowCfg
   Else
      strID = "-1"
   End If
   '
   ' Load it up
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "mwcWorkflow", goSession.DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdTable
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.Load Configuration, no mwcWorkflow Configuration records found."
      LoadConfiguration = False
      goSession.CloseRecordset loRs
      Exit Function
   End If
   '
   ' Post 0392 patch ?
   '
      '
   If UCase(Trim(loRs.Fields(1).Name)) = "WORKFLOWID" Then
      If strID <> "-1" Then
         loRs.Filter = "WorkflowID='" & strID & "'"
      End If
      mIsPost0392 = True
   Else
      If strID <> "-1" Then
         loRs.Filter = "ID='" & strID & "'"
      End If
      mIsPost0392 = False
   End If
   
   
   'sSQL = "select * from mwcWorkflow where ID='" & strID & "'"
   'loRS.Open sSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   '
   ' 3/25/2002 ms - Extended Formtemplate Properties...
       
   If Not moExtendedKeys Is Nothing Then
      Set moExtendedKeys = Nothing
   End If
   Set moExtendedKeys = New ConfigKeys
   For Each loField In loRs.Fields
      Select Case UCase(loField.Name)
         
         Case Is = UCase("ID")
            If mIsPost0392 Then
               mIDKey = loRs!ID
            End If
         Case Is = UCase("WorkflowID")
            If Not IsNull(loRs!WorkflowID) Then
               mWorkflowID = loRs!WorkflowID
            End If
         Case Is = UCase("Description")
            If Not IsNull(loRs!Description) Then
               mDescription = loRs!Description
            End If
         Case Is = UCase("PollingFolder")
            If Not IsNull(loRs!PollingFolder) Then
               mPollingFolder = loRs!PollingFolder
            End If
         Case Is = UCase("SendByMediaFolder")
            If Not IsNull(loRs!SendByMediaFolder) Then
               mSendByMediaFolder = loRs!SendByMediaFolder
            End If
         Case Is = UCase("ScriptFolder")
            If Not IsNull(loRs!ScriptFolder) Then
               mScriptFolder = loRs!ScriptFolder
            End If
         'Case Is = UCase("CommonScriptFilename")
         '   If Not IsNull(loRs!CommonScriptFilename) Then
         '      mCommonScriptFilename = loRs!CommonScriptFilename
         '   End If
         Case Is = UCase("PollingIntervalMinutes")
            If Not IsNull(loRs!PollingIntervalMinutes) Then
               mPollingIntervalMinutes = loRs!PollingIntervalMinutes
            End If
         Case Is = UCase("IsDebugMode")
            If Not IsNull(loRs!IsDebugMode) Then
               mIsDebugMode = loRs!IsDebugMode
            End If
            
         Case Is = UCase("IsDebugMemory")
            mIsDebugMemory = BoolNull(loRs!IsDebugMemory)

         Case Is = UCase("IsSilentErrors")
            If Not IsNull(loRs!IsSilentErrors) Then
               mIsSilentErrors = loRs!IsSilentErrors
            End If
         Case Is = UCase("IsNotifyPkgReceipt")
            If Not IsNull(loRs!IsNotifyPkgReceipt) Then
               mIsNotifyPkgReceipt = loRs!IsNotifyPkgReceipt
            End If
         Case Is = UCase("IsRunOnceMode")
            If Not IsNull(loRs!IsRunOnceMode) Then
               mIsRunOnceMode = loRs!IsRunOnceMode
            End If
         Case Is = UCase("IsShowAllDatagramErrors")
            If Not IsNull(loRs!IsShowAllDatagramErrors) Then
               mIsShowAllDatagramErrors = loRs!IsShowAllDatagramErrors
            End If
         Case Is = UCase("MaxErrorsBeforeQuit")
            If Not IsNull(loRs!MaxErrorsBeforeQuit) Then
               mMaxErrorsBeforeQuit = loRs!MaxErrorsBeforeQuit
            End If
         Case Is = UCase("MaxErrorsBeforeNotify")
            If Not IsNull(loRs!MaxErrorsBeforeNotify) Then
               mMaxErrorsBeforeNotify = loRs!MaxErrorsBeforeNotify
            End If
         Case Is = UCase("WorkingFolder")
            If Not IsNull(loRs!WorkingFolder) Then
               mWorkingFolder = loRs!WorkingFolder
            End If
         Case Is = UCase("ArchiveOkZipFile")
            If Not IsNull(loRs!ArchiveOkZipFile) Then
               mArchiveOkZipFile = loRs!ArchiveOkZipFile
            End If
         Case Is = UCase("ArchiveErrorZipFile")
            If Not IsNull(loRs!ArchiveErrorZipFile) Then
               mArchiveErrorZipFile = loRs!ArchiveErrorZipFile
            End If
         Case Is = UCase("IsFormsProcessingAllowed")
            If Not IsNull(loRs!IsFormsProcessingAllowed) Then
               mIsFormsProcessingAllowed = loRs!IsFormsProcessingAllowed
            End If
         Case Is = UCase("IsInboxAllowed")
            If Not IsNull(loRs!IsInboxAllowed) Then
               mIsInboxAllowed = loRs!IsInboxAllowed
            End If
         Case Is = UCase("IsPackageAllowed")
            If Not IsNull(loRs!IsPackageAllowed) Then
               mIsPackageAllowed = loRs!IsPackageAllowed
            End If
         Case Is = UCase("IsManualMode")
            If Not IsNull(loRs!IsManualMode) Then
               mIsManualMode = loRs!IsManualMode
            End If
         Case Is = UCase("InboundQueueFolder")
            If Not IsNull(loRs!InboundQueueFolder) Then
               mInboundQueueFolder = loRs!InboundQueueFolder
            End If
         Case Is = UCase("SendByCdromFolder")
            If Not IsNull(loRs!SendByCdromFolder) Then
               mSendByCdromFolder = loRs!SendByCdromFolder
            End If
         Case Is = UCase("IsZipWfFiles")
            If Not IsNull(loRs!IsZipWfFiles) Then
               mIsZipWfFiles = loRs!IsZipWfFiles
            End If
         Case Is = UCase("IsZipEncryptWfFiles")
            If Not IsNull(loRs!IsZipEncryptWfFiles) Then
               mIsZipEncryptWfFiles = loRs!IsZipEncryptWfFiles
            End If
         Case Is = UCase("IsSendDataChangesAllowed")
            If Not IsNull(loRs!IsSendDataChangesAllowed) Then
               mIsSendDataChangesAllowed = loRs!IsSendDataChangesAllowed
            End If
         Case Is = UCase("IsSendBlobChangesAllowed")
            If Not IsNull(loRs!IsSendBlobChangesAllowed) Then
               mIsSendBlobChangesAllowed = loRs!IsSendBlobChangesAllowed
            End If
         Case Is = UCase("IsSendErrorAsAlert")
            If Not IsNull(loRs!IsSendErrorAsAlert) Then
               mIsSendErrorAsAlert = loRs!IsSendErrorAsAlert
            End If
         Case Is = UCase("mwcRoleTypeKeyAlert")
            If Not IsNull(loRs!mwcRoleTypeKeyAlert) Then
               mRoleTypeKeyAlert = loRs!mwcRoleTypeKeyAlert
            End If
         Case Is = UCase("StartBlackout")
            If Not IsNull(loRs!StartBlackout) Then
               mStartBlackout = loRs!StartBlackout
            End If
         Case Is = UCase("BlackoutMinutes")
            If Not IsNull(loRs!BlackoutMinutes) Then
               mBlackoutMinutes = loRs!BlackoutMinutes
            End If
         Case Is = UCase("IsSchedulerAllowed")
            If Not IsNull(loRs!IsSchedulerAllowed) Then
               mIsSchedulerAllowed = loRs!IsSchedulerAllowed
            End If
         Case Is = UCase("TemplateWorkFolder")
            If Not IsNull(loRs!TemplateWorkFolder) Then
               mTemplateWorkFolder = loRs!TemplateWorkFolder
            End If
         Case Is = UCase("TemplatePlacesID")
            If Not IsNull(loRs!TemplatePlacesID) Then
               mTemplatePlacesID = loRs!TemplatePlacesID
            End If
         Case Is = UCase("SignedFormsPlaceID")
            If Not IsNull(loRs!SignedFormsPlaceID) Then
               mSignedFormsPlaceID = loRs!SignedFormsPlaceID
            End If
         Case Is = UCase("IsResendRequested")
            If Not IsNull(loRs!IsResendRequested) Then
               mIsResendRequested = loRs!IsResendRequested
            End If
         Case Is = UCase("IsDeferAckUntilMoreData")
            If Not IsNull(loRs!IsDeferAckUntilMoreData) Then
               mIsDeferAckUntilMoreData = loRs!IsDeferAckUntilMoreData
            End If
         Case Is = UCase("IsManualArchiveAllowed")         '---PBT-109
            If Not IsNull(loRs!IsManualArchiveAllowed) Then
               mIsManualArchiveAllowed = loRs!IsManualArchiveAllowed
            End If
         Case Is = UCase("IsStopUpdateDBOnError")         '---PBT-1567
            If Not IsNull(loRs!IsStopUpdateDBOnError) Then
               mIsStopUpdateDBOnError = loRs!IsStopUpdateDBOnError
            End If
         Case Is = UCase("IsHideSystemTray")
            If Not IsNull(loRs!IsHideSystemTray) Then
               mIsHideSystemTray = loRs!IsHideSystemTray
            End If
         Case Is = UCase("IsNotifyManualRevisionUpdate")         '---PBT-1567
            If Not IsNull(loRs!IsNotifyManualRevisionUpdate) Then
               mIsNotifyManualRevisionUpdate = loRs!IsNotifyManualRevisionUpdate
            End If
         Case Is = UCase("IsAutoSendFCChangesAllowed")         '---PBT-1742
            If Not IsNull(loRs!IsAutoSendFCChangesAllowed) Then
               mIsAutoSendFCChangesAllowed = loRs!IsAutoSendFCChangesAllowed
            End If
         Case Is = UCase("IsSendDAWithInThresholdLimit")
            If Not IsNull(loRs!IsSendDAWithInThresholdLimit) Then
               mIsSendDAWithInThresholdLimit = loRs!IsSendDAWithInThresholdLimit
            End If
         Case Is = UCase("IsStripOffExtendedAsciiChars")
            If Not IsNull(loRs!IsStripOffExtendedAsciiChars) Then
               mIsStripOffExtendedAsciiChars = loRs!IsStripOffExtendedAsciiChars
            End If
         Case Else
            '
            ' Otherwise, Add to the Collection of configkeys... (extended properties)
            '
            'Debug.Print loField.Name
            If Not IsNull(loField.value) Then
            
               If goSession.IsOracle And loField.Type = adNumeric And loField.Precision = 1 Then
                  If loField.value = 0 Then
                     moExtendedKeys.Add UCase(loField.Name), "False", UCase(loField.Name)
                  Else
                     moExtendedKeys.Add UCase(loField.Name), "True", UCase(loField.Name)
                  End If
               Else
                  moExtendedKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
               End If
            End If
      End Select
   Next
   '
   ' Close up shop
   '
   goSession.CloseRecordset loRs
   '
   ' Create WorkingSessionFolder...
   '
   If Not moFso.FolderExists(mWorkingFolder) Then
      On Error Resume Next
      moFso.CreateFolder (mWorkingFolder)
      If Err Then
         MsgBox "Unlogged error in mwSession.mwWorkflow.LoadConfiguration. Unable to Create Working folder, system will not operate correctly."
            LoadConfiguration = False
            Exit Function
         End If
   End If
   '
   ' Create the Working Session Folder
   '
   IsSessionFolder = False
   Do While Not IsSessionFolder
      mWorkingSessionFolder = mWorkingFolder & "\" & goSession.User.RoleType & Format(Now(), "yymmddhhmmss")
      If moFso.FolderExists(mWorkingSessionFolder) Then
         mWorkingSessionFolder = mWorkingSessionFolder & "1"
         If Len(mWorkingSessionFolder) > 60 Then
            MsgBox "Error in mwSession.mwWorkflow.LoadConfiguration, Cannot create Working Session Folder."
            LoadConfiguration = False
            Exit Function
         End If
      Else
         moFso.CreateFolder mWorkingSessionFolder
         IsSessionFolder = True
      End If
   Loop
   mIsLoadWorkflow = False
   LoadConfiguration = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.LoadConfiguration.", Err.Number, Err.Description
   LoadConfiguration = False
   goSession.CloseRecordset loRs
End Function



Public Function OpenNewPackage() As mwPackage
   Set OpenNewPackage = New mwPackage
   Set OpenNewPackage.Parent = Me
   Set OpenNewPackage.Session = moParent
End Function

Public Function OpenNewWorkflowItem() As mwWorkflowItem
   Set OpenNewWorkflowItem = New mwWorkflowItem
End Function




Public Function ProcessWorkflowForms() As Boolean
   Dim loFolder As Folder
   Dim loFile As file
   Dim loform As mwForm
   Dim lIsErrors As Boolean
   Dim lHasFileFailed As Boolean
   Dim strSourceFile As String
   Dim strTargetFile As String
   Dim strEmailSubject As String
   Dim strEmailTextBody As String
   Dim strVbDLL As String
   Dim sNotifySubject As String
   Dim sNotifyBody As String
   Dim loRs As Recordset
   Dim loZip As Zipper
   Dim sThisExt As String
   
   goSession.MemoryLogIt "ProcessWorkflowForms 1"
   
   '
   ' Preliminary Checks...
   '
   On Error GoTo FunctionError
   ' Raise error on an exception basis
   ProcessWorkflowForms = True
   If Trim(mPollingFolder) = "" Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowForms," & _
        "No Polling folder specified in mwcWorkflow table."
      ProcessWorkflowForms = False
      Exit Function
   End If
   If Trim(mWorkingSessionFolder) = "" Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowForms," & _
        "Can not Process DataForm Messages, No Working Session Folder specified in mwcWorkflow table."
      ProcessWorkflowForms = False
      Exit Function
   End If
   
   LicenseValidationCheck
   
   Set loRs = GetSortedInboundRS()
   Do While Not loRs.EOF
      goSession.MemoryLogIt "ProcessWorkflowForms 2"
      sThisExt = UCase(moFso.GetExtensionName(loRs!Path))
      '
      ' Files with ".dgm" extension... or ".dat" in Version 2 Compatibility Mode
      '
      If mIsFormsProcessingAllowed = True And _
         (sThisExt = "DGM" Or (sThisExt = "DAT" And GetExtendedProperty("Version2Compatible") = "True")) Then
         '
         ' Attempt to Process DataGram
         '
         lIsErrors = False
         Set loform = goSession.OpenNewForm
         
         If loform.OpenForm(loRs!Path) Then
            '
            ' Error Notify
            '
            If loform.TemplateProperties.TemplateID = "ERROR_NOTIFY" Then
               '
               ' Convert to email and force out...
               '

               If Trim(goSession.Site.WorkflowErrorNotifyAddress) <> "" Then
                  sNotifySubject = goSession.Site.GetShipProperty(loform.FormProperties.SiteID, "SiteName") & ": " & _
                                    loform.DataForm.Records(2).ConfigKeys.GetKeyValue("Subject")
                  sNotifyBody = Replace(loform.DataForm.Records(2).ConfigKeys.GetKeyValue("TxtBody"), "|", vbCrLf)
                  goSession.SendNotification sNotifySubject, sNotifyBody, goSession.Site.WorkflowErrorNotifyAddress
                  goSession.RaiseWorkflowStatus "Logging in mwSession.mwWorkflow.ProcessWorkflowForms, Error Notification Email Address sent for form: " & loform.FormProperties.FormID
               Else
                  goSession.RaiseWorkflowStatus "Logging in mwSession.mwWorkflow.ProcessWorkflowForms, Error Notification datagram ignored, no Error Notify Email Address defined."
               End If
            Else
               '
               ' Test TargetSiteID
               '
               If loform.FormType = mwft_DATA_FORM Then
                  If loform.FormProperties.TargetSiteID <> "" And loform.FormProperties.TargetSiteID _
                    <> goSession.Site.SiteID Then
                     '
                     ' Error - Datagram delivered to wrong site...
                     '
                     lIsErrors = True
                     goSession.RaiseWorkflowStatus "Error in mwSession.mwWorkflow.ProcessWorkflowForms, attempt to processe datagram for wrong TargetSiteID. Form ID: " & loform.FormProperties.FormID
                  ElseIf loform.FormProperties.IsFleetTransmission And loform.FormProperties.SiteID = goSession.Site.SiteID Then
                     '
                     ' Error - Datagram delivered to wrong site...
                     '
                     lIsErrors = True
                     goSession.RaiseWorkflowStatus "Error in mwSession.mwWorkflow.ProcessWorkflowForms, attempt to processe fleet transmission datagram from originating site."
                  Else
                     '
                     ' VRS Intercepts - Translations
                     '
                     If UCase(Left(loform.TemplateProperties.TemplateID, 3)) = "VRS" Then
                        vrsProcessTranslations loform
                        '
                        ' Compatibility Layer for old VRS...
                        '
                        If Trim(loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK1")) <> "" And _
                         Trim(loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK")) = "" Then
                           vrsProcessOldRemarks loform
                        End If
                     End If
                     '
                     ' External DLL and VbScript
                     '
                     lIsErrors = Not ExecuteDllOnProcess(loform)
                  End If
               End If
            End If
            goSession.MemoryLogIt "ProcessWorkflowForms 3"
            If Not lIsErrors Then
               ' Go for it
               ' Process as Recordset...
               If loform.WorkflowProperties.IsUpdateAsRecordset Then
                  If loform.DataForm.UpdateAsRecordset() Then
                     goSession.RaiseWorkflowStatus "Successfully Processed Data Form Message: " & _
                       loform.FormProperties.FormID
                  Else
                     goSession.RaiseWorkflowStatus "Error(s) Processing Data Form Message: " & loRs!Path
                     lIsErrors = True
                  End If
               End If
               '
               ' Datagrams route to folder ? Not implemented yet...
               '
               
               '
               ' Convert Datagram to Email Option ?
               '
               If loform.GetExtendedProperty("ConvertDgToEmail") = "True" Then
                  SendDatagramEmails loform
               End If
               
            End If
         Else
            goSession.RaiseWorkflowStatus "Unable to Open Form: " & loRs!Path
            lIsErrors = True
         End If
      '
      ' Spreadsheets
      '
      ElseIf mIsFormsProcessingAllowed = True And _
            (sThisExt = "XLS" Or sThisExt = "XLSX" Or sThisExt = "XLSM") Then
         '
         ' Excel Forms:
         '
         Set loform = goSession.OpenNewForm
         If loform.OpenForm(loRs!Path) Then
            '
            ' Run Script if there...
            '
            '
            ' External DLL and VbScript
            '
            lIsErrors = Not ExecuteDllOnProcess(loform)
            If Not lIsErrors Then
               ' Go for it
               ' Process as Recordset...
               '
               ' 8/2010 ms BGLNG-4
               If MaybeProcessReturningForm(loform) Then
                  goSession.RaiseWorkflowStatus "Successfully processed returning external form: " & loform.FormProperties.FormID
               ' 2/2011 ms BGLNG-18 Excel Forms Mapped to Marine Assurance Question Sets
               ElseIf MaybeProcessQuestionSetForm(loform) Then
                  goSession.RaiseWorkflowStatus "Processed Marine Assurance Question Set Spreadsheet for form: " & loform.FormProperties.FormID
               Else
                  goSession.MemoryLogIt "ProcessWorkflowForms 4"
                  If Trim(loform.WorkflowProperties.MapScriptFile) <> "" Then
                     If Not loform.WorkflowProperties.IsRouteToMailboxExclusive Then
                        If loform.ExcelForm.UpdateAsRecordset() Then
                           goSession.RaiseWorkflowStatus "Successfully Processed Excel spreadsheet : " & loform.FormProperties.FormID
                           'loForm.CloseForm
                           'ArchiveWorkflowItem lors!Path, ArchivedOK, loForm.TemplateProperties.DeleteAfterTransmit
                        Else
                           goSession.RaiseWorkflowStatus "Error(s) Processing Excel Form Message: " & loRs!Path
                           'loForm.CloseForm
                           'ArchiveWorkflowItem lors!Path, ArchiveError, True
                           ProcessWorkflowForms = False
                           lIsErrors = True
                        End If
                     Else
                        goSession.RaiseWorkflowStatus "Reroute Form Exclusive encoutered, skipping UpdateAsRecordset: " & loRs!Path
                     End If
                  End If
               End If
            End If
         Else
            goSession.RaiseWorkflowStatus "Unable to Open Form: " & loRs!Path
            'ArchiveWorkflowItem lors!Path, ArchiveError, True
            ProcessWorkflowForms = False
            lIsErrors = True
         End If
      
      ElseIf mIsFormsProcessingAllowed = True And _
             (sThisExt = "DOC" Or sThisExt = "DOCX" Or sThisExt = "DOCM") Then
         '
         ' Word Forms:
         '
         Set loform = goSession.OpenNewForm
         If loform.OpenForm(loRs!Path) Then
            '
            ' External DLL and VbScript
            '
            lIsErrors = Not ExecuteDllOnProcess(loform)
            
         Else
            goSession.RaiseWorkflowStatus "Unable to Open Form: " & loRs!Path
            'ArchiveWorkflowItem lors!Path, ArchiveError, True
            ProcessWorkflowForms = False
            lIsErrors = True
         End If
      ElseIf moKeys.GetBoolKeyValue("ProcessZipFiles") = True And sThisExt = "ZIP" Then
         '
         ' Unzip file and archive it...
         '
         goSession.MemoryLogIt "ProcessWorkflowForms 5"
         Set loZip = New Zipper
         If Not loZip.OpenZip(loRs!Path) Then
            goSession.RaiseWorkflowStatus "Unable to open ZIP File: " & moFso.GetFileName(loRs!Path)
            ArchiveWorkflowItem loRs!Path, ArchiveError, True
         Else
            goSession.RaiseWorkflowStatus "Unzipping: " & moFso.GetFileName(loRs!Path)
            If loZip.GetFromZip("*.*", goSession.Workflow.PollingFolder) Then
               ArchiveWorkflowItem loRs!Path, ArchivedOK, True
            Else
               ArchiveWorkflowItem loRs!Path, ArchiveError, True
            End If
         End If
            
         loZip.TerminateClass
         Set loZip = Nothing
         
      End If
      '
      '
      ' Process the physical file.
      '
      goSession.MemoryLogIt "ProcessWorkflowForms 6"
      If Not loform Is Nothing Then
         '
         ' After Process DLL call
         '
         ExecuteDllAfterProcess loform
         
         ' 8/2010 ms BGLNG-4
         If MaybeProcessReturningForm(loform) Then
            goSession.RaiseWorkflowStatus "Successfully processed returning external form: " & loform.FormProperties.FormID
            
         Else
            '
            ' Check for Route to Folder...
            '
            If Trim(loform.WorkflowProperties.RouteToFolder) = "" Then
               loform.CloseForm
            Else
               If Not moFso.FolderExists(loform.WorkflowProperties.RouteToFolder) Then
                  goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowItems, Route To Folder missing: " & loform.WorkflowProperties.RouteToFolder
                  goSession.RaiseWorkflowStatus "Skipping Route To Folder due to an error..."
               Else
                  '
                  ' RouteToFolder designed for email processing, but may be used in SendByMedia or Demo Loopback
                  '
                  strTargetFile = GetRouteToFolder(loform)
                  strTargetFile = loform.WorkflowProperties.RouteToFolder & "\" & moFso.GetFileName(loRs!Path)
                  loform.CloseForm
                  If Not moFso.FileExists(strTargetFile) Then
                     moFso.CopyFile loRs!Path, strTargetFile
                  End If
               End If
            End If
         End If
         '
         ' If route to email, and TransportMetod is not email... do it here instead.
         '
         If goSession.User.DefaultTransport <> mwt_ElectronicMail And _
          Trim(loform.WorkflowProperties.RouteToEmailAddress) <> "" Then
            'sIsRouteToEmailAddress = loForm.WorkflowProperties.RouteToEmailAddress
            'sIsRouteToFilename = loForm.FormProperties.FullFilename
            loform.SendRouteToEmail
         End If
         '
         ' Archive the form
         '
         strSourceFile = loRs!Path
         goSession.TearDownForm loform
         If gIsLeaveTransFile Then
            ' Leave Out of Sequence Batch File for next cycle...
            gIsLeaveTransFile = False
         Else
            If lIsErrors Then
               ArchiveWorkflowItem loRs!Path, ArchiveError, True
               lHasFileFailed = True
            Else
               ArchiveWorkflowItem loRs!Path, ArchivedOK, True
            End If
         End If
      End If
      loRs.MoveNext
   Loop
   '
   ' Close up shop...
   goSession.KillObject loFolder
   goSession.KillObject loFile
   goSession.TearDownForm loform
   If lHasFileFailed = True Then
      ProcessWorkflowForms = False
   Else
      ProcessWorkflowForms = True
   End If
   goSession.MemoryLogIt "ProcessWorkflowForms Exit"
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ProcessWorkflowForms.", Err.Number, Err.Description
   goSession.KillObject loFolder
   goSession.KillObject loFile
   goSession.TearDownForm loform
   ProcessWorkflowForms = False
End Function



Private Function KillFolder(wFolder As String) As Boolean
   Dim oFolder As Folder
   Dim oFile As file
   On Error GoTo FunctionError
   If moFso.FolderExists(wFolder) Then
      Set oFolder = moFso.GetFolder(wFolder)
      For Each oFile In oFolder.Files
         oFile.Attributes = Normal
      Next
      moFso.DeleteFolder wFolder
   End If
   'moFso.CreateFolder wFolder
   goSession.KillObject oFolder
   goSession.KillObject oFile
   KillFolder = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.KillFolder.", Err.Number, Err.Description
   KillFolder = False
   goSession.KillObject oFolder
   goSession.KillObject oFile
End Function


Public Function RenameFileNowExtension(FullFilename As String) As Boolean
   Dim strTarget As String
   On Error GoTo FunctionError
   
   If moFso.FileExists(FullFilename) Then
      ' Set attributes to normal
      moFso.GetFile(FullFilename).Attributes = Normal
      strTarget = FullFilename & "." & Format(Now(), "YYYYMMDDhhmmss")
      Do While moFso.FileExists(strTarget)
         strTarget = strTarget & ".1"
      Loop
      moFso.MoveFile FullFilename, strTarget
   End If
   RenameFileNowExtension = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwWorkflow.RenameFileNowExtension file= " & FullFilename, Err.Number, Err.Description
   RenameFileNowExtension = False
End Function

Friend Function RaiseStatusMessage(Msg As String)
   RaiseEvent Status(Msg)
End Function

Public Function ArchiveWorkflowItem(FullFilename As String, ArchiveStatus As mwSession.mw_ArchiveStatus, _
  DeleteAfterArchive As Boolean, Optional IgnoreErrors As Boolean) As Boolean
   Dim oZip As Zipper
   Dim Error As String
   Dim strZipFile As String
   Dim strTargetFile As String
   Dim iCount As Integer
   On Error GoTo FunctionError
   Set oZip = New Zipper
   If DeleteAfterArchive Then
      oZip.DeleteAfterZip = True
   Else
      oZip.DeleteAfterZip = False
   End If
   ' Different Zip file for Processed OK vs. Error forms...
   If ArchiveStatus = ArchivedOK Then
      strZipFile = mArchiveOkZipFile
   ElseIf ArchiveStatus = ArchiveError Then
      strZipFile = mArchiveErrorZipFile
   Else
      strZipFile = mInboundQueueFolder
   End If
   '
   ' Do we zip or simply move files...
   '
   If UCase(moFso.GetExtensionName(strZipFile)) = "ZIP" Then
      '
      ' Does Zip Folder Exist...
      '
      If Not moFso.FolderExists(moFso.GetParentFolderName(strZipFile)) Then
         If Not IgnoreErrors Then
            goSession.RaiseError "Error in mwSession.Workflow.ArchiveWorkflowItem, Folder doesn't exist: " _
              & moFso.GetParentFolderName(strZipFile)
         Else
            goSession.RaiseWorkflowStatus "Error in mwSession.Workflow.ArchiveWorkflowItem, Folder doesn't exist: " _
              & moFso.GetParentFolderName(strZipFile)
         End If
         ArchiveWorkflowItem = False
         oZip.TerminateClass
         Set oZip = Nothing
         Exit Function
      End If
      
      
      ' Move or Copy to the zip
      oZip.OpenZip strZipFile
      '
      ' Check if item already in zip ?
      '
      ' Leave for future...
      
      If Not oZip.AddToZip(FullFilename) Then
         If Not IgnoreErrors Then
            goSession.RaiseError "Error in mwSession.mwWorkflow.ArchiveWorkflowItem, Zip Error: " & oZip.ErrorCode
         Else
            goSession.RaiseWorkflowStatus "Error in mwSession.mwWorkflow.ArchiveWorkflowItem, Zip Error: " & oZip.ErrorCode
         End If
         ArchiveWorkflowItem = False
      Else
         ArchiveWorkflowItem = True
      End If
      '
   Else
      '
      ' No zipping, just move it...
      '
      If Not moFso.FolderExists(strZipFile) Then
         If Not IgnoreErrors Then
            goSession.RaiseError "Error in mwSession.Workflow.ArchiveWorkflowItem, Folder doesn't exist: " _
              & strZipFile
         Else
            goSession.RaiseWorkflowStatus "Error in mwSession.Workflow.ArchiveWorkflowItem, Folder doesn't exist: " _
              & strZipFile
         End If
         ArchiveWorkflowItem = False
         oZip.TerminateClass
         Set oZip = Nothing
         Exit Function
      End If
      strTargetFile = strZipFile & "\" & moFso.GetFileName(FullFilename)
      If moFso.FileExists(FullFilename) Then
         If moFso.FileExists(strTargetFile) Then
            RenameFileNowExtension (strTargetFile)
         End If
      End If
      If oZip.DeleteAfterZip Then
         moFso.MoveFile FullFilename, strTargetFile
      Else
         moFso.CopyFile FullFilename, strTargetFile
      End If
      ArchiveWorkflowItem = True
   End If
   oZip.TerminateClass
   Set oZip = Nothing
   Exit Function
FunctionError:
   If Err.Number = 53 And Not IgnoreErrors Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Warning in mwSession.mwWorkflow.ArchiveWorkflowItem, file not found to archive: " & FullFilename
   ElseIf Not IgnoreErrors Then
      goSession.RaiseError "General error in mwSession.mwWorkflow.ArchiveWorkflowItem.", Err.Number, Err.Description
   End If
   oZip.TerminateClass
   Set oZip = Nothing
   ArchiveWorkflowItem = False
End Function


Public Function KillWorkingSessionFolder()
   '
   ' Delete folder created for the session
   '
   On Error GoTo SubError
   moFso.DeleteFolder mWorkingSessionFolder, True
   Exit Function
SubError:
   '
End Function


Public Function SendByDemoLoopback(FullFilename As String, Optional TargetSiteKey As Long) As Boolean
   On Error GoTo FunctionError
   moFso.CopyFile FullFilename, goSession.Workflow.PollingFolder & "\"
   SendByDemoLoopback = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendByDemoLoopback.", Err.Number, Err.Description
   SendByDemoLoopback = False
End Function


Public Function SendByMedia(FullFilename As String, Optional Zipit As Boolean, _
  Optional ExplicitPath As String, Optional ChooseTargetPath As Boolean, _
  Optional TargetSiteID As String, _
  Optional IsSendByCDROM As Boolean, Optional TargetSiteKey As Long) As Boolean
   Dim oForm As frmLocation
   Dim strTarget As String
   Dim sSendFolder As String
   Dim sFileToSend As String
   On Error GoTo FunctionError
   '
   '
   If Trim(ExplicitPath) <> "" Then
      sSendFolder = ExplicitPath
   ElseIf IsSendByCDROM Then
      If mSendByCdromFolder = "" Then
         sSendFolder = mSendByMediaFolder
      Else
         sSendFolder = mSendByCdromFolder
      End If
   Else
      sSendFolder = mSendByMediaFolder
   End If
   
   If FullFilename = "" Then
      ' stop here
      goSession.RaiseError "Error in mwSession.mwWorkflow.SendByMedia, No File Specified."
      SendByMedia = False
      Exit Function
   ElseIf Not moFso.FileExists(FullFilename) Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.SendByMedia, File not found: " & FullFilename
      SendByMedia = False
      Exit Function
   End If
   '
   ' zip it good ?
   '
   If mIsZipWfFiles Then
      sFileToSend = CreateZipEncryptFile(FullFilename)
   Else
      sFileToSend = FullFilename
   End If
   If Trim(sFileToSend) = "" Then
      SendByMedia = False
      Exit Function
   End If
   
   If Trim(ExplicitPath) <> "" And Not ChooseTargetPath Then
      '
      ' Explicit Path specified
      '
      If Right(ExplicitPath, 1) <> "\" Then ExplicitPath = ExplicitPath & "\"
      If Not moFso.FolderExists(ExplicitPath) Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwWorkflow.SendByMedia, folder missing: " & ExplicitPath
         SendByMedia = False
         Exit Function
      End If
      strTarget = ExplicitPath & moFso.GetFileName(sFileToSend)
   ElseIf Trim(sSendFolder) <> "" And Not ChooseTargetPath Then
      '
      ' System determines SendByMedia Folder
      '
      If Right(sSendFolder, 1) <> "\" Then sSendFolder = sSendFolder & "\"
      If Not moFso.FolderExists(sSendFolder) Then
         goSession.RaiseWorkflowStatus "Error in mwSession.mwWorkflow.SendByMedia, SendByMedia folder missing: " & ExplicitPath
         SendByMedia = False
         Exit Function
      End If
      '
      ' Is TargetSite folder to be used ?
      '
      If TargetSiteID = "" Then
         strTarget = sSendFolder & moFso.GetFileName(sFileToSend)
      Else
         '
         ' Create folder if it is missing...
         '
         strTarget = sSendFolder & TargetSiteID
         If Not moFso.FolderExists(strTarget) Then
            moFso.CreateFolder strTarget
         End If
         strTarget = strTarget & "\" & moFso.GetFileName(sFileToSend)
      End If
   Else
      '
      ' OK, let the user decide which is right, and which is an illusion
      '
      Set oForm = New frmLocation
      oForm.SetParent moParent
      oForm.Show vbModal
      '
      If oForm.IsCancelled Then
         goSession.RaiseWorkflowStatus "Send By Media Cancelled By User..."
         SendByMedia = False
         Unload oForm
         Set oForm = Nothing
         Exit Function
      End If
      strTarget = oForm.GetPath
      If Right(strTarget, 1) <> "\" Then
         strTarget = strTarget & "\" & moFso.GetFileName(sFileToSend)
      Else
         strTarget = strTarget & moFso.GetFileName(sFileToSend)
      End If
      Unload oForm
      Set oForm = Nothing
   End If
   
   If TargetSiteKey > 0 Then
      sFileToSend = MaybeCreateZipInZipFile(sFileToSend, TargetSiteKey)
   End If
   
   strTarget = moFso.GetParentFolderName(strTarget) & "\" & moFso.GetFileName(sFileToSend)
   
   On Error Resume Next
   '
   ' Copy File to Target
   '
   moFso.CopyFile sFileToSend, strTarget
   If Err Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.SendByMedia. Error Copying File: " & sFileToSend & " to " & strTarget, _
        Err.Number, Err.Description
      SendByMedia = False
   Else
      goSession.RaiseWorkflowStatus "mwWorkflow.SendByMedia: Successfully Copied File: " _
        & sFileToSend & " to " & strTarget
      SendByMedia = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendByDemoLoopback.", Err.Number, Err.Description
   SendByMedia = False
End Function

Private Sub Class_Initialize()
   On Error GoTo SubError
   mIsFormsProcessingAllowed = True
   Set moKeys = goSession.GetEventSecurityKeys(MW_EVENT_WORKFLOW_AGENT)
   Set moFso = New FileSystemObject
   mLocalDecimalSeparator = goSession.API.GetLocaleString(LOCALE_DECIMAL_SEPARATOR)
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.Class_Initialize. ", Err.Number, Err.Description
End Sub
Private Sub Class_Terminate()
   Set moFso = Nothing
   Set moKeys = Nothing
   If Not moRecords Is Nothing Then Set moRecords = Nothing
End Sub

' SiteID omitted, send to all sites, RoleTypeID omitted send to all role types
Public Function SendRoleTypeNotification(Msg As String, SiteID As String, RoleTypeID As String) As Boolean
   Dim oRsRT As Recordset
   Dim strSQL As String
   Dim loform As mwForm
   On Error GoTo FunctionError
   '
   ' Construct the Datagram
   '
   Set loform = goSession.OpenNewForm()
   If Not loform.MakeNewForm("SYSTEM_CHANNEL") Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.SendRoleTypeNotification. SYSTEM_CHANNEL Form Template missing."
      SendRoleTypeNotification = False
      goSession.TearDownForm loform
      Exit Function
   End If
   With loform.DataForm.AddRecord
      .Add "Table", "mwwfRoleTypeNotification"
      .Add "Operation", "A"
      .Add "Message", Replace(Msg, vbCrLf, "|")
      .Add "DateCreated", Format(Now(), "Short Date")
      .Add "RoleTypeID", RoleTypeID
   End With
   '
   ' Now send the datagram
   '
   If Not loform.DataForm.WriteDataForm() Then
      SendRoleTypeNotification = False
      goSession.TearDownForm loform
      goSession.KillObject loform
      Exit Function
   End If
   '
   ' Transmit the message
   '
   loform.DeleteAfterSend = True
   If Not loform.Send(goSession.Site.GetShipWorkflowSendToAddress(SiteID)) Then
      ' Error sending email
      SendRoleTypeNotification = False
      goSession.TearDownForm loform
      goSession.KillObject loform
      Exit Function
   End If
   goSession.TearDownForm loform
   SendRoleTypeNotification = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendRoleTypeNotification. ", Err.Number, Err.Description
   SendRoleTypeNotification = False
   goSession.TearDownForm loform
   goSession.KillObject loform

End Function


Public Function ReceiveRoleTypeNotification() As Boolean
   Dim oRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError:
   '
   If Not goSession.IsLoggedIn Then Exit Function
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   strSQL = "select * from mwwfRoleTypeNotification where RoleTypeID = '" & _
     goSession.User.RoleType & "' and IsDisplayed=0"
   'strSQL = "select * from mwwfRoleTypeNotification where RoleTypeID = '" & _
   '  goSession.User.RoleType & "' and NOT IsDisplayed"
   oRs.Open strSQL, goCon, adOpenStatic, adLockOptimistic
   Do While Not oRs.EOF
      MsgBox Replace(oRs!Message, "|", vbCrLf), vbInformation, "ShipNet Fleet Notification Service"
      oRs!IsDisplayed = True
      oRs!DateDisplayed = Now()
      oRs.Update
      oRs.MoveNext
   Loop
   goSession.CloseRecordset oRs
   ReceiveRoleTypeNotification = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.mwWorkflow.ReceiveRoleTypeNotification", Err.Number, Err.Description
   ReceiveRoleTypeNotification = False
End Function

Private Function ConvertTF(str As String) As String
   ConvertTF = Replace(str, "|", vbCrLf)
End Function

Public Function SendAnyFile(FullFilename As String, Optional SendToAddress As Variant, _
 Optional Direction As mw_WhichWorkflowDirection, Optional EmailBody As String) As Boolean
   Dim strSendTo As String
   Dim strTarget As String
   On Error GoTo FunctionError
   strTarget = FullFilename
   '
   ' Does file exist ?
   '
   If Not moFso.FileExists(strTarget) Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.SendAnyFile. file missing: " & strTarget
      SendAnyFile = False
      Exit Function
   End If
   '
   ' Send the form...
   '
   If goSession.User.DefaultTransport = mwt_DemoLoopback And IsMissing(SendToAddress) Then
      If goSession.Workflow.SendByDemoLoopback(strTarget) Then
         SendAnyFile = True
      End If
   ElseIf goSession.User.DefaultTransport = mwt_Media And IsMissing(SendToAddress) Then
      If goSession.Workflow.SendByMedia(strTarget) Then
         SendAnyFile = True
      End If
   ElseIf goSession.User.DefaultTransport = mwt_TransportContainer Then
      '
      goSession.RaiseError "Sorry, temporarily closed for construction."
      SendAnyFile = False
   Else
      '
      ' Must be mail...
      '
      '
      ' Get email address from 3 possible places...
      '
      If IsMissing(SendToAddress) Then
         If Direction = 0 Or Direction = mw_SendToShip Then
            '
            ' Sending to the shore...
            '
            If Trim(goSession.Site.WorkflowSendToAddress) = "" Then
               goSession.RaiseError "Error in mwSession.mwWorkflow.SendAnyFile, no workflow email address found."
               SendAnyFile = False
               Exit Function
            Else
               strSendTo = goSession.Site.WorkflowSendToAddress
            End If
         Else
            '
            ' Sending to the Ship...
            '
            If Trim(goSession.Site.ShipWorkflowSendToAddress) = "" Then
               goSession.RaiseError "Error in mwSession.mwWorkflow.SendAnyFile, no workflow email address found."
               SendAnyFile = False
               Exit Function
            Else
               strSendTo = goSession.Site.ShipWorkflowSendToAddress
            End If
         End If
      Else
         strSendTo = SendToAddress
      End If
      '
      ' Send the File...
      '
      With goSession.Message
         If mIsZipAttachmentFile Then
            .ZipAttachment = True
         End If
         .AttachmentFile = strTarget
         .Server = goSession.User.MailServerName
         .SendFromAddress = goSession.User.GetExtendedProperty("FromEmailAddress")
         If Trim(strSendTo) <> "" Then
            .SendToAddress = strSendTo
         End If
         If EmailBody = "" Then
            .Body = strTarget & "File Attached. Sent by ShipNet Fleet."
            .subject = "Shipnet Fleet File sent from: " & goSession.Site.SiteName
         Else
            .Body = EmailBody
            .subject = Left(EmailBody, 70)
      End If
         .AttachmentOptions = mw_LEAVE_SOURCE
         '
         If .Send() Then
            SendAnyFile = True
         End If
      End With
   End If
   '
   ' Follow-up
   '
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendAnyFile.", Err.Number, Err.Description
   SendAnyFile = False
End Function



Private Function vrsProcessTranslations(ByRef loform As mwForm) As Boolean
   'Dim loRec As ConfigKeys
   Dim loKey As ConfigKey
   Dim loRs As Recordset
   Dim strOldValue As String
   Dim strSiteID As String
   Dim iRec As Integer
   Dim strSQL As String
   Dim strMessage As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSiteID = loform.DataForm.Records(2).ConfigKeys("VESSEL_CODE").KeyValue
   '
   strSQL = "select * from VRS_Translate where IncomingSiteID='" & strSiteID & "'"
   loRs.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.CloseRecordset loRs
      vrsProcessTranslations = True
      Exit Function
   End If
   '
   ' OK we have translations
   '
   
   Do While Not loRs.EOF
      '
      ' Translate ?
      '
      For iRec = 2 To loform.DataForm.Records.Count
         For Each loKey In loform.DataForm.Records(iRec).ConfigKeys
            Select Case loKey.KeyName
               Case Is = "VESSEL_CODE"
                  If loRs!TranslateType = 1 Then
                     If UCase(loKey.KeyValue) = UCase(loRs!IncomingCode) Then
                        strOldValue = loKey.KeyValue
                        loKey.KeyValue = loRs!TranslateToCode
                        goSession.RaiseWorkflowStatus "VRS Translation: VESSEL_ID, from " & strOldValue & _
                          " to " & loKey.KeyValue
                        '
                        ' Modify the vrs_Message Component...(which is imbedded in an Insert statement...)
                        '
                        If loform.DataForm.Records(loform.DataForm.Records.Count).ConfigKeys("1") Is Nothing Then
                           loform.DataForm.Records(loform.DataForm.Records.Count).ConfigKeys("VESSEL_CODE").KeyValue = loRs!TranslateToCode
                           ' 10/23/2002 for EWG Translate...
                        Else
                           strMessage = loform.DataForm.Records(loform.DataForm.Records.Count).ConfigKeys("1").KeyValue
                           strMessage = Replace(strMessage, strOldValue, loRs!TranslateToCode)
                           loform.DataForm.Records(loform.DataForm.Records.Count).ConfigKeys("1").KeyValue = strMessage
                        End If
                        loform.FormProperties.SiteID = loRs!TranslateToCode
                        loform.FormProperties.SiteKey = goSession.Site.GetSiteKey(loRs!TranslateToCode)
                     End If
                  End If
               
               Case Is = "VOYAGE_NO"
                  If loRs!TranslateType = 2 Then
                     If UCase(loKey.KeyValue) = UCase(loRs!IncomingCode) Then
                        strOldValue = loKey.KeyValue
                        loKey.KeyValue = loRs!TranslateToCode
                        goSession.RaiseWorkflowStatus "VRS Translation: VOYAGE_NO, from " & strOldValue & _
                          " to " & loKey.KeyValue
                     End If
                  End If
               Case Is = "PORTID"
                  If loRs!TranslateType = 3 Then
                     If UCase(loKey.KeyValue) = UCase(loRs!IncomingCode) Then
                        strOldValue = loKey.KeyValue
                        loKey.KeyValue = loRs!TranslateToCode
                        goSession.RaiseWorkflowStatus "VRS Translation: PORT ID, from " & strOldValue & _
                          " to " & loKey.KeyValue
                     End If
                  End If
               Case Is = "DEST_PORT"
                  If loRs!TranslateType = 3 Then
                     If UCase(loKey.KeyValue) = UCase(loRs!IncomingCode) Then
                        strOldValue = loKey.KeyValue
                        loKey.KeyValue = loRs!TranslateToCode
                        goSession.RaiseWorkflowStatus "VRS Translation: DEST PORT ID, from " & strOldValue & _
                          " to " & loKey.KeyValue
                     End If
                  End If
               Case Is = "PORT_NAME"
                  If loRs!TranslateType = 3 Then
                     If UCase(loKey.KeyValue) = UCase(loRs!IncomingCode) Then
                        strOldValue = loKey.KeyValue
                        loKey.KeyValue = loRs!TranslateToCode
                        goSession.RaiseWorkflowStatus "VRS Translation: PORT NAME, from " & strOldValue & _
                          " to " & loKey.KeyValue
                     End If
                  End If
            End Select
         Next loKey
      Next iRec
      loRs.MoveNext
   Loop
   goSession.CloseRecordset loRs
   vrsProcessTranslations = True
   Exit Function
FunctionError:
   If Err.Number = 91 Then
      ' Assume Sqlbase MakeInsertStatement - keep moving...
      'goSession..RaiseWorkflowStatus "Error in mwSession.mwWorkflow.vrsProcessTranslations, assume SqlBase Insert statement - keep trucking."
      goSession.CloseRecordset loRs
      vrsProcessTranslations = True
   Else
      goSession.RaiseError "General Error in mwSession.mwWorkflow.vrsProcessTranslations.", Err.Number, Err.Description
      vrsProcessTranslations = False
   End If
End Function


Public Function GetExtendedProperty(KeyName As String) As String
   On Error GoTo FunctionError
   'If Not mIsOpen Then
      'GetExtendedProperty = ""
      'Exit Function
   'End If
   GetExtendedProperty = moExtendedKeys.GetKeyValue(UCase(KeyName))
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(KeyName As String, KeyValue As String, Optional ID As String) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   'If Not mIsOpen Then
      SetExtendedProperty = False
      Exit Function
   'End If
   If Trim(ID) = "" Then
      If mIsPost0392 Then
         strSQL = "select * from mwcWorkflow where ID=" & mIDKey
      Else
         strSQL = "select * from mwcWorkflow where ID='" & mIDKey & "'"
      End If
   Else
      If mIsPost0392 Then
         strSQL = "select * from mwcWorkflow where ID=" & ID
      Else
         strSQL = "select * from mwcWorkflow where ID='" & ID & "'"
      End If
   End If
   If goSession.UpdateExtendedKeyData(KeyName, KeyValue, strSQL) Then
      moExtendedKeys(UCase(KeyName)).KeyValue = KeyValue
   End If
   SetExtendedProperty = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SetExtendedProperty. Invalid Key ?: " & KeyName, Err.Number, Err.Description
   SetExtendedProperty = False
End Function


Private Function ExecuteDllOnProcess(ByRef loform As mwForm) As Boolean
   Dim strVbDLL As String
   Dim loFormSession As Object
   On Error GoTo FunctionError
   '
   ' Run external VB DLL - if it exists...
   '
   ExecuteDllOnProcess = True
   strVbDLL = loform.WorkflowProperties.VbFormEventDLL
   If Trim(strVbDLL) <> "" Then
      strVbDLL = strVbDLL & ".FormSession"
      Set loFormSession = CreateObject(strVbDLL)
      Set loFormSession.ActiveMwSession = moParent
      Set loFormSession.ActiveMwForm = loform
      loFormSession.TriggerFormEvent ("OnProcess")
      If loFormSession.IsCancel Then
         '
         ' Cancel the process
         '
         goSession.RaiseWorkflowStatus "OnProcess DLL Code has cancelled workflow processing for: " & loform.FormProperties.FormID
         ExecuteDllOnProcess = False
      End If
      Set loFormSession = Nothing
   End If

   Exit Function
FunctionError:
   If Err.Number = 429 Or Err.Number = 13 Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.ExecuteDllOnProcess, unable to launch DLL: " & strVbDLL
   Else
      goSession.RaiseError "General Error in mwSession.mwWorkflow.ExecuteDllOnProcess. ", Err.Number, Err.Description
   End If
   ExecuteDllOnProcess = False

End Function

'Private Function SendDatagramEmails(ByRef loform As mwForm) As Boolean
'   Dim strEmailSubject As String
'   Dim strEmailTextBody As String
'   Dim loRs As Recordset
'   Dim strSQL As String
'   On Error GoTo FunctionError
'   '
'   ' Build the message
'   '
'   strEmailSubject = goSession.Site.GetShipProperty(loform.FormProperties.SiteID, "SiteName") & _
'     "-  Received: " & loform.TemplateProperties.TemplateID
'   strEmailTextBody = loform.DataForm.ConvertToEmailTextBody
'   If Trim(strEmailTextBody) <> "" And Trim(loform.GetExtendedProperty("ConvertDgToEmailAddress")) <> "" Then
'      goSession.SendNotification strEmailSubject, strEmailTextBody, _
'        loform.GetExtendedProperty("ConvertDgToEmailAddress")
'      goSession.RaiseWorkflowStatus "Datagram Email sent to: " & _
'        loform.GetExtendedProperty("ConvertDgToEmailAddress")
'   End If
'
'   strSQL = "SELECT mwwfFormTemplateEmailList.EmailAddress " & _
'    " FROM mwcFleets , mwcFleetSites, mwwfFormTemplateEmailList " & _
'    " WHERE mwcFleets.ID = mwcFleetSites.mwcFleetsKey AND " & _
'    " mwcFleets.ID = mwwfFormTemplateEmailList.FleetKey AND " & _
'    " mwwfFormTemplateEmailList.TemplateID = 'vrs_POSASS' AND " & _
'    " mwcFleetSites.mwcSitesKey = " & loform.FormProperties.SiteKey & " AND " & _
'    " mwcFleets.mwcFleetTypeKey = 1"
'
'
'   Set loRs = New Recordset
'   loRs.CursorLocation = adUseClient
'   loRs.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
'   Do While Not loRs.EOF
'      goSession.SendNotification strEmailSubject, strEmailTextBody, loRs!EmailAddress
'      goSession.RaiseWorkflowStatus "Datagram Email sent to: " & loRs!EmailAddress
'      loRs.MoveNext
'   Loop
'   goSession.CloseRecordset loRs
'   SendDatagramEmails = True
'   Exit Function
'FunctionError:
'   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendDatagramEmails. ", Err.Number, Err.Description
'   SendDatagramEmails = False
'End Function

Private Function SendDatagramEmails(ByRef loform As mwForm) As Boolean
   Dim sEmailSubject As String
   Dim sEmailTemplate As String
   Dim sVbDLL As String
   Dim nSiteKey As Long
   Dim loFormSession As Object
   Dim sEmailTextBody As String
   Dim loRs As Recordset
   Dim ts As TextStream
   Dim sSQL As String
   On Error GoTo FunctionError
   '
   ' Build the message
   '
   nSiteKey = goSession.Site.GetSiteKey(loform.FormProperties.SiteID)
   sSQL = "SELECT mwwfFormTemplateEmailList.* FROM (mwcFleets INNER JOIN mwwfFormTemplateEmailList ON " & _
     " mwcFleets.ID = mwwfFormTemplateEmailList.FleetKey) INNER JOIN mwcFleetSites ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey " & _
     " WHERE IsShipsideEmail=0 and mwcFleetSites.mwcSitesKey=" & nSiteKey & _
     " AND mwwfFormTemplateEmailList.TemplateID='" & loform.TemplateProperties.TemplateID & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      If Not IsNull(loRs!EmailTemplatePath) Then
         sEmailTemplate = goSession.Workflow.ScriptFolder & "\" & loRs!EmailTemplatePath
      End If
      sEmailSubject = goSession.Site.GetShipProperty(loform.FormProperties.SiteID, "SiteName") & _
        "-  Received: " & loform.TemplateProperties.TemplateID
      '
      ' Custom DLL option...
      '
      If Trim(loform.WorkflowProperties.VbFormEventDLL) <> "" Then
         '
         ' Run external VB DLL - if it exists...
         '
         sVbDLL = loform.WorkflowProperties.VbFormEventDLL & ".FormSession"
         Set loFormSession = CreateObject(sVbDLL)
         Set loFormSession.ActiveMwSession = goSession
         Set loFormSession.ActiveMwForm = loform
         loFormSession.TriggerFormEvent ("OnDatagramEmail")
         If loFormSession.IsCancel Then
            '
            ' Cancel the process
            '
            'goSession.RaiseWorkflowStatus "OnDatagramEmail DLL Code has cancelled workflow processing for: " & loForm.FormProperties.FormID
            sEmailTextBody = ""
         Else
            sEmailTextBody = loFormSession.GetEmailTextBody
         End If
         '
         Set loFormSession = Nothing
         
      Else
         '
         ' Read in Template...
         '
         sEmailTextBody = ""
         If moFso.FileExists(sEmailTemplate) Then
            Set ts = moFso.OpenTextFile(sEmailTemplate, ForReading, False)
            sEmailTextBody = ts.ReadAll
            ts.Close
            goSession.KillObject ts
         End If
      End If
      '
      ' Basic Datagram translation...
      '
      If Trim(sEmailTextBody) <> "" Then
         sEmailTextBody = loform.DataForm.ConvertBasicDgText(sEmailTextBody)
      
         If UCase(Left(loform.TemplateProperties.TemplateID, 4)) = "VRS_" Then
            sEmailTextBody = ConvertCustomVrsDgText(sEmailTextBody, loform.DataForm.Records)
            
            ' strip vrs unused pairs from email .textbody
            sEmailTextBody = StripUnusedEmailFields(sEmailTextBody)
                        
         End If
      End If
      If Trim(sEmailTextBody) <> "" And Not IsNull(loRs!EmailAddress) Then
         '
         ' Send the email...
         '
         goSession.SendNotification sEmailSubject, sEmailTextBody, loRs!EmailAddress
         goSession.RaiseWorkflowStatus "Datagram Email sent to: " & loRs!EmailAddress
      Else
         goSession.RaiseWorkflowStatus "SKIPPED Datagram Email, either Email Text Body or email address are blank."
      End If
      loRs.MoveNext
   Loop
   goSession.CloseRecordset loRs
   SendDatagramEmails = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendDatagramEmails. ", Err.Number, Err.Description
   SendDatagramEmails = False
End Function


Public Function ProcessWorkflowSchedule() As Boolean
   Dim loform As Object
   Dim loCmd As Command
   Dim loPar As Parameter
   Dim loRs As Recordset
   Dim strVbDLL As String
   Dim loScheduleSession As Object
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goSession.DBConnection
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT mwwfSchedule.* FROM mwwfSchedule, mwcWorkflow, mwcWorkflowSched " & _
          " where mwcWorkflow.ID = mwcWorkflowSched.mwcWorkflowKey and " & _
          " mwwfSchedule.ID = mwcWorkflowSched.mwwfScheduleKey and " & _
          " mwcWorkflow.ID=" & mIDKey & " and NextScheduled < ? and IsEnabled<>0 and" & _
          " mwcWorkflowSched.IsActive<>0  "
   If goSession.IsSqlServer Then
      sSQL = sSQL & "ORDER BY ISNULL(ExecutionSequence,99999)"
   ElseIf goSession.IsOracle Then
      sSQL = sSQL & "ORDER BY NVL(ExecutionSequence,99999)"
   Else
      sSQL = sSQL & "ORDER BY IIF(ISNULL(ExecutionSequence),99999,ExecutionSequence)"
   End If
   'sSQL = "select * from mwwfSchedule where NextScheduled < ? and IsEnabled<>0"
   loCmd.CommandText = sSQL
   Set loPar = loCmd.CreateParameter("NowDate", adDate, adParamInput, , Now())
   loCmd.Parameters.Append loPar
   loRs.Open loCmd, , adOpenDynamic, adLockOptimistic
   If loRs.RecordCount < 1 Then
      goSession.CloseRecordset loRs
      ProcessWorkflowSchedule = True
      Exit Function
   End If
   '
   ' OK, let's do it...
   '
   Do While Not loRs.EOF
      '
      '
      If Not IsNull(loRs!VbClassObject) Then
         strVbDLL = loRs!VbClassObject
         If Trim(strVbDLL) = "" Then
            goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowSchedule, missing VbClassObject value in the Event Type table."
         Else
            strVbDLL = strVbDLL & ".ScheduleSession"
            On Error Resume Next
            Set loScheduleSession = CreateObject(strVbDLL)
            If Err Then
               goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowSchedule, Unable to launch: " & strVbDLL
               On Error GoTo FunctionError
            Else
               On Error GoTo FunctionError
               Set loScheduleSession.ActiveMwSession = moParent
               
               If Trim(loRs!VbDefaultForm) <> "" Then
                  loScheduleSession.TriggerScheduleEvent (loRs!VbDefaultForm)
               Else
                  If Not loScheduleSession.TriggerScheduleEvent() Then
                     mIsWorkflowError = True
                  End If
               End If
            End If
         End If
      End If
      '
      ' Update Last Scheduled, Increment NextScheduled...
      '
      goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwSession.mwWorkflow.ProcessWorkflowSchedule, Successfully executed: " & _
        loRs!ID & " - " & loRs!Description
        
      UpdateWorkflowSchedule loRs!ID, Now(), DateAdd("n", ZeroNull(loRs!IntervalMinutes), DateAdd("h", ZeroNull(loRs!IntervalHours), Now()))
'      loRs!LastCompleted = Now()
'      loRs!NextScheduled = DateAdd("n", ZeroNull(loRs!IntervalMinutes), DateAdd("h", ZeroNull(loRs!IntervalHours), Now()))
'      loRs.Update
      loRs.MoveNext
   Loop
   goSession.CloseRecordset loRs
   KillObject loScheduleSession
   ProcessWorkflowSchedule = True
   Exit Function
FunctionError:
   If Err.Number = 13 Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.ProcessWorkflowSchedule, unable to launch DLL: " & strVbDLL
   Else
      goSession.RaiseError "General Error in mwSession.mwWorkflow.ProcessWorkflowSchedule. ", Err.Number, Err.Description
   End If
   goSession.CloseRecordset loRs
   ProcessWorkflowSchedule = False
End Function

Public Sub UpdateWorkflowSchedule(mwwfScheduleKey As Long, LastCompleted As Date, NextScheduled As Date)
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo SubError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * FROM mwwfSchedule WHERE ID = " & mwwfScheduleKey
   
   loRs.Open sSQL, goSession.DBConnection, adOpenDynamic, adLockPessimistic
   
   If loRs.RecordCount < 1 Then
      goSession.CloseRecordset loRs
      Exit Sub
   End If
   loRs!LastCompleted = LastCompleted
   loRs!NextScheduled = NextScheduled
   loRs.Update
   
   goSession.CloseRecordset loRs
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.UpdateWorkflowSchedule. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
End Sub

Private Function ExecuteDllAfterProcess(ByRef loform As mwForm) As Boolean
   Dim strVbDLL As String
   Dim loFormSession As Object
   On Error GoTo FunctionError
   '
   ' Run external VB DLL - if it exists...
   '
   ExecuteDllAfterProcess = True
   strVbDLL = loform.WorkflowProperties.VbFormEventDLL
   If Trim(strVbDLL) <> "" Then
      strVbDLL = strVbDLL & ".FormSession"
      Set loFormSession = CreateObject(strVbDLL)
      Set loFormSession.ActiveMwSession = moParent
      Set loFormSession.ActiveMwForm = loform
      loFormSession.TriggerFormEvent ("AfterProcess")
      Set loFormSession = Nothing
   End If
   Exit Function
FunctionError:
   'If err.Number = 429 Or err.Number = 13 Then
   '   goSession.RaiseError "Error in mwSession.mwWorkflow.ExecuteDllAfterProcess, unable to launch DLL: " & strVbDLL
   'Else
   '   goSession.RaiseError "General Error in mwSession.mwWorkflow.ExecuteDllAfterProcess. ", err.Number, err.Description
   'End If
   ExecuteDllAfterProcess = False
   Set loFormSession = Nothing
End Function


Public Function ForceScheduleDate(NewTime As Date, Optional ScheduleKey As Long) As Boolean
   'Dim loRS As Recordset
   'Dim sSQL As String
   Dim loCmd As Command
   Dim loPar As Parameter
   On Error GoTo FunctionError
   'Set loRS = New Recordset
   'loRS.CursorLocation = adUseClient
   'sSQL = "select * from mwwfSchedule where IsEnabled<>0"
   'If ScheduleKey > 0 Then
   '   sSQL = sSQL & " and ID=" & ScheduleKey
   'End If
   'loRS.Open sSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   'Do While Not loRS.EOF
   '   loRS.Fields("NextScheduled") = NewTime
   '   loRS.Update
   '   loRS.MoveNext
   'Loop
   'CloseRecordset loRS
   Set loCmd = New Command
   Set loCmd.ActiveConnection = goCon
   loCmd.CommandType = adCmdText
   loCmd.CommandText = "update mwwfSchedule set NextScheduled=? where IsEnabled<>0 and (IsSkipForceNow is null or IsSkipForceNow = 0)"
   
   loCmd.Parameters.Append loCmd.CreateParameter("NewDate", adDate, adParamInput, , NewTime)
   loCmd.Execute
   Set loCmd = Nothing
   goSession.API.SleepApp 800
   ForceScheduleDate = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ForceScheduleDate. ", Err.Number, Err.Description
   ForceScheduleDate = False
End Function


Private Function vrsProcessOldRemarks(ByRef loform As mwForm) As Boolean
   Dim i As Integer
   Dim s1 As String
   Dim s2 As String
   On Error GoTo FunctionError
   s1 = loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK1")
   '
   ' Merge old remarks into new Remarks memo field...
   '
   s2 = loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK2")
   If s2 <> "" Then
      s1 = s1 & "|" & s2
   End If
   
   s2 = loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK3")
   If s2 <> "" Then
      s1 = s1 & "|" & s2
   End If
   
   s2 = loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK4")
   If s2 <> "" Then
      s1 = s1 & "|" & s2
   End If
   
   s2 = loform.DataForm.Records(2).ConfigKeys.GetKeyValue("REMARK5")
   If s2 <> "" Then
      s1 = s1 & "|" & s2
   End If
   s1 = s1 & "|"
   '
   ' Add Remarks to the little bugger...
   '
   loform.DataForm.Records(2).ConfigKeys.Add "REMARKS", s1, "REMARKS"
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ForceScheduleDate. ", Err.Number, Err.Description
End Function



Public Function GetTargetFolder(Optional DefaultFolder As String, Optional OptCaption As String) As String
   Dim loform As frmLocation
   Dim PromptForFolder As Boolean
   
   On Error GoTo FunctionError
   '
   ' OK, let the user decide which is right, and which is an illusion
   '
   Set loform = New frmLocation
   loform.SetParent moParent
   If DefaultFolder <> "" And moFso.FolderExists(DefaultFolder) Then
      loform.SetDrive moFso.GetDriveName(DefaultFolder)
      loform.SetFolder DefaultFolder
   End If
   If OptCaption <> "" Then
      loform.SetOptCaption OptCaption
   End If
   
   loform.SetIsNotSendByMedia
   PromptForFolder = True
   
   While PromptForFolder = True
   
      loform.Show vbModal
      '
      If loform.IsCancelled Then
         GetTargetFolder = ""
         PromptForFolder = False
      Else
         GetTargetFolder = loform.GetPath
         If moFso.FolderExists(GetTargetFolder) = False Then
            MsgBox "Folder " & GetTargetFolder & " does not exist. Try again please. Press the Cancel button to use the default location.", vbOKOnly, "Folder does not exist"
            GetTargetFolder = ""
         Else
            PromptForFolder = False
         End If
      End If
   Wend
   
   Unload loform
   Set loform = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendByDemoLoopback.", Err.Number, Err.Description
   GetTargetFolder = False
End Function


Private Function UpdateUserData(field As String, value As Variant)
   Dim strSQL As String
   Dim oRs As Recordset
   On Error GoTo FunctionError
   '
   ' Don't Update User when loading user...
   '
   If mIsLoadWorkflow Then
      Exit Function
   End If
   '
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   strSQL = "select * from mwcWorkflow where WorkflowID='" & mWorkflowID & "'"
   oRs.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockPessimistic
   If oRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwSession.mwWorkflow.UpdateUserData, Workflow record missing: " & mWorkflowID, Err.Number, Err.Description
   Else
      oRs.Fields(field) = value
      oRs.Update
   End If
   goSession.CloseRecordset oRs
  Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mWorkflowID.UpdateUserData. ", Err.Number, Err.Description
   goSession.CloseRecordset oRs
End Function

Public Function SendByEmail(FullFilename As String, Optional TextBodyOverride As String, _
                Optional EmailAddress As String, Optional SubmitEmailAsDraft As Boolean, _
                Optional TargetSiteKey As Long) As Boolean
   Dim sSendTo As String
   Dim sFileToSend As String
   Dim nTransportMethod As Long
   On Error GoTo FunctionError
   '
   ' Get email address from 3 possible places...
   '
   goSession.MemoryLogIt "mwWorkflow.SendByEmail 1 "
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      '
      ' Sending to the shore...
      '
      If EmailAddress <> "" Then
         sSendTo = EmailAddress
      ElseIf Trim(goSession.Site.WorkflowSendToAddress) = "" Then
         goSession.RaiseError "Error in mwSession.mwWorkflow.Send, no workflow email address found."
         SendByEmail = False
         Exit Function
      Else
         sSendTo = goSession.Site.WorkflowSendToAddress
      End If
   Else
      '
      ' Sending to the Ship...
      '
      If EmailAddress <> "" Then
         sSendTo = EmailAddress
      ElseIf Trim(goSession.Site.ShipWorkflowSendToAddress) = "" Then
         goSession.RaiseError "Error in mwSession.MwForm.Send, no workflow email address found."
         SendByEmail = False
         Exit Function
      Else
         sSendTo = goSession.Site.ShipWorkflowSendToAddress
      End If
   End If
   '
   ' zip it good ?
   '
   goSession.MemoryLogIt "mwWorkflow.SendByEmail 2 "
   If mIsZipWfFiles Then
      sFileToSend = CreateZipEncryptFile(FullFilename)
   Else
      sFileToSend = FullFilename
   End If
   
   If TargetSiteKey > 0 Then
      sFileToSend = MaybeCreateZipInZipFile(sFileToSend, TargetSiteKey)
   End If
   
   If Trim(sFileToSend) = "" Then
      SendByEmail = False
      Exit Function
   End If
'   If goSession.User.DefaultTransport = mwt_Media Then
'      '
'      '
'      '
'      SendByEmail = SendByMedia(FullFilename, True, , , , SEND_CDROM_FALSE)
'   Else
      '
      ' send it good
      '
   goSession.MemoryLogIt "mwWorkflow.SendByEmail 3 "
      
      With goSession.Message
         .AttachmentFile = sFileToSend
         .Server = goSession.User.MailServerName
'         .SendFromAddress = goSession.User.GetExtendedProperty("FromEmailAddress")
         .SendFromAddress = goSession.Site.WorkflowSendFromAddress
         .SendToAddress = sSendTo
         If SubmitEmailAsDraft Then
            .UserResolveAddress = True
         Else
            .UserResolveAddress = False
         End If
         If Trim(TextBodyOverride) = "" Then
            .Body = "<Datagram Attached For Processing>"
            .subject = "Workflow File from: " & goSession.Site.SiteName & " received"
         Else
            .Body = TextBodyOverride
            .subject = TextBodyOverride
         End If
         '.docID = moFormProperties.FormID
         .AttachmentOptions = mw_LEAVE_SOURCE
         '
         If .Send() Then
            SendByEmail = True
         End If
      End With
'   End If
   goSession.MemoryLogIt "mwWorkflow.SendByEmail exit "
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.SendByEmail.", Err.Number, Err.Description
   goSession.MemoryLogIt "mwWorkflow.SendByEmail err exit "
End Function

Public Function CreateZipEncryptFile(FullFilename As String) As String
   Dim sNewFile As String
   Dim sNewExt As String
   Dim loZip As Zipper
   Dim loFso As FileSystemObject
   
   On Error GoTo FunctionError
   
   goSession.MemoryLogIt "mwWorkflow.CreateZipEncryptFile 1 "
   Set loFso = New FileSystemObject
   ' validation
   If Trim(FullFilename) = "" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.CreateZipEncryptFile, blank filename provided to function."
      Exit Function
   ElseIf Not loFso.FileExists(FullFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.CreateZipEncryptFile, file not found: " & FullFilename
      Exit Function
   End If
   ' Create new Filename
   Select Case UCase(loFso.GetExtensionName(FullFilename))
      Case Is = "DGM"
         sNewExt = "DGZ"
      Case Is = "PKG"
         sNewExt = "PKZ"
      Case Is = "FCB"
         sNewExt = "FCZ"
      Case Is = "FMB"
         sNewExt = "FMZ"
      Case Is = "DAB"
         sNewExt = "DAZ"
      Case Else
         CreateZipEncryptFile = FullFilename
         Exit Function
   End Select
   sNewFile = loFso.GetParentFolderName(FullFilename) & "\" & loFso.GetBaseName(FullFilename) & "." & sNewExt
   ' zip it good
   goSession.MemoryLogIt "mwWorkflow.CreateZipEncryptFile 2 "
   Set loZip = New Zipper
   If mIsZipEncryptWfFiles Then
      loZip.IsEncypt = True
   End If
   loZip.OpenZip (sNewFile)
   If Not loZip.IsOpen Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.CreateZipEncryptFile, unable to create new file: " & sNewFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.AddToZip(FullFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.CreateZipEncryptFile, unable to add file to new file: " & sNewFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   loZip.TerminateClass
   Set loZip = Nothing
   Set loFso = Nothing
   
   CreateZipEncryptFile = sNewFile
   goSession.MemoryLogIt "mwWorkflow.CreateZipEncryptFile exit "
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.CreateZipEncryptFile.", Err.Number, Err.Description
End Function
Public Function MaybeCreateZipInZipFile(FullFilename As String, mwcSitesKey As Long) As String
   Dim sNewFile As String
   Dim sIntermediateFile As String
   Dim loZip As Zipper
   Dim IsZipInZip As Boolean
   Dim sExt As String
   
   On Error GoTo FunctionError
   
   goSession.MemoryLogIt "mwWorkflow.MaybeCreateZipInZipFile 1 "
   
   ' validation
   If Trim(FullFilename) = "" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, blank filename provided to function."
      Exit Function
   ElseIf Not moFso.FileExists(FullFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, file not found: " & FullFilename
      Exit Function
   ElseIf mwcSitesKey <= 0 Then
      Exit Function
   Else
      If goSession.Site.GetShipProperty("", "IsZipInZip", mwcSitesKey) = "" Then
         IsZipInZip = False
      Else
         IsZipInZip = goSession.Site.GetShipProperty("", "IsZipInZip", mwcSitesKey)
      End If
         
      If IsZipInZip = False Then
         MaybeCreateZipInZipFile = FullFilename
         Exit Function
      End If
   End If
   
   sExt = UCase(moFso.GetExtensionName(FullFilename))
   
   ' Is it already a ZIZ? No need to encapsulate it again.
   If sExt = "ZIZ" Then
      MaybeCreateZipInZipFile = FullFilename
      Exit Function
   End If
   
   ' Create new Filename
   sIntermediateFile = moFso.GetParentFolderName(FullFilename) & "\" & moFso.GetBaseName(FullFilename) & ".EZZ"
   
   ' zip it good
   Set loZip = New Zipper
   loZip.IsEncypt = True
   loZip.OpenZip (sIntermediateFile)
   If Not loZip.IsOpen Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, unable to create new file: " & sIntermediateFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.AddToZip(FullFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, unable to add file to new file: " & sIntermediateFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   
   loZip.TerminateClass
   Set loZip = Nothing
   
   sNewFile = moFso.GetParentFolderName(FullFilename) & "\" & moFso.GetBaseName(FullFilename) & ".ZIZ"
   
   ' zip it good
   Set loZip = New Zipper
   loZip.IsEncypt = False
   loZip.DeleteAfterZip = True
   
   loZip.OpenZip (sNewFile)
   If Not loZip.IsOpen Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, unable to create new file: " & sNewFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.AddToZip(sIntermediateFile) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile, unable to add file to new file: " & sIntermediateFile
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   loZip.TerminateClass
   Set loZip = Nothing
   
   MaybeCreateZipInZipFile = sNewFile
   goSession.MemoryLogIt "mwWorkflow.MaybeCreateZipInZipFile 1 "
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.MaybeCreateZipInZipFile.", Err.Number, Err.Description
End Function

Public Function UnpackZipEncryptFile(FullZipFilename As String) As Boolean
   Dim sNewFile As String
   Dim sExt As String
   Dim loZip As Zipper
   On Error GoTo FunctionError
   ' validation
   If Trim(FullZipFilename) = "" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, blank filename provided to function."
      UnpackZipEncryptFile = False
      Exit Function
   ElseIf Not moFso.FileExists(FullZipFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, file not found: " & FullZipFilename
      UnpackZipEncryptFile = False
      Exit Function
   End If
   ' Create new Filename
   sExt = UCase(moFso.GetExtensionName(FullZipFilename))
   
'   If mID$(sExt, 1, 3) = "SNO" Then
'      sExt = mID$(sExt, 4)
'   End If
   
   If sExt <> "DGZ" And sExt <> "PKZ" And sExt <> "FCZ" And sExt <> "FMZ" And sExt <> "DAZ" And sExt <> "EZZ" And sExt <> "ZIZ" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, only DGZ, PKZ, FCZ, and DAZ file types supported for compression."
      UnpackZipEncryptFile = False
      Exit Function
   End If
   ' zip it good
   Set loZip = New Zipper
   If mIsZipEncryptWfFiles Then
      loZip.IsEncypt = True
      loZip.Password = ENCRYPT_PASSWORD
   End If
   
   
   If Not loZip.OpenZip(FullZipFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, unable to open file: " & FullZipFilename
      UnpackZipEncryptFile = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.GetFromZip("*.*", moFso.GetParentFolderName(FullZipFilename)) Then
   
      If mIsZipEncryptWfFiles = False Then
         loZip.IsEncypt = True
         loZip.Password = ENCRYPT_PASSWORD
      
         goSession.LogIt mwl_Workstation, mwl_Warning, "Error unZipping, trying again with Encryption ON"
      Else
         loZip.IsEncypt = False
         loZip.Password = ""
      
         goSession.LogIt mwl_Workstation, mwl_Warning, "Error unZipping, trying again with Encryption OFF"
      End If
      If Not loZip.GetFromZip("*.*", moFso.GetParentFolderName(FullZipFilename)) Then
         goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, unable to extract file: " & FullZipFilename
         UnpackZipEncryptFile = False
         loZip.TerminateClass
         Set loZip = Nothing
         Exit Function
      End If
'      Else
'         goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile, unable to extract file: " & FullZipFilename
'         UnpackZipEncryptFile = False
'         loZip.TerminateClass
'         Set loZip = Nothing
'         Exit Function
'      End If
   End If
   loZip.TerminateClass
   Set loZip = Nothing
   UnpackZipEncryptFile = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipEncryptFile.", Err.Number, Err.Description
   UnpackZipEncryptFile = False
End Function

Public Function UnpackZipInZipFile(FullZipFilename As String) As Boolean
   Dim sNewFile As String
   Dim sExt As String
   Dim loZip As Zipper
   On Error GoTo FunctionError
   
   ' validation
   If Trim(FullZipFilename) = "" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, blank filename provided to function."
      UnpackZipInZipFile = False
      Exit Function
   ElseIf Not moFso.FileExists(FullZipFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, file not found: " & FullZipFilename
      UnpackZipInZipFile = False
      Exit Function
   End If
   
   ' Create new Filename
   sExt = UCase(moFso.GetExtensionName(FullZipFilename))
   
'   If mID(sExt, 1, 3) = "SNO" Then
'      sExt = mID(sExt, 4)
'   End If
   
   If sExt <> "ZIZ" Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, only ZIZ file types supported."
      UnpackZipInZipFile = False
      Exit Function
   End If
   
   ' Extract the intermediate EZZ file.
   Set loZip = New Zipper
   loZip.IsEncypt = False
   
   If Not loZip.OpenZip(FullZipFilename) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, unable to open file: " & FullZipFilename
      UnpackZipInZipFile = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.GetFromZip("*.*", moFso.GetParentFolderName(FullZipFilename)) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, unable to extract file: " & FullZipFilename
      UnpackZipInZipFile = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   loZip.TerminateClass
   Set loZip = Nothing
   
   sNewFile = moFso.GetParentFolderName(FullZipFilename) & "\" & moFso.GetBaseName(FullZipFilename) & ".EZZ"
   
   If Not moFso.FileExists(sNewFile) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, file not found: " & sNewFile
      UnpackZipInZipFile = False
      Exit Function
   End If
   
   Set loZip = New Zipper
      
   loZip.IsEncypt = True
   loZip.Password = ENCRYPT_PASSWORD
   
   If Not loZip.OpenZip(sNewFile) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, unable to open file: " & sNewFile
      UnpackZipInZipFile = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.GetFromZip("*.*", moFso.GetParentFolderName(sNewFile)) Then
      goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile, unable to extract file: " & sNewFile
      UnpackZipInZipFile = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   loZip.TerminateClass
   Set loZip = Nothing
   
   moFso.DeleteFile sNewFile
   
   UnpackZipInZipFile = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.UnpackZipInZipFile.", Err.Number, Err.Description
   UnpackZipInZipFile = False
End Function

Private Function GetRouteToFolder(oForm As mwForm) As String
   Dim sOption As String
   Dim sFolder As String
   On Error GoTo FunctionError
   sOption = Trim(oForm.GetExtendedProperty("RouteToFolderOption"))
   If sOption = "" Then
      GetRouteToFolder = oForm.WorkflowProperties.RouteToFolder
      Exit Function
   ElseIf oForm.FormProperties.SiteID = "" Then
      goSession.RaiseWorkflowStatus "Warning in mwSession.mwWorkflow.GetRouteToFolder, cannot make site folder (SiteID is empty) for Form: " & oForm.FormProperties.FormID
      GetRouteToFolder = oForm.WorkflowProperties.RouteToFolder
      Exit Function
   End If
   sFolder = oForm.WorkflowProperties.RouteToFolder & "\" & oForm.FormProperties.SiteID
   If Not moFso.FolderExists(sFolder) Then
      moFso.CreateFolder sFolder
   End If
   If sOption = "2" Then 'Voyage Folder
      If Trim(goSession.Site.CurrentVoyage) = "" Then
         goSession.RaiseWorkflowStatus "Warning in mwSession.mwWorkflow.GetRouteToFolder, cannot make site\voy folder (CurrentVoyage is empty) for ship: " & goSession.Site.SiteName
      Else
         sFolder = sFolder & "\" & LTrim(goSession.Site.CurrentVoyage)
         If Not moFso.FolderExists(sFolder) Then
            moFso.CreateFolder sFolder
         End If
      End If
   ElseIf sOption = "3" Then 'Month Folder
      sFolder = sFolder & "\" & LTrim(Format(Now(), "MMM"))
      If Not moFso.FolderExists(sFolder) Then
         moFso.CreateFolder sFolder
      End If
   End If
   GetRouteToFolder = sFolder
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.GetRouteToFolder.", Err.Number, Err.Description
End Function

Private Function ConvertCustomVrsDgText(TextBody As String, loRecords As ConfigGroups) As String
   Dim sTextBody As String
   Dim sShipCode As String
   Dim sNGradeBunkers As String
   Dim sPortCallKey As String
   Dim sAtSeaPerfKey As String
   Const EtPORTCALL = 15
   Const EtAtSea = 3
   
   Dim sTemplateID As String
   Dim nDetailKey As Long
   Dim nDetailType As Integer
   
   On Error GoTo FunctionError
   
   sTextBody = TextBody
   sTextBody = Replace(sTextBody, "[VRS_BUNKERS1_LABEL]", goSession.Site.GetExtendedProperty("vrsBunkers1Label"))
   sTextBody = Replace(sTextBody, "[VRS_BUNKERS2_LABEL]", goSession.Site.GetExtendedProperty("vrsBunkers2Label"))
   sTextBody = Replace(sTextBody, "[VRS_ROB1_LABEL]", goSession.Site.GetExtendedProperty("vrsROB1Label"))
   sTextBody = Replace(sTextBody, "[VRS_ROB2_LABEL]", goSession.Site.GetExtendedProperty("vrsROB2Label"))
   ' Vessel Code and Vessel name as well...
   If Not loRecords(2).ConfigKeys("vessel_code") Is Nothing Then
      sShipCode = loRecords(2).ConfigKeys("vessel_code").KeyValue
      sTextBody = Replace(sTextBody, "[VESSEL_CODE]", sShipCode)
      sTextBody = Replace(sTextBody, "[VESSEL_NAME]", goSession.Site.GetShipProperty(sShipCode, "SiteName"))
   End If

   If moRecords Is Nothing Then
      Set moRecords = Nothing
   End If
   
   Set moRecords = loRecords
   
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
   
   ElseIf sTemplateID = "VRS_POSASS" Then
   
      If Not moRecords(2).ConfigKeys("POSASSKEY") Is Nothing Then
         nDetailKey = moRecords(2).ConfigKeys("POSASSKEY").KeyValue
         nDetailType = DT_ATSEA
      End If
   
   End If
   
   
   If nDetailKey > 0 Then
   
      FetchNGradeBunkers nDetailKey, nDetailType
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADEBUNKERS]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROBCONS]")
   
      CloseNGradeBunkers
   End If
   
   
   'sTextBody = ConvertCustomVrsNGrades(sTextBody)
   Set moRecords = Nothing
   
   
   ' Strip out unused remarks...
   sTextBody = Replace(sTextBody, "[REMARK1]", "")
   sTextBody = Replace(sTextBody, "[REMARK2]", "")
   sTextBody = Replace(sTextBody, "[REMARK3]", "")
   sTextBody = Replace(sTextBody, "[REMARK4]", "")
   sTextBody = Replace(sTextBody, "[REMARK5]", "")
   ConvertCustomVrsDgText = sTextBody
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ConvertCustomVrsDgText. ", Err.Number, Err.Description
   ConvertCustomVrsDgText = TextBody
End Function

Private Function GetSortedInboundRS() As Recordset
   Dim loRs As Recordset
   Dim loFolder As Folder
   Dim loFile As file
   Dim nCount As Long
   On Error GoTo FunctionError
   Set loFolder = moFso.GetFolder(goSession.Workflow.PollingFolder)
   Set loRs = New Recordset
   loRs.Fields.Append "ID", adInteger, 4
   loRs.Fields.Append "Status", adVarWChar, 50
   loRs.Fields.Append "Path", adVarWChar, 255
   loRs.Fields.Append "Name", adVarWChar, 255
   loRs.Fields.Append "ChangeID", adVarWChar, 50
   loRs.Fields.Append "Description", adVarWChar, 255
   loRs.Fields.Append "Message", adVarWChar, 255
   loRs.Open
   nCount = 1
   For Each loFile In loFolder.Files
      loRs.AddNew
      loRs!ID = nCount
      nCount = nCount + 1
      loRs!Status = ""
      loRs!Path = loFile.Path
      loRs!Name = loFile.Name
      loRs!ChangeID = ""
      loRs!Description = ""
      loRs!Message = ""
      loRs.Update
   Next loFile
   If loRs.RecordCount > 0 Then
      loRs.MoveFirst
      loRs.Sort = "path"
   End If
   Set GetSortedInboundRS = loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.GetSortedInboundRS. ", Err.Number, Err.Description
End Function

Private Sub MakeFolderExist(strFolder As String)
   
   On Error GoTo ErrorHandler
   
   If Len(strFolder) > 0 Then
      If Not moFso.FolderExists(strFolder) Then
         On Error Resume Next
         moFso.CreateFolder (strFolder)
         If Err Then
            MsgBox "Unlogged error in mwSession.mwWorkflow.MakeFolderExist. Unable to Create Working folder " & strFolder & ", system will not operate correctly."
         End If
      End If
   End If
   Exit Sub
ErrorHandler:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.MakeFolderExist. ", Err.Number, Err.Description
End Sub

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
   
   ElseIf sTemplateID = "VRS_POSASS" Then
   
      If Not moRecords(2).ConfigKeys("POSASSKEY") Is Nothing Then
         nDetailKey = moRecords(2).ConfigKeys("POSASSKEY").KeyValue
         nDetailType = DT_ATSEA
      End If
   
   End If
   
   
   If nDetailKey > 0 Then
   
      FetchNGradeBunkers nDetailKey, nDetailType
      
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADEBUNKERS]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROB]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_CONS]")
      sTextBody = ProcessNGradeBunkers(sTextBody, "[VRS_NGRADE_ROBCONS]")
   
      CloseNGradeBunkers
   End If
   
   ConvertCustomVrsNGrades = sTextBody
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ConvertCustomVrsNGrades. ", Err.Number, Err.Description
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
      & " Where vrsBunker.DetailKey = " & IdKey2 & " And vrsBunker.DetailType = " & DetailTypeEnum _
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
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.FetchNGradeBunkers ", Err.Number, Err.Description
   CloseRecordset lorsB
End Sub

Private Sub CloseNGradeBunkers()
   
   On Error GoTo SubError
   
   CloseRecordset moNGradeRS
   KillObject moNGradeRS
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.CloseNGradeBunkers ", Err.Number, Err.Description
   
End Sub

Private Function ProcessNGradeBunkers(TextBody As String, DGVariable As String) As String
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
         moNGradeRS.Filter = adFilterNone
         sNGradeBunkers = DisplayNGradeBunkers
         
      ElseIf DGVariable = "[VRS_NGRADE_ROB]" Then
          ' Show just ROB
         moNGradeRS.Filter = "vrsBunkPurposeKey = 1"
         sNGradeBunkers = DisplayNGradeBunkers
          
      ElseIf DGVariable = "[VRS_NGRADE_CONS]" Then
          ' Show everything except ROB
         moNGradeRS.Filter = "vrsBunkPurposeKey > 1"
         sNGradeBunkers = DisplayNGradeBunkers
          
      ElseIf DGVariable = "[VRS_NGRADE_ROBCONS]" Then
          
         moNGradeRS.Filter = "vrsBunkPurposeKey = 1"
         sNGradeBunkers = DisplayNGradeBunkers
              
         moNGradeRS.Filter = "vrsBunkPurposeKey > 1"
         sNGradeBunkers = sNGradeBunkers & DisplayNGradeBunkers
              
      End If
      
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
   goSession.RaiseError "General Error in mwSession.mwWorkflow.ProcessNGradeBunkers. ", Err.Number, Err.Description
   ProcessNGradeBunkers = TextBody
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
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.DisplayNGradeBunkers ", Err.Number, Err.Description
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
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.FabricateRS", Err.Number, Err.Description
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
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.GetUnitType", Err.Number, Err.Description
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
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.GetBunkerPurposeDescription", Err.Number, Err.Description
   CloseRecordset loRs
End Function


' 8/2010 ms BGLNG-4
Private Function MaybeProcessReturningForm(loform As mwForm) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim fso As FileSystemObject
   Dim loWork As mwAlertWork
   Dim bSentStatus As Boolean
   Dim sUserEmail As String
   Dim sRefSite As String
   Dim lRT As Long
   On Error GoTo FunctionError
   If loform Is Nothing Then
      MaybeProcessReturningForm = False
      Exit Function
   ElseIf Not loform.IsOpen Or Trim(loform.FormProperties.FormID) = "" Then
      MaybeProcessReturningForm = False
      Exit Function
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT mwEventFormLog.FullFilename, mwEventFormLog.mwFormHeaderKey, mwFormHeader.mwcSitesKey, " & _
         " mwFormHeader.FormEmailedDT, mwFormHeader.FormReturnedDT, mwFormHeader.FormEmailedUserKey" & _
         " FROM mwFormHeader INNER JOIN mwEventFormLog ON mwFormHeader.ID = mwEventFormLog.mwFormHeaderKey " & _
         " WHERE mwEventFormLog.FormID= '" & loform.FormProperties.FormID & _
         "' AND mwFormHeader.FormEmailedDT Is Not Null AND mwFormHeader.mwFormStatusKey = 1"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If IsRecordLoaded(loRs) Then
      '
      ' Move file to folder...
      '
      Set fso = New FileSystemObject
      If Not fso.FileExists(BlankNull(loRs!FullFilename)) Then
         '
         goSession.RaiseError "Error in mwSession.Workflow.MaybeProcessReturningForm, file not found: " & BlankNull(loRs!FullFilename)
         MaybeProcessReturningForm = False
      Else
         RenameFileNowExtension loRs!FullFilename
         'loForm.ExcelForm.CloseApplication
         fso.CopyFile loform.FormProperties.FullFilename, loRs!FullFilename
         UpdateDateFormReturned ZeroNull(loRs!mwFormHeaderKey)
         MaybeProcessReturningForm = True
         sRefSite = goSession.Site.GetSiteName(ZeroNull(loRs!mwcSitesKey))
         ' 10/2011 ms - Send email to user...
         'Send Alert
         If Not IsNull(loRs!FormEmailedUserKey) Then
            ' Send email to the user
            sUserEmail = FetchUserEmail(loRs!FormEmailedUserKey)
            If sUserEmail <> "" Then
               With goSession.Message
                  .SendToAddress = sUserEmail
                  .subject = "EXTERNAL FORM RETURNED"
                  .Body = "Form: " & loform.FormProperties.FormID & " has been returned from an external party. " & vbCrLf & _
                  " Reference Site: " & sRefSite & vbCrLf & vbCrLf & _
                          "It was emailed out on: " & loRs!FormEmailedDT
                  ' what about HTML...?????
                  .AttachmentFile = ""
                  bSentStatus = .Send()
               End With
               If bSentStatus Then
                  goSession.RaiseWorkflowStatus "  - Returned Form Alert email sent to User email: " & sUserEmail
               Else
                  goSession.RaiseWorkflowStatus "  - FAILED: Returned Form Alert email NOT SENT to User: " & sUserEmail
               End If
            Else
               goSession.RaiseWorkflowStatus "  - FAILED: Returned Form Alert email NOT SENT. No email address for User Key: " & loRs!FormEmailedUserKey
            End If
            
         ElseIf Trim(loform.FormProperties.RoleType) <> "" And IsNumeric(loform.FormProperties.RoleType) Then
            'older mechanism before mwFormHeader was updated with User Key.
            lRT = CLng(loform.FormProperties.RoleType)
            Set loWork = New mwAlertWork
            loWork.CreateAlert loform.FormProperties.SiteKey, "EXTERNAL FORM RETURNED", lRT, _
              "Form: " & loform.FormProperties.FormID & " has been returned from an external party. " & vbCrLf & vbCrLf & _
              " Reference Site: " & sRefSite & vbCrLf & vbCrLf & _
              "It was emailed out on: " & loRs!FormEmailedDT
            Set loWork = Nothing
         Else
            goSession.RaiseWorkflowStatus "Warning in mwSession.mwWorkflow.MaybeProcessReturningForm, No alert sent because Role Type Key missing in form: " & _
              loform.FormProperties.FormID
         End If
      End If
   Else
      ' Nope - not a returning form
      MaybeProcessReturningForm = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.MaybeProcessReturningForm", Err.Number, Err.Description
   CloseRecordset loRs
End Function


Private Function UpdateDateFormReturned(HeaderKey As Long) As Boolean
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Update mwFormHeader...
   Set moRsFormHeader = New Recordset
   moRsFormHeader.CursorLocation = adUseClient
   If Not IsNull(HeaderKey) Then
      sSQL = "SELECT * from mwFormHeader where ID=" & HeaderKey
      
      moRsFormHeader.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      If IsRecordLoaded(moRsFormHeader) Then
         moRsFormHeader!FormReturnedDT = Now()
         moRsFormHeader.Update
         CloseRecordset moRsFormHeader
         UpdateDateFormReturned = True
      End If
   Else
      UpdateDateFormReturned = False
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwEventForms_ocx.UpdateDateFormReturned. ", Err.Number, Err.Description
End Function

Private Sub moRsFormHeader_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwFormHeader") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow..moRsFormHeader_WillChangeRecord. ", Err.Number, Err.Description

End Sub

' 2/2011 ms BGLNG-18
Private Function MaybeProcessQuestionSetForm(loform As mwForm) As Boolean
   Dim sSQL As String
   Dim nFTKey As Long
   Dim loRs As Recordset
   On Error GoTo FunctionError
   If loform Is Nothing Then
      MaybeProcessQuestionSetForm = False
      Exit Function
   ElseIf Not loform.IsOpen Then
      MaybeProcessQuestionSetForm = False
      Exit Function
   Else
   
      '
      ' Template linked to a QuestionSetForm?
      '
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      sSQL = "select ID from mwwFFormTemplate where TemplateID='" & loform.TemplateProperties.TemplateID & "'"
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If Not IsRecordLoaded(loRs) Then
         CloseRecordset loRs
         MaybeProcessQuestionSetForm = False
         Exit Function
      End If
      nFTKey = loRs!ID
      loRs.Close
      sSQL = "select count(*) from maQuestionSet where mwwfFormTemplateKey=" & nFTKey
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If loRs.Fields(0).value < 1 Then
         CloseRecordset loRs
         MaybeProcessQuestionSetForm = False
         Exit Function
      End If
   End If
   '
   ' OK, Process Form into Marine Assurance System
   '
   ProcessQuestionSetForm loform, nFTKey
   MaybeProcessQuestionSetForm = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.MaybeProcessQuestionSetForm", Err.Number, Err.Description
   'CloseRecordset loRs
End Function




Private Function ProcessQuestionSetForm(loform As mwForm, TemplateKey As Long) As Boolean
   Dim sSQL As String
   Dim nFTKey As Long
   Dim nASKey As Long
   Dim loRs As Recordset
   Dim loRs2 As Recordset
   Dim loRsQS As Recordset
   Dim loRsQSQ As Recordset
   Dim loRsAS As Recordset
   Dim loRsASA As Recordset
   Dim loMaWork As Object
   Dim sCell As String
   Dim nWS As Long
   Dim sCellValPK As String
   Dim sCellValSL As String
   Dim nScopePK As Long
   
   On Error GoTo FunctionError
   Set loRsQS = New Recordset
   loRsQS.CursorLocation = adUseClient
   sSQL = "select * from maQuestionSet where mwwfFormTemplateKey=" & TemplateKey
   loRsQS.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set loMaWork = CreateObject("mwSafety4.smMarineAssWork")
   loMaWork.InitSession goSession
   '
   ' Check for duplicates
   '
   If Not IsNull(loRsQS!maQsqPkKey) Then
      ' Fetch QSQ to get worksheet/cell reference
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      sSQL = "select ExcelWorksheetOffset, ExcelCellRef from maQuestionSetQues where ExcelWorksheetOffset is not null and " & _
             "ExcelCellRef is not null and ID=" & loRsQS!maQsqPkKey
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If IsRecordLoaded(loRs) Then
         sCellValPK = loform.ExcelForm.GetCell(loRs!ExcelCellRef, loRs!ExcelWorksheetOffset)
         loRs.Close
         If Len(sCellValPK) > 0 Then
            sSQL = "select ID from maAnswerSet where VisualDescription='" & sCellValPK & "'"
            loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
            If IsRecordLoaded(loRs) Then
               'Duplicate!
               goSession.RaisePublicError "Error in mwSession.mwWorkflow.ProcessQuestionSetForm, Spreadsheet already processed. PK Value: " & _
                                          sCellValPK & " - " & loform.FormProperties.FileName
               CloseRecordset loRs
               KillObject loMaWork
               ProcessQuestionSetForm = True
               Exit Function
            End If
         End If
      End If
      CloseRecordset loRs
   End If
   '
   ' Check for scope record lookup requirement
   '
   If Not IsNull(loRsQS!QSScopeEnum) And Not IsNull(loRsQS!maQsqLookupKey) Then
      ' Fetch QSQ to get worksheet/cell reference
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      sSQL = "select ExcelWorksheetOffset, ExcelCellRef from maQuestionSetQues where ExcelWorksheetOffset is not null and " & _
             "ExcelCellRef is not null and ID=" & loRsQS!maQsqLookupKey
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If IsRecordLoaded(loRs) Then
         ' Fetch lookup value from spreadsheet
         sCellValSL = loform.ExcelForm.GetCell(loRs!ExcelCellRef, loRs!ExcelWorksheetOffset)
         loRs.Close
         If Len(sCellValSL) > 0 Then
            Select Case loRsQS!QSScopeEnum
               Case Is = 0 'MA_QS_SCOPE_SITE
                  nScopePK = loMaWork.FetchScopeSite(sCellValSL)
               Case Is = 1 'MA_QS_SCOPE_TECH_OP
                  nScopePK = loMaWork.FetchScopeTechOp(sCellValSL)
               Case Is = 2 'MA_QS_SCOPE_PORT
                  nScopePK = loMaWork.FetchScopePort(sCellValSL)
               Case Is = 3 'MA_QS_SCOPE_TERM_BERTH
                  'Poor design, doesn't work
                  'nScopePK = loMaWork.FetchScopeTermBerth(sCellValSL)
               Case Is = 4 'MA_QS_SCOPE_COMPANY
                  nScopePK = loMaWork.FetchScopeCompany(sCellValSL)
            End Select
            
            If nScopePK < 1 Then
               'Duplicate!
               goSession.RaisePublicError "Error in mwSession.mwWorkflow.ProcessQuestionSetForm, Scope Lookup Failed for: " & _
                                          sCellValPK & " - " & loform.FormProperties.FileName
               CloseRecordset loRs
               KillObject loMaWork
               ProcessQuestionSetForm = True
               Exit Function
            End If
         End If
      End If
      CloseRecordset loRs
   End If
   ' sCellValPK contains the VisualDescription for the maAnswerSet record...
   ' nScopePK contains the Scope PK for the maAnswerSet record...
   goCon.BeginTrans
      Do While Not loRsQS.EOF
      '
      ' OK, add the Answerset...
      '
         nASKey = loMaWork.AddAnswerSet(loRsQS!ID, loform.FormProperties.SiteKey, 0, 0, 0, True)
         '
         ' Update Answer Set, if necessary...
         '
         If Len(sCellValPK) > 0 Or nScopePK > 0 Then
            '
            ' Fetch and update the Answer Set
            '
            Set loRsAS = New Recordset
            loRsAS.CursorLocation = adUseClient
            sSQL = "select * from maAnswerset where ID=" & nASKey
            loRsAS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
            If Len(sCellValPK) > 0 Then
               loRsAS!VisualDescription = sCellValPK
            End If
            If nScopePK > 0 Then
               Select Case loRsQS!QSScopeEnum
                  Case Is = 0 'MA_QS_SCOPE_SITE
                     loRsAS!mwcSitesKey = nScopePK
                  Case Is = 1 'MA_QS_SCOPE_TECH_OP
                     loRsAS!mwcOperatorKey = nScopePK
                  Case Is = 2 'MA_QS_SCOPE_PORT
                     loRsAS!vrsPortsKey = nScopePK
                  Case Is = 3 'MA_QS_SCOPE_TERM_BERTH
                     'Poor design, doesn't work
                     'loRsAS!vrsPortTermKey = nScopePK
                  Case Is = 4 'MA_QS_SCOPE_COMPANY
                     loRsAS!mwContactKey = nScopePK
               End Select
               loRsAS.Update
               CloseRecordset loRsAS
            End If
         End If
         '
         ' Fetch all the answers to update
         '
         Set loRsASA = New Recordset
         loRsASA.CursorLocation = adUseClient
         sSQL = "SELECT maAnswerSetAns.*, maQuestionSetQues.ExcelWorksheetOffset, maQuestionSetQues.ExcelCellRef, " & _
                " maCoreQuestion.DataTypeEnum, maQuestionSetQues.ExcelCommentRef, maQuestionSetQues.ExcelBasicReasonRef, " & _
                " maQuestionSetQues.ErrorCondition, maQuestionSetQues.WarningCondition " & _
                " FROM maCoreQuestion, maQuestionSetQues, maAnswerSetAns where " & _
                " maQuestionSetQues.ID = maAnswerSetAns.maQuestionSetQuesKey and " & _
                " maCoreQuestion.ID = maQuestionSetQues.maCoreQuestionKey  " & _
                " and maQuestionSetQues.ExcelWorksheetOffset Is Not Null " & _
                " AND maQuestionSetQues.ExcelCellRef Is Not Null " & _
                " AND maAnswerSetAns.maAnswerSetKey=" & nASKey
         loRsASA.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
         '
         ' Update Answer Set Answer values from the spreadsheet...
         '
         Do While Not loRsASA.EOF
            '
            '
            ' Set worksheet
            '
            sCell = loform.ExcelForm.GetCell(loRsASA!ExcelCellRef, loRsASA!ExcelWorksheetOffset)
            loRsASA!AnswerText = Left(sCell, 199)
            '
            ' Update data type specific value
            '
            'MA_ADT_INTEGER = 1
            'MA_ADT_DECIMAL = 2
            'MA_ADT_BOOLEAN = 3
            'MA_ADT_DATETIME = 4
            'MA_ADT_STRING = 5
            'MA_ADT_MEMO = 6
            Select Case loRsASA!DataTypeEnum
               Case Is = 1
                  If IsNumeric(sCell) Then
                     loRsASA!AnswerNumeric = CLng(sCell)
                  End If
               Case Is = 2
                  If IsNumeric(sCell) Then
                     loRsASA!AnswerNumeric = CDbl(sCell)
                  End If
               Case Is = 3
                  If UCase(sCell) = "TRUE" Or UCase(sCell) = "YES" Then
                     loRsASA!AnswerBool = True
                  Else
                     loRsASA!AnswerBool = False
                  End If
               Case Is = 4
                  If IsDate(sCell) Then
                     loRsASA!AnswerDateTime = CDate(sCell)
                  End If
               Case Is = 6
                  loRsASA!AnswerMemo = sCell
            End Select
            '
            ' Comments adn Basic Reason Update
            '
            If Not IsNull(loRsASA!ExcelCommentRef) Then
               loRsASA!ExcelIntComments = BlankNull(loform.ExcelForm.GetCell(loRsASA!ExcelCommentRef, loRsASA!ExcelWorksheetOffset))
            End If
            If Not IsNull(loRsASA!ExcelBasicReasonRef) Then
               loRsASA!BasicReason = BlankNull(loform.ExcelForm.GetCell(loRsASA!ExcelBasicReasonRef, loRsASA!ExcelWorksheetOffset))
            End If
            loRsASA.Update
            loRsASA.MoveNext
         Loop
         CloseRecordset loRsASA
         loRsQS.MoveNext
      Loop
   goCon.CommitTrans
   ProcessQuestionSetForm = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwWorkflow.ProcessQuestionSetForm", Err.Number, Err.Description
   On Error Resume Next
   goCon.RollbackTrans
   CloseRecordset loRsQS
   CloseRecordset loRsQSQ
   CloseRecordset loRsASA
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
   goSession.RaiseError "General Error in mwSession.mwWorkflow.StripUnusedEmailFields: ", Err.Number, Err.Description
   StripUnusedEmailFields = str
End Function



Private Function FetchUserEmail(UserKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ' Update mwFormHeader...
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT FromEmailAddress from mwcUsers where ID=" & UserKey
   
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   If IsRecordLoaded(loRs) Then
      If IsNull(loRs!FromEmailAddress) Then
         FetchUserEmail = ""
      Else
         FetchUserEmail = loRs!FromEmailAddress
      End If
   Else
      FetchUserEmail = ""
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwWorkflow.FetchUserEmail. ", Err.Number, Err.Description
End Function

Private Sub LicenseValidationCheck()
   Dim loLicenseCheck As Object
   
   On Error GoTo FunctionError

   Set loLicenseCheck = CreateObject("mwIntegrate.mwLicenseCheck")
   loLicenseCheck.InitSession goSession
   
   loLicenseCheck.LicenseValidationCheck
   
   KillObject loLicenseCheck
   
   
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwWorkflow.LicenseValidationCheck. ", Err.Number, Err.Description

End Sub


Public Function RenameFileNowExtensionNew(sFullFilename As String, sStrTarget As String) As Boolean
   On Error GoTo FunctionError
   
   If moFso.FileExists(sFullFilename) Then
      ' Set attributes to normal
      moFso.GetFile(sFullFilename).Attributes = Normal
      sStrTarget = sFullFilename & "." & Format(Now(), "YYYYMMDDhhmmss")
      Do While moFso.FileExists(sStrTarget)
         sStrTarget = sStrTarget & ".1"
      Loop
      moFso.MoveFile sFullFilename, sStrTarget
   End If
   RenameFileNowExtensionNew = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwWorkflow.RenameFileNowExtensionNew file= " & sFullFilename, Err.Number, Err.Description
   RenameFileNowExtensionNew = False
End Function

