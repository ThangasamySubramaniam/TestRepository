VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventFormWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwEventFormWork - Support functions for managing Event Forms
' 10/28/02 ms
'

Option Explicit
Const RS_TYPE_mwEventFormType_ID = 0
Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1
Dim WithEvents moRsFormHeader As Recordset
Attribute moRsFormHeader.VB_VarHelpID = -1
Dim WithEvents moRsFormDetailRs As Recordset
Attribute moRsFormDetailRs.VB_VarHelpID = -1


Private mSiteKey As Long
Dim mSig1Required As Boolean
Dim mSig2Required As Boolean

Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventFormTypeKey = 3
Const RS_FormID = 4
Const RS_Subject = 5
Const RS_FullFilename = 6
Const RS_IsCreated = 7
Const RS_IsPrinted = 8
Const RS_IsSubmitted = 9
Const RS_SubmittedBy = 10
Const RS_SubmittedDateTime = 11
Const RS_mwcSitesKey = 12
Const RS_mwcRoleTypeKey = 13
Const RS_mwFormHeaderKey = 14
Const RS_IsRequiredForm = 15

Const RS_IsSignedOut = 16
Const RS_SignedOutToUser = 17
Const RS_SignOutDate = 18

Const RS_T_ID = 19
Const RS_T_EventTypeKey = 20
Const RS_T_TemplateID = 21
Const RS_T_IsMandatoryForm = 22
Const RS_T_IsSuggestedForm = 23
Const RS_T_IsAutoSubmit = 24
Const RS_T_IsPrintOnSubmit = 25
Const RS_T_DisplaySequence = 26
Const RS_T_DisplayIcon = 27
Const RS_T_SubmitEmailAsDraft = 28
Const RS_T_mwFormCabinetKey = 29


'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


Public Function AutoCreate(EventType As Long, EventDetailKey As Long, Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loRsType As Recordset
   On Error GoTo FunctionError
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   Set loRsType = FetchTemplatesRS(EventType, SiteKey)
   Do While Not loRsType.EOF
      '
      If loRsType!IsMandatoryForm Or loRsType!IsSuggestedForm Then
         ' Create Event Form record...
         CreateRecord EventType, EventDetailKey, loRsType.Fields(RS_TYPE_mwEventFormType_ID).value, SiteKey, loRsType!IsMandatoryForm
      End If
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   AutoCreate = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.AutoCreate. ", Err.Number, Err.Description
   AutoCreate = False
End Function

Public Function CreateRecord(EventType As Long, EventDetailKey As Long, _
 mwEventFormTypeKey As Long, Optional SiteKey As Long, Optional IsRequired As Boolean) As Long
   Dim sSQL As String
   Dim loform As mwForm
   Dim loEqptHistory As Object
   Dim sFormName As String
   On Error GoTo FunctionError
   mSiteKey = SiteKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFormLog where ID=-1"
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   moRS.AddNew
   moRS!ID = goSession.MakePK("mwEventFormLog")
   moRS!mwEventTypeKey = EventType
   moRS!mwEventDetailKey = EventDetailKey
   moRS!mwEventFormTypeKey = mwEventFormTypeKey
   sFormName = Left(Trim(FetchFormTemplateDescription(mwEventFormTypeKey)), 50)
   moRS!subject = sFormName
   moRS!IsCreated = 0
   moRS!IsPrinted = 0
   moRS!IsSubmitted = 0
   moRS!IsSignedOut = 0
   If SiteKey > 0 Then
      moRS!mwcSitesKey = SiteKey
   End If
   moRS!IsRequiredForm = IsRequired
   moRS.Update
   CreateRecord = moRS!ID
   If EventType = SW_EVENT_HISTORY Then
      Set loEqptHistory = CreateObject("mwWorks6.swEqptHistoryWork")
      If loEqptHistory.IsClosedWorkOrder(EventDetailKey, SiteKey) Then
         Call loEqptHistory.UpdateClosedWOEventHistory(EventDetailKey, "New Form", "New Form Added", "Form Name:" & sFormName, SiteKey)
      End If
      KillObject loEqptHistory
   End If
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.CreateRecord. ", Err.Number, Err.Description
   CreateRecord = -1
   KillObject loEqptHistory
End Function
Public Function CreateCopyRecord(mwEventFormLogKey As Long, mwFormHeaderKey As Long) As Long
   Dim sSQL As String
   Dim OldRs As Recordset
   
   On Error GoTo FunctionError
   
   Set OldRs = New Recordset
   OldRs.CursorLocation = adUseClient
   sSQL = "select * from mwEventFormLog where ID= " & mwEventFormLogKey
   OldRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(OldRs) Then
      Set moRS = New Recordset
      moRS.CursorLocation = adUseClient
      sSQL = "select * from mwEventFormLog where ID=-1"
      moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      
      
      moRS.AddNew
      moRS!ID = goSession.MakePK("mwEventFormLog")
      moRS!mwEventTypeKey = OldRs!mwEventTypeKey
      moRS!mwEventDetailKey = OldRs!mwEventDetailKey
      moRS!mwEventFormTypeKey = OldRs!mwEventFormTypeKey
      moRS!subject = GetFormHeaderDescription(mwFormHeaderKey)
      moRS!IsCreated = OldRs!IsCreated
      moRS!IsPrinted = 0
      moRS!IsSubmitted = 0
      moRS!IsSignedOut = 0
      moRS!mwcSitesKey = OldRs!mwcSitesKey
      moRS!IsRequiredForm = OldRs!IsRequiredForm
      moRS!FormID = GetFormHeaderFormID(mwFormHeaderKey)
      moRS!mwFormHeaderKey = mwFormHeaderKey
      moRS!FullFilename = GetFormHeaderFileName(mwFormHeaderKey)
      
      moRS.Update
      CreateCopyRecord = moRS!ID
   End If
   
   goSession.CloseRecordset OldRs
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.CreateCopyRecord. ", Err.Number, Err.Description
   CreateCopyRecord = -1
End Function
Private Function GetFormHeaderFileName(mwFormHeaderKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select FullFileName from mwFormHeader where ID= " & mwFormHeaderKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic

   If IsRecordLoaded(loRs) Then
      GetFormHeaderFileName = BlankNull(loRs!FullFilename)
   Else
      GetFormHeaderFileName = ""
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.GetFormHeaderFileName. ", Err.Number, Err.Description
   GetFormHeaderFileName = ""
   
End Function
Private Function GetFormHeaderDescription(mwFormHeaderKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select Description from mwFormHeader where ID= " & mwFormHeaderKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic

   If IsRecordLoaded(loRs) Then
      GetFormHeaderDescription = BlankNull(loRs!Description)
   Else
      GetFormHeaderDescription = ""
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.GetFormHeaderDescription. ", Err.Number, Err.Description
   GetFormHeaderDescription = ""
   
End Function
Private Function GetFormHeaderFormID(mwFormHeaderKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   Dim fso As FileSystemObject
   
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject

   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select FullFilename from mwFormHeader where ID= " & mwFormHeaderKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic

   If IsRecordLoaded(loRs) Then
      
      If IsNull(loRs!FullFilename) Then
         GetFormHeaderFormID = ""
      Else
         GetFormHeaderFormID = fso.GetBaseName(loRs!FullFilename)
      End If
   Else
      GetFormHeaderFormID = ""
   End If
   goSession.CloseRecordset loRs
   Set fso = Nothing
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.GetFormHeaderFormID. ", Err.Number, Err.Description
   GetFormHeaderFormID = ""
   
End Function

Public Function MakeEventForm(oInterface As mwEventInterface, FormLogKey As Long, _
 Optional SiteKey As Long) As mwForm
   Dim sSQL As String
   Dim fso As FileSystemObject
   Dim loform As mwForm
   Dim loEvWork As mwEventWork
   Dim strPath As String
   Dim strMonth As String
   Dim strTemplate As String
   Dim nFormCabKey As Long
   Dim nFormFolderKey As Long
   Dim nKey As Long
   Dim nTemplateKey As Long
   Static loWork As Object
   Dim nFormHeaderKey As Long

   On Error GoTo FunctionError
   
   '
   ' Fetch Voyage Event Form record
   '
   Set moRS = FetchFormRecord(FormLogKey)
   If moRS.RecordCount < 1 Then
      goSession.RaiseError "Error in mwEventFormWork.MakeEventForm, Voyage Event Form record missing."
      goSession.CloseRecordset moRS
      Exit Function
   End If
   '
   ' Create Voyage folder, if necessary...
   '
   Set loEvWork = New mwEventWork
   strPath = loEvWork.GetVoyageFolder(oInterface.EventType)
   Set loEvWork = Nothing
   If strPath = "" Then
      goSession.RaiseError "Error in mwEventFormWork.MakeEventForm, Voyage Form Folder missing. "
      Set fso = Nothing
      CloseRecordset moRS
      MakeEventForm = False
      Exit Function
   End If
   
   '
   ' Create Event Form...update event form record...
   '
   '---PBT-1547
   If SiteKey > 0 And BoolNull(moRS.Fields("IsCopySubmittedForm").value) Then
      If MsgBox("Do you want to create New Form from Submitted Form", vbYesNo, "Create New Form") = vbYes Then
         nFormHeaderKey = GetEventSubmittedForm(oInterface.EventType, moRS.Fields("TemplateID").value, SiteKey)
         If nFormHeaderKey < 1 Then
            Exit Function
         End If
      End If
   End If
   '---PBT-1547
   Set loform = goSession.OpenNewForm
   '
   ' Place Version 4 form properties...
   '
   loform.FormProperties.EventType = oInterface.EventType
   loform.FormProperties.EventRecordKey = oInterface.EventDetailKey

   If loform.MakeNewForm(moRS.Fields("TemplateID").value, strPath) Then
      goCon.BeginTrans
      '
      ' Update Voyage Form record...
      '
      moRS!FullFilename = loform.FormProperties.FullFilename
      moRS!IsCreated = True
      moRS!FormID = loform.FormProperties.FormID
      moRS!IsCreated = True
      If SiteKey > 0 Then
         moRS!mwcSitesKey = SiteKey
      End If
      moRS.Update
      If Not IsNull(moRS.Fields("mwFormCabinetKey").value) And Not IsNull(moRS.Fields(RS_mwEventFormTypeKey).value) Then
         '
         ' Create File Cabinet record ?
         '
         nFormCabKey = GetFormCabKey(moRS.Fields(RS_mwEventFormTypeKey).value)
         '
         ' Create Form Folder record...
         '
         Set loWork = CreateObject("mwManuals.mwFormMaintenanceWork")
         loWork.InitSession goSession
         If Not IsNull(moRS.Fields("mwFormFoldersKey").value) Then
            nFormFolderKey = moRS.Fields("mwFormFoldersKey").value
         Else
            nFormFolderKey = loWork.SelectFormFolder(nFormCabKey, SiteKey)
         End If
         If nFormFolderKey < 1 Then
            MsgBox "No form folder selected, cancelling Open Form.", vbInformation, "Cancel Open Form"
            goCon.RollbackTrans
            goSession.TearDownForm loform
            Exit Function
         End If
      
         nTemplateKey = GetTemplateKey(moRS.Fields(RS_T_TemplateID).value)
         nKey = MakeFormCabinetRecord(nTemplateKey, _
          nFormFolderKey, loform.FormProperties.FullFilename, SiteKey)
         If nKey > 0 Then
            moRS!mwFormHeaderKey = nKey
            moRS.Update
         Else
            MsgBox "No form record created, cancelling Open Form.", vbInformation, "Canel Open Form"
            goCon.RollbackTrans
            goSession.TearDownForm loform
            Exit Function
         End If
         Set loWork = Nothing
      End If
      '---PBT-1547
      If nFormHeaderKey > 0 Then
         loform.InstantiateForm moRS.Fields(RS_T_TemplateID).value, loform.FormProperties.FullFilename, nFormHeaderKey, True
      End If
      '---PBT-1547
      CloseRecordset moRS
      Set MakeEventForm = loform
      goCon.CommitTrans
      Set loform = Nothing
   Else
      goSession.RaiseError "Error in mwEventFormWork.MakeEventForm, Unable to create form with TemplateID: " & moRS.Fields("TemplateID").value
      goSession.TearDownForm loform
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.MakeEventForm. ", Err.Number, Err.Description
   CloseRecordset moRS
   goSession.TearDownForm loform
   On Error Resume Next
   goCon.RollbackTrans
End Function

Public Function FetchFormsRS(EventType As Long, EventDetailKey As Long, Optional NotSubmitted As Boolean) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   If NotSubmitted Then
'      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
'        " FROM mwEventFormType INNER JOIN mwEventFormLog ON mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey" & _
'        " WHERE mwEventFormLog.mwEventTypeKey=" & EventType & " AND mwEventFormLog.mwEventDetailKey=" & EventDetailKey & _
'        " AND mwEventFormLog.IsCreated<>0 AND mwEventFormLog.IsSubmitted=0;"
'   Else
'      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
'        " FROM mwEventFormType INNER JOIN mwEventFormLog ON mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey" & _
'        " WHERE mwEventFormLog.mwEventTypeKey)=" & EventType & " AND mwEventFormLog.mwEventDetailKey=" & EventDetailKey
'   End If

   If NotSubmitted Then
      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
        " FROM mwEventFormLog, mwEventFormType " & _
        " WHERE mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey AND " & _
        " mwEventFormLog.mwEventTypeKey = " & EventType & " AND " & _
        " mwEventFormLog.mwEventDetailKey = " & EventDetailKey & " AND " & _
        " mwEventFormLog.IsCreated <> 0 AND mwEventFormLog.IsSubmitted = 0"
   Else
      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
        " FROM mwEventFormLog, mwEventFormType " & _
        " WHERE mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey AND " & _
        " mwEventFormLog.mwEventTypeKey = " & EventType & " AND " & _
        " mwEventFormLog.mwEventDetailKey = " & EventDetailKey
   End If

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFormLog"
   Set FetchFormsRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.FetchFormsRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchTemplatesRS(EventType As Long, Optional SiteKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If UseFleetFormEvents And SiteKey > 0 Then
      sSQL = "SELECT mwEventFormType.*, mwwfFormTemplate.*" & _
        " FROM mwwfFormTemplate INNER JOIN ((mwcFleets INNER JOIN mwcFleetSites " & _
        " ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey) " & _
        " INNER JOIN (mwEventFormType INNER JOIN mwEventFormTypeFleet " & _
        " ON mwEventFormType.ID = mwEventFormTypeFleet.mwEventFormTypeKey) " & _
        " ON mwcFleets.ID = mwEventFormTypeFleet.mwcFleetsKey) " & _
        " ON mwwfFormTemplate.TemplateID = mwEventFormType.TemplateID " & _
        " Where mwcFleets.mwcFleetTypeKey = 145 And mwcFleetSites.mwcSitesKey=" & SiteKey & _
        " and mwEventFormType.mwEventTypeKey=" & EventType & _
        " AND (mwEventFormType.IsObsolete Is Null or mwEventFormType.IsObsolete = 0)"
   Else
      sSQL = "SELECT mwEventFormType.*, mwwfFormTemplate.* " & _
        " FROM mwwfFormTemplate , mwEventFormType " & _
        " WHERE mwwfFormTemplate.TemplateID = mwEventFormType.TemplateID AND " & _
        " mwEventFormType.mwEventTypeKey=" & EventType & _
        " AND (mwEventFormType.IsObsolete Is Null or mwEventFormType.IsObsolete = 0)"
   End If
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchTemplatesRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.FetchTemplatesRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchFormRecord(EventFormKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
'     " FROM mwEventFormType RIGHT JOIN mwEventFormLog ON mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey " & _
'     " WHERE mwEventFormLog.ID=" & EventFormKey

   If goSession.IsOracle Then
      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
         " FROM mwEventFormType , mwEventFormLog " & _
         " WHERE mwEventFormType.ID(+) = mwEventFormLog.mwEventFormTypeKey AND " & _
         " mwEventFormLog.ID = " & EventFormKey
   Else
      sSQL = "SELECT mwEventFormLog.*, mwEventFormType.* " & _
         " FROM mwEventFormType RIGHT JOIN mwEventFormLog ON mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey " & _
         " WHERE mwEventFormLog.ID=" & EventFormKey
   End If
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFormLog"
   Set FetchFormRecord = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.FetchFormRecord. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function



Public Function PrintForms(EventType As Long, EventDetailKey As Long, Optional PrintAllForms As Boolean) As Boolean
   Dim sSQL As String
   Dim moRS As Recordset
   Dim loform As mwForm
   Dim i As Integer
   On Error GoTo FunctionError
   sSQL = "select * from mwEventFormLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventDetailKey
   moRS.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If moRS.RecordCount < 1 Then
      CloseRecordset moRS
      PrintForms = True
      Exit Function
   End If
   '
   ' Loop through all forms
   '
   Do While Not moRS.EOF
      '
      If Not IsNull(moRS!mwcSitesKey) Then
         mSiteKey = moRS!mwcSitesKey
      Else
         mSiteKey = 0
      End If
      If moRS!IsMandatoryForm Or moRS!IsSuggestedForm Then
         '
         ' Create Event Form record...
         '
         If Not IsNull(moRS!FullFilename) Then
            If (Not moRS!IsPrinted) Or PrintAllForms Then
               Set loform = goSession.OpenNewForm
               loform.OpenForm moRS!FullFilename
               loform.PrintForm
               i = MsgBox("Press OK to if form printed, and you wish to continue printing next form.", vbOKCancel)
               goSession.TearDownForm loform
               Set loform = Nothing
            End If
         End If
      End If
      If i <> vbOK Then
         Exit Do
      End If
      moRS!IsPrinted = True
      moRS.Update
      moRS.MoveNext
   Loop
   CloseRecordset moRS
   
   PrintForms = True
   Set moRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.PrintForms. ", Err.Number, Err.Description
   CloseRecordset moRS
End Function


Public Function ChooseFormType(EventType As Long) As Long
   Dim loform As frmEventChooser
   On Error GoTo FunctionError
   Set loform = New frmEventChooser
   Load loform
   If loform.FormInitChooser(EventType) Then
      loform.Show vbModal
   Else
      Unload loform
      Set loform = Nothing
      ChooseFormType = -1
      Exit Function
   End If
   '
   ' Cancel operation ?
   '
   If loform.IsCancelled() Then
      ChooseFormType = -1
   Else
      ChooseFormType = loform.FetchFormTypeKey()
   End If
   Unload loform
   Set loform = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.ChooseFormType. ", Err.Number, Err.Description
   ChooseFormType = -1
End Function



Public Function EditForm(FullFilename As String) As Boolean
   Dim loform As mwForm
   On Error GoTo FunctionError
   Set loform = goSession.OpenNewForm
   loform.OpenForm FullFilename
   '
   If Not loform.IsOpen Then
      goSession.TearDownForm loform
      EditForm = False
      Exit Function
   End If
   If loform.IsReadOnly Then
      MsgBox "Form is set to Read Only, you will not be able to save any changes.", vbExclamation
   End If
   '
   loform.RunFormDll "OnActivate"
   If loform.EditForm Then
      '
      ' True option leaves Excel Application open, while ridding ourselves of the object...
      '
      goSession.TearDownForm loform, True
      EditForm = True
   Else
      goSession.TearDownForm loform
      EditForm = False
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.EditForm: " & FullFilename, Err.Number, Err.Description
   goSession.TearDownForm loform
   EditForm = False
   goSession.TearDownForm loform
   
End Function

Public Function SubmitAllForms(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loform As mwForm
   Dim strFile As String
   On Error GoTo FunctionError
   Set moRS = FetchFormsRS(EventType, EventDetailKey, True)
   ' Any forms ?
   If moRS.RecordCount < 1 Then
      SubmitAllForms = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   '
   ' any mandatory, not created ?
   '
   moRS.Filter = "IsCreated = False and IsMandatoryForm = True"
   If moRS.RecordCount > 0 Then
      '
      goSession.RaiseError "Error in mwEventFormWork.SubmitAllForms, a Mandatory form has not been created."
      SubmitAllForms = False
      goSession.CloseRecordset moRS
   End If
   moRS.Filter = adFilterNone
   '
   ' Loop through each form
   '
   moRS.MoveFirst
   Do While Not moRS.EOF
      '
      If Not IsNull(moRS!mwcSitesKey) Then
         mSiteKey = moRS!mwcSitesKey
      Else
         mSiteKey = 0
      End If
      Set loform = goSession.OpenNewForm
      If loform.OpenForm(moRS!FullFilename) Then
         'Print Form ?
            
         If moRS.Fields("IsPrinted").value = False _
           And moRS.Fields("IsPrintOnSubmit").value = True Then
            loform.PrintForm
            moRS!IsPrinted = True
         End If
         ' Submit form
         If loform.SubmitForm() Then
            moRS!IsSubmitted = True
            moRS!SubmittedBy = goSession.User.UserID
            moRS!SubmittedDateTime = Now()
            moRS.Update
         End If
      Else
         goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwEventFormWork.SubmitAllForms, OpenForm Failed for: " & moRS!FullFilename
      End If
      strFile = loform.FormProperties.FullFilename
      goSession.TearDownForm loform
      SetFormReadOnly strFile
      moRS.MoveNext
   Loop
   '
   ' Close up shop
   '
   SubmitAllForms = True
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.SubmitAllForms. ", Err.Number, Err.Description
   CloseRecordset moRS
End Function

Public Function SubmitForm(EventFormKey As Long, Optional SubmitEmailAsDraft As Boolean) As Boolean
   Dim loform As mwForm
   Dim lIsSent As Boolean
   Dim strFile As String
   Dim sSQL As String
   Dim nSig1UserKey  As Long
   Dim nSig2UserKey As Long
   Dim loSigWork As Object
   On Error GoTo FunctionError
   '
   ' Fetch Event Form Record
   '
   Set moRS = FetchFormRecord(EventFormKey)
   ' Any forms ?
   If Not IsRecordLoaded(moRS) Then
      goSession.RaiseError "Error in mwEventFormWork.SubmitForm, Event Form record Missing: " & EventFormKey
      SubmitForm = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   '
   ' Signatures Required?
   '
   If IsSignaturesRequired() Then
      Set loSigWork = CreateObject("mwSecurity.mwElecSignature")
      loSigWork.InitSession goSession
      If loSigWork.GetSignature(mSig2Required) Then
         '
         nSig1UserKey = loSigWork.Sig1UserKey
         nSig2UserKey = loSigWork.Sig2UserKey
      Else
         Set loSigWork = Nothing
         Exit Function
      End If
   End If
   
   
   
   ' OK, go for it
   If Not IsNull(moRS!mwcSitesKey) Then
      mSiteKey = moRS!mwcSitesKey
   Else
      mSiteKey = 0
   End If
   
   Set loform = goSession.OpenNewForm
   If loform.OpenForm(moRS!FullFilename, , ZeroNull(moRS!mwFormHeaderKey)) Then
      ' Print Form ?
      If Not IsNull(moRS!mwFormHeaderKey) Then
         '
         ' Run external VB DLL - if it exists... OnSubmit form event...
         '
         If Not loform.RunFormDll("OnSubmit") Then
            goSession.LogIt mwl_User_Defined, mwl_Information, "Submit has been cancelled by the OnSubmit DLL Code."
            SubmitForm = False
            Exit Function
         End If
      
         '
         ' Form Cabinet Submit
         '
         If loform.GetFormDetailData(moRS!mwFormHeaderKey, moRS!mwcSitesKey) Then
            SetFormHeaderSubmitStatus moRS!mwFormHeaderKey, nSig1UserKey, nSig2UserKey
            lIsSent = True
         Else
            lIsSent = False
         End If
      Else
         '
         ' Original Form Submit
         '
         If loform.SubmitForm() Then
            '
            '
            lIsSent = True
            '
            SetFormReadOnly moRS!FullFilename
         Else
            lIsSent = False
         End If
      End If
      If lIsSent Then
         moRS!IsSubmitted = True
         moRS!SubmittedBy = goSession.User.UserID
         moRS!SubmittedDateTime = Now()
         moRS.Update
      End If
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwEventFormWork.SubmitForm, OpenForm Failed for: " & moRS!FullFilename
   End If
   goSession.TearDownForm loform
   'If lIsSent Then
   '   SetFormReadOnly strFile
   'End If
   '
   ' Close up shop
   '
   SubmitForm = lIsSent
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.SubmitForm. ", Err.Number, Err.Description
End Function

Public Function SetFormReadOnly(FullFilename As String) As Boolean

   Dim fso As FileSystemObject
   Dim loFile As file
   '
   ' Set file to Read Only...
   '
   Set fso = New FileSystemObject
   Set loFile = fso.GetFile(FullFilename)
   loFile.Attributes = ReadOnly
   goSession.KillObject loFile
   goSession.KillObject fso
   SetFormReadOnly = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.SetFormReadOnly. ", Err.Number, Err.Description
   goSession.KillObject loFile
   goSession.KillObject fso
   SetFormReadOnly = False
End Function

Public Function DeleteUnusedForms(EventType As Long, EventDetailKey As Long) As Boolean
   On Error GoTo FunctionError
   'Dim fso As FileSystemObject
   Dim sSQL As String
   'Dim loRS As Recordset
   On Error GoTo FunctionError
   'Set loRS = New Recordset
   'loRS.CursorLocation = adUseClient
   'sSQL = "select * from mwEventFormLog where mwEventTypeKey=" & EventType & _
   '  " and mwEventDetailKey=" & EventDetailKey
   'loRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   'If loRS.RecordCount < 1 Then
   '   DeleteUnusedForms = True
   '   goSession.CloseRecordset loRS
   '   Exit Function
   'End If
   'Set fso = New FileSystemObject
   ' delete forms if not submitted
   'Do While Not loRS.EOF
   '   If Not loRS!IsSubmitted Then
   '      If loRS!IsCreated Then
   '         fso.DeleteFile (loRS!FullFilename)
   '         goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwEventFormWork, deleting form: " & loRS!FullFilename
   '      End If
   '      loRS.Delete
   '   Else
   '      goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventFormWork, skip delete submitted form: " & loRS!FullFilename
   '   End If
   '   loRS.MoveNext
   'Loop
   'Set fso = Nothing
   'goSession.CloseRecordset loRS
'   sSQL = "delete from mwEventFormLog where IsCreated=0 and mwEventTypeKey=" & EventType & _
'     " and mwEventDetailKey=" & EventDetailKey
'   goCon.Execute sSQL
   ' add replication ok, now deletes similar
   sSQL = "SELECT * FROM mwEventFormLog WHERE IsCreated=0 AND mwEventTypeKey=" & EventType & _
     " AND mwEventDetailKey=" & EventDetailKey
   goSession.CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If goSession.GUI.IsRecordLoaded(moRS) Then
      Do While Not moRS.EOF
         moRS.Delete
         moRS.MoveNext
      Loop
   End If
   
   CloseRecordset moRS
   DeleteUnusedForms = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.DeleteUnusedForms. ", Err.Number, Err.Description
   DeleteUnusedForms = False
End Function

Public Function DeleteForm(EventFormKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim fso As FileSystemObject
   Dim sSQL As String
   Dim moRS As Recordset
   On Error GoTo FunctionError
   sSQL = "select * from mwEventFormLog where ID= & EventFormKey"
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteForm = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   If Not IsNull(moRS!mwcSitesKey) Then
      mSiteKey = moRS!mwcSitesKey
   Else
      mSiteKey = 0
   End If
   
   Set fso = New FileSystemObject
   ' delete forms if not submitted
   If Not moRS!IsSubmitted Then
      If moRS!IsCreated Then
         fso.DeleteFile (moRS!FullFilename)
         goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwEventFormWork, deleting form: " & moRS!FullFilename
      End If
      moRS.Delete
      DeleteForm = True
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventFormWork, skip delete submitted form: " & moRS!FullFilename
      DeleteForm = False
   End If
   Set fso = Nothing
   goSession.CloseRecordset moRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.DeleteForm. ", Err.Number, Err.Description
   DeleteForm = False
End Function


Public Function IsSubmittedForms(EventType As Long, EventDetailKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select count(*) from mwEventFormLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventDetailKey & " and (IsSubmitted is not null and IsSubmitted <> 0)"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   ' Got Forms ?
   If loRs.Fields(0).value > 0 Then
      IsSubmittedForms = True
   Else
      IsSubmittedForms = False
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.IsSubmittedForms. ", Err.Number, Err.Description
   IsSubmittedForms = False
End Function




Private Function FetchFormTemplateDescription(FormTypeKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   Dim strTemplateID As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select TemplateID from mwEventFormType where ID=" & FormTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      FetchFormTemplateDescription = ""
      CloseRecordset loRs
      Exit Function
   End If
   strTemplateID = loRs!TemplateID
   loRs.Close
   sSQL = "select Description from mwwfFormTemplate where TemplateID='" & strTemplateID & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      FetchFormTemplateDescription = ""
      CloseRecordset loRs
      Exit Function
   End If
   FetchFormTemplateDescription = loRs!Description
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.FetchFormTemplateDescription. ", Err.Number, Err.Description
   CloseRecordset loRs
   FetchFormTemplateDescription = ""
End Function


Public Function IsEventFormTypesExist(EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select count(*) from mwEventFormType where mwEventTypeKey=" & EventType
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.Fields(0).value > 0 Then
      IsEventFormTypesExist = True
   Else
      IsEventFormTypesExist = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormsWork.IsEventFormTypesExist. ", Err.Number, Err.Description
End Function


Public Function AddRecords(EventType As Long, EventDetail As Long, Optional SiteKey As Long) As Long
   Dim loform As frmEventChooser
   Dim iFormType As Long
   Dim iNewFormKey As Long
   Dim strDummy As String
   On Error GoTo FunctionError
   Set loform = New frmEventChooser
   Load loform
   If loform.FormInitChooser(EventType, SiteKey) Then
      loform.Show vbModal
   Else
      Unload loform
      Set loform = Nothing
      AddRecords = -1
      Exit Function
   End If
   ' Cancel operation ?
   If loform.IsCancelled() Then
      AddRecords = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   '
   '
   ' May 2003 - Allow Multi Row select
   '
   Do While loform.IsSelectedRows
      iFormType = loform.FetchNextSelected(strDummy)
      If iFormType < 1 Then
         Exit Do
      End If
      iNewFormKey = CreateRecord(EventType, EventDetail, iFormType, SiteKey)
   Loop
   AddRecords = iNewFormKey
   Unload loform
   Set loform = Nothing
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.AddRecords. ", Err.Number, Err.Description
   AddRecords = -1
End Function

Public Function IsMandatoryCreated(EventType As Long, EventDetailKey As Long, _
  ByRef ErrorString As String) As Boolean
   Dim lIsError As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   '
   lIsError = False
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT Subject, TemplateID  " & _
'     " FROM mwEventFormType INNER JOIN mwEventFormLog ON mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey" & _
'     " WHERE mwEventFormLog.mwEventTypeKey=" & EventType & " AND mwEventFormLog.mwEventDetailKey=" & EventDetailKey & _
'     " AND mwEventFormLog.IsCreated=0 AND mwEventFormType.IsMandatoryForm<>0;"

   sSQL = "SELECT Subject, TemplateID  " & _
     " FROM mwEventFormType , mwEventFormLog " & _
     " WHERE mwEventFormType.ID = mwEventFormLog.mwEventFormTypeKey AND " & _
     " mwEventFormLog.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFormLog.mwEventDetailKey = " & EventDetailKey & _
     " AND mwEventFormLog.IsCreated = 0 AND mwEventFormType.IsMandatoryForm <> 0 " & _
     " AND (mwEventFormType.IsObsolete Is Null or mwEventFormType.IsObsolete = 0)"

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      IsMandatoryCreated = True
   Else
      IsMandatoryCreated = False
      Do While Not loRs.EOF
         ErrorString = ErrorString & "Mandatory Form not created: " & loRs!TemplateID & "-" & loRs!subject & vbCrLf
         loRs.MoveNext
      Loop
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFacts.IsMandatoryCreated. ", Err.Number, Err.Description
End Function




Public Function DeleteAllForms(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   On Error GoTo FunctionError
   
   sSQL = "select * from mwEventFormLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteAllForms = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   ' delete facts not used...
   Do While Not moRS.EOF
      If Not IsNull(moRS!mwcSitesKey) Then
         mSiteKey = moRS!mwcSitesKey
      Else
         mSiteKey = 0
      End If
      moRS.Delete
      moRS.MoveNext
   Loop
   goSession.CloseRecordset moRS
   DeleteAllForms = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactsWork.DeleteAllForms. ", Err.Number, Err.Description
   DeleteAllForms = False
End Function


Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwEventFormLog") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmSmOccurrence.moRS_WillChangeRecord. ", Err.Number, Err.Description
End Sub



Private Function UseFleetFormEvents() As Boolean
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = goCon.Execute("select count(*) from mwEventFormTypeFleet")
   If loRs.Fields(0).value > 0 Then
      UseFleetFormEvents = True
   Else
      UseFleetFormEvents = False
   End If
   loRs.Close
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.UseFleetFormEvents. ", Err.Number, Err.Description
End Function


Private Function MakeFormCabinetRecord(TemplateKey As Long, FormFolderKey As Long, FullFilename As String, _
 SiteKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   '
   ' Fetch template
   sSQL = "select * from mwwfFormTemplate where ID=" & TemplateKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwEventFormWork.MakeFormCabinetRecord, Template not found: " & TemplateKey
      CloseRecordset loRs
      Exit Function
   End If
   
   
   CloseRecordset moRsFormHeader
   sSQL = " SELECT * FROM mwFormHeader WHERE ID = -1"
   Set moRsFormHeader = New Recordset
   moRsFormHeader.CursorLocation = adUseClient
   moRsFormHeader.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   'If moRsFormHeader.RecordCount < 1 Then
   moRsFormHeader.AddNew
   moRsFormHeader!ID = goSession.MakePK("mwFormHeader")
   moRsFormHeader!mwFormFoldersKey = FormFolderKey
   moRsFormHeader!FormName = loRs!TemplateID
   moRsFormHeader!Description = loRs!Description
   moRsFormHeader!mwFormStatusKey = FORM_HEADER_STATUS_PENDING
   moRsFormHeader!mwwfFormTemplateKey = loRs!ID
   moRsFormHeader!TemplateRevisionNo = GetLatestTemplateVersion(loRs!TemplateID)
   moRsFormHeader!mwcSitesKey = SiteKey
   moRsFormHeader!OrigSiteKey = goSession.Site.SiteKey
   moRsFormHeader!CreateDate = Now()
'      moRsFormHeader!UpdateDate = Now()
   moRsFormHeader!FullFilename = FullFilename
   
   moRsFormHeader!CreatedByUser = goSession.User.UserKey
   moRsFormHeader!CreatedByRole = goSession.User.RoleTypeKey
   moRsFormHeader.Update
   MakeFormCabinetRecord = moRsFormHeader!ID
      
   'End If
   CloseRecordset moRsFormHeader

   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.MakeFormCabinetRecord. ", Err.Number, Err.Description
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
   goSession.RaisePublicError "General Error in mwEventFormWork.moRsFormHeader_WillChangeRecord. ", Err.Number, Err.Description

End Sub
Private Sub moRsFormDetailRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwFormDetail") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwEventFormWork.moRsFormDetailRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub


Private Function GetFormCabKey(mwEventFormTypeKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   '
   ' Fetch template
   sSQL = "select mwFormCabinetKey from mwEventFormType where ID=" & mwEventFormTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetFormCabKey = 0
   ElseIf IsNull(loRs!mwFormCabinetKey) Then
      GetFormCabKey = 0
   Else
      GetFormCabKey = loRs!mwFormCabinetKey
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.MakeFormCabinetRecord. ", Err.Number, Err.Description
End Function

Public Function GetLatestTemplateVersion(TemplateID As String) As Long

   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   '
   ' Fetch template
   sSQL = "SELECT DistributedRevisionNo FROM mwwfFormTemplate " & _
     " WHERE TemplateID='" & TemplateID & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetLatestTemplateVersion = 1
   ElseIf IsNull(loRs!DistributedRevisionNo) Then
      GetLatestTemplateVersion = 1
   Else
      GetLatestTemplateVersion = loRs!DistributedRevisionNo
   End If
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.GetLatestTemplateVersion. ", Err.Number, Err.Description
End Function

Private Function SetFormHeaderSubmitStatus(FormHeaderKey As Long, Optional Sig1UserKey As Long, Optional Sig2UserKey As Long) As Boolean
   On Error GoTo FunctionError
   
   SetFormHeaderSubmitStatus = SetFormHeaderStatus(FormHeaderKey, FORM_HEADER_STATUS_SUBMITTED, Sig1UserKey, Sig2UserKey)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.SetFormHeaderSubmitStatus. ", Err.Number, Err.Description
   SetFormHeaderSubmitStatus = False
End Function
Private Function SetFormHeaderStatus(FormHeaderKey As Long, Status As Integer, Optional Sig1UserKey As Long, Optional Sig2UserKey As Long) As Boolean
   Dim loWork As Object
   On Error GoTo FunctionError
   
   CloseRecordset moRsFormHeader
   
   Set loWork = CreateObject("mwManuals.mwFormMaintenanceWork")
   loWork.InitSession goSession
   
   SetFormHeaderStatus = loWork.SetFormHeaderStatus(FormHeaderKey, Status, Sig1UserKey, Sig2UserKey)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.SetFormHeaderStatus ", Err.Number, Err.Description
   CloseRecordset moRsFormHeader
End Function

Private Function GetTemplateKey(TemplateID As String) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   '
   sSQL = " SELECT ID FROM mwwfFormTemplate WHERE TemplateID ='" & TemplateID & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      GetTemplateKey = loRs!ID
   End If
   CloseRecordset loRs

   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.SetFormHeaderSubmitStatus. ", Err.Number, Err.Description
End Function

Private Function IsSignaturesRequired() As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   If Not IsRecordLoaded(moRS) Then
      mSig1Required = False
      mSig2Required = False
      Exit Function
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select Sig1Required, Sig2Required from mwwfFormTemplate where TemplateID='" & moRS!TemplateID & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If IsRecordLoaded(loRs) Then
      mSig1Required = loRs!Sig1Required
      mSig2Required = loRs!Sig2Required
   Else
      mSig1Required = False
      mSig2Required = False
   End If
   loRs.Close
   Set loRs = Nothing
   If mSig1Required Or mSig2Required Then
      IsSignaturesRequired = True
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwManuals.frmFormList.SetFormHeaderStatus ", Err.Number, Err.Description
End Function

Public Function FetchSubmittedFormsRS(EventType As Long, TemplateID As String, SiteKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT mwFormHeader.ID ,mwFormHeader.FormName ,mwEventFormLog.Subject,mwEventFormLog.SubmittedDateTime,mwEventFormLog.FormID " & _
        " FROM mwEventFormLog INNER JOIN (mwFormHeader INNER JOIN mwwfFormTemplate ON mwFormHeader.FormName = mwwfFormTemplate.TemplateID AND mwFormHeader.TemplateRevisionNo = mwwfFormTemplate.DistributedRevisionNo) ON mwFormHeader.ID = mwEventFormLog.mwFormHeaderKey " & _
        " WHERE mwFormHeader.FormName = '" & TemplateID & "'" & _
        " AND mwEventFormLog.mwEventTypeKey = " & EventType & _
        " AND mwEventFormLog.mwcSitesKey = " & SiteKey & _
        " AND mwEventFormLog.IsSubmitted <> 0 ORDER BY mwEventFormLog.SubmittedDateTime Desc"

     loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchSubmittedFormsRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.FetchSubmittedFormsRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function GetEventSubmittedForm(EventType As Long, TemplateID As String, SiteKey As Long) As Long
   On Error GoTo FunctionError
   Dim loSubmitform As frmEventFormChooser
   Dim nFormHeaderKey As Long

   nFormHeaderKey = -1
   Set loSubmitform = New frmEventFormChooser
   Load loSubmitform
   If loSubmitform.FormInitChooser(EventType, TemplateID, SiteKey) Then
      loSubmitform.Show vbModal
   
      If Not loSubmitform.IsCancelled() Then
         nFormHeaderKey = loSubmitform.FetchSelected
         If nFormHeaderKey > 0 Then
            GetEventSubmittedForm = nFormHeaderKey
         End If
      End If
   End If
   Unload loSubmitform
   Set loSubmitform = Nothing
   
   GetEventSubmittedForm = nFormHeaderKey
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFormWork.GetEventSubmittedForm. ", Err.Number, Err.Description
   GetEventSubmittedForm = -1
End Function

Public Function FetchFormTemplatesRS(EventType As Long, Optional SiteKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If UseFleetFormEvents And SiteKey > 0 Then
      sSQL = "SELECT mwEventFormType.ID, mwEventFormType.TemplateID, mwEventFormType.DisplayIcon, mwwfFormTemplate.Description " & _
        " FROM mwwfFormTemplate INNER JOIN ((mwcFleets INNER JOIN mwcFleetSites " & _
        " ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey) " & _
        " INNER JOIN (mwEventFormType INNER JOIN mwEventFormTypeFleet " & _
        " ON mwEventFormType.ID = mwEventFormTypeFleet.mwEventFormTypeKey) " & _
        " ON mwcFleets.ID = mwEventFormTypeFleet.mwcFleetsKey) " & _
        " ON mwwfFormTemplate.TemplateID = mwEventFormType.TemplateID " & _
        " Where mwcFleets.mwcFleetTypeKey = 145 And mwcFleetSites.mwcSitesKey=" & SiteKey & _
        " and mwEventFormType.mwEventTypeKey=" & EventType & _
        " AND (mwEventFormType.IsObsolete Is Null or mwEventFormType.IsObsolete = 0)"
   Else
      sSQL = "SELECT mwEventFormType.ID, mwEventFormType.TemplateID, mwEventFormType.DisplayIcon, mwwfFormTemplate.Description " & _
        " FROM mwwfFormTemplate , mwEventFormType " & _
        " WHERE mwwfFormTemplate.TemplateID = mwEventFormType.TemplateID AND " & _
        " mwEventFormType.mwEventTypeKey=" & EventType & _
        " AND (mwEventFormType.IsObsolete Is Null or mwEventFormType.IsObsolete = 0)"
   End If
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchFormTemplatesRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFormWork.FetchFormTemplatesRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function


Public Function IsMandatoryEventForm(mwEventFormTypeKey As Long) As Boolean
'PBT 1677
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select IsMandatoryForm from mwEventFormType where ID =" & mwEventFormTypeKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If BoolNull(loRs!IsMandatoryForm) = True Then
        IsMandatoryEventForm = True
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwEventFormWork.IsMandatoryEventForm. ", Err.Number, Err.Description
   CloseRecordset loRs
   IsMandatoryEventForm = False
End Function

