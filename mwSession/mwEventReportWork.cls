VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventReportWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public WithEvents moCrWork As mwCrystalWork
Attribute moCrWork.VB_VarHelpID = -1


' mwEventReportWork - Support functions for managing Event Crystal Reports
' 17 May 2003
'


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


Public Function AutoCreateEventReportRecords(EventType As Long, EventDetailKey As Long) As Boolean
   Dim strSQL As String
   Dim loRsType As Recordset
   On Error GoTo FunctionError
   Set loRsType = FetchVoyTemplatesRS(EventType)
   Do While Not loRsType.EOF
      '
      If loRsType!IsMandatoryForm Or loRsType!IsSuggestedForm Then
         ' Create Event Form record...
         CreateEventReportRecord EventType, EventDetailKey, loRsType!ID
      End If
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   AutoCreateEventReportRecords = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.AutoCreateEventReportRecords. ", err.Number, err.Description
   AutoCreateEventReportRecords = False
End Function

Public Function CreateEventReportRecord(EventType As Long, EventDetailKey As Long, mwEventReportTypeKey As Long) As Long
   Dim loRS As Recordset
   Dim loForm As mwForm
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   loRS.Open "mwEventReportLog", goSession.DBConnection, adOpenDynamic, adLockOptimistic, adCmdTable
   loRS.AddNew
   loRS!ID = goSession.MakePK("mwEventReportLog")
   loRS!mwEventTypeKey = EventType
   loRS!mwEventDetailKey = EventDetailKey
   loRS!mwEventReportTypeKey = mwEventReportTypeKey
   loRS!subject = FetchFormTemplateDescription(mwEventReportTypeKey)
   loRS!IsCreated = 0
   loRS!IsPrinted = 0
   loRS!IsSubmitted = 0
   loRS.Update
   CreateEventReportRecord = loRS!ID
   goSession.CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.CreateEventReportRecords. ", err.Number, err.Description
   CreateEventReportRecord = -1
End Function


Public Function MakeEventReport(oInterface As mwEventInterface, FormLogKey As Long) As mwForm
   Dim strSQL As String
   Dim fso As FileSystemObject
   Dim loRS As Recordset
   Dim loForm As mwForm
   Dim loEvWork As mwEventWork
   Dim strPath As String
   Dim strMonth As String
   Dim strTemplate As String
   On Error GoTo FunctionError
   
   '
   ' Fetch Voyage Event Form record
   '
   Set loRS = FetchVoyFormRecord(FormLogKey)
   If loRS.RecordCount < 1 Then
      goSession.RaiseError "Error in mwEventReportWork.MakeEventReport, Voyage Event Form record missing."
      goSession.CloseRecordset loRS
      Exit Function
   End If
   '
   ' Create Voyage folder, if necessary...
   '
   Set loEvWork = New mwEventWork
   strPath = loEvWork.GetVoyageFolder(oInterface.EventType)
   Set loEvWork = Nothing
   If strPath = "" Then
      goSession.RaiseError "Error in mwEventReportWork.MakeEventReport, Voyage Form Folder missing. "
      Set fso = Nothing
      CloseRecordset loRS
      MakeEventReport = False
      Exit Function
   End If
   
   '
   ' Create Event Form...update event form record...
   '
   Set loForm = goSession.OpenNewForm
   '
   ' Place Version 4 form properties...
   '
   loForm.FormProperties.EventType = oInterface.EventType
   loForm.FormProperties.EventRecordKey = oInterface.EventDetailKey
   If loForm.MakeNewForm(loRS.Fields("TemplateID").value, strPath) Then
      '
      ' Update Voyage Form record...
      '
      loRS!FullFilename = loForm.FormProperties.FullFilename
      loRS!IsCreated = True
      loRS!FormID = loForm.FormProperties.FormID
      loRS!IsCreated = True
      loRS.Update
      '
      ' OnCreate stuff...
      '
      'loForm.PopulateForm
      '
      ' Populate with extended recordset...
      
      'If Not loRS Is Nothing Then
      '   loForm.PopulateEventReport loRS.Fields("mwEventTypeKey").value, , loRS
      'End If
      CloseRecordset loRS
      Set MakeEventReport = loForm
      Set loForm = Nothing
   Else
      goSession.RaiseError "Error in mwEventReportWork.MakeEventReport, Unable to create form with TemplateID: " & loRS.Fields("TemplateID").value
      goSession.TearDownForm loForm
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.MakeEventReport. ", err.Number, err.Description
   CloseRecordset loRS
   goSession.TearDownForm loForm
End Function

Public Function FetchEventReportsRS(EventType As Long, EventDetailKey As Long, Optional NotSubmitted As Boolean) As Recordset
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   If NotSubmitted Then
      strSQL = "SELECT mwEventReportLog.*, mwEventReportType.* " & _
        " FROM mwEventReportType INNER JOIN mwEventReportLog ON mwEventReportType.ID = mwEventReportLog.mwEventReportTypeKey" & _
        " WHERE mwEventReportLog.mwEventTypeKey=" & EventType & " AND mwEventReportLog.mwEventDetailKey=" & EventDetailKey & _
        " AND mwEventReportLog.IsCreated<>0 AND mwEventReportLog.IsSubmitted=0;"
   Else
      strSQL = "SELECT mwEventReportLog.*, mwEventReportType.* " & _
        " FROM mwEventReportType INNER JOIN mwEventReportLog ON mwEventReportType.ID = mwEventReportLog.mwEventReportTypeKey" & _
        " WHERE mwEventReportLog.mwEventTypeKey)=" & EventType & " AND mwEventReportLog.mwEventDetailKey=" & EventDetailKey
   End If
   loRS.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   loRS.Properties("Unique Table").value = "mwEventReportLog"
   Set FetchEventReportsRS = loRS
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.FetchEventReportsRS. ", err.Number, err.Description
   CloseRecordset loRS
End Function

Public Function FetchVoyFormRecord(EventReportKey As Long) As Recordset
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "SELECT mwEventReportLog.*, mwEventReportType.* " & _
     " FROM mwEventReportType RIGHT JOIN mwEventReportLog ON mwEventReportType.ID = mwEventReportLog.mwEventReportTypeKey " & _
     " WHERE mwEventReportLog.ID=" & EventReportKey
   loRS.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   loRS.Properties("Unique Table").value = "mwEventReportLog"
   Set FetchVoyFormRecord = loRS
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.FetchVoyFormRecord. ", err.Number, err.Description
   CloseRecordset loRS
End Function



Public Function PrintVoyForms(EventType As Long, EventDetailKey As Long, Optional PrintAllForms As Boolean) As Boolean
   Dim strSQL As String
   Dim loRS As Recordset
   Dim loForm As mwForm
   Dim i As Integer
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select * from mwEventReportLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventDetailKey
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.RecordCount < 1 Then
      CloseRecordset loRS
      PrintVoyForms = True
      Exit Function
   End If
   '
   ' Loop through all forms
   '
   Do While Not loRS.EOF
      '
      If loRS!IsMandatoryForm Or loRS!IsSuggestedForm Then
         '
         ' Create Event Form record...
         '
         If Not IsNull(loRS!FullFilename) Then
            If (Not loRS!IsPrinted) Or PrintAllForms Then
               Set loForm = goSession.OpenNewForm
               loForm.OpenForm loRS!FullFilename
               loForm.PrintForm
               i = MsgBox("Press OK to if form printed, and you wish to continue printing next form.", vbOKCancel)
               goSession.TearDownForm loForm
               Set loForm = Nothing
            End If
         End If
      End If
      If i <> vbOK Then
         Exit Do
      End If
      loRS!IsPrinted = True
      loRS.Update
      loRS.MoveNext
   Loop
   CloseRecordset loRS
   
   PrintVoyForms = True
   Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.PrintVoyForms. ", err.Number, err.Description
   CloseRecordset loRS
End Function


Public Function ChooseFormType(EventType As Long) As Long
   Dim loForm As frmEventChooser
   On Error GoTo FunctionError
   Set loForm = New frmEventChooser
   Load loForm
   If loForm.FormInitChooser(EventType) Then
      loForm.Show vbModal
   Else
      Unload loForm
      Set loForm = Nothing
      ChooseFormType = -1
      Exit Function
   End If
   '
   ' Cancel operation ?
   '
   If loForm.IsCancelled() Then
      ChooseFormType = -1
   Else
      ChooseFormType = loForm.FetchFormTypeKey()
   End If
   Unload loForm
   Set loForm = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.ChooseFormType. ", err.Number, err.Description
   ChooseFormType = -1
End Function



Public Function EditVoyForm(FullFilename As String) As Boolean
   Dim loForm As mwForm
   On Error GoTo FunctionError
   Set loForm = goSession.OpenNewForm
   loForm.OpenForm FullFilename
   '
   If Not loForm.IsOpen Then
      goSession.TearDownForm loForm
      EditVoyForm = False
      Exit Function
   End If
   If loForm.IsReadOnly Then
      MsgBox "Form is set to Read Only, you will not be able to save any changes.", vbExclamation
   End If
   '
   loForm.RunFormDll "OnActivate"
   If loForm.EditForm Then
      '
      ' True option leaves Excel Application open, while ridding ourselves of the object...
      '
      goSession.TearDownForm loForm, True
      EditVoyForm = True
   Else
      goSession.TearDownForm loForm
      EditVoyForm = False
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.EditVoyForm: " & FullFilename, err.Number, err.Description
   goSession.TearDownForm loForm
   EditVoyForm = False
   goSession.TearDownForm loForm
   
End Function

Public Function SubmitVoyForms(EventType As Long, EventDetailKey As Long) As Boolean
   Dim strSQL As String
   Dim loRS As Recordset
   Dim loForm As mwForm
   Dim strFile As String
   On Error GoTo FunctionError
   Set loRS = FetchEventReportsRS(EventType, EventDetailKey, True)
   ' Any forms ?
   If loRS.RecordCount < 1 Then
      SubmitVoyForms = True
      goSession.CloseRecordset loRS
      Exit Function
   End If
   '
   ' any mandatory, not created ?
   '
   loRS.Filter = "IsCreated = False and IsMandatoryForm = True"
   If loRS.RecordCount > 0 Then
      '
      goSession.RaiseError "Error in mwEventReportWork.SubmitVoyForms, a Mandatory form has not been created."
      SubmitVoyForms = False
      goSession.CloseRecordset loRS
   End If
   loRS.Filter = adFilterNone
   '
   ' Loop through each form
   '
   loRS.MoveFirst
   Do While Not loRS.EOF
      '
      Set loForm = goSession.OpenNewForm
      If loForm.OpenForm(loRS!FullFilename) Then
         'Print Form ?
            
         If loRS.Fields("IsPrinted").value = False _
           And loRS.Fields("IsPrintOnSubmit").value = True Then
            loForm.PrintForm
            loRS!IsPrinted = True
         End If
         ' Submit form
         If loForm.SubmitForm() Then
            loRS!IsSubmitted = True
            loRS!SubmittedBy = goSession.User.UserID
            loRS!SubmittedDateTime = Now()
            loRS.Update
         End If
      Else
         goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwEventReportWork.SubmitVoyForms, OpenForm Failed for: " & loRS!FullFilename
      End If
      strFile = loForm.FormProperties.FullFilename
      goSession.TearDownForm loForm
      SetVoyFormReadOnly strFile
      loRS.MoveNext
   Loop
   '
   ' Close up shop
   '
   SubmitVoyForms = True
   goSession.CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.SubmitVoyForms. ", err.Number, err.Description
   CloseRecordset loRS
End Function

Public Function SubmitVoyForm(EventReportKey As Long) As Boolean
   Dim loForm As mwForm
   Dim lIsSent As Boolean
   Dim loRS As Recordset
   Dim strFile As String
   Dim strSQL As String
   On Error GoTo FunctionError
   '
   ' Fetch Event Form Record
   '
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   
   strSQL = "select * from mwEventReportLog where ID=" & EventReportKey
   loRS.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   ' Any forms ?
   If loRS.RecordCount < 1 Then
      goSession.RaiseError "Error in mwEventReportWork.SubmitVoyForm, Event Form Missing: " & EventReportKey
      SubmitVoyForm = True
      goSession.CloseRecordset loRS
      Exit Function
   End If
   ' Already Submitted ?
   If loRS!IsSubmitted Then
      goSession.RaiseError "Error in mwEventReportWork.SubmitVoyForm, Form Already Submitted: " & EventReportKey
      SubmitVoyForm = True
      goSession.CloseRecordset loRS
      Exit Function
   End If
   ' OK, go for it
   Set loForm = goSession.OpenNewForm
   If loForm.OpenForm(loRS!FullFilename) Then
      ' Print Form ?
      
      If loForm.SubmitForm() Then
         lIsSent = True
         loRS!IsSubmitted = True
         loRS!SubmittedBy = goSession.User.UserID
         loRS!SubmittedDateTime = Now()
         loRS.Update
      End If
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwEventReportWork.SubmitVoyForm, OpenForm Failed for: " & loRS!FullFilename
   End If
   strFile = loForm.FormProperties.FullFilename
   goSession.TearDownForm loForm
   If lIsSent Then
      SetVoyFormReadOnly strFile
   End If
   '
   ' Close up shop
   '
   SubmitVoyForm = lIsSent
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.SubmitVoyForm. ", err.Number, err.Description
End Function

Public Function SetVoyFormReadOnly(FullFilename As String) As Boolean

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
   SetVoyFormReadOnly = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.SetVoyFormReadOnly. ", err.Number, err.Description
   goSession.KillObject loFile
   goSession.KillObject fso
   SetVoyFormReadOnly = False
End Function

Public Function DeleteNotSubmittedVoyForms(EventType As Long, EventDetailKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim fso As FileSystemObject
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select * from mwEventReportLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventDetailKey
   loRS.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If loRS.RecordCount < 1 Then
      DeleteNotSubmittedVoyForms = True
      goSession.CloseRecordset loRS
      Exit Function
   End If
   Set fso = New FileSystemObject
   ' delete forms if not submitted
   Do While Not loRS.EOF
      If Not loRS!IsSubmitted Then
         If loRS!IsCreated Then
            fso.DeleteFile (loRS!FullFilename)
            goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwEventReportWork, deleting form: " & loRS!FullFilename
         End If
         loRS.Delete
      Else
         goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventReportWork, skip delete submitted form: " & loRS!FullFilename
      End If
      loRS.MoveNext
   Loop
   Set fso = Nothing
   goSession.CloseRecordset loRS
   DeleteNotSubmittedVoyForms = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.DeleteNotSubmittedVoyForms. ", err.Number, err.Description
   DeleteNotSubmittedVoyForms = False
End Function

Public Function DeleteVoyForm(EventReportKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim fso As FileSystemObject
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select * from mwEventReportLog where ID= & EventReportKey"
   loRS.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If loRS.RecordCount < 1 Then
      DeleteVoyForm = True
      goSession.CloseRecordset loRS
      Exit Function
   End If
   Set fso = New FileSystemObject
   ' delete forms if not submitted
   If Not loRS!IsSubmitted Then
      If loRS!IsCreated Then
         fso.DeleteFile (loRS!FullFilename)
         goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwEventReportWork, deleting form: " & loRS!FullFilename
      End If
      loRS.Delete
      DeleteVoyForm = True
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventReportWork, skip delete submitted form: " & loRS!FullFilename
      DeleteVoyForm = False
   End If
   Set fso = Nothing
   goSession.CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.DeleteVoyForm. ", err.Number, err.Description
   DeleteVoyForm = False
End Function


Public Function IsSubmittedForms(EventType As Long, EventDetailKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim strSQL As String
   Dim loRS As Recordset
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select count(*) from mwEventReportLog where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventDetailKey & " and IsSubmitted=True"
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   ' Got Forms ?
   If loRS.Fields(0).value > 0 Then
      IsSubmittedForms = True
   Else
      IsSubmittedForms = False
   End If
   goSession.CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventReportWork.IsSubmittedForms. ", err.Number, err.Description
   IsSubmittedForms = False
End Function

Public Function DeleteUnusedFormRecords() As Boolean
   DeleteUnusedFormRecords = True
   MsgBox "stump code"
End Function


Private Function FetchFormTemplateDescription(FormTypeKey As Long) As String
   Dim loRS As Recordset
   Dim strSQL As String
   Dim strTemplateID As String
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select TemplateID from mwEventReportType where ID=" & FormTypeKey
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.RecordCount < 1 Then
      FetchFormTemplateDescription = ""
      CloseRecordset loRS
      Exit Function
   End If
   strTemplateID = loRS!TemplateID
   loRS.Close
   strSQL = "select Description from mwwfFormTemplate where TemplateID='" & strTemplateID & "'"
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.RecordCount < 1 Then
      FetchFormTemplateDescription = ""
      CloseRecordset loRS
      Exit Function
   End If
   FetchFormTemplateDescription = loRS!Description
   CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventActivity.FetchFormTemplateDescription. ", err.Number, err.Description
   CloseRecordset loRS
   FetchFormTemplateDescription = ""
End Function


Public Function IsEventReportTypesExist(EventType As Long) As Boolean
   Dim loRS As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   Set loRS = New Recordset
   loRS.CursorLocation = adUseClient
   strSQL = "select count(*) from mwEventReportType where mwEventTypeKey=" & EventType
   loRS.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRS.Fields(0).value > 0 Then
      IsEventReportTypesExist = True
   Else
      IsEventReportTypesExist = False
   End If
   CloseRecordset loRS
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventReportsWork.IsEventReportTypesExist. ", err.Number, err.Description
End Function




