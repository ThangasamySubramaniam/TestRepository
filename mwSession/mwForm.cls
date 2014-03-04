VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Top Level Class for Document processing
' 3/6/2001 ms
' July/August 2001 ms - Overhaul Object Model
'
Option Explicit

'Version 2 Forms
Private Const MW_CAT_NUMFIELDS = 9
Private Const MW_CAT_EXTERNAL_NUMFIELDS = 6
Private Const MW_CAT_TEMPLATE_ID = "0"
Private Const MW_CAT_CURRENT_WORKFLOW_STAGE = "5"
Private Const MW_CAT_SITE_ID = "6"
Private Const MW_CAT_ROLE_TYPE = "7"
Private Const MW_CAT_FORM_ID = "8"

'Version 3 Forms
Private Const MW_CAT3_NUMFIELDS = 5
Private Const MW_CAT3_EXTERNAL_NUMFIELDS = 1
Private Const MW_CAT3_TEMPLATE_ID = "0"
Private Const MW_CAT3_CURRENT_WORKFLOW_STAGE = "1"
Private Const MW_CAT3_SITE_ID = "2"
Private Const MW_CAT3_ROLE_TYPE = "3"
Private Const MW_CAT3_FORM_ID = "4"

Private Const MW_CAT4_NUMFIELDS = 7
Private Const MW_CAT4_EVENT_TYPE = "5"
Private Const MW_CAT4_EVENT_RECORD_KEY = "6"
Private Const MW_CAT4_EVENT_VOYAGE_ID = "7"



Private Const MW_FORMS_ARCHIVEFILE = "SentFormArchive.zip"
Private Const MW_FORMS_ERROR_FILE = "SentFormError.zip"

Public Enum mw_WhichSiteCode
   mw_SiteID = 1
   mw_ShortSiteID = 2
   mw_YetAnotherSiteID = 3
End Enum

Public Enum mw_WhichWorkflowDirection
   mw_SendToShip
   mw_SendToShore
End Enum


Public Enum mw_FormType
   mwft_EXCEL = 1
   mwft_WORD = 2
   mwft_DATA_FORM = 3
   mwft_EXCEL_XLSX = 11
   mwft_EXCEL_XLSM = 12
   mwft_WORD_DOCX = 21
   mwft_WORD_DOCM = 22
End Enum

Public Enum mw_WorkflowActions
   mw_UNWRAP_CONTAINER
   mw_PROCESS_PACKAGE
   mw_PROCESS_RS_DATAFORM
   mw_PROCESS_SCRIPT
End Enum

' Class Properties
Private mIsNewForm As Boolean
Private mIsOpen As Boolean
Private mIsTemplateOpen As Boolean
Private mIsOfficeTemplate As Boolean
Private mDeleteAfterSend As Boolean
Private mSuppressOpenError As Boolean
Private mFullFilename As String
Private mTextBodyOverride As String
Private mIsIncompleteForm As Boolean
Private mIsReadOnly As Boolean
Private mIsScriptCancelSubmit As Boolean
Private mIsScriptCancelWorkflowProcessing As Boolean
Private mFormVersion As Integer
Private mIsPocketPcFormCompatible As Boolean
Private mIsSuppressGetSubject As Boolean
Private mIsPartialSubmit As Boolean
Private mIsResourceOpenError As Boolean
Private mVoyageID As String
'
Private mIsExternalDbConnection As Boolean
Private moExternalCon As Connection
Private mFormHeadersKey As Long
Private mFormHeadersStatus As Long

' Working Objects

Private moFormProperties As mwFormProperties
Private moTemplateProperties As mwFormTemplateProperties
Private moWorkflowProperties As mwFormWorkflowProperties
Private moExcelForm As mwExcelForm
Private moWordForm As mwWordForm
Private moDataForm As mwDataForm
Private moParent As Session
'Private moDataConnection As Connection
Private moExtendedKeys As ConfigKeys
Private moRsXmlConversions As Recordset
Private moRsEventDetail As Recordset



Friend Property Set Parent(oParent As mwSession.Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Friend Property Get Parent() As mwSession.Session
   Set Parent = moParent
End Property


Public Property Get FormProperties() As mwFormProperties
    Set FormProperties = moFormProperties
End Property

Public Property Set FormProperties(Data As mwFormProperties)
    Set moFormProperties = Data
End Property

Public Property Get TemplateProperties() As mwFormTemplateProperties
    Set TemplateProperties = moTemplateProperties
End Property

Public Property Let TemplateProperties(Data As mwFormTemplateProperties)
    Set moTemplateProperties = Data
End Property

Public Property Get WorkflowProperties() As mwFormWorkflowProperties
    Set WorkflowProperties = moWorkflowProperties
End Property

Public Property Let WorkflowProperties(Data As mwFormWorkflowProperties)
    Set moWorkflowProperties = Data
End Property

Public Property Get IsResourceOpenError() As Boolean
    IsResourceOpenError = mIsResourceOpenError
End Property

Public Property Let IsResourceOpenError(ByVal Data As Boolean)
    mIsResourceOpenError = Data
End Property


Friend Property Let IsOpen(Data As Boolean)
    mIsOpen = Data
End Property


Public Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

Public Property Get ExcelForm() As mwExcelForm
    Set ExcelForm = moExcelForm
End Property

Public Property Get DataForm() As mwDataForm
    Set DataForm = moDataForm
End Property

Public Property Get WordForm() As mwWordForm
    Set WordForm = moWordForm
End Property


Public Property Get FormType() As mw_FormType
    FormType = moTemplateProperties.FormType
End Property

Friend Property Let IsTemplateOpen(ByVal Data As Boolean)
    mIsTemplateOpen = Data
End Property

Public Property Get IsTemplateOpen() As Boolean
    IsTemplateOpen = mIsTemplateOpen
End Property

Public Property Let DeleteAfterSend(ByVal Data As Boolean)
    mDeleteAfterSend = Data
End Property

Public Property Get DeleteAfterSend() As Boolean
    DeleteAfterSend = mDeleteAfterSend
End Property

Public Property Let SuppressOpenError(ByVal Data As Boolean)
    mSuppressOpenError = Data
End Property

Public Property Get SuppressOpenError() As Boolean
    SuppressOpenError = mSuppressOpenError
End Property

Public Property Let FormHeadersKey(ByVal Data As Long)
    mFormHeadersKey = Data
End Property

Public Property Get FormHeadersKey() As Long
    FormHeadersKey = mFormHeadersKey
End Property

Public Property Let FormHeadersStatus(ByVal Data As Long)
    mFormHeadersStatus = Data
End Property

Public Property Get FormHeadersStatus() As Long
    FormHeadersStatus = mFormHeadersStatus
End Property

Public Property Let VoyageID(ByVal Data As String)
    mVoyageID = Data
End Property

Public Property Get VoyageID() As String
    VoyageID = mVoyageID
End Property


Public Property Let TextBodyOverride(ByVal Data As String)
    mTextBodyOverride = Data
End Property

Public Property Get TextBodyOverride() As String
    TextBodyOverride = mTextBodyOverride
End Property

Public Property Let IsIncompleteForm(ByVal Data As Boolean)
    mIsIncompleteForm = Data
End Property

Public Property Get IsIncompleteForm() As Boolean
    IsIncompleteForm = mIsIncompleteForm
End Property

Public Property Let IsReadOnly(ByVal Data As String)
    mIsReadOnly = Data
End Property

Public Property Get IsReadOnly() As String
    IsReadOnly = mIsReadOnly
End Property

Public Property Let IsScriptCancelSubmit(ByVal Data As Boolean)
    mIsScriptCancelSubmit = Data
End Property

Public Property Get IsScriptCancelSubmit() As Boolean
    IsScriptCancelSubmit = mIsScriptCancelSubmit
End Property

Public Property Let IsScriptCancelWorkflowProcessing(ByVal Data As Boolean)
    mIsScriptCancelWorkflowProcessing = Data
End Property

Public Property Get IsScriptCancelWorkflowProcessing() As Boolean
    IsScriptCancelWorkflowProcessing = mIsScriptCancelWorkflowProcessing
End Property

Public Property Get FormVersion() As Integer
    FormVersion = mFormVersion
End Property

Friend Property Let FormVersion(ByVal Data As Integer)
    mFormVersion = Data
End Property

Public Property Let IsPocketPcFormCompatible(ByVal Data As String)
    mIsPocketPcFormCompatible = Data
End Property

Public Property Get IsPocketPcFormCompatible() As String
    IsPocketPcFormCompatible = mIsPocketPcFormCompatible
End Property

Public Property Let IsSuppressGetSubject(ByVal Data As Boolean)
    mIsSuppressGetSubject = Data
End Property

Public Property Get IsSuppressGetSubject() As Boolean
    IsSuppressGetSubject = mIsSuppressGetSubject
End Property

Public Property Let IsPartialSubmit(ByVal Data As Boolean)
    mIsPartialSubmit = Data
End Property

Public Property Get IsPartialSubmit() As Boolean
    IsPartialSubmit = mIsPartialSubmit
End Property

Public Property Get EventDetailRS() As Recordset
   On Error GoTo PropertyError
   If Not moRsEventDetail Is Nothing Then
      Set EventDetailRS = moRsEventDetail
   End If
   Exit Property
PropertyError:
   goSession.RaiseError "General Error in mwSession.mwForm.Get.EventDetailRS.", Err.Number, Err.Description
End Property

Public Property Set EventDetailRS(Data As Recordset)
   On Error GoTo PropertyError
   Set moRsEventDetail = Data
   Exit Property
PropertyError:
   goSession.RaiseError "General Error in mwSession.mwForm.Set.EventDetailRS.", Err.Number, Err.Description
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

Public Function OpenForm(Optional FullFilename As String, _
  Optional LeaveApplicationOpen As Boolean, _
  Optional FormHeadersKey As Long) As Boolean
   ' Figure out the Form Type (based on extension), and open it
   ' Create a working object when it is open, exposing content
   ' Raise error flag if document is invalid...
   Dim fso As FileSystemObject
   Dim strExtension As String
   Dim strSiteKey As String
   Dim nStatus As Long
   On Error GoTo FormError
   DoEvents
   If FormHeadersKey > 0 Then
      mFormHeadersKey = FormHeadersKey
      mFormHeadersStatus = GetFormHeadersStatus(mFormHeadersKey)
   End If
      
      
   Set fso = New FileSystemObject
   strExtension = UCase(fso.GetExtensionName(FullFilename))
   If strExtension = "DOC" Then
      moTemplateProperties.FormType = mwft_WORD
      Set moWordForm = New mwWordForm
      Set moWordForm.Parent = Me
      mIsOpen = moWordForm.OpenForm(FullFilename)
   ElseIf strExtension = "DOCX" Then
      moTemplateProperties.FormType = mwft_WORD_DOCX
      Set moWordForm = New mwWordForm
      Set moWordForm.Parent = Me
      mIsOpen = moWordForm.OpenForm(FullFilename)
   ElseIf strExtension = "DOCM" Then
      moTemplateProperties.FormType = mwft_WORD_DOCM
      Set moWordForm = New mwWordForm
      Set moWordForm.Parent = Me
      mIsOpen = moWordForm.OpenForm(FullFilename)
   ElseIf strExtension = "XLS" Then
      moTemplateProperties.FormType = mwft_EXCEL
      Set moExcelForm = New mwExcelForm
      Set moExcelForm.Parent = Me
      mIsOpen = moExcelForm.OpenForm(FullFilename, LeaveApplicationOpen)
   ElseIf strExtension = "XLSX" Then
      moTemplateProperties.FormType = mwft_EXCEL_XLSX
      Set moExcelForm = New mwExcelForm
      Set moExcelForm.Parent = Me
      mIsOpen = moExcelForm.OpenForm(FullFilename, LeaveApplicationOpen)
   ElseIf strExtension = "XLSM" Then
      moTemplateProperties.FormType = mwft_EXCEL_XLSM
      Set moExcelForm = New mwExcelForm
      Set moExcelForm.Parent = Me
      mIsOpen = moExcelForm.OpenForm(FullFilename, LeaveApplicationOpen)
   ElseIf strExtension = "DGM" Then
      moTemplateProperties.FormType = mwft_DATA_FORM
      Set moDataForm = New mwDataForm
      Set moDataForm.Parent = Me
      mIsOpen = moDataForm.OpenForm(FullFilename)
   ElseIf strExtension = "DAT" And goSession.Workflow.GetExtendedProperty("Version2Compatible") = "True" Then
      '
      ' 6/2002 Eastwind Group Version 2 Compatibility Mode
      '
      moTemplateProperties.FormType = mwft_DATA_FORM
      Set moDataForm = New mwDataForm
      Set moDataForm.Parent = Me
      mIsOpen = moDataForm.OpenForm(FullFilename)
   Else
      If Not mSuppressOpenError Then
         goSession.RaiseError "Error in mwSession.mwForm.OpenForm. Invalid Form extension: " & _
           strExtension, Err.Number, Err.Description
      End If
      mIsOpen = False
      'OpenForm = False
      'Exit Function
   End If
   If mIsOpen Then
      '
      ' Look for Override Template
      '
      moFormProperties.FullFilename = FullFilename
      strSiteKey = goSession.Site.GetShipProperty(moFormProperties.SiteID, "ID")
      If IsNumeric(strSiteKey) Then
         moFormProperties.SiteKey = CLng(goSession.Site.GetShipProperty(moFormProperties.SiteID, "ID"))
      Else
         '
         ' Invalid template ?
         '
'         goSession.RaiseWorkflowStatus "Error in mwSession.mwForm.OpenForm: " & moFormProperties.FormID & ", Invalid SiteID: " & moFormProperties.SiteID
'         OpenForm = False
'         Exit Function
      End If
      mDeleteAfterSend = moTemplateProperties.DeleteAfterTransmit
      If Trim(moFormProperties.FormID) <> "" Then
         'goSession.RaiseWorkflowStatus "Form Opened:  Template: " & _
          moTemplateProperties.TemplateID & ", Form ID: " & moFormProperties.FormID
      End If
      If OpenTemplateOverRides() Then
         goSession.RaiseWorkflowStatus "Form Opened with Override Properties for Site: " & _
           moFormProperties.SiteID
      End If
   End If
   '
   ' Fetch SiteKey...
   '
   OpenForm = mIsOpen
   Exit Function
FormError:
   goSession.RaiseError "General Error in mwSession.mwForm.OpenForm. ", Err.Number, Err.Description
   OpenForm = False
   goSession.KillObject fso
End Function

Public Function CloseForm() As Boolean
   On Error GoTo FunctionError
   If mIsOpen Then
      Select Case moTemplateProperties.FormType
         Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
            moWordForm.CloseWordForm
            
         Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
            moExcelForm.CloseExcelForm
            
         Case Is = mwft_DATA_FORM
            '
         
      End Select
   End If
   mIsTemplateOpen = False
   CloseForm = True
   mIsResourceOpenError = False
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.CloseForm. ", Err.Number, Err.Description
   CloseForm = False
End Function

Public Function GetCell(cellref As String) As String
   If mIsOpen And _
      (moTemplateProperties.FormType = mwft_EXCEL Or _
       moTemplateProperties.FormType = mwft_EXCEL_XLSX Or _
       moTemplateProperties.FormType = mwft_EXCEL_XLSM) Then
      GetCell = moExcelForm.GetCell(cellref)
   Else
      GetCell = ""
   End If
End Function

Public Function GetField(FieldName, Optional RecordNumber As Variant) As String
   Dim strFieldName As String
   On Error GoTo FunctionError:
   strFieldName = UCase(FieldName)
   If mIsOpen And moTemplateProperties.FormType = mwft_DATA_FORM Then
      If IsMissing(RecordNumber) Then
         If moDataForm.Records(2).ConfigKeys(strFieldName) Is Nothing Then
            GetField = ""
         Else
            GetField = moDataForm.Records(2).ConfigKeys(strFieldName).KeyValue
         End If
      Else
         If Not IsNumeric(RecordNumber) Then
            goSession.RaiseError "Error in mwSession.GetField, RecordNumber must contain a number."
            GetField = ""
         Else
            'if RecordNumber > moDataForm.
            GetField = moDataForm.Records(RecordNumber).ConfigKeys(strFieldName).KeyValue
         End If
      End If
   Else
      GetField = ""
   End If
   Exit Function
FunctionError:
   'goSession.RaiseError "General Error in mwSession.GetField.", err.Number, err.Description
   GetField = ""
End Function

'
' Extract form information into properties...
'
Public Function ExtractMsFormInfo(category As String) As Boolean
   Dim strSplit() As String
   Dim iFields As Integer
   Dim oRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   
   strSplit = Split(category, "|")
   iFields = UBound(strSplit)
   If iFields = 0 Then
      ' Form doesn't have pipeline...
      ReDim strSplit(1)
      strSplit(0) = category
      iFields = 1
   End If
   Select Case iFields
      Case Is = MW_CAT_NUMFIELDS
         mFormVersion = 2
         mIsIncompleteForm = False
         moFormProperties.FormID = strSplit(MW_CAT_FORM_ID)
         moFormProperties.SiteID = strSplit(MW_CAT_SITE_ID)
         moFormProperties.RoleType = strSplit(MW_CAT_ROLE_TYPE)
         If Trim(moFormProperties.FormID) = "" Then
            ExtractMsFormInfo = False
            Exit Function
         End If
         moWorkflowProperties.WFStage = strSplit(MW_CAT_CURRENT_WORKFLOW_STAGE)
         If Trim(moWorkflowProperties.WFStage) = "" Then
            ExtractMsFormInfo = False
            Exit Function
         End If
      Case Is = MW_CAT_EXTERNAL_NUMFIELDS
         mFormVersion = 2
         mIsIncompleteForm = True
         moFormProperties.FormID = " "
         moFormProperties.SiteID = " "
         moFormProperties.RoleType = " "
      Case Is = MW_CAT3_NUMFIELDS
         '
         ' 3.x - OfficeForms Version
         '
         mFormVersion = 3
         mIsIncompleteForm = False
         If Not mIsNewForm Then
            moFormProperties.FormID = strSplit(MW_CAT3_FORM_ID)
            moFormProperties.SiteID = strSplit(MW_CAT3_SITE_ID)
            moFormProperties.RoleType = strSplit(MW_CAT3_ROLE_TYPE)
            If Trim(moFormProperties.FormID) = "" Then
               ExtractMsFormInfo = False
               Exit Function
            End If
            moWorkflowProperties.WFStage = strSplit(MW_CAT3_CURRENT_WORKFLOW_STAGE)
            If Trim(moWorkflowProperties.WFStage) = "" Then
               ExtractMsFormInfo = False
               Exit Function
            End If
         End If
      Case Is = MW_CAT3_EXTERNAL_NUMFIELDS
         mFormVersion = 3
         mIsIncompleteForm = True
      Case Is = MW_CAT4_NUMFIELDS, MW_CAT4_NUMFIELDS + 1
      '
      ' Voyage Forms Version
      '
         mFormVersion = 4
         mIsIncompleteForm = False
         If Not mIsNewForm Then
            moFormProperties.FormID = strSplit(MW_CAT3_FORM_ID)
            moFormProperties.SiteID = strSplit(MW_CAT3_SITE_ID)
            moFormProperties.RoleType = strSplit(MW_CAT3_ROLE_TYPE)
            ' Voyage Forms
            If IsNumeric(strSplit(MW_CAT4_EVENT_TYPE)) Then
               moFormProperties.EventType = strSplit(MW_CAT4_EVENT_TYPE)
            Else
               goSession.RaiseError "Error in mwForm.ExtractMsInfo, Invalid Event Type: " & strSplit(MW_CAT4_EVENT_TYPE)
            End If
            If IsNumeric(strSplit(MW_CAT4_EVENT_RECORD_KEY)) Then
               moFormProperties.EventRecordKey = strSplit(MW_CAT4_EVENT_RECORD_KEY)
            Else
               goSession.RaiseError "Error in mwForm.ExtractMsInfo, Invalid Event Record Key: " & strSplit(MW_CAT4_EVENT_RECORD_KEY)
            End If
            If Trim(moFormProperties.FormID) = "" Then
               ExtractMsFormInfo = False
               Exit Function
            End If
            moWorkflowProperties.WFStage = strSplit(MW_CAT3_CURRENT_WORKFLOW_STAGE)
            If Trim(moWorkflowProperties.WFStage) = "" Then
               ExtractMsFormInfo = False
               Exit Function
            End If
            If iFields = MW_CAT4_NUMFIELDS + 1 Then
               mVoyageID = strSplit(MW_CAT4_EVENT_VOYAGE_ID)
            End If
         End If
      Case Else
         'goSession.LogIt mwl_Workstation, mwl_Warning, "Logging error in mwForm.ExtractMsInfo, Category string = " & category
         If Len(moTemplateProperties.TemplateID) > 0 Then
            If Not mIsTemplateOpen Then
               If Not OpenTemplate(moTemplateProperties.TemplateID) Then
                  ExtractMsFormInfo = False
               Else
                  ExtractMsFormInfo = True
               End If
            Else
               ExtractMsFormInfo = True
            End If
         Else
            ExtractMsFormInfo = False
         End If
         Exit Function
   End Select
   '
   If Len(moTemplateProperties.TemplateID) = 0 Then
      moTemplateProperties.TemplateID = strSplit(MW_CAT_TEMPLATE_ID)
   End If
   If Not mIsTemplateOpen Then
      If Not OpenTemplate(moTemplateProperties.TemplateID) Then
         ExtractMsFormInfo = False
      Else
         ExtractMsFormInfo = True
      End If
   Else
      ExtractMsFormInfo = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.ExtractFormInfo.", Err.Number, Err.Description
   ExtractMsFormInfo = False
End Function

'
' Prepares the "Category" field for update into the Document's record.
'
Public Function EncodeMsOfficeFormInfo() As String
   On Error GoTo FunctionError
   If Not mIsOpen Then
      EncodeMsOfficeFormInfo = ""
      Exit Function
   End If
   If moFormProperties.EventType > 0 Then
      mFormVersion = 4
   End If
   If mFormVersion = 2 Then
      '
      ' Version 2 coding
      '
      EncodeMsOfficeFormInfo = moTemplateProperties.TemplateID & "|1.0|" & moWorkflowProperties.WFStage _
        & "|N/A|N/A|N/A|" & moFormProperties.SiteID & "|" & moFormProperties.RoleType _
        & "|" & moFormProperties.FormID & "|"
   ElseIf mFormVersion = 3 Then
      '
      ' Version 3 coding
      '
      EncodeMsOfficeFormInfo = moTemplateProperties.TemplateID & "|" & moWorkflowProperties.WFStage _
        & "|" & moFormProperties.SiteID & "|" & moFormProperties.RoleType _
        & "|" & moFormProperties.FormID & "|"
   ElseIf mFormVersion = 4 Then
      '
      ' Version 4 coding
      '
      EncodeMsOfficeFormInfo = moTemplateProperties.TemplateID & "|" & moWorkflowProperties.WFStage _
        & "|" & moFormProperties.SiteID & "|" & moFormProperties.RoleType _
        & "|" & moFormProperties.FormID & "|" & moFormProperties.EventType & "|" & moFormProperties.EventRecordKey
      If Trim(mVoyageID) = "" Then
         EncodeMsOfficeFormInfo = EncodeMsOfficeFormInfo & "|" & goSession.Site.CurrentVoyage & "|"
      Else
         EncodeMsOfficeFormInfo = EncodeMsOfficeFormInfo & "|" & mVoyageID & "|"
      End If
        
   End If
   ' Once new form gets encoded, no longer considered "external"
   mIsIncompleteForm = False
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwForm.EnCodeMsOfficeFormInfo.", Err.Number, Err.Description
   EncodeMsOfficeFormInfo = ""
End Function




Private Sub Class_Initialize()
   Set moFormProperties = New mwFormProperties
   Set moFormProperties.Parent = Me
   Set moTemplateProperties = New mwFormTemplateProperties
   Set moWorkflowProperties = New mwFormWorkflowProperties
End Sub

Private Sub Class_Terminate()
   '
   '  Cleanup message object...
   '
   'goSession.Message.AttachmentFile = ""
   '
   On Error GoTo SubError
   Set moFormProperties = Nothing
   Set moTemplateProperties = Nothing
   Set moWorkflowProperties = Nothing
   
   If Not moExcelForm Is Nothing Then Set moExcelForm = Nothing
   If Not moWordForm Is Nothing Then Set moWordForm = Nothing
   If Not moDataForm Is Nothing Then Set moDataForm = Nothing
   If Not moRsXmlConversions Is Nothing Then
      moRsXmlConversions.Close
      Set moRsXmlConversions = Nothing
   End If
   '
   ' If external connection, close it now..
   '
   If Not moExternalCon Is Nothing Then
      If moExternalCon.State = adStateOpen Then
         moExternalCon.Close
      End If
      Set moExternalCon = Nothing
   End If
   CloseRecordset moRsEventDetail
   Exit Sub
SubError:
   goSession.RaiseError "Error in msSession.mwForm.Class_Terminate.", Err.Number, Err.Description
End Sub

Public Function IsFileExists(ByVal FileName As String) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   If fso.FileExists(FileName) Then
      IsFileExists = True
   Else
      IsFileExists = False
   End If
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "Error in msSession.mwForm.IsFileExists.", Err.Number, Err.Description
   IsFileExists = False
End Function


Public Function OpenTemplate(TemplateID As String) As Boolean
   Dim strSQL As String
   Dim strPlace As String
   Dim oRs As Recordset
   Dim loField As field
   On Error GoTo FunctionError
   If moParent Is Nothing Then
      goSession.RaiseError "Error in mwForm.OpenTemplate, improper use - use the the mwSession.OpenNewForm to instantiate a mwForm object."
      mIsTemplateOpen = False
      OpenTemplate = False
   End If
   If mIsTemplateOpen Then
      goSession.RaiseError "Error in mwSession.mwForm.OpenTemplate. Template already open."
      mIsTemplateOpen = False
      OpenTemplate = False
   End If
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
'   If App.Revision > 597 Then
      strSQL = "Select * from mwwfFormTemplate where TemplateID='" & TemplateID & "'"
'   Else
'      strSQL = "Select * from mwwfFormTemplates where TemplateID='" & TemplateID & "'"
'   End If
   
   oRs.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If oRs.RecordCount < 1 Then
      If Not SuppressOpenError Then
         goSession.RaiseError "Error in mwSession.mwForm.OpenTemplate, Template Not Found: " & TemplateID
      End If
      goSession.CloseRecordset oRs, True
      mIsTemplateOpen = False
      OpenTemplate = False
      Exit Function
   End If
   '
   ' 3/25/2002 ms - Extended Formtemplate Properties...
       
   If Not moExtendedKeys Is Nothing Then
      Set moExtendedKeys = Nothing
   End If
   Set moExtendedKeys = New ConfigKeys
   For Each loField In oRs.Fields
      Select Case UCase(loField.Name)
         
         Case Is = UCase("DefaultEmailAddress")
            If Not IsNull(oRs!DefaultEmailAddress) Then
               moTemplateProperties.DefaultEmailAddress = oRs!DefaultEmailAddress
            End If
         Case Is = UCase("Description")
            If Not IsNull(oRs!Description) Then
               moTemplateProperties.Description = oRs!Description
            End If
         Case Is = UCase("FilePrefix")
            If Not IsNull(oRs!FilePrefix) Then
               moTemplateProperties.FilePrefix = oRs!FilePrefix
            End If
         Case Is = UCase("FormType")
            If Not IsNull(oRs!FormType) Then
               moTemplateProperties.FormType = oRs!FormType
            End If
         Case Is = UCase("TemplateID")
            If Not IsNull(oRs!TemplateID) Then
               moTemplateProperties.TemplateID = oRs!TemplateID
            End If
         Case Is = UCase("DeleteAfterTransmit")
            If Not IsNull(oRs!DeleteAfterTransmit) Then
               moTemplateProperties.DeleteAfterTransmit = oRs!DeleteAfterTransmit
            End If
         Case Is = UCase("DeleteAfterTransmit")
            If Not IsNull(oRs!DeleteAfterTransmit) Then
               moTemplateProperties.DeleteAfterTransmit = oRs!DeleteAfterTransmit
            End If
         Case Is = UCase("MakeInsertStatement")
            If Not IsNull(oRs!MakeInsertStatement) Then
               moTemplateProperties.MakeInsertStatement = oRs!MakeInsertStatement
            End If
         Case Is = UCase("ExcludeExternalForms")
            If Not IsNull(oRs!ExcludeExternalForms) Then
               moTemplateProperties.ExcludeExternalForms = oRs!ExcludeExternalForms
            End If
         Case Is = UCase("ConvertToDatagram")
            If Not IsNull(oRs!ConvertToDatagram) Then
               moTemplateProperties.ConvertToDatagram = oRs!ConvertToDatagram
            End If
         Case Is = UCase("IsShipboardOnly")
            If Not IsNull(oRs!IsShipboardOnly) Then
               moTemplateProperties.IsShipboardOnly = oRs!IsShipboardOnly
            End If
         Case Is = UCase("DatagramTemplateID")
            If Not IsNull(oRs!DatagramTemplateID) Then
               moTemplateProperties.DatagramTemplateID = oRs!DatagramTemplateID
            End If
         Case Is = UCase("SourceFilename")
            If Not IsNull(oRs!SourceFilename) Then
               moTemplateProperties.SourceFilename = oRs!SourceFilename
            End If
         Case Is = UCase("DBConnectString")
'            If Not IsNull(oRs!DBConnectString) Then
'               moWorkflowProperties.DBConnectString = oRs!DBConnectString
'            End If
         Case Is = UCase("MapScriptFile")
            If Not IsNull(oRs!MapScriptFile) Then
               moWorkflowProperties.MapScriptFile = oRs!MapScriptFile
            End If
         Case Is = UCase("VbFormEventDLL")
            If Not IsNull(oRs!VbFormEventDLL) Then
               moWorkflowProperties.VbFormEventDLL = oRs!VbFormEventDLL
            End If
         Case Is = UCase("NotifyEmailError")
            If Not IsNull(oRs!NotifyEmailError) Then
               moWorkflowProperties.NotifyEmailError = oRs!NotifyEmailError
            End If
         Case Is = UCase("NotifyEmailOK")
            If Not IsNull(oRs!NotifyEmailOK) Then
               moWorkflowProperties.NotifyEmailOK = oRs!NotifyEmailOK
            End If
         Case Is = UCase("IsUpdateAsRecordset")
            If Not IsNull(oRs!IsUpdateAsRecordset) Then
               moWorkflowProperties.IsUpdateAsRecordset = oRs!IsUpdateAsRecordset
            End If
         Case Is = UCase("TolerateDuplicateRecords")
            If Not IsNull(oRs!TolerateDuplicateRecords) Then
               moWorkflowProperties.TolerateDuplicateRecords = oRs!TolerateDuplicateRecords
            End If
         Case Is = UCase("TolerateMissingColumns")
            If Not IsNull(oRs!TolerateMissingColumns) Then
               moWorkflowProperties.TolerateMissingColumns = oRs!TolerateMissingColumns
            End If
         Case Is = UCase("TolerateOtherErrors")
            If Not IsNull(oRs!TolerateOtherErrors) Then
               moWorkflowProperties.TolerateOtherErrors = oRs!TolerateOtherErrors
            End If
         Case Is = UCase("RouteToFolder")
            If Not IsNull(oRs!RouteToFolder) Then
               moWorkflowProperties.RouteToFolder = oRs!RouteToFolder
            End If
         Case Is = UCase("RouteToEmailAddress")
            If Not IsNull(oRs!RouteToEmailAddress) Then
               moWorkflowProperties.RouteToEmailAddress = oRs!RouteToEmailAddress
            End If
         Case Is = UCase("IsRouteToFolderExclusive")
            If Not IsNull(oRs!IsRouteToFolderExclusive) Then
               moWorkflowProperties.IsRouteToFolderExclusive = oRs!IsRouteToFolderExclusive
            End If
         Case Is = UCase("IsRouteToMailboxExclusive")
            If Not IsNull(oRs!IsRouteToMailboxExclusive) Then
               moWorkflowProperties.IsRouteToMailboxExclusive = oRs!IsRouteToMailboxExclusive
            End If
         Case Is = UCase("IsProtectWorkbook")
            If Not IsNull(oRs!IsProtectWorkbook) Then
               moTemplateProperties.IsProtectWorkbook = oRs!IsProtectWorkbook
            End If
         Case Is = UCase("ConnectionID")
            If Not IsNull(oRs!ConnectionID) Then
               moWorkflowProperties.ConnectionID = oRs!ConnectionID
            End If
         Case Else
            '
            ' Otherwise, Add to the Collection of configkeys... (extended properties)
            '
            If IsNull(loField.value) Then
               moExtendedKeys.Add UCase(loField.Name), "", UCase(loField.Name)
            Else
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
   goSession.CloseRecordset oRs, True
   mIsTemplateOpen = True
   OpenTemplate = True
   Exit Function
FunctionError:
   If Not SuppressOpenError Then
      goSession.RaiseError "Error in msSession.mwForm.OpenTemplate.", Err.Number, Err.Description
   End If
   OpenTemplate = False
End Function

'
' 1/17/2002 ms - Diferent FormTemplate properties based on a mwcSitesKey
' Returns True if Overrides were found, False if none were found
'
Public Function OpenTemplateOverRides() As Boolean
   Dim strSQL As String
   Dim strPlace As String
   Dim oRs As Recordset
   Dim loField As field
   On Error GoTo FunctionError
   If moParent Is Nothing Then
      'MsgBox "Error: Use the the mwSession.OpenNewForm to instantiate a mwForm object."
      OpenTemplateOverRides = False
      Exit Function
   End If
   '
   ' Look for Override of FormTemplate, based on mwcSitesKey...
   '
   If Not mIsOpen = True Then
      OpenTemplateOverRides = False
      Exit Function
   End If
   '
   ' Look for Override of FormTemplate, based on mwcSitesKey...
   '
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
'   If App.Revision > 597 Then
'      strSQL = "SELECT mwcFleetSites.mwcSitesKey, mwwfFormTemplate.TemplateID, mwwfFleetFormTemplate.* " & _
'        "FROM (mwwfFormTemplate INNER JOIN mwwfFleetFormTemplate ON mwwfFormTemplate.TemplateID = " & _
'        " mwwfFleetFormTemplate.TemplateID) INNER JOIN (mwcFleets INNER JOIN mwcFleetSites ON mwcFleets.ID = " & _
'        "mwcFleetSites.mwcFleetsKey) ON mwwfFleetFormTemplate.FleetKey = mwcFleets.ID " & _
'        "WHERE mwcFleetSites.mwcSitesKey='" & moFormProperties.SiteKey & "' AND mwwfFormTemplate.TemplateID='" & _
'        moTemplateProperties.TemplateID & "';"
'   Else
'      strSQL = "SELECT mwcFleetSites.mwcSitesKey, mwwfFormTemplates.TemplateID, mwwfFleetFormTemplates.* " & _
'        "FROM (mwwfFormTemplates INNER JOIN mwwfFleetFormTemplates ON mwwfFormTemplates.ID = " & _
'        " mwwfFleetFormTemplates.FormTemplateKey) INNER JOIN (mwcFleets INNER JOIN mwcFleetSites ON mwcFleets.ID = " & _
'        "mwcFleetSites.mwcFleetsKey) ON mwwfFleetFormTemplates.FleetKey = mwcFleets.ID " & _
'        "WHERE mwcFleetSites.mwcSitesKey='" & moFormProperties.SiteKey & "' AND mwwfFormTemplates.TemplateID='" & _
'        moTemplateProperties.TemplateID & "';"
'   End If
 
'   If App.Revision > 597 Then
      strSQL = "SELECT mwcFleetSites.mwcSitesKey, mwwfFormTemplate.TemplateID, mwwfFleetFormTemplate.* " & _
         " FROM mwwfFormTemplate , mwwfFleetFormTemplate, mwcFleets, mwcFleetSites " & _
         " WHERE mwwfFormTemplate.TemplateID = mwwfFleetFormTemplate.TemplateID AND " & _
         " mwcFleets.ID = mwcFleetSites.mwcFleetsKey AND " & _
         " mwwfFleetFormTemplate.FleetKey = mwcFleets.ID AND " & _
         " mwcFleetSites.mwcSitesKey = " & moFormProperties.SiteKey & " AND " & _
         " mwwfFormTemplate.TemplateID = '" & moTemplateProperties.TemplateID & "'"
'   Else
'      strSQL = "SELECT mwcFleetSites.mwcSitesKey, mwwfFormTemplates.TemplateID, mwwfFleetFormTemplates.* " & _
'         " FROM mwwfFormTemplates , mwwfFleetFormTemplates, mwcFleets, mwcFleetSites " & _
'         " WHERE mwwfFormTemplates.ID = mwwfFleetFormTemplates.FormTemplateKey AND " & _
'         " mwcFleets.ID = mwcFleetSites.mwcFleetsKey AND " & _
'         " mwwfFleetFormTemplates.FleetKey = mwcFleets.ID AND " & _
'         " mwcFleetSites.mwcSitesKey = '" & moFormProperties.SiteKey & "' AND " & _
'         " mwwfFormTemplates.TemplateID = '" & moTemplateProperties.TemplateID & "'"
'   End If
   
   oRs.Open strSQL, goSession.DBConnection, adOpenStatic, adLockReadOnly
   If oRs.RecordCount < 1 Then
      goSession.CloseRecordset oRs, True
      OpenTemplateOverRides = False
      Exit Function
   End If
   '
   ' OK we have one...
   '
   moTemplateProperties.IsOverride = True
   
   For Each loField In oRs.Fields
      Select Case UCase(loField.Name)
         
         Case Is = UCase("DefaultEmailAddress")
            If Not IsNull(oRs!DefaultEmailAddress) Then
               moTemplateProperties.DefaultEmailAddress = oRs!DefaultEmailAddress
            End If
         Case Is = UCase("Description")
            If Not IsNull(oRs!Description) Then
               moTemplateProperties.Description = oRs!Description
            End If
         Case Is = UCase("FilePrefix")
            If Not IsNull(oRs!FilePrefix) Then
               moTemplateProperties.FilePrefix = oRs!FilePrefix
            End If
         Case Is = UCase("FormType")
            If Not IsNull(oRs!FormType) Then
               moTemplateProperties.FormType = oRs!FormType
            End If
         Case Is = UCase("TemplateID")
            If Not IsNull(oRs!TemplateID) Then
               moTemplateProperties.TemplateID = oRs!TemplateID
            End If
         Case Is = UCase("DeleteAfterTransmit")
            If Not IsNull(oRs!DeleteAfterTransmit) Then
               moTemplateProperties.DeleteAfterTransmit = oRs!DeleteAfterTransmit
            End If
         Case Is = UCase("DeleteAfterTransmit")
            If Not IsNull(oRs!DeleteAfterTransmit) Then
               moTemplateProperties.DeleteAfterTransmit = oRs!DeleteAfterTransmit
            End If
         Case Is = UCase("MakeInsertStatement")
            If Not IsNull(oRs!MakeInsertStatement) Then
               moTemplateProperties.MakeInsertStatement = oRs!MakeInsertStatement
            End If
         Case Is = UCase("ExcludeExternalForms")
            If Not IsNull(oRs!ExcludeExternalForms) Then
               moTemplateProperties.ExcludeExternalForms = oRs!ExcludeExternalForms
            End If
         Case Is = UCase("ConvertToDatagram")
            If Not IsNull(oRs!ConvertToDatagram) Then
               moTemplateProperties.ConvertToDatagram = oRs!ConvertToDatagram
            End If
         Case Is = UCase("IsShipboardOnly")
            If Not IsNull(oRs!IsShipboardOnly) Then
               moTemplateProperties.IsShipboardOnly = oRs!IsShipboardOnly
            End If
         Case Is = UCase("DatagramTemplateID")
            If Not IsNull(oRs!DatagramTemplateID) Then
               moTemplateProperties.DatagramTemplateID = oRs!DatagramTemplateID
            End If
         Case Is = UCase("SourceFilename")
            If Not IsNull(oRs!SourceFilename) Then
               moTemplateProperties.SourceFilename = oRs!SourceFilename
            End If
         Case Is = UCase("DBConnectString")
'            If Not IsNull(oRs!DBConnectString) Then
'               moWorkflowProperties.DBConnectString = oRs!DBConnectString
'            End If
         Case Is = UCase("MapScriptFile")
            If Not IsNull(oRs!MapScriptFile) Then
               moWorkflowProperties.MapScriptFile = oRs!MapScriptFile
            End If
         Case Is = UCase("VbFormEventDLL")
            If Not IsNull(oRs!VbScriptFile) Then
               moWorkflowProperties.VbFormEventDLL = oRs!VbFormEventDLL
            End If
         Case Is = UCase("NotifyEmailError")
            If Not IsNull(oRs!NotifyEmailError) Then
               moWorkflowProperties.NotifyEmailError = oRs!NotifyEmailError
            End If
         Case Is = UCase("NotifyEmailOK")
            If Not IsNull(oRs!NotifyEmailOK) Then
               moWorkflowProperties.NotifyEmailOK = oRs!NotifyEmailOK
            End If
         Case Is = UCase("IsUpdateAsRecordset")
            If Not IsNull(oRs!IsUpdateAsRecordset) Then
               moWorkflowProperties.IsUpdateAsRecordset = oRs!IsUpdateAsRecordset
            End If
         Case Is = UCase("TolerateDuplicateRecords")
            If Not IsNull(oRs!TolerateDuplicateRecords) Then
               moWorkflowProperties.TolerateDuplicateRecords = oRs!TolerateDuplicateRecords
            End If
         Case Is = UCase("TolerateMissingColumns")
            If Not IsNull(oRs!TolerateMissingColumns) Then
               moWorkflowProperties.TolerateMissingColumns = oRs!TolerateMissingColumns
            End If
         Case Is = UCase("TolerateOtherErrors")
            If Not IsNull(oRs!TolerateOtherErrors) Then
               moWorkflowProperties.TolerateOtherErrors = oRs!TolerateOtherErrors
            End If
         Case Is = UCase("RouteToFolder")
            If Not IsNull(oRs!RouteToFolder) Then
               moWorkflowProperties.RouteToFolder = oRs!RouteToFolder
            End If
         Case Is = UCase("RouteToEmailAddress")
            If Not IsNull(oRs!RouteToEmailAddress) Then
               moWorkflowProperties.RouteToEmailAddress = oRs!RouteToEmailAddress
            End If
         Case Is = UCase("IsRouteToFolderExclusive")
            If Not IsNull(oRs!IsRouteToFolderExclusive) Then
               moWorkflowProperties.IsRouteToFolderExclusive = oRs!IsRouteToFolderExclusive
            End If
         Case Is = UCase("IsRouteToMailboxExclusive")
            If Not IsNull(oRs!IsRouteToMailboxExclusive) Then
               moWorkflowProperties.IsRouteToMailboxExclusive = oRs!IsRouteToMailboxExclusive
            End If
         Case Else
            '
            ' Otherwise, determine if already in collection and either add it or replace it...
            '
            If moExtendedKeys(loField.Name) Is Nothing Then
               ' New key in override table only
               If IsNull(loField.value) Then
                  moExtendedKeys.Add UCase(loField.Name), "", UCase(loField.Name)
               Else
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
            Else
               ' Replace value
               If Not IsNull(loField.value) Then
                  If Trim(loField.value) <> "" Then
                     If goSession.IsOracle And loField.Type = adNumeric And loField.Precision = 1 Then
                        If loField.value = 0 Then
                           moExtendedKeys(loField.Name).KeyValue = "False"
                        Else
                           moExtendedKeys(loField.Name).KeyValue = "True"
                        End If
                     Else
                        moExtendedKeys(loField.Name).KeyValue = loField.value
                     End If
                  End If
               End If
            End If
            
      End Select
   Next
   
   '
   ' Popoulate the FormTemplate and WorkflowProperties object properties...
   '
   'If Not IsNull(oRs!DefaultEmailAddress) Then
   '   moTemplateProperties.DefaultEmailAddress = oRs!DefaultEmailAddress
   'End If
   'If Not IsNull(oRs!Description) Then
   '   moTemplateProperties.Description = oRs!Description
   'End If
   'If Not IsNull(oRs!FilePrefix) Then
   '   moTemplateProperties.FilePrefix = oRs!FilePrefix
   'End If
  '
  ' If Not IsNull(oRs!DBConnectString) Then
  '    moWorkflowProperties.DBConnectString = oRs!DBConnectString
  ' End If
  '
  ' If Not IsNull(oRs!MapScriptFile) Then
  '    moWorkflowProperties.MapScriptFile = oRs!MapScriptFile
  ' End If
  '
  ' If Not IsNull(oRs!VbScriptFile) Then
  '    moWorkflowProperties.VbScriptFile = oRs!VbScriptFile
  ' End If
  '
  ' If Not IsNull(oRs!NotifyEmailError) Then
  '    moWorkflowProperties.NotifyEmailError = oRs!NotifyEmailError
  ' End If
  '
  ' If Not IsNull(oRs!NotifyEmailOK) Then
  '    moWorkflowProperties.NotifyEmailOK = oRs!NotifyEmailOK
  ' End If
  '
  ' If Not IsNull(oRs!RouteToFolder) Then
  '    moWorkflowProperties.RouteToFolder = oRs!RouteToFolder
  ' End If
  '
  ' If Not IsNull(oRs!RouteToEmailAddress) Then
  '    moWorkflowProperties.RouteToEmailAddress = oRs!RouteToEmailAddress
  ' End If
   '
   '
   goSession.CloseRecordset oRs, True
   '
   ' ok, it's done...
   '
   goSession.CloseRecordset oRs, True
   OpenTemplateOverRides = True
   Exit Function

FunctionError:
    goSession.RaiseError "General error in mwForm.OpenTemplateOverRides. ", Err.Number, Err.Description
    goSession.CloseRecordset oRs, True
    OpenTemplateOverRides = False
    
End Function


Public Function CreateFormIdentity(Optional NewFormID As String) As Boolean
   Dim iSeq As Integer
   Dim sShortID As String
   Dim strExtension As String
   On Error GoTo FunctionError
   If Not mIsTemplateOpen Then
      goSession.RaiseError "Error in mwSession.mwForm.CreateFormIdentity, Must Open a Form Template First."
      CreateFormIdentity = ""
      Exit Function
   End If
   sShortID = goSession.User.GetExtendedProperty("ShortUserID")
   Select Case moTemplateProperties.FormType
      Case Is = mwft_EXCEL
         strExtension = ".xls"
      Case Is = mwft_EXCEL_XLSX
         strExtension = ".xlsX"
      Case Is = mwft_EXCEL_XLSM
         strExtension = ".xlsM"
      Case Is = mwft_WORD
         strExtension = ".doc"
      Case Is = mwft_WORD_DOCX
         strExtension = ".docX"
      Case Is = mwft_WORD_DOCM
         strExtension = ".docM"
      Case Is = mwft_DATA_FORM
         strExtension = ".dgm"
   End Select
   '
   If Trim(NewFormID) = "" Then
      
      iSeq = goSession.User.GetNextFormSequenceNumber
      
      If sShortID = "" Then
         moFormProperties.FormID = UCase(Trim(moTemplateProperties.FilePrefix)) & UCase(Trim(goSession.Site.ShortSiteID)) _
        & UCase(Trim(goSession.User.UserID)) & LTrim(Trim(str(iSeq)))
      Else
         moFormProperties.FormID = UCase(Trim(moTemplateProperties.FilePrefix)) & UCase(Trim(goSession.Site.ShortSiteID)) _
           & sShortID & Format(iSeq, "0000")
      End If
   Else
      moFormProperties.FormID = NewFormID
   End If
   moFormProperties.FileName = moFormProperties.FormID & strExtension
   CreateFormIdentity = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.CreateFormIdentity.", Err.Number, Err.Description
   CreateFormIdentity = False
End Function


Public Function MakeNewForm(Optional TemplateID As String, _
  Optional TargetFolder As String, Optional SourceFilename As String, _
  Optional NewFormID As String, Optional TargetSiteID As String) As Boolean
   Dim strNewFormName As String
   Dim strTempFilename As String
   Dim TemplateSourceFName As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   ' Only one form open per instance...
   If mIsOpen Then
      goSession.RaiseError "Error in mwForm.MakeNewForm, Form already open."
      MakeNewForm = False
      Exit Function
   End If
   If TemplateID <> "" Then
      moTemplateProperties.TemplateID = TemplateID
   End If
   If Trim(moTemplateProperties.TemplateID) = "" Then
      goSession.RaiseError "Error in mwForm.MakeNewForm, No Template Specified."
      MakeNewForm = False
      Exit Function
   Else
      TemplateID = moTemplateProperties.TemplateID
   End If
   '
   ' open the template
   '
   If Not mIsTemplateOpen Then
      If Not OpenTemplate(TemplateID) Then
         'goSession.RaiseError "Error in mwSession.mwForm.MakeNewForm, Template Not Found in Database:" _
           & TemplateID
         MakeNewForm = False
         Exit Function
      End If
   End If
   If moTemplateProperties.FormType < 1 Then
      goSession.RaiseError "Error in mwForm.MakeNewForm, No FormType specified for: " & TemplateID
      MakeNewForm = False
      Exit Function
   End If
   '
   ' Optional TargetSiteID
   '
   If TargetSiteID <> "" Then
      moFormProperties.TargetSiteID = TargetSiteID
      If moFormProperties.TargetSiteKey < 1 Then
         moFormProperties.TargetSiteKey = goSession.Site.GetSiteKey(TargetSiteID)
      End If
   End If
   '
   ' Source File must exist if it is specified...
   '
   Set fso = New FileSystemObject
   
   If moTemplateProperties.FormType <> mwft_DATA_FORM Then

      TemplateSourceFName = GetTemplateFileName(0)
   '   If TemplateSourceFName = "" And goSession.Workflow.TemplatePlacesID = 0 Then
      If TemplateSourceFName = "" Then
         TemplateSourceFName = moTemplateProperties.SourceFilename
      End If
      If Trim(TemplateSourceFName) <> "" Then
         If Not fso.FileExists(TemplateSourceFName) Then
            goSession.RaiseError "Error in mwForm.MakeNewForm, Template: " & TemplateID & " Source file is missing: " & TemplateSourceFName & vbCrLf & vbCrLf & "The Form Template may not have been distributed yet."
            mIsTemplateOpen = False
            MakeNewForm = False
            goSession.KillObject fso
            Exit Function
         End If
      Else
         ' Should error out here ?
         
      End If
   End If
   
   '
   ' Create Form Filename...
   '
   If Trim(NewFormID) = "" Then
      If Not CreateFormIdentity() Then
         MakeNewForm = False
         goSession.KillObject fso
         Exit Function
      End If
   Else
      If Not CreateFormIdentity(NewFormID) Then
         MakeNewForm = False
         goSession.KillObject fso
         Exit Function
      End If
   End If
   '
   ' Populate Form Properties...
   '
   moFormProperties.CreatedDate = Now()
   '8/2010 ms BGLNG-4 Populate RoleTypeKey instead of RoleType (deprecated years ago)
   moFormProperties.RoleType = CStr(goSession.User.RoleTypeKey)
   moFormProperties.SiteID = goSession.Site.SiteID
   moFormProperties.SiteKey = goSession.Site.SiteKey
   If Trim(moFormProperties.subject) = "" Then
      moFormProperties.subject = moTemplateProperties.Description
   End If
   moWorkflowProperties.WFStage = "1"
   '
   ' 30 March 2002 - Eliminate use of TargetFolder -
   '
   If Trim(TargetFolder) = "" Then
      ' Set default
      TargetFolder = goSession.Workflow.WorkingSessionFolder
   End If
   
   If Not fso.FolderExists(TargetFolder) Then
      goSession.RaiseError "Error in mwForm.MakeNewForm, Target Folder does not exist: " & TargetFolder
      mIsTemplateOpen = False
      MakeNewForm = False
      goSession.KillObject fso
      Exit Function
   
   End If
   '
   ' Copy Template to new file, create duplicate, or create new file...
   '
   If moTemplateProperties.FormType <> mwft_DATA_FORM Then
      '
      ' Copy from Template ?
      '
      ' target file
      moFormProperties.FullFilename = TargetFolder & "\" & moFormProperties.FileName
      If Trim(SourceFilename) = "" Then
         ' Copy MS Office Template...
         strTempFilename = TemplateSourceFName
      Else
         ' copy from existing form
         strTempFilename = SourceFilename
      End If
      '
      ' Copy file...
      '
      fso.CopyFile strTempFilename, moFormProperties.FullFilename
      
      '
      ' set attributes
      '
      fso.GetFile(moFormProperties.FullFilename).Attributes = Normal
      
   End If
   '
   ' Open Document
   '
   ' When opening form in Make mode, we need to do some different stuff...
   mIsNewForm = True
   Select Case moTemplateProperties.FormType
         
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         If moWordForm Is Nothing Then
            Set moWordForm = New mwWordForm
            Set moWordForm.Parent = Me
         End If
         If Not moWordForm.OpenForm() Then
            MakeNewForm = False
         Else
            If Not moWordForm.StuffNewFormProperties() Then
               MakeNewForm = False
            End If
         End If
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         If moExcelForm Is Nothing Then
            Set moExcelForm = New mwExcelForm
            Set moExcelForm.Parent = Me
         End If
         If Not moExcelForm.OpenForm() Then
            MakeNewForm = False
         Else
            If Not moExcelForm.StuffNewFormProperties() Then
               MakeNewForm = False
            Else
               If mIsPocketPcFormCompatible Then
                  ' Same copy of FormProperties to disk
                  MakePocketPCBackupProperties
               End If
            End If
         End If
      Case Is = mwft_DATA_FORM
         If moDataForm Is Nothing Then
            Set moDataForm = New mwDataForm
            Set moDataForm.Parent = Me
         End If
         ' Data Form has preliminary staging, instead of copying a template...
         ' It also "opens" the form in memory...
         moDataForm.MakeDataForm
         mIsOpen = True
   End Select
   '
   If Not mIsOpen Then
      MakeNewForm = False
      If fso.FileExists(moFormProperties.FullFilename) Then
         goSession.API.SleepApp 1000
         fso.DeleteFile moFormProperties.FullFilename
      End If
      goSession.KillObject fso
      Exit Function
   End If
   mIsNewForm = False
   MakeNewForm = True
   mIsOpen = True
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.MakeNewForm.", Err.Number, Err.Description
   If Not fso Is Nothing Then
      If fso.FileExists(moFormProperties.FullFilename) Then
         On Error Resume Next
         fso.DeleteFile moFormProperties.FullFilename
      End If
   End If
   goSession.KillObject fso
   MakeNewForm = False
   mIsNewForm = False
End Function

Public Function InstantiateForm(TemplateID As String, FormName As String, mwFormHeaderKey As Long, Optional IsEditInPlace As Boolean) As Boolean
'   Dim strNewFormName As String
   Dim strTempFilename As String
   Dim sTarget As String
   Dim fso As FileSystemObject
   Dim FormType As mw_FormType
   Dim nStatus As Integer
   
   On Error GoTo FunctionError
   
'   mIsOpen = False
   InstantiateForm = False
   nStatus = 1
   
   ' Only one form open per instance...
'   If mIsOpen Then
'      goSession.RaiseError "Error in mwForm.InstantiateForm, Form already open."
'      Exit Function
'   End If
   If TemplateID <> "" Then
      moTemplateProperties.TemplateID = TemplateID
   End If
   If Trim(moTemplateProperties.TemplateID) = "" Then
      goSession.RaiseError "Error in mwForm.InstantiateForm, No Template Specified."
      Exit Function
   Else
      TemplateID = moTemplateProperties.TemplateID
   End If
   '
   ' open the template
   '
   If Not mIsTemplateOpen Then
      If Not OpenTemplate(TemplateID) Then
         'goSession.RaiseError "Error in mwSession.mwForm.InstantiateForm, Template Not Found in Database:" _
           & TemplateID
         Exit Function
      End If
   End If
   If moTemplateProperties.FormType < 1 Then
      goSession.RaiseError "Error in mwForm.InstantiateForm, No FormType specified for: " & TemplateID
      Exit Function
   End If
   nStatus = 2
   
   '
   ' Copy from Template
   '
   ' target file
   Set fso = New FileSystemObject
   moFormProperties.FullFilename = FormName
   
   If IsEditInPlace = True Then
      nStatus = 3
      strTempFilename = GetTemplateFileName(mwFormHeaderKey)
      sTarget = FormName
      If fso.FileExists(sTarget) Then
         fso.GetFile(sTarget).Attributes = Normal
      Else
      
         fso.CopyFile strTempFilename, sTarget
      End If
   Else
      nStatus = 4
      strTempFilename = GetTemplateFileName(mwFormHeaderKey)
      sTarget = goSession.Workflow.WorkingSessionFolder & "\" & fso.GetFileName(FormName)
      '
      ' Copy file...
      '
      If fso.FileExists(sTarget) Then
         fso.GetFile(sTarget).Attributes = Normal
      End If
      
      fso.CopyFile strTempFilename, sTarget
   End If
   moFormProperties.FullFilename = sTarget
   
   '
   ' set attributes
   '
   nStatus = 5
   FormType = GetFormType(FormName)
   nStatus = 6
   Select Case FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         nStatus = 7
         moTemplateProperties.FormType = FormType
         If mIsOpen = False Then
            Set moWordForm = New mwWordForm
            Set moWordForm.Parent = Me
            mIsOpen = moWordForm.OpenForm(sTarget)
         End If
         If mIsOpen = True Then
            moWordForm.SetFormDetailData mwFormHeaderKey
         End If
         nStatus = 8
         
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         '
         nStatus = 9
         moTemplateProperties.FormType = FormType
         If mIsOpen = False Then
            Set moExcelForm = New mwExcelForm
            Set moExcelForm.Parent = Me
            mIsOpen = moExcelForm.OpenForm(sTarget)
         End If
         If mIsOpen = True Then
            moExcelForm.SetFormDetailData mwFormHeaderKey
         End If
         nStatus = 10
      
      Case Is = mwft_DATA_FORM
'         moTemplateProperties.FormType = mwft_DATA_FORM
'         Set moDataForm = New mwDataForm
'         Set moDataForm.Parent = Me
'         mIsOpen = moDataForm.OpenForm(sTarget)
         
   End Select
   nStatus = 11
   If IsEditInPlace = False Then
      ' Command fails under Novell Operating System
      On Error Resume Next
      fso.GetFile(sTarget).Attributes = ReadOnly
   End If
   nStatus = 12
   InstantiateForm = mIsOpen
   
   KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.InstantiateForm, Status= " & nStatus, Err.Number, Err.Description
   If IsEditInPlace = False And fso.FileExists(moFormProperties.FullFilename) Then
      On Error Resume Next
      fso.DeleteFile moFormProperties.FullFilename
   End If
   KillObject fso
   InstantiateForm = False
End Function

Public Function PreviewForm() As Boolean
   Dim fso As FileSystemObject
   Dim oFile As file
   Dim strWorkingFolder
   Dim strTargetFile As String
   On Error GoTo FunctionError
   PreviewForm = False
   '
   ' Template must be open to proceed
   '
   If Not mIsTemplateOpen Then
      goSession.RaiseError "Error in mwSession.mwForm.PreviewForm. Template must be opened first."
      PreviewForm = False
      Exit Function
   End If
   
   If moTemplateProperties.FormType = mwft_DATA_FORM Then
      goSession.RaiseError "Error in mwSession.mwForm.PreviewForm. Preview Mode Not Available for DataForms."
      PreviewForm = False
      Exit Function
   End If
   
   '
   ' Copy Template to WorkingFolder
   '
   strTargetFile = goSession.Workflow.WorkingSessionFolder & "\" & _
     moTemplateProperties.SourceFilename
   Set fso = New FileSystemObject
   '
   ' Delete if found on working folder
   '
   If fso.FileExists(strTargetFile) Then
      Set oFile = fso.GetFile(strTargetFile)
      oFile.Attributes = Hidden
      oFile.Delete
      Set oFile = Nothing
   End If
   '
   fso.CopyFile moTemplateProperties.TemplatePlace.GetSubFolderPath & "\" & _
     moTemplateProperties.SourceFilename, strTargetFile, True
   '
   ' set to read only
   '
   Set oFile = fso.GetFile(strTargetFile)
   oFile.Attributes = ReadOnly
   Set oFile = Nothing
   '
   ' Open Document
   '
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         If moWordForm.PreviewForm(strTargetFile) Then PreviewForm = True
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         If moExcelForm.PreviewForm(strTargetFile) Then PreviewForm = True
   End Select
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.", Err.Number, Err.Description
   PreviewForm = False
End Function

Public Function Send(Optional SendToAddress As Variant, Optional Direction As mw_WhichWorkflowDirection, _
  Optional TargetSiteID As String, Optional NoArchive As Boolean, Optional SubmitEmailAsDraft As Boolean, _
  Optional TargetSiteKey As Long) As Boolean
   Dim strSendTo As String
   Dim strTarget As String
   Dim nTransport As Long
   Dim IsEmailAddress As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwSession.mwForm.Send. no form is open to send."
      Send = False
      Exit Function
   End If
   goSession.MemoryLogIt "MwForm.Send  "
   strTarget = moFormProperties.FullFilename
   '
   ' Does form file exist ?
   '
   Set fso = New FileSystemObject
   If Not fso.FileExists(strTarget) Then
      goSession.RaiseError "Error in mwSession.mwForm.Send. Form file missing: " & strTarget
      Send = False
      goSession.KillObject fso
      Exit Function
   End If
   '
   ' Send the form...
   '
   If IsNumeric(goSession.Site.GetShipProperty(TargetSiteID, "DefaultTransport")) Then
      nTransport = CLng(goSession.Site.GetShipProperty(TargetSiteID, "DefaultTransport"))
   End If
   If nTransport < 1 Then
      nTransport = goSession.User.DefaultTransport
   End If
   If nTransport = mwt_DemoLoopback Then
      If goSession.Workflow.SendByDemoLoopback(strTarget) Then
         Send = True
      End If
   ElseIf nTransport = mwt_Media Then
      If GetExtendedProperty("IsChooseTargetOnSendByMedia") = "True" Then
         If goSession.Workflow.SendByMedia(strTarget, , , True, , , TargetSiteKey) Then
            Send = True
         End If
      Else
         If TargetSiteID <> "" Then
            If goSession.Workflow.SendByMedia(strTarget, , , False, TargetSiteID, , TargetSiteKey) Then
               Send = True
            End If
         Else
            If goSession.Workflow.SendByMedia(strTarget, , , , , , TargetSiteKey) Then
               Send = True
            End If
         End If
      End If
   ElseIf nTransport = mwt_TransportContainer Then
      '
      goSession.RaiseError "Sorry, Feature has been retired and is no longer supported."
      Send = False
   Else
      '
      ' Must be mail...
      '
      '
      ' 8/2005 - Must be switch to go by email...
      '
      If GetExtendedProperty("SubmitFormByEmail") <> "True" And _
         (moTemplateProperties.FormType = mwft_EXCEL Or moTemplateProperties.FormType = mwft_EXCEL_XLSX Or moTemplateProperties.FormType = mwft_EXCEL_XLSM Or _
          moTemplateProperties.FormType = mwft_WORD Or moTemplateProperties.FormType = mwft_WORD_DOCX Or moTemplateProperties.FormType = mwft_WORD_DOCM) Then
         Send = True
         Exit Function
      End If
      '
      ' Get email address from 3 possible places...
      '
      If IsMissing(SendToAddress) Then
         IsEmailAddress = False
      ElseIf Trim(SendToAddress) = "" Then
         IsEmailAddress = False
      Else
         IsEmailAddress = True
      End If
      If Not IsEmailAddress Then
         If Trim(moWorkflowProperties.RouteToEmailAddress) <> "" Then
            strSendTo = moWorkflowProperties.RouteToEmailAddress
         ElseIf Trim(moTemplateProperties.DefaultEmailAddress) = "" Then
            If Direction = mw_SendToShore Or goSession.Site.SiteType = SITE_TYPE_SHIP Then
               '
               ' Sending to the shore...
               '
               If Trim(goSession.Site.WorkflowSendToAddress) = "" Then
                  goSession.RaiseError "Error in mwSession.MwForm.Send, no Site.WorkflowSendToAddress found."
                  Send = False
                  Exit Function
               Else
                  strSendTo = goSession.Site.WorkflowSendToAddress
               End If
            ElseIf Direction = mw_SendToShip And goSession.Site.SiteType = SITE_TYPE_SHORE Then
               '
               ' Shore Sending Package to the shore...
               '
               If Trim(goSession.Site.WorkflowSendToAddress) = "" Then
                  goSession.RaiseError "Error in mwSession.MwForm.Send, no Site.WorkflowSendToAddress found."
                  Send = False
                  Exit Function
               Else
                  strSendTo = goSession.Site.WorkflowSendToAddress
               End If
            Else
               '
               ' Sending to the Ship...
               '
               If TargetSiteID = "" Then
                  goSession.RaiseError "Error in mwSession.MwForm.Send, no TargetSiteID found for Shore to Ship Transmit."
                  Send = False
                  Exit Function
               Else
               
                  If Trim(goSession.Site.GetShipWorkflowSendToAddress(TargetSiteID)) = "" Then
                     goSession.RaiseError "Error in mwSession.MwForm.Send, no Site.ShipWorkflowSendToAddress found for TargetSiteID = " & TargetSiteID
                     Send = False
                     Exit Function
                  Else
                     strSendTo = goSession.Site.GetShipWorkflowSendToAddress(TargetSiteID)
                  End If
               End If
            End If
         Else
            strSendTo = moTemplateProperties.DefaultEmailAddress
         End If
      Else
         strSendTo = SendToAddress
      End If
      '
      ' Data Form persistence check...
      '
      If (moTemplateProperties.FormType = mwft_DATA_FORM) And moDataForm Is Nothing Then
         goSession.RaiseError "Error in mwSession.MwForm.Send, Datagram has not been written, so can't send it."
         Send = False
         Exit Function
      End If
      '
      ' Send the form...
      '
      If mTextBodyOverride = "" Then
         mTextBodyOverride = "Workflow Form: " & moTemplateProperties.TemplateID & ", Filename: " & fso.GetFileName(strTarget)
      End If
      
      Send = goSession.Workflow.SendByEmail(strTarget, mTextBodyOverride, strSendTo, , TargetSiteKey)
      
   End If
   '
   ' Follow-up
   '
   If Send And Not NoArchive Then
      goSession.Workflow.ArchiveWorkflowItem FormProperties.FullFilename, ArchivedOK, moTemplateProperties.DeleteAfterTransmit
   End If
   goSession.KillObject fso
   goSession.MemoryLogIt "MwForm.Send exit "
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.MwForm.Send.", Err.Number, Err.Description
   Send = False
   goSession.KillObject fso
   goSession.MemoryLogIt "MwForm.Send err exit "
End Function

Public Function SendMultipleSites(SiteIDs() As Variant, WhichSiteID As mw_WhichSiteCode, _
  Optional Direction As mw_WhichWorkflowDirection) As Boolean
   Dim i As Integer
   Dim IsSend As Boolean
   Dim oRs As Recordset
   Dim strSite As String
   Dim mwcSitesKey As Long
   
   Dim CheckedAllSites As Boolean
   On Error GoTo FunctionError
   If UBound(SiteIDs) = 0 Then
      SendMultipleSites = False
   End If
   Exit Function
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   oRs.Open "mwcSites", goSession.DBConnection, adOpenStatic, adLockReadOnly
   '
   'Loop though array of SiteID's...
   '
   For i = 0 To UBound(SiteIDs) - 1
      oRs.MoveFirst
      CheckedAllSites = False
      IsSend = False
      '
      'Loop though all site records...
      '
      Do While Not CheckedAllSites
         strSite = UCase(Trim(SiteIDs(i)))
         Select Case WhichSiteID
            Case Is = mw_ShortSiteID
               If strSite = oRs!ShortSiteID Then
                  mwcSitesKey = oRs!ID
                  IsSend = True
               End If
            Case Is = mw_SiteID
               If strSite = oRs!SiteID Then
                  mwcSitesKey = oRs!ID
                  IsSend = True
               End If
            Case Is = mw_YetAnotherSiteID
               If strSite = oRs!YetAnotherSiteID Then
                  mwcSitesKey = oRs!ID
                  IsSend = True
               End If
         End Select
         If IsSend Then
            If Direction = mw_SendToShip Then
               If Not Send(oRs!WorkflowSendToAddress, , , , , mwcSitesKey) Then
                  goSession.RaiseWorkflowStatus "Error in mwForm.SendMultipleSites WorkflowSendToAddress not found for Site: " & oRs!SiteName & _
                    "Email: " & oRs!WorkflowSendToAddress
                  SendMultipleSites = False
                  goSession.CloseRecordset oRs
                  Exit Function
               End If
            Else
               If Not Send(oRs!ShipWorkflowSendToAddress, , , , , mwcSitesKey) Then
                  goSession.RaiseWorkflowStatus "Error in mwForm.SendMultipleSites ShipWorkflowSendToAddress not found for Site: " & oRs!SiteName & _
                    "Email: " & oRs!ShipWorkflowSendToAddress
                  SendMultipleSites = False
                  goSession.CloseRecordset oRs
                  Exit Function
               End If
            End If
            CheckedAllSites = True
         Else
            oRs.MoveNext
            If oRs.EOF Then CheckedAllSites = True
         End If
      Loop
   Next i
   goSession.CloseRecordset oRs
   SendMultipleSites = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwForm.SendMultipleSites", Err.Number, Err.Description
   goSession.CloseRecordset oRs
   SendMultipleSites = False
End Function


Public Function TearDown(Optional LeaveAppOpen As Boolean)
   If Not moExcelForm Is Nothing Then
      moExcelForm.TearDown LeaveAppOpen
      Set moExcelForm = Nothing
   End If
   If Not moDataForm Is Nothing Then
      moDataForm.TearDown
      Set moDataForm = Nothing
   End If
   If Not moWordForm Is Nothing Then
      moWordForm.TearDown LeaveAppOpen
      Set moWordForm = Nothing
   End If
   If Not moParent Is Nothing Then Set moParent = Nothing
   If Not moFormProperties Is Nothing Then Set moFormProperties = Nothing
   If Not moTemplateProperties Is Nothing Then Set moTemplateProperties = Nothing
   If Not moWorkflowProperties Is Nothing Then Set moWorkflowProperties = Nothing
End Function


Public Function GetSubject() As Boolean
   Dim loDescription As frmOfficeDescription
   On Error GoTo FunctionError
   Set loDescription = New frmOfficeDescription
   Load loDescription
   If Trim(moFormProperties.subject) = "" Then
      loDescription.txtSubject.Text = moTemplateProperties.Description
   Else
      loDescription.txtSubject.Text = moFormProperties.subject
   End If
   loDescription.Show vbModal
   If Trim(loDescription.txtSubject.Text) = "" Then
      moFormProperties.subject = moTemplateProperties.Description
   Else
      moFormProperties.subject = loDescription.txtSubject.Text
   End If
   Unload loDescription
   Set loDescription = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwForm.GetSubject", Err.Number, Err.Description
End Function

Public Function GetExtendedProperty(KeyName As String) As String
   On Error GoTo FunctionError
   'If Not mIsOpen Then
      'GetExtendedProperty = ""
      'Exit Function
   'End If
   GetExtendedProperty = moExtendedKeys(UCase(KeyName)).KeyValue
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(KeyName As String, KeyValue As String, Optional TemplateID As String) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   If Not mIsOpen Then
      SetExtendedProperty = False
      Exit Function
   End If
'   If App.Revision > 597 Then
      If Trim(TemplateID) = "" Then
         strSQL = "select * from mwwfFormTemplate where TemplateID='" & moTemplateProperties.TemplateID & "'"
      Else
         strSQL = "select * from mwwfFormTemplate where TemplateID='" & TemplateID & "'"
      End If
'   Else
'      If Trim(TemplateID) = "" Then
'         strSQL = "select * from mwwfFormTemplates where TemplateID='" & moTemplateProperties.TemplateID & "'"
'      Else
'         strSQL = "select * from mwwfFormTemplates where TemplateID='" & TemplateID & "'"
'      End If
'   End If
   If goSession.UpdateExtendedKeyData(KeyName, KeyValue, strSQL) Then
      moExtendedKeys(UCase(KeyName)).KeyValue = KeyValue
   End If
   SetExtendedProperty = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.SetExtendedProperty. Invalid Key ?: " & KeyName, Err.Number, Err.Description
   SetExtendedProperty = False
End Function


Public Function MakePocketPCBackupProperties() As Boolean
   Dim strFile As String
   Dim fso As FileSystemObject
   Dim loCGS As ConfigGroups
   Dim loCG As ConfigGroup
   
   On Error GoTo FunctionError
   If Not mIsOpen Then
      MakePocketPCBackupProperties = False
      Exit Function
   End If
   Set fso = New FileSystemObject
   strFile = Left(moFormProperties.FullFilename, Len(moFormProperties.FullFilename) - 4) & ".pbak"
   If fso.FileExists(strFile) Then
      fso.DeleteFile (strFile)
   End If
   Set loCGS = New ConfigGroups
   Set loCG = loCGS.Add("POCKET_PC_PROPERTIES")
   If moTemplateProperties.FormType = mwft_EXCEL Or moTemplateProperties.FormType = mwft_EXCEL_XLSX Or moTemplateProperties.FormType = mwft_EXCEL_XLSM Then
      '
      ' Excel Forms
      '
      moExcelForm.StuffPocketPcProperties loCG
      
   ElseIf moTemplateProperties.FormType = mwft_WORD Or moTemplateProperties.FormType = mwft_WORD_DOCX Or moTemplateProperties.FormType = mwft_WORD_DOCM Then
      '
      ' Word Forms
      '
      moWordForm.StuffPocketPcProperties loCG
   Else
      MakePocketPCBackupProperties = False
      goSession.KillObject fso
      Exit Function
   End If
   If goSession.WriteConfigGroupsFile(loCGS, strFile) Then
      MakePocketPCBackupProperties = True
   Else
      MakePocketPCBackupProperties = False
   End If
   goSession.KillObject fso
   MakePocketPCBackupProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.MakePocketPCBackupProperties. ", Err.Number, Err.Description
   MakePocketPCBackupProperties = False
End Function

Public Function RestorePocketPCBackupProperties() As Boolean
   Dim strFile As String
   Dim fso As FileSystemObject
   Dim loCGS As ConfigGroups
   Dim loCG As ConfigGroup
   
   On Error GoTo FunctionError
   'If Not mIsOpen Then
   '   RestorePocketPCBackupProperties = False
   '   Exit Function
   'End If
   Set fso = New FileSystemObject
   strFile = Left(moFormProperties.FullFilename, Len(moFormProperties.FullFilename) - 4) & ".pbak"
   If Not fso.FileExists(strFile) Then
      goSession.KillObject fso
      RestorePocketPCBackupProperties = False
      Exit Function
   End If
   Set loCGS = goSession.LoadConfigGroupsFile(strFile)
   Set loCG = loCGS.Item(1)
   If moTemplateProperties.FormType = mwft_EXCEL Or moTemplateProperties.FormType = mwft_EXCEL_XLSX Or moTemplateProperties.FormType = mwft_EXCEL_XLSM Then
      '
      ' Excel Form Properties
      '
      moExcelForm.FetchPocketPcProperties loCG
      RestorePocketPCBackupProperties = True
   ElseIf moTemplateProperties.FormType = mwft_WORD Or moTemplateProperties.FormType = mwft_WORD_DOCX Or moTemplateProperties.FormType = mwft_WORD_DOCM Then
      '
      ' Word Form properties
      '
      moWordForm.FetchPocketPcProperties loCG
      RestorePocketPCBackupProperties = True
   Else
      RestorePocketPCBackupProperties = False
      RestorePocketPCBackupProperties = False
   End If
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.RestorePocketPCBackupProperties. ", Err.Number, Err.Description
   RestorePocketPCBackupProperties = False
End Function

Public Function DiscardPocketPCBackupProperties() As Boolean
   Dim strFile As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   strFile = Left(moFormProperties.FullFilename, Len(moFormProperties.FullFilename) - 4) & ".pbak"
   If fso.FileExists(strFile) Then
      fso.DeleteFile (strFile)
   End If
   Set fso = Nothing
   DiscardPocketPCBackupProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.DiscardPocketPCBackupProperties. ", Err.Number, Err.Description
   DiscardPocketPCBackupProperties = False

End Function

Public Function ReplaceXmlCharacters(strValue As String) As String
   Dim strTemp As String
   Static IsLoaded As Boolean
   On Error GoTo FunctionError
   If IsLoaded Then
      moRsXmlConversions.MoveFirst
   Else
      Set moRsXmlConversions = New Recordset
      moRsXmlConversions.CursorLocation = adUseClient
      moRsXmlConversions.Open "mwcXmlConvertChars", goSession.DBConnection, adOpenForwardOnly, adLockReadOnly, adCmdTable
      IsLoaded = True
   End If
   strTemp = strValue
   Do While Not moRsXmlConversions.EOF
      strTemp = Replace(strTemp, Chr(moRsXmlConversions!AsciiValue), moRsXmlConversions!ConvertString)
      moRsXmlConversions.MoveNext
   Loop
   ReplaceXmlCharacters = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.ReplaceXmlCharacters. ", Err.Number, Err.Description
   ReplaceXmlCharacters = strValue
End Function

Public Function ChangeToCustomFormID(NewFormID As String) As Boolean
   Dim fso As FileSystemObject
   Dim strNewFile As String
   Dim strOldFile As String
   On Error GoTo FunctionError
   If Not mIsOpen = True Then
      ChangeToCustomFormID = False
      goSession.RaiseError "Error in mwForm.ChangeToCustomFormID, form must be open before changing Form ID."
      Exit Function
   End If
   '
   ' Only Word and Excel Forms
   '
      Select Case moTemplateProperties.FormType
         Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
            ' Only Word and Excel Forms
         Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
            ' Only Word and Excel Forms
         Case Else
            ChangeToCustomFormID = False
            goSession.RaiseError "Error in mwForm.ChangeToCustomFormID, only Excel and Word forms can be reassigned (now)."
            Exit Function
      End Select
   
'   If moTemplateProperties.FormType <> mwft_EXCEL And moTemplateProperties.FormType <> mwft_WORD Then
'      ChangeToCustomFormID = False
'      goSession.RaiseError "Error in mwForm.ChangeToCustomFormID, only Excel and Word forms can be reassigned (now)."
'      Exit Function
'   End If
   '
   ' Does new file already exist ?
   '
   Set fso = New FileSystemObject
   strOldFile = moFormProperties.FullFilename
   strNewFile = fso.GetParentFolderName(strOldFile) & "\" & NewFormID & "." & fso.GetExtensionName(strOldFile)
   If fso.FileExists(strNewFile) Then
      ChangeToCustomFormID = False
      goSession.RaiseError "Error in mwForm.ChangeToCustomFormID, File already exists: " & strNewFile
      Set fso = Nothing
   End If
   '
   ' OK, do the job...
   '
   ' a. Change properties in Object (FormID, Filename, and FullFilename)
   ' b. Change properties in Excel/Word file
   ' c. Close file, rename it, open it again
   moFormProperties.FormID = UCase(NewFormID)
   moFormProperties.FileName = NewFormID & "." & fso.GetExtensionName(strOldFile)
   moFormProperties.FullFilename = strNewFile
   
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         moWordForm.StuffNewFormProperties
         moWordForm.CloseWordForm
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         moExcelForm.StuffNewFormProperties
         moExcelForm.CloseExcelForm
   End Select
   
'   If moTemplateProperties.FormType = mwft_EXCEL Then
'      moExcelForm.StuffNewFormProperties
'      moExcelForm.CloseExcelForm
'   Else
'      moWordForm.StuffNewFormProperties
'      moWordForm.CloseWordForm
'   End If
   '
   ' Rename physical form file
   '
   fso.MoveFile strOldFile, strNewFile
   '
   ' Reopen the form...
   '
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         moWordForm.OpenForm
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         moExcelForm.OpenForm
   End Select
   
'   If moTemplateProperties.FormType = mwft_EXCEL Then
'      moExcelForm.OpenForm
'   Else
'      moWordForm.OpenForm
'   End If
   ChangeToCustomFormID = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.ReplaceXmlCharacters. ", Err.Number, Err.Description
   ChangeToCustomFormID = False
End Function

Public Function PopulateForm(Optional Worksheet As Integer, Optional IsRepopulate As Boolean) As Boolean
   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwForm.PopulateForm, form must be open."
      PopulateForm = False
      Exit Function
   End If
   
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         PopulateForm = moWordForm.PopulateForm(IsRepopulate)
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         PopulateForm = moExcelForm.PopulateForm(Worksheet, IsRepopulate)
      
      Case Else
         PopulateForm = True
      
   End Select
   
'   If moTemplateProperties.FormType = mwft_EXCEL Then
'      PopulateForm = moExcelForm.PopulateForm(Worksheet, IsRepopulate)
'   ElseIf moTemplateProperties.FormType = mwft_WORD Then
'      PopulateForm = moWordForm.PopulateForm(IsRepopulate)
'   Else
'      PopulateForm = True
'   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.PopulateForm. ", Err.Number, Err.Description
   PopulateForm = False
End Function

'
' Function supports "overload", either pass RecordeKeyValue to look up, or recordset reference - but not both
'
Public Function PopulateEventForm(EventType As Long, Optional RecordKeyValue As Long, Optional ByRef loDataRefRS As Recordset) As Boolean
   Dim loCon As Connection
   Dim loDW As mwDataWork
   Dim loRs As Recordset
   Dim loRSDataRecord As Recordset
   Dim loRsColumns As Recordset
   Dim strSQL As String

   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwForm.PopulateEventForm, form must be open."
      PopulateEventForm = False
      Exit Function
   End If
   
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         ' ok
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         ' ok
      Case Else
         goSession.RaiseError "Error in mwForm.PopulateEventForm, Invalid Form Type - Function only available with Excel or Word forms."
         PopulateEventForm = False
         Exit Function
      
   End Select
   
'   If moTemplateProperties.FormType <> mwft_EXCEL And moTemplateProperties.FormType <> mwft_WORD Then
'      goSession.RaiseError "Error in mwForm.PopulateEventForm, Invalid Form Type - Function only available with Excel or Word forms."
'      PopulateEventForm = False
'      Exit Function
'   End If

   ' Verify not over-overloaded...
   If RecordKeyValue > 0 And (Not loDataRefRS Is Nothing) Then
      goSession.RaiseError "Error in mwForm.PopulateEventForm, Invalid Overload - either RecordKeyValue can be filled, or a pointer to a recordset passed - but not both."
      PopulateEventForm = False
      Exit Function
   End If
   '
   ' If recordset - must be a record in there... (beware of server side cursors - not supported here)
   '
   If Not loDataRefRS Is Nothing Then
      If loDataRefRS.RecordCount < 1 Then
         goSession.RaiseError "Error in mwForm.PopulateEventForm, passed in record set has no records in the recordset."
         PopulateEventForm = False
         Exit Function
      End If
      '
      Set moRsEventDetail = loDataRefRS
   End If
   '
   ' Fetch "populate table" record...
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwwfFormPopulateTable where mwEventTypeKey=" & EventType & _
     "and TemplateID='" & TemplateProperties.TemplateID & "'"
   loRs.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.CloseRecordset loRs
      PopulateEventForm = True
      Exit Function
   End If
   '
   ' Either fetch the "data record" or use the pointer to the recordset...
   '
   If RecordKeyValue > 0 Then
      Set loDW = New mwDataWork
      Set loCon = loDW.OpenConnection(loRs!ConnectID)
      Set loRSDataRecord = New Recordset
      loRSDataRecord.CursorLocation = adUseClient
      ' Fetch record (if it is there)
      strSQL = "select * from " & loRs!TableName & " where ID=" & RecordKeyValue
      loRSDataRecord.Open strSQL, loCon, adOpenForwardOnly, adLockReadOnly
      If loRSDataRecord.RecordCount < 1 Then
         ' No data record found...
         
         goSession.CloseRecordset loRs
         goSession.CloseRecordset loRSDataRecord
         goSession.CloseConnection loCon
         goSession.KillObject loDW
         PopulateEventForm = True
         Exit Function
      End If
   Else
      Set loRSDataRecord = loDataRefRS
   End If
   '
   ' Loop through records
   '
   Do While Not loRs.EOF
      '
      ' Next up, find the columns to rifle through...
      '
      Set loRsColumns = New Recordset
      loRsColumns.CursorLocation = adUseClient
      strSQL = "select * from mwwfFormPopulateColumn where fptkey = " & loRs!ID
      loRsColumns.Open strSQL, goSession.DBConnection, adOpenForwardOnly, adLockReadOnly
      If loRsColumns.RecordCount > 0 Then
         '
         ' Pass control to Word or Excel to loop through the columns...
         '
         Select Case moTemplateProperties.FormType
            Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
               moWordForm.PopulateEventForm loRsColumns, loRSDataRecord
               
            Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
               moExcelForm.PopulateEventForm loRsColumns, loRSDataRecord
               
         End Select
         
'         If moTemplateProperties.FormType = mwft_EXCEL Then
'            moExcelForm.PopulateEventForm loRsColumns, loRSDataRecord
'               '
'         ElseIf moTemplateProperties.FormType = mwft_WORD Then
'            moWordForm.PopulateEventForm loRsColumns, loRSDataRecord
'         End If
      End If
      goSession.CloseRecordset loRsColumns
      loRs.MoveNext
   Loop
   '
   ' Done, close up shop
   '
   'goSession.CloseRecordset loRs
   Set loRSDataRecord = Nothing
   goSession.CloseConnection loCon
   goSession.KillObject loDW
   PopulateEventForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.PopulateEventForm. ", Err.Number, Err.Description
   goSession.CloseRecordset loRsColumns
   goSession.CloseRecordset loRs
   goSession.CloseRecordset loRSDataRecord
   goSession.CloseConnection loCon
   goSession.KillObject loDW
   PopulateEventForm = False
End Function




Public Function PrintForm() As Boolean
   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwForm.PrintForm, form must be open."
      PrintForm = False
      Exit Function
   End If
   
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         PrintForm = moWordForm.PrintForm()
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         PrintForm = moExcelForm.PrintForm()
      Case Else
         PrintForm = True
   End Select
   
'   If moTemplateProperties.FormType = mwft_EXCEL Then
'      PrintForm = moExcelForm.PrintForm()
'   ElseIf moTemplateProperties.FormType = mwft_WORD Then
'      PrintForm = moWordForm.PrintForm()
'   Else
'      PrintForm = True
'   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.PrintForm. ", Err.Number, Err.Description
   PrintForm = False

End Function

Public Function PreviewForm2() As Boolean
   Dim fso As FileSystemObject
   Dim oFile As file
   Dim strTargetFile As String
   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwForm.PreviewForm2, form must be open."
      PreviewForm2 = False
      Exit Function
   End If
   
   strTargetFile = goSession.Workflow.WorkingSessionFolder & "\" & moFormProperties.FileName & Format(Now(), "yymmddhhmmss")
   Set fso = New FileSystemObject
     
   fso.CopyFile moFormProperties.FullFilename, strTargetFile
   Set oFile = fso.GetFile(strTargetFile)
   oFile.Attributes = ReadOnly
   Set oFile = Nothing
   
   Select Case moTemplateProperties.FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         PreviewForm2 = moWordForm.PreviewForm2(strTargetFile)
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         PreviewForm2 = moExcelForm.PreviewForm2(strTargetFile)
      Case Else
         MsgBox "Preview is not supported for this Form Type " & moTemplateProperties.FormType & " ", vbInformation, "Preview Form"
   End Select
   
   PreviewForm2 = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.PreviewForm2. ", Err.Number, Err.Description
   PreviewForm2 = False
End Function

Public Function RunFormDll(WhichEvent As String) As Boolean
   Dim strVbDLL As String
   Dim loFormSession As Object
   
   On Error GoTo FunctionError
   If Not IsOpen Then
      goSession.RaiseError "Error in mwForm.RunFormDll, form not open. ", Err.Number, Err.Description
      RunFormDll = False
      Exit Function
   End If
   '
   ' Run external VB DLL - if it exists...
   '
   strVbDLL = moWorkflowProperties.VbFormEventDLL
   If Trim(strVbDLL) <> "" Then
      goSession.LogIt mwl_Workstation, mwl_Information, "RunFormDLL underway for : " & WhichEvent
      strVbDLL = strVbDLL & ".FormSession"
      On Error Resume Next
      Set loFormSession = CreateObject(strVbDLL)
      If Err Then
         If Err.Number = 429 Or Err.Number = 13 Then
            goSession.RaiseError "Error in mwForm.RunFormDLL, Unable to launch DLL: " & strVbDLL
         Else
            goSession.RaiseError "General Error in mwForm.RunFormDLL, launching DLL: " & strVbDLL, _
              Err.Number, Err.Description
         End If
         RunFormDll = False
         Exit Function
      End If
      On Error GoTo FunctionError
      Set loFormSession.ActiveMwSession = moParent
      Set loFormSession.ActiveMwForm = Me
      loFormSession.TriggerFormEvent (WhichEvent)
      If loFormSession.IsCancel Then
         '
         ' Cancel the process
         '
         goSession.KillObject loFormSession
         RunFormDll = False
         Exit Function
      End If
      goSession.KillObject loFormSession
   End If
   RunFormDll = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.RunFormDll. ", Err.Number, Err.Description
   RunFormDll = False
End Function

Public Function SubmitForm(Optional SubmitEmailAsDraft As Boolean) As Boolean
   Dim lIsCancelled As Boolean
   Dim loDatagramForm As mwForm
   Dim IsSent As Boolean
   
   On Error GoTo FunctionError
   If Not IsOpen Then
      goSession.RaiseError "Error in mwForm.RunFormDll, form not open. ", Err.Number, Err.Description
      SubmitForm = False
      Exit Function
   End If
   goSession.LogIt mwl_Workstation, mwl_Information, "Submit form underway for Form: " & _
     FormProperties.FormID
   '
   ' Run external VB DLL - if it exists... OnSubmit form event...
   '
   If Not RunFormDll("OnSubmit") Then
      goSession.LogIt mwl_User_Defined, mwl_Information, "Submit has been cancelled by the OnSubmit DLL Code."
      SubmitForm = False
      Exit Function
   End If
   '
   ' Process Local On Submit ?
   '
   If GetExtendedProperty("ProcessLocalOnSubmit") = "True" Then
      If FormType = mwft_EXCEL Or FormType = mwft_EXCEL_XLSX Or FormType = mwft_EXCEL_XLSM Then
         If Not ExcelForm.UpdateAsRecordset() Then
            goSession.RaiseError "Error Occurred Updating locally. Form has not been submitted."
            SubmitForm = False
            Exit Function
         End If
      End If
   End If
   '
   If TemplateProperties.ConvertToDatagram = True And mFormHeadersKey < 1 Then
      '
      ' Convert To Datagram
      '
      Set loDatagramForm = ExcelForm.ConvertToDatagram()
      If loDatagramForm Is Nothing Then
         goSession.RaiseError "Error in mwForm.SubmitForm, unable to convert to Datagram. Spreadsheet has not been submitted."
         SubmitForm = False
         Exit Function
      Else
         If loDatagramForm.Send() Then
            IsSent = True
         End If
      End If
   ElseIf mFormHeadersKey > 0 Then
      '
      ' New Form cabinet submit
      '
      'GetFormDetailData mFormHeadersKey,
      IsSent = True
      
      
   Else
      If moTemplateProperties.IsShipboardOnly Then
         IsSent = True
      Else
         If Me.GetExtendedProperty("submitemailaddress") <> "" Then
            If Send(Me.GetExtendedProperty("submitemailaddress")) Then
               IsSent = True
            End If
         Else
            If Send(, , , , SubmitEmailAsDraft) Then
               IsSent = True
            End If
         End If
      End If
   End If
   
   If IsSent Then
      RunFormDll "AfterSubmit"
      SubmitForm = True
   Else
      SubmitForm = False
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.SubmitForm. ", Err.Number, Err.Description
   SubmitForm = False
   
End Function

Public Function EditForm() As Boolean

   On Error GoTo FunctionError
   '
   ' Enter Edit the form...
   '
   If Not IsOpen Then
      goSession.RaiseError "Error in mwForm.EditForm, Form is not open. "
      EditForm = False
      Exit Function
   End If
   
   Select Case FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         WordForm.EditForm
         EditForm = True
         
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         ExcelForm.EditForm
         EditForm = True
         
      Case Else
         goSession.RaiseError "Error in mwForm.EditForm, only Word and Excel form types may be edited. ", Err.Number, Err.Description
         EditForm = False
      
   End Select

   
'   If FormType = mwft_EXCEL Then
'      ExcelForm.EditForm
'      EditForm = True
'   ElseIf FormType = mwft_WORD Then
'      WordForm.EditForm
'      EditForm = True
'   Else
'      goSession.RaiseError "Error in mwForm.EditForm, only Word and Excel form types may be edited. ", err.Number, err.Description
'      EditForm = False
'
'   End If

   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.EditForm. ", Err.Number, Err.Description
   EditForm = False
End Function


Friend Function InsertLogo(CellBookmarkValue As String) As Boolean
   Dim nKey As Long
   Dim loRs As Recordset
   Dim strSQL As String
   Dim strFilename As String
   On Error GoTo FunctionError
   nKey = goSession.Site.GetExtendedProperty("mwcCommercialOperatorKey")
   If nKey < 1 Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logged error in mwSession.mwExcelForm.InsertLogo, mwcCommercialOperatorKey value missing."
      InsertLogo = False
      Exit Function
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwcCommercialOperator where ID=" & nKey
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logged error in mwSession.mwExcelForm.InsertLogo, mwcCommercialOperator record missing."
      CloseRecordset loRs
      InsertLogo = False
      Exit Function
   End If
   '
   ' Load picture into spreadsheet.
   '
   If IsNull(loRs!LogoFilename) Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logged error in mwSession.mwExcelForm.InsertLogo, mwcCommercialOperator.LogoFilename is null."
      CloseRecordset loRs
      InsertLogo = False
      Exit Function
   End If
   strFilename = goSession.GetAppPath() & "\icons\32x32\" & loRs!LogoFilename
   
   Select Case FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         moWordForm.LoadPicture strFilename, CellBookmarkValue
         
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
         moExcelForm.LoadPicture strFilename, CellBookmarkValue
         
   End Select

   
'   If FormType = mwft_EXCEL Then
'      moExcelForm.LoadPicture strFilename, CellBookmarkValue
'   ElseIf FormType = mwft_WORD Then
'      moWordForm.LoadPicture strFilename, CellBookmarkValue
'   End If

   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwExcelForm.InsertLogo", Err.Number, Err.Description
End Function

Public Function SendRouteToEmail() As Boolean
   Dim sRouteTo As String
   Dim sTarget As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   ' Form Open ?
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwSession.mwForm.SendRouteToEmail. no form is open to send."
      SendRouteToEmail = False
      Exit Function
   End If
   sTarget = moFormProperties.FullFilename
   '
   ' Does form file exist ?
   '
   Set fso = New FileSystemObject
   If Not fso.FileExists(sTarget) Then
      goSession.RaiseError "Error in mwSession.mwForm.SendRouteToEmail. Form file missing: " & sTarget
      SendRouteToEmail = False
      goSession.KillObject fso
      Exit Function
   End If
   '
   If Trim(moWorkflowProperties.RouteToEmailAddress) = "" Then
      goSession.RaiseError "Error in mwSession.mwForm.SendRouteToEmail, no Route To Email Address found for this form." & sTarget
      SendRouteToEmail = False
      goSession.KillObject fso
      Exit Function
   End If
   '
   ' Data Form exists check...
   '
   If (moTemplateProperties.FormType = mwft_DATA_FORM) And moDataForm Is Nothing Then
      goSession.RaiseError "Error in mwSession.mwForm.SendRouteToEmail, Datagram has not been written, so can't send it."
      SendRouteToEmail = False
      Exit Function
   End If
   '
   ' Send the form...
   '
   With goSession.Message
      .AttachmentFile = sTarget
      .Server = goSession.User.MailServerName
      .SendFromAddress = goSession.User.MailUserID
      .SendToAddress = moWorkflowProperties.RouteToEmailAddress
      If Trim(mTextBodyOverride) = "" Then
         .Body = "<The attached form has been routed to you by the ShipNet Fleet Workflow Agent>"
      Else
         .Body = mTextBodyOverride
      End If
      .docID = moFormProperties.FormID
      .AttachmentOptions = mw_LEAVE_SOURCE
      .subject = "MWS Form from: " & goSession.Site.SiteName & ", Message Type: " & moTemplateProperties.TemplateID
      '
      If .Send() Then
         SendRouteToEmail = True
         goSession.RaiseWorkflowStatus "Incoming Form Rerouted, " & moFormProperties.FileName & ", to email address: " & moWorkflowProperties.RouteToEmailAddress
      Else
         SendRouteToEmail = False
         goSession.RaiseWorkflowStatus "Error occurred rerouting form, " & moFormProperties.FormID & ", to email address: " & moWorkflowProperties.RouteToEmailAddress
      End If
   End With
   '
   ' Follow-up
   '
   'If Send Then
   '   goSession.Workflow.ArchiveWorkflowItem FormProperties.FullFilename, ArchivedOK, moTemplateProperties.DeleteAfterTransmit
   'End If
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.SendRouteToEmail.", Err.Number, Err.Description
   SendRouteToEmail = False
   goSession.KillObject fso
End Function


Public Function GetDataConnection() As Connection
   Dim loDataWork As mwDataWork
   Dim sGet_SqlServerDATEFORMAT As String
   Dim loCon As Connection
   On Error GoTo FunctionError
   
   If Trim(moWorkflowProperties.ConnectionID) <> "" Then
      '
      ' Connection is in mwcDBConnection...
      '
      Set loDataWork = New mwDataWork
      loDataWork.LeaveOpenOnDestroy = True
      Set loCon = loDataWork.OpenConnection(moWorkflowProperties.ConnectionID)
      If loCon Is Nothing Then
         Exit Function
      End If
      Set moExternalCon = loCon
      Set GetDataConnection = moExternalCon
      '
      ' Set date format to specialized setting for date handling - using SET DATEFORMAT comand
      '
      sGet_SqlServerDATEFORMAT = loDataWork.Get_SqlServerDATEFORMAT(moWorkflowProperties.ConnectionID)
      If sGet_SqlServerDATEFORMAT <> "" Then
         loCon.Execute "SET DATEFORMAT " & sGet_SqlServerDATEFORMAT
      End If
      KillObject loDataWork
      mIsExternalDbConnection = True
   Else
      ' Default is use main connection
      Set GetDataConnection = goCon
      mIsExternalDbConnection = False
   End If
   
   'Set loCon = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwForm.GetDataConnection. ", Err.Number, Err.Description
End Function

Private Function GetFormHeadersStatus(FormHeadersKey) As Long
   Dim loRs As Recordset
   Dim nKey As Long
   Dim sSQL As String
   On Error GoTo FunctionError
     
   Set loRs = New Recordset
   
   sSQL = "SELECT mwFormStatusKey, mwwfFormTemplateKey FROM mwFormHeader WHERE ID= " & FormHeadersKey
   
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If IsNull(loRs!mwFormStatusKey) Then
         GetFormHeadersStatus = FORM_HEADER_STATUS_NULL
      Else
         GetFormHeadersStatus = loRs!mwFormStatusKey
      End If
      nKey = ZeroNull(loRs!mwwfFormTemplateKey)
   Else
         GetFormHeadersStatus = FORM_HEADER_STATUS_NULL
   End If
   loRs.Close
   sSQL = "select TemplateID from mwwfFormTemplate where ID=" & nKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If IsRecordLoaded(loRs) Then
      moTemplateProperties.TemplateID = BlankNull(loRs!TemplateID)
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwForm.GetFormHeadersStatus ", Err.Number, Err.Description
   CloseRecordset loRs
End Function
Private Function GetFormType(Optional FullFilename As String) As mw_FormType
   
   Dim fso As FileSystemObject
   Dim strExtension As String
   On Error GoTo FormError
   
   Set fso = New FileSystemObject
   strExtension = UCase(fso.GetExtensionName(FullFilename))
   If strExtension = "DOC" Or strExtension = "DOCX" Or strExtension = "DOCM" Then
      GetFormType = mwft_WORD
   ElseIf strExtension = "DOCX" Then
      GetFormType = mwft_WORD_DOCX
   ElseIf strExtension = "DOCM" Then
      GetFormType = mwft_WORD_DOCM
      
   ElseIf strExtension = "XLS" Then
      GetFormType = mwft_EXCEL
   ElseIf strExtension = "XLSX" Then
      GetFormType = mwft_EXCEL_XLSX
   ElseIf strExtension = "XLSM" Then
      GetFormType = mwft_EXCEL_XLSM
   ElseIf strExtension = "DGM" Then
      GetFormType = mwft_DATA_FORM
   ElseIf strExtension = "DAT" And goSession.Workflow.GetExtendedProperty("Version2Compatible") = "True" Then
      '
      ' 6/2002 Eastwind Group Version 2 Compatibility Mode
      '
      GetFormType = mwft_DATA_FORM
   Else
      If Not mSuppressOpenError Then
         goSession.RaiseError "Error in mwSession.mwForm.GetFormType. Invalid Form extension: " & _
           strExtension, Err.Number, Err.Description
      End If
      GetFormType = 0
   End If
   goSession.KillObject fso
   
   Exit Function
FormError:
   goSession.RaiseError "General Error in mwSession.mwForm.GetFormType.", Err.Number, Err.Description
   GetFormType = 0
   goSession.KillObject fso
End Function

Public Function GetFormDetailData(mwFormHeaderKey As Long, mwcSitesKey As Long) As Boolean
'   If FormType = mwft_WORD Then
'      GetFormDetailData = WordForm.GetFormDetailData(mwFormHeaderKey, mwcSitesKey)
   On Error GoTo FunctionError
   GetFormDetailData = False
   
   Select Case FormType
      Case Is = mwft_WORD, mwft_WORD_DOCX, mwft_WORD_DOCM
         GetFormDetailData = moWordForm.GetFormDetailData(mwFormHeaderKey, mwcSitesKey)
         
      Case Is = mwft_EXCEL, mwft_EXCEL_XLSX, mwft_EXCEL_XLSM
      
         GetFormDetailData = moExcelForm.GetFormDetailData(mwFormHeaderKey, mwcSitesKey)
      
      Case Is = mwft_DATA_FORM
      
'         GetFormDetailData = moDataForm.GetFormDetailData(mwFormHeaderKey, mwcSitesKey)
         
   End Select
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.GetFormDetailData.", Err.Number, Err.Description
      
End Function

Private Function GetTemplateFileName(mwFormHeaderKey As Long) As String
   Dim RevNo As Long
   Dim StrRevNo As String
   Dim SourceFName As String
   Dim SourceFExt As String
   Dim mwcPlacesKey As Long
   
   On Error GoTo FunctionError
   
   If mwFormHeaderKey > 0 Then
      RevNo = GetFormHeadersRevNo(mwFormHeaderKey)
   Else
      StrRevNo = GetExtendedProperty("DistributedRevisionNo")
      If Len(StrRevNo) > 0 Then
         RevNo = CLng(StrRevNo)
      Else
         RevNo = 1
      End If
      
      If RevNo < 1 Then
         RevNo = 1
      End If
   End If
   
   SourceFName = GetExtendedProperty("SourceFName")
   SourceFExt = GetExtendedProperty("SourceFExt")
   
   If Len(SourceFName) > 0 And Len(SourceFExt) > 0 Then
      mwcPlacesKey = goSession.Workflow.TemplatePlacesID
      
      If RevNo > 1 Then
         GetTemplateFileName = BuildPlacesPath(mwcPlacesKey) & SourceFName & "_R" & Format(RevNo, "000") & "." & SourceFExt
      Else
         GetTemplateFileName = BuildPlacesPath(mwcPlacesKey) & SourceFName & "." & SourceFExt
      End If
   Else
      GetTemplateFileName = ""
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwForm.GetTemplateFileName.", Err.Number, Err.Description
   GetTemplateFileName = ""
End Function

Private Function GetFormHeadersRevNo(mwFormHeaderKey As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
     
   Set loRs = New Recordset
   
   sSQL = "SELECT TemplateRevisionNo FROM mwFormHeader WHERE ID= " & mwFormHeaderKey
   
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If IsNull(loRs!TemplateRevisionNo) Then
         GetFormHeadersRevNo = 1
      Else
         GetFormHeadersRevNo = loRs!TemplateRevisionNo
      End If
   Else
      GetFormHeadersRevNo = 1
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwForm.GetFormHeadersRevNo ", Err.Number, Err.Description
   CloseRecordset loRs
   GetFormHeadersRevNo = 1
End Function

Private Function BuildPlacesPath(mwcPlacesKey As Long) As String
   Dim oRs As Recordset
   On Error GoTo FunctionError
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   
   If mwcPlacesKey <= 0 Then
      goSession.RaiseError "mwSession.mwForm.BuildPlacesPath: Workflow Profile Template Place is empty, form can not be opened. "
      BuildPlacesPath = ""
   Else
      oRs.Open "select * from mwcPlaces where ID=" & mwcPlacesKey, goCon, adOpenForwardOnly, adLockReadOnly
      If oRs.RecordCount < 1 Then
         BuildPlacesPath = ""
      Else
         If Not IsNull(oRs!DriveLetter) Then
            BuildPlacesPath = BuildPlacesPath & oRs!DriveLetter & ":\"
         End If
         If Not IsNull(oRs!RootPath) Then
            BuildPlacesPath = BuildPlacesPath & oRs!RootPath & "\"
         End If
         If Not IsNull(oRs!SubFolderPath) Then
            BuildPlacesPath = BuildPlacesPath & oRs!SubFolderPath & "\"
         End If
      End If
   End If
   
   CloseRecordset oRs
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwForm.BuildPlacesPath. ", Err.Number, Err.Description
   BuildPlacesPath = ""
   CloseRecordset oRs
End Function

