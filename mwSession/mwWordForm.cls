VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwWordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'mwWordForm - Class supports manipulation of Excel Spreadsheets
' Started: 3/6/2001 ms / Overhauled 7/28/2001
Option Explicit

'   Public WithEvents momwFormHeaderRs As Recordset
   Public WithEvents momwFormDetailRs As Recordset
Attribute momwFormDetailRs.VB_VarHelpID = -1


' Microsoft Word BuiltinDocumentProperties constants...
Const mw_wdPropertyCategory = 18
Const mw_wdPropertyTitle = 1
Const mw_wdPropertyAuthor = 3
Const mw_wdPropertyTimeCreated = 11
Const mw_wdPropertyTimeLastSaved = 12
Const mw_wdPropertySubject = 2
Const mw_wdPropertyManager = 20
Const mw_wdFieldFormCheckBox = 71
Const mw_wdFieldFormDropDown = 83

Public Enum mwWordProtectOffsets
   mw_wdAllowOnlyRevisions = 0
   mw_wdAllowOnlyComments = 1
   mw_wdAllowOnlyFormFields = 2
End Enum
   

' Word Window States
Const MW_wdWindowStateMaximize = 1
Const MW_wdWindowStateMinimize = 2
Const MW_wdWindowStateNormal = 0
Const MW_wdDoNotSaveChanges = 0
Const MW_wdSaveChanges = -1

Private moWord As Object
Private moDoc As Object
'Private moWord As Word.Application
'Private moDoc As Word.Document

Private moParent As mwForm
Private mIsOpen As Boolean
Private mIsApplicationOpen As Boolean
Private mIsLeaveApplicationOpen As Boolean

Public Property Set Parent(oParent As mwForm)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As mwForm
   Set Parent = moParent
End Property

'#If LATE_BIND Then
   Public Property Set WordApp(ByVal Data As Object)
       Set moWord = Data
   End Property
   
   Public Property Get WordApp() As Object
      Set WordApp = moWord
   End Property
   
   Public Property Set WordDoc(ByVal Data As Object)
       Set moDoc = Data
   End Property

   Public Property Get WordDoc() As Object
      Set WordDoc = moDoc
   End Property



'#Else
'   Public Property Set WordApp(ByVal Data As Word.Application)
'       Set moWord = Data
'   End Property
'
'   Public Property Set WordDoc(ByVal Data As Word.Document)
'       Set moDoc = Data
'   End Property
'#End If




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

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property


Public Function CloseWordForm() As Boolean
   On Error GoTo FunctionError
   If mIsOpen Then
      If Not moDoc Is Nothing Then
         moDoc.Close MW_wdDoNotSaveChanges
         goSession.RaiseWorkflowStatus "Closed Word Form: " & moParent.FormProperties.FormID
      End If
      If Not mIsLeaveApplicationOpen Then
         moWord.Quit MW_wdDoNotSaveChanges
         goSession.KillObject moWord
         mIsApplicationOpen = False
      End If
   End If
   CloseWordForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.CloseWordForm.", Err.Number, Err.Description
   CloseWordForm = False
End Function


Public Function OpenForm(Optional FileName As Variant, Optional LeaveApplicationOpen As Boolean) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   If mIsOpen Then
      goSession.RaiseError "Error in mwWordForm.OpenForm, Word Form already Open."
      OpenForm = True
      Exit Function
   End If
   Set fso = New FileSystemObject
   '
   ' File exists ?
   '
   If IsMissing(FileName) Then
      FileName = moParent.FormProperties.FullFilename
   Else
      moParent.FormProperties.FullFilename = FileName
      moParent.FormProperties.FileName = fso.GetBaseName(FileName)
   End If
   If Not moParent.IsFileExists(FileName) Then
      If Not moParent.SuppressOpenError Then
         goSession.RaiseError "Error in mwSession.mwWordForm.OpenForm, File not found: " & FileName
      Else
         goSession.LogIt mwl_Workstation, mwl_Warning, "Suppressed Error in mwSession.mwWordForm.OpenForm, File not found: " & FileName
      End If
      OpenForm = False
      goSession.KillObject fso
      Exit Function
   End If
   moParent.FormProperties.Modified = fso.GetFile(FileName).DateLastModified
      
   If Not mIsApplicationOpen Then
      Set moWord = CreateObject("Word.Application")
      mIsApplicationOpen = True
      If LeaveApplicationOpen Then
         mIsLeaveApplicationOpen = True
      End If
   End If
   moWord.Documents.Open FileName
   Set moDoc = moWord.ActiveDocument
   moParent.FormProperties.subject = moDoc.BuiltinDocumentProperties(mw_wdPropertySubject)

   '
   ' Set Read Only in Parent object...
   '
   moParent.IsReadOnly = moDoc.ReadOnly
   '
   ' Extract Form properties...
   '
   If moParent.ExtractMsFormInfo(moDoc.BuiltinDocumentProperties(mw_wdPropertyCategory)) Then
      mIsOpen = True
      OpenForm = True
      moParent.IsOpen = True
   Else
      mIsOpen = False
      OpenForm = False
      moParent.IsOpen = False
      moWord.Quit MW_wdDoNotSaveChanges
   End If
   goSession.KillObject fso
   Exit Function
FunctionError:
   If Err.Number = 429 Then
      goSession.RaiseError "Error in mwSession.mwWordForm.OpenForm, Could not create Word Application object."
      moParent.IsResourceOpenError = True
   Else
      goSession.RaiseError "General Error in mwSession.mwWordForm.OpenForm.", Err.Number, Err.Description
   End If
   goSession.KillObject fso
   moParent.IsOpen = False
End Function

Friend Function PreviewForm(FileName As String) As Boolean
   On Error GoTo FunctionError
   If moParent.IsOpen Then
      goSession.RaiseError "Error in mwSession.mwWordForm.PreviewForm, Cannot preview an open form."
      PreviewForm = False
      Exit Function
   End If
   #If LATE_BINDING Then
      Set moWord = New Word.Application
   #Else
      Set moWord = CreateObject("Word.Application")
   #End If
   moWord.Documents.Open FileName
   'Set moDoc = loWord.ActiveDocument
   If goSession.User.StartOfficeMaximized Then
      moWord.WindowState = MW_wdWindowStateMaximize
   End If
   moWord.Visible = True
   moWord.Activate
   PreviewForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwWordForm.PreviewForm.", Err.Number, Err.Description
   PreviewForm = False
End Function


Public Function EditForm(Optional FileName As Variant) As Boolean
   On Error GoTo FunctionError
   '
   ' If form isn't open already, open it...
   '
   If Not moParent.IsOpen Then
      If IsMissing(FileName) Then
         If Not OpenForm() Then
            EditForm = False
            Exit Function
         End If
      Else
         If Not OpenForm(FileName) Then
            EditForm = False
            Exit Function
         End If
      End If
   End If
   '
   ' Enter edit Mode in Word...
   '
   If goSession.User.StartOfficeMaximized Then
      moWord.WindowState = MW_wdWindowStateMaximize
   End If
   moWord.Visible = True
   moWord.Activate
   SetWordOnTop moDoc.ActiveWindow.Caption
   
   EditForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWordForm.EditForm.", Err.Number, Err.Description
   EditForm = False
   moParent.IsOpen = False
End Function


Private Sub SetWordOnTop(appTitle As String)
   Dim dummyVariable As Long
   Dim lenTitle As Integer
   Dim winTitle As String * 256
   Dim appHandle As Long
 
   'initialize the function return as False
   If appTitle = "" Then
      appTitle = "Microsoft Word"
   End If
   
   lenTitle = Len(appTitle)

   'Get the handle of the first child of the desktop window
   appHandle = GetTopWindow(0)

   'Loop through all top-level windows and search for the sub-string
   'in the Window title
   Do Until appHandle = 0
      dummyVariable = GetWindowText(appHandle, winTitle, 255)
      If Left(winTitle, lenTitle) = appTitle Then
      
'         SetWindowPos appHandle, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
         SetWindowPos appHandle, HWND_TOP, 0, 0, 0, 0, FLAGS
         Exit Sub
      Else
         appHandle = GetWindow(appHandle, GW_HWNDNEXT)
      End If
   Loop
 End Sub

Public Function StuffNewFormProperties() As Boolean
   On Error GoTo FunctionError
   If Trim(moParent.FormProperties.subject) = "" Then
      moParent.FormProperties.subject = moParent.TemplateProperties.Description
   End If
   
   moDoc.BuiltinDocumentProperties(mw_wdPropertyCategory) = moParent.EncodeMsOfficeFormInfo
   moDoc.BuiltinDocumentProperties(mw_wdPropertySubject) = moParent.FormProperties.subject
   moDoc.BuiltinDocumentProperties(mw_wdPropertyTitle) = moParent.FormProperties.FormID
   moDoc.BuiltinDocumentProperties(mw_wdPropertyAuthor) = goSession.User.UserID
   
   On Error Resume Next
   
   moParent.FormProperties.Modified = moDoc.BuiltinDocumentProperties(mw_wdPropertyTimeLastSaved)
   If Err Then
      moParent.FormProperties.Modified = Now()
   End If
   
   On Error GoTo FunctionError
   
   moDoc.Save
   StuffNewFormProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWordForm.StuffNewFormProperties.", Err.Number, Err.Description
   StuffNewFormProperties = False
End Function

Public Function SaveDocument() As Boolean
   On Error GoTo FunctionError
   moDoc.Save
   SaveDocument = True
   Exit Function
FunctionError:
   SaveDocument = False
End Function

Public Function TearDown(Optional LeaveAppOpen As Boolean)
   On Error GoTo FunctionError
   If Not moWord Is Nothing Then
      If Not LeaveAppOpen Then
         On Error Resume Next
         moWord.Quit MW_wdDoNotSaveChanges
      End If
   End If
   If Not moDoc Is Nothing Then
      Set moDoc = Nothing
   End If
   If Not moParent Is Nothing Then Set moParent = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.TearDown.", Err.Number, Err.Description

End Function

Friend Function StuffPocketPcProperties(ByRef loCG As ConfigGroup) As Boolean
   On Error GoTo FunctionError
   loCG.ConfigKeys.Add "Category", moDoc.BuiltinDocumentProperties(mw_wdPropertyCategory)
   If Trim(moParent.FormProperties.subject) = "" Then
      loCG.ConfigKeys.Add "Subject", moDoc.BuiltinDocumentProperties(mw_wdPropertySubject)
   Else
      loCG.ConfigKeys.Add "Subject", moParent.FormProperties.subject
   End If
   loCG.ConfigKeys.Add "Title", moDoc.BuiltinDocumentProperties(mw_wdPropertyTitle)
   loCG.ConfigKeys.Add "Author", moDoc.BuiltinDocumentProperties(mw_wdPropertyAuthor)
   StuffPocketPcProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.StuffPocketPcProperties.", Err.Number, Err.Description
   StuffPocketPcProperties = False
End Function

Friend Function FetchPocketPcProperties(ByRef loCG As ConfigGroup) As Boolean
   On Error GoTo FunctionError
   moDoc.BuiltinDocumentProperties(mw_wdPropertyCategory) = loCG.ConfigKeys.GetKeyValue("Category")
   moDoc.BuiltinDocumentProperties(mw_wdPropertySubject) = loCG.ConfigKeys.GetKeyValue("Subject")
   moParent.FormProperties.subject = loCG.ConfigKeys.GetKeyValue("Subject")
   moDoc.BuiltinDocumentProperties(mw_wdPropertyTitle) = loCG.ConfigKeys.GetKeyValue("Title")
   moDoc.BuiltinDocumentProperties(mw_wdPropertyAuthor) = loCG.ConfigKeys.GetKeyValue("Author")
   FetchPocketPcProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.FetchPocketPcProperties.", Err.Number, Err.Description
   FetchPocketPcProperties = False
End Function


Public Function SetBookMark(BookmarkName As String, value As String) As Boolean
   On Error GoTo FunctionError
   moDoc.Bookmarks(BookmarkName).Range.Text = value
   SetBookMark = True
   Exit Function
FunctionError:
   SetBookMark = False
End Function

Public Function SetFormField(FieldName As String, value As String) As Boolean
   Dim objListEntry As Object
   
   On Error GoTo FunctionError
   'MsgBox "debug code: Doc: " & moDoc.Name
   'MsgBox "form fields count: " & moDoc.formfields.Count
   If value <> "" Then
'      moDoc.FormFields(CVar(FieldName)).Result = value
      If moDoc.FormFields(CVar(FieldName)).Type = mw_wdFieldFormCheckBox Then
         If value = "1" Then
            moDoc.FormFields(CVar(FieldName)).CheckBox.value = True
         Else
            moDoc.FormFields(CVar(FieldName)).CheckBox.value = False
         End If
      ElseIf moDoc.FormFields(CVar(FieldName)).Type = mw_wdFieldFormDropDown Then
         For Each objListEntry In moDoc.FormFields(CVar(FieldName)).DropDown.ListEntries
            If objListEntry.Name = value Then
               moDoc.FormFields(CVar(FieldName)).DropDown.value = objListEntry.Index
               Exit For
            End If
         Next
      Else
         moDoc.FormFields(CVar(FieldName)).Range.Fields(1).Result.Text = value
'         moDoc.FormFields(CVar(FieldName)).Result = Left(value, 250)
'         moDoc.FormFields(CVar(FieldName)).Result = value
         
'         If Len(value) < 255 Then
'            moDoc.FormFields(CVar(FieldName)).Result = value
'         Else
'            moDoc.FormFields(CVar(FieldName)).Result = "||||"
'
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            moWord.Selection.Collapse
'            moWord.Selection.MoveRight wdCharacter, 1
'            moWord.Selection.TypeText value
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            With moWord.Selection.Find
'               .Execute FindText:="|", replacewith:="", Replace:=wdReplaceAll
'            End With
'         End If
         
      End If
   End If
   SetFormField = True
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaiseError "Error in mwWordForm.SetFormField, Field Missing: " & FieldName
   Else
      goSession.RaiseError "General Error in mwWordForm.SetFormField, on field: " & FieldName, Err.Number, Err.Description
   End If
   SetFormField = False
End Function

Private Function SetFormFieldByFieldNum(FieldNum As Long, value As String) As Boolean
   On Error GoTo FunctionError
   'MsgBox "debug code: Doc: " & moDoc.Name
   'MsgBox "form fields count: " & moDoc.formfields.Count
   If value <> "" Then
'      moDoc.FormFields(CVar(FieldName)).Result = value
      If moDoc.FormFields(FieldNum).Type = mw_wdFieldFormCheckBox Then
         If value = "1" Then
            moDoc.FormFields(FieldNum).CheckBox.value = True
         Else
            moDoc.FormFields(FieldNum).CheckBox.value = False
         End If
      Else
         moDoc.FormFields(FieldNum).Range.Fields(1).Result.Text = value
'         moDoc.FormFields(CVar(FieldName)).Result = Left(value, 250)
'         moDoc.FormFields(CVar(FieldName)).Result = value
         
'         If Len(value) < 255 Then
'            moDoc.FormFields(CVar(FieldName)).Result = value
'         Else
'            moDoc.FormFields(CVar(FieldName)).Result = "||||"
'
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            moWord.Selection.Collapse
'            moWord.Selection.MoveRight wdCharacter, 1
'            moWord.Selection.TypeText value
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            With moWord.Selection.Find
'               .Execute FindText:="|", replacewith:="", Replace:=wdReplaceAll
'            End With
'         End If
         
      End If
   End If
   SetFormFieldByFieldNum = True
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaiseError "Error in mwWordForm.SetFormFieldByFieldNum, Field Missing: Field # " & FieldNum
   Else
      goSession.RaiseError "General Error in mwWordForm.SetFormFieldByFieldNum, on field: Field # " & FieldNum, Err.Number, Err.Description
   End If
   SetFormFieldByFieldNum = False
End Function

Public Function GetBookMark(BookmarkName As String) As String
   On Error GoTo FunctionError
   GetBookMark = moDoc.Bookmarks(BookmarkName).Range.Text
   Exit Function
FunctionError:
   GetBookMark = ""
End Function

Public Function GetFormField(FieldName As String) As String
   On Error GoTo FunctionError
   GetFormField = moDoc.Fields(FieldName).Result
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaiseError "Error in mwWordForm.GetFormField, Field Missing: " & FieldName
   Else
      goSession.RaiseError "General Error in mwWordForm.GetFormField, on field: " & FieldName, Err.Number, Err.Description
   End If
   GetFormField = ""
End Function

Public Function PopulateForm(Optional IsRepopulate As Boolean) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   Dim IsFormField As Boolean
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwWFFormPopulate where TemplateID='" & moParent.TemplateProperties.TemplateID & "'"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount > 0 Then
      '
      ' Populate...
      '
      UnProtectDocument
      Do While Not loRs.EOF
         If loRs!IsWordFormField Then
            '
            ' FormField Interface
            '
            If Not IsNull(loRs!VesselCode) Then
               SetFormField loRs!VesselCode, goSession.Site.SiteID
            End If
            If Not IsNull(loRs!VesselName) Then
               SetFormField loRs!VesselName, goSession.Site.SiteName
            End If
            If Not IsNull(loRs!Voyage_NO) Then
               SetFormField loRs!Voyage_NO, goSession.Site.CurrentVoyage
            End If
            If Not IsNull(loRs!CurrentDate) Then
               SetFormField loRs!CurrentDate, Format(Now(), "Short Date")
            End If
            If Not IsNull(loRs!CurrentMaster) Then
               SetFormField loRs!CurrentMaster, goSession.Site.CurrentMaster
            End If
            If Not IsNull(loRs!CurrentCheng) Then
               SetFormField loRs!CurrentCheng, goSession.Site.GetExtendedProperty("CurrentChiefEngineer")
            End If
            If Not IsNull(loRs!CurrentPort) Then
               SetFormField loRs!CurrentPort, goSession.Site.NextPort
            End If
            
            'MOL-143 Adding Ship IMO Number
            If Not IsNull(loRs!MtmlShipImoNumber) Then
               SetFormField loRs!MtmlShipImoNumber, goSession.Site.ShipImoNumber
            End If
            
            '
            ' Populate Logo ?
            '
            If Not IsRepopulate Then
               If loRs!IsInsertLogo Then
                  If Not (IsNull(loRs!LogoCellBookmark) Or goSession.Site.GetExtendedProperty("mwcCommercialOperatorKey") = "") Then
                     'MsgBox "about to insert logo"
                     moParent.InsertLogo loRs!LogoCellBookmark
                  End If
               End If
            End If
         Else
            '
            ' Bookmark Interface
            '
            If Not IsNull(loRs!VesselCode) Then
               SetBookMark loRs!VesselCode, goSession.Site.SiteID
            End If
            If Not IsNull(loRs!VesselName) Then
               SetBookMark loRs!VesselName, goSession.Site.SiteName
            End If
            If Not IsNull(loRs!Voyage_NO) Then
               SetBookMark loRs!Voyage_NO, goSession.Site.CurrentVoyage
            End If
            If Not IsNull(loRs!CurrentDate) Then
               SetBookMark loRs!CurrentDate, Format(Now(), "Short Date")
            End If
            If Not IsNull(loRs!CurrentMaster) Then
               SetBookMark loRs!CurrentMaster, goSession.Site.CurrentMaster
            End If
            If Not IsNull(loRs!CurrentCheng) Then
               SetBookMark loRs!CurrentCheng, goSession.Site.GetExtendedProperty("CurrentChiefEngineer")
            End If
            If Not IsNull(loRs!CurrentPort) Then
               SetBookMark loRs!CurrentPort, goSession.Site.NextPort
            End If
            
            'MOL-143 Adding Ship IMO Number
            If Not IsNull(loRs!MtmlShipImoNumber) Then
               SetBookMark loRs!MtmlShipImoNumber, goSession.Site.ShipImoNumber
            End If
            
            '
            ' Populate Logo ?
            '
            If loRs!IsInsertLogo Then
               If Not (IsNull(loRs!LogoCellBookmark) Or goSession.Site.GetExtendedProperty("mwcCommercialOperatorKey") = "") Then
                  'MsgBox "about to insert logo"
                  moParent.InsertLogo loRs!LogoCellBookmark
               End If
            End If
         End If
         loRs.MoveNext
      Loop
   End If
   ProtectDocument mw_wdAllowOnlyFormFields
   SaveDocument
   goSession.CloseRecordset loRs
   PopulateForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.PopulateForm. ", Err.Number, Err.Description
   PopulateForm = False
End Function

Friend Function PopulateEventForm(ByRef oRsColumns As Recordset, ByRef oRsData As Recordset) As Boolean
   Dim strField As String
   On Error GoTo FunctionError
   Do While Not oRsColumns.EOF
      '
      'store to string - eliminate pointer...
      strField = oRsColumns!ColumnName
      If Not IsNull(oRsData.Fields(strField).value) Then
         If oRsColumns!IsWordFormField Then
            If IsNull(oRsColumns!FormatString) Then
               SetFormField oRsColumns!CellFieldReference, oRsData.Fields(strField).value
            Else
               SetFormField oRsColumns!CellFieldReference, Format(oRsData.Fields(strField).value, oRsColumns!FormatString)
            End If
         Else
            If IsNull(oRsColumns!FormatString) Then
               SetBookMark oRsColumns!CellFieldReference, oRsData.Fields(strField).value
            Else
               SetBookMark oRsColumns!CellFieldReference, Format(oRsData.Fields(strField).value, oRsColumns!FormatString)
            End If
         End If
      End If
      oRsColumns.MoveNext
   Loop
   PopulateEventForm = True
   Exit Function
FunctionError:
   If Err.Number = 3265 Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Warning in mwWordForm..PopulateEventForm, column not found in database: " & strField
   Else
      goSession.RaiseError "General Error in mwWordForm.PopulateEventForm. ", Err.Number, Err.Description
   End If
   PopulateEventForm = False
End Function



Public Function OpenApplication() As Boolean
   On Error GoTo FunctionError
   If mIsApplicationOpen Then
      OpenApplication = True
      Exit Function
   End If
   Set moWord = CreateObject("Word.Application")
   mIsLeaveApplicationOpen = True
   mIsApplicationOpen = True
   OpenApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.OpenApplication. ", Err.Number, Err.Description
   OpenApplication = False
End Function

Public Function CloseApplication() As Boolean
   On Error GoTo FunctionError
   If Not mIsApplicationOpen Then
      CloseApplication = True
      Exit Function
   End If
   
   If Not moDoc Is Nothing Then
      Set moDoc = Nothing
   End If
   If Not moWord Is Nothing Then
      moWord.Quit MW_wdDoNotSaveChanges
      Set moWord = Nothing
   End If
   mIsLeaveApplicationOpen = False
   mIsApplicationOpen = False
   CloseApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.CloseApplication. ", Err.Number, Err.Description
   CloseApplication = False

End Function

Private Sub Class_Terminate()
   On Error Resume Next
   If Not moDoc Is Nothing Then
      'moDoc.Close MW_wdDoNotSaveChanges
      Set moDoc = Nothing
   End If
   If Not moWord Is Nothing Then
      'moWord.Quit MW_wdDoNotSaveChanges
      Set moWord = Nothing
   End If
End Sub


Public Function PrintForm() As Boolean
   On Error GoTo FunctionError
   moDoc.PrintOut
   PrintForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.PrintForm. ", Err.Number, Err.Description
   PrintForm = False

End Function

Public Function PreviewForm2(FileName As String) As Boolean
   Dim loDoc As Object
   Dim loWord As Object
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   If fso.FileExists(FileName) Then
      Set loWord = CreateObject("Word.Application")
      loWord.Documents.Open FileName
      Set loDoc = loWord.ActiveDocument
      loDoc.Application.Visible = True
      loDoc.PrintPreview
   End If
   
   KillObject fso
   PreviewForm2 = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordForm.PreviewForm2. ", Err.Number, Err.Description
   KillObject fso
   PreviewForm2 = False
End Function

Public Function LoadPicture(FileName As String, Bookmark As String) As Boolean
   On Error GoTo FunctionError
    
      'If moDoc.InlineShapes.Count < 1 Then
      moDoc.InlineShapes.AddPicture FileName, True, True, moDoc.Bookmarks(Bookmark).Range
      'End If
   LoadPicture = True
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaisePublicError "Error in mwSession.mwWordForm.LoadPicture, Bookmark (" _
        & Bookmark & ") missing in Template: " & moParent.TemplateProperties.TemplateID
   Else
      goSession.RaisePublicError "General error in mwSession.mwWordForm.LoadPicture", Err.Number, Err.Description
   End If
   LoadPicture = False
End Function

Public Function ProtectDocument(ProtectionType As mwWordProtectOffsets, Optional Password As String) As Boolean
   Dim strPassword As String
   Dim strPW As String
   On Error GoTo FunctionError
   If Password <> "" Then
      strPW = Password
   ElseIf Trim(moParent.GetExtendedProperty("Protect")) <> "" Then
      strPW = moParent.GetExtendedProperty("Protect")
   Else
      ' No action without password
      Exit Function
   End If
   moDoc.Protect ProtectionType, True, strPW
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwWordForm.ProtectDocument, password in configuration file is incorrect."
   ElseIf Err.Number = 4605 Then
      ' Document already protected. Ignore it.
   Else
      goSession.RaiseError "General Error in mwWordForm.ProtectDocument. ", Err.Number, Err.Description
   End If
End Function

Public Function UnProtectDocument(Optional Password As String)
   Dim strPW As String
   Dim iType As Long
   On Error GoTo FunctionError
   If Password <> "" Then
      strPW = Password
   ElseIf Trim(moParent.GetExtendedProperty("Protect")) <> "" Then
      strPW = moParent.GetExtendedProperty("Protect")
   Else
      Exit Function
   End If
   moDoc.Unprotect strPW
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwWordForm.UnProtectDocument, password in configuration file is incorrect."
   ElseIf Err.Number = 4605 Then
      ' Document not protected. Ignore it.
   Else
      goSession.RaiseError "General Error in mwWordForm.UnProtectDocument. ", Err.Number, Err.Description
   End If
End Function

' SetFormDetailData - Read data from the database and store it in the form

Public Function SetFormDetailData(mwFormHeaderKey As Long) As Boolean
   Dim strSQL As String
   Dim IsFormField As Boolean
   On Error GoTo FunctionError
   
   CloseRecordset momwFormDetailRs
   Set momwFormDetailRs = New Recordset
   momwFormDetailRs.CursorLocation = adUseClient
   strSQL = "select * from mwFormDetail WHERE mwFormHeaderKey=" & mwFormHeaderKey & " ORDER BY DetailSeq"
   momwFormDetailRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If momwFormDetailRs.RecordCount > 0 Then
      '
      ' Populate...
      '
      UnProtectDocument
      Do While Not momwFormDetailRs.EOF
      
' CREATE TABLE mwFormDetail (
' ID long NOT NULL CONSTRAINT mwFormDetail_PK PRIMARY KEY,
' mwFormHeaderKey long,
' DetailSeq long,
' FieldName varchar(32),
' ShortData varchar(250),
' MemoData memo,
' IsShortData bit)
      
            '
            ' FormField Interface
            '
            
            ' SetFormFieldByFieldNum
            
            If IsNull(momwFormDetailRs!FieldName) Then
               If CBool(momwFormDetailRs!IsStringData) = True Then
                  If Not IsNull(momwFormDetailRs!StringData) Then
                     SetFormFieldByFieldNum momwFormDetailRs!DetailSeq, momwFormDetailRs!StringData
                  End If
               Else
                  If Not IsNull(momwFormDetailRs!MemoData) Then
                     SetFormFieldByFieldNum momwFormDetailRs!DetailSeq, momwFormDetailRs!MemoData
                  End If
               End If
            Else
               If CBool(momwFormDetailRs!IsStringData) = True Then
                  If Not IsNull(momwFormDetailRs!StringData) Then
                     SetFormField momwFormDetailRs!FieldName, momwFormDetailRs!StringData
                  End If
               Else
                  If Not IsNull(momwFormDetailRs!MemoData) Then
                     SetFormField momwFormDetailRs!FieldName, momwFormDetailRs!MemoData
                  End If
               End If
            End If
         momwFormDetailRs.MoveNext
      Loop
   End If
   CloseRecordset momwFormDetailRs
   
   ProtectDocument mw_wdAllowOnlyFormFields
   SaveDocument
   SetFormDetailData = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwWordForm.SetFormDetailData. ", Err.Number, Err.Description
   SetFormDetailData = False
End Function

' GetFormDetailData - Read data from the form and store it in the database

Public Function GetFormDetailData(mwFormHeaderKey As Long, mwcSitesKey As Long) As Boolean
   Dim strSQL As String
'   Dim loField As FormField
   Dim loField As Object
   Dim loCnt As Long
   
   On Error GoTo FunctionError
   
   CloseRecordset momwFormDetailRs
   Set momwFormDetailRs = New Recordset
   momwFormDetailRs.CursorLocation = adUseClient
   strSQL = "select * from mwFormDetail WHERE mwFormHeaderKey=" & mwFormHeaderKey & " ORDER BY DetailSeq"
   momwFormDetailRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If momwFormDetailRs.RecordCount > 0 Then
'      UnProtectDocument
      Do While Not momwFormDetailRs.EOF
      
         momwFormDetailRs.Delete
         momwFormDetailRs.MoveNext
      Loop
   End If
   
      
' CREATE TABLE mwFormDetail (
' ID long NOT NULL CONSTRAINT mwFormDetail_PK PRIMARY KEY,
' mwFormHeaderKey long,
' DetailSeq long,
' FieldName varchar(32),
' ShortData varchar(250),
' MemoData memo,
' IsShortData bit)
   
   loCnt = 1
   For Each loField In moDoc.FormFields
      
      If Len(loField.Result) > 0 And Not (loField.Type = mw_wdFieldFormCheckBox And loField.Result = "0") Then
         momwFormDetailRs.AddNew
         momwFormDetailRs!ID = goSession.MakePK("mwFormDetail")
         momwFormDetailRs!mwFormHeaderKey = mwFormHeaderKey
         momwFormDetailRs!DetailSeq = loCnt
         
'         If Len(loField.Name) <= 0 Then
'            goSession.RaisePublicError "General Error in mwWordForm.GetFormDetailData: Field " & loCnt & " in the Word form does not have a Field Name."
'            GetFormDetailData = False
'            momwFormDetailRs.CancelUpdate
'            CloseRecordset momwFormDetailRs
'            Exit Function
'         End If
         
         If Len(loField.Name) > 0 Then
            momwFormDetailRs!FieldName = loField.Name
         End If
         
         momwFormDetailRs!mwcSitesKey = mwcSitesKey
         
         If Len(loField.Result) > 250 Then
            momwFormDetailRs!IsMemoData = True
            momwFormDetailRs!MemoData = loField.Result
         Else
            momwFormDetailRs!IsStringData = True
            momwFormDetailRs!StringData = loField.Result
         End If
         
         momwFormDetailRs.Update
      End If
      loCnt = loCnt + 1
   Next loField
   
   CloseRecordset momwFormDetailRs
   
   GetFormDetailData = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwWordForm.GetFormDetailData. ", Err.Number, Err.Description
   GetFormDetailData = False
End Function

'Private Sub momwFormHeaderRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   Static IsBeginAdd As Boolean
'   Static IsBeginDelete As Boolean
'   On Error GoTo SubError
'   If adReason = adRsnAddNew Then
'      IsBeginAdd = True
'   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
'      IsBeginDelete = False
'   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
'      goSession.ReplicateWork.LogAddChange MWRT_mwFormHeader, momwFormHeaderRs!ID, momwFormHeaderRs!mwcSitesKey, momwFormHeaderRs.Fields
'      IsBeginAdd = False
'   ElseIf adReason = adRsnDelete Then
'      goSession.ReplicateWork.LogDeleteChange MWRT_mwFormHeader, momwFormHeaderRs!ID, momwFormHeaderRs!mwcSitesKey
'      IsBeginDelete = True
'   ElseIf adReason <> adRsnFirstChange Then
'      goSession.ReplicateWork.LogModifyChange MWRT_mwFormHeader, momwFormHeaderRs.Fields, momwFormHeaderRs!mwcSitesKey
'   End If
'   Exit Sub
'SubError:
'   goSession.RaisePublicError "General Error in mwManuals.frmFormExplorer.momwFormHeaderRs_WillChangeRecord. ", Err.Number, Err.Description
'
'End Sub
Private Sub momwFormDetailRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
   goSession.RaisePublicError "General Error in mwManuals.frmFormExplorer.momwFormDetailRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub


