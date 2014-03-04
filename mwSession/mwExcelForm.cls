VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwExcelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
'mwExcelForm - Class supports manipulation of Excel Spreadsheets
' Started: 3/6/2001 ms / Overhauled 7/28/2001

' All tables used to store Form Data must have 4 fields defined:
'   ID long                - Primary Key of the table
'   mwcSitesKey long       - SiteKey where record was created
'   mwFormHeaderKey long   - Key of mwFormHeader record that points to this record
'   XAPSection long        - Section number of the XAP record that created this record
'   LRRowOffset long       - For LR sections, the CurrentOffset for this record

Option Explicit
Private moCon As Connection
Private moRS As Recordset

Private WithEvents moFormHeaderRs As Recordset
Attribute moFormHeaderRs.VB_VarHelpID = -1
Private WithEvents moFormDataRs As Recordset
Attribute moFormDataRs.VB_VarHelpID = -1

Private mwrChangeTableKey As Long
Private mwFormHeaderKey As Long
Private mwFormDetailCnt As Long

Private mIsApplicationOpen As Boolean
Private mIsLeaveApplicationOpen As Boolean

Private miCurrentOffset As Integer
Private mCurrentParentRecordID As Long
Private mCurrentParentTable As String
Private mCurrentSectionNumber As Long
Private mCurrentTable As String
Private mwcSitesKey As Long

Private mLocalGeneralDateFormat As String
Private mLocalShortDateFormat As String
Private mLocalDecimalSeparator As String

Private mLrp As LrParameters

Private moXapRecHeader As ConfigKeys
Private moXapRecLines As ConfigKeys

Private moXAP As ConfigGroups
Private mIsXapOpen As Boolean
Private mXapCurrentRecord As Integer

Const MW_TOTAL_LR_PARAMETERS = 9

Const MW_xlMinimized = &HFFFFEFD4
Const MW_xlMaximized = &HFFFFEFD7
Const MW_xlNormal = &HFFFFEFD1

Const MTML_INDENT = "  "
'Const MTML_QUOTE = """"

Dim moExcel As Object
Dim moWorkbook As Object
'Dim moExcel As Excel.Application
'Dim moWorkbook As Workbook

Dim MTML_QUOTE As String
Dim mIsShipnetMtml As Boolean
Dim miHeaderWorksheet As Integer

' working objects...
Private moParent As mwForm

Public Enum mw_XapMapType
   XAP_DATAGRAM = 1
   XAP_MTML_REQ = 2
End Enum


Type XLProtect
   IsAllowDeletingColumns As Boolean
   IsAllowDeletingRows As Boolean
   IsAllowFiltering As Boolean
   IsAllowFormattingCells As Boolean
   IsAllowFormattingColumns As Boolean
   IsAllowFormattingRows As Boolean
   IsAllowInsertingColumns As Boolean
   IsAllowInsertingHyperlinks As Boolean
   IsAllowInsertingRows As Boolean
   IsAllowSorting As Boolean
   IsAllowUsingPivotTables As Boolean
End Type
 
Dim arrXLProtect() As XLProtect

' Looping Rows Parameters
Type LrParameters
   OffsetIncrement As Integer
   MaxRowOffset As Integer
   NullTestCell As String
   SkipRowTestCell As String
   TotalNonLoopingRowFields As Integer
   TotalLrParamters As Integer
   Worksheet As Integer
End Type

Public Property Set Parent(oParent As mwForm)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As mwForm
   Set Parent = moParent
End Property

Public Property Set ExcelApp(ByVal Data As Object)
    Set moExcel = Data
End Property

Public Property Get ExcelApp() As Object
   If Not moExcel Is Nothing Then
      Set ExcelApp = moExcel
   End If
End Property

Public Property Set Workbook(ByVal Data As Object)
    Set moWorkbook = Data
End Property

Public Property Get Workbook() As Object
   If Not moWorkbook Is Nothing Then
      Set Workbook = moWorkbook
   End If
End Property

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

Public Function CloseExcelForm() As Boolean
   On Error GoTo FunctionError
   moWorkbook.Close False, moParent.FormProperties.FileName
   '
   ' Win 9x bullshit
   '
   If gIsAllowWin9xDelay Then
      If gWin9xMilliseconds > 0 Then
         goSession.API.SleepApp gWin9xMilliseconds
      End If
   End If
   '
   If Not mIsLeaveApplicationOpen Then
      moExcel.Quit
      goSession.KillObject moExcel
      '
      ' Win 9x bullshit
      '
      If gIsAllowWin9xDelay Then
         If gWin9xMilliseconds > 0 Then
            goSession.API.SleepApp gWin9xMilliseconds
         End If
      End If
      '
      mIsApplicationOpen = False
   End If
   CloseExcelForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.CloseForm. ", Err.Number, Err.Description
End Function

Public Function OpenApplication() As Boolean
   On Error GoTo FunctionError
   If mIsApplicationOpen Then
      OpenApplication = True
      Exit Function
   End If
   Set moExcel = CreateObject("Excel.Application")
   mIsLeaveApplicationOpen = True
   mIsApplicationOpen = True
   OpenApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.OpenApplication. ", Err.Number, Err.Description
   OpenApplication = False
End Function

Public Function CloseApplication() As Boolean
   On Error GoTo FunctionError
   If Not mIsApplicationOpen Then
      CloseApplication = True
      Exit Function
   End If
   
   If Not moWorkbook Is Nothing Then
      Set moWorkbook = Nothing
   End If
   If Not moExcel Is Nothing Then
      moExcel.Quit
      Set moExcel = Nothing
   End If
   mIsLeaveApplicationOpen = False
   mIsApplicationOpen = False
   CloseApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.CloseApplication. ", Err.Number, Err.Description
   CloseApplication = False

End Function



Public Function OpenForm(Optional FileName As Variant, Optional LeaveApplicationOpen As Boolean) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   
   mwrChangeTableKey = 0
   
   If IsMissing(FileName) Then
      goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwExcelForm.OpenForm, opening form file: " & moParent.FormProperties.FullFilename
   Else
      goSession.LogIt mwl_Workstation, mwl_Information, "Logging in mwExcelForm.OpenForm, opening form file: " & FileName
   End If
   Set fso = New FileSystemObject
   'goSession.LogIt mwl_Workstation, mwl_Information, "Debug logging in mwExcelForm.OpenForm, after fso object declaration."
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
      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenForm, File not found: " & FileName
      OpenForm = False
      goSession.KillObject fso
      Exit Function
   End If
      
   'Set moExcel = New Excel.Application
   If Not mIsApplicationOpen Then
      'goSession.LogIt mwl_Workstation, mwl_Information, "Debug logging in mwExcelForm.OpenForm, before Excel object declaration."
      Set moExcel = CreateObject("Excel.Application")
      'goSession.LogIt mwl_Workstation, mwl_Information, "Debug logging in mwExcelForm.OpenForm, before Excel object declaration."
      mIsApplicationOpen = True
      If LeaveApplicationOpen Then
         mIsLeaveApplicationOpen = True
      End If
   End If
   moParent.FormProperties.Modified = fso.GetFile(FileName).DateLastModified
  
   Set moWorkbook = moExcel.Workbooks.Open(FileName)
   '
   '
   ' Pocket PC form stripped of its glory ?
   '
   If moParent.IsPocketPcFormCompatible Then
      ' If this fails, may be new form and proeprties will be in the spreadsheet...
      If Not moParent.RestorePocketPCBackupProperties() Then
         moParent.FormProperties.subject = moWorkbook.BuiltinDocumentProperties("Subject")
      End If
   Else
      '
      ' Refresh subject, in case user has changed it in the form...
      '
      moParent.FormProperties.subject = moWorkbook.BuiltinDocumentProperties("Subject")
   End If
   
   If moParent.GetExtendedProperty("IsRetainExcelProtectOptions") = "True" Then
      CopyExcelProtectOptions
   End If
   
   moWorkbook.Worksheets(1).Activate
   '
   ' Extract Form properties...
   '
   ' 9/17/2001 - old vs. new...
   
   moParent.IsReadOnly = moWorkbook.ReadOnly
   
   
   '
   ' If form is read only, disable the Save and Save As buttons...
   '
   'If moWorkbook.ReadOnly Then
   '   moExcel.CommandBars("File").Controls("Save As...").Enabled = False
   '   moExcel.CommandBars("File").Controls("Save").Enabled = False
   'End If
      
   
   If moParent.ExtractMsFormInfo(moWorkbook.BuiltinDocumentProperties("Category")) Then
      'if not mIsFormIncomplete
      OpenForm = True
      moParent.IsOpen = True
   Else
      OpenForm = False
      moParent.IsOpen = False
      CloseExcelForm
   End If
   Exit Function
FunctionError:
   If Err.Number = 429 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenForm, probably Win9.x out of resources.", Err.Number, Err.Description
      moParent.IsResourceOpenError = True
   ElseIf Err.Number = 424 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenForm: Could not open form " & FileName & vbCrLf & "Make sure the drive is connected and usable."
      moParent.IsResourceOpenError = True
   Else
      goSession.RaiseError "General Error in mwSession.mwExcelForm.OpenForm.", Err.Number, Err.Description
   End If
   moParent.IsOpen = False
   
End Function

Friend Function PreviewForm(FileName As String) As Boolean
   On Error GoTo FunctionError
   If moParent.IsOpen Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.PreviewForm, Cannot preview an open form."
      PreviewForm = False
      Exit Function
   End If
         
   If Not mIsApplicationOpen Then
      Set moExcel = CreateObject("Excel.Application")
      mIsApplicationOpen = True
   End If
   Set moWorkbook = moExcel.Workbooks.Open(FileName)
   moWorkbook.Worksheets(1).Activate
   If goSession.User.StartOfficeMaximized Then
      moExcel.WindowState = MW_xlMaximized
   End If
   moExcel.EnableEvents = True
   moExcel.ScreenUpdating = True
   moExcel.Visible = True
   PreviewForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwExcelForm.OpenForm.", Err.Number, Err.Description
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
      moExcel.WindowState = MW_xlMaximized
   End If
   EditForm = True
   moExcel.EnableEvents = True
   moExcel.ScreenUpdating = True
   moExcel.Visible = True
   moExcel.UserControl = True
   
   moWorkbook.Activate
'   SetWindowPos moExcel.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
' WinHandle = goSession.API.RunShellExecute(FileName)
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwExcelForm.EditForm.", Err.Number, Err.Description
   EditForm = False
   moParent.IsOpen = False
End Function

'Public Function GetAllUnprotectedCells(Optional Worksheet As Variant) As String
'   Dim Rng As Range
'
'   On Error GoTo FunctionError
'   If IsMissing(Worksheet) Then Worksheet = 1
'   If moParent.IsOpen Then
'      moExcel.Worksheets(Worksheet).Activate
'      For Each Rng In moExcel.Worksheets(Worksheet).Cells
'         If Rng.Locked = False Then
'
'         End If
'
'      Next Rng
'      GetCell = moExcel.Range(Cell)
'   Else
'      GetCell = ""
'   End If
'   Exit Function
'FunctionError:
'   goSession.RaiseError "General Error in mwSession.mwExcelForm.GetCell, Cell= " & Cell & ": ", Err.Number, Err.Description
'   GetCell = ""
'End Function


Public Function GetCell(Cell As String, Optional Worksheet As Variant) As String
   On Error GoTo FunctionError
   If IsMissing(Worksheet) Then Worksheet = 1
   If moParent.IsOpen Then
      moExcel.Worksheets(Worksheet).Activate
      GetCell = moExcel.Range(Cell)
   Else
      GetCell = ""
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwExcelForm.GetCell, Cell= " & Cell & ": ", Err.Number, Err.Description
   GetCell = ""
End Function


Public Function SetCell(Cell As String, NewValue As String, Optional Worksheet As Variant)
   Dim sExcelFormFormula As String
   On Error GoTo FunctionError
   If IsMissing(Worksheet) Or IsNull(Worksheet) Then Worksheet = 1
   If moParent.IsOpen Then
      moExcel.Worksheets(Worksheet).Activate
' IsDate fails with comma separater decimal (non-US regional)
'      If IsDate(NewValue) Then
'         On Error Resume Next
'         moExcel.Range(Cell) = CDate(NewValue)
'         If err Then
'            On Error GoTo FunctionError
'            moExcel.Range(Cell) = NewValue
'         End If
'      Else
'         moExcel.Range(Cell) = NewValue
'      End If
       sExcelFormFormula = moExcel.Range(Cell).Formula
       moExcel.Range(Cell) = NewValue
       If sExcelFormFormula <> "" Then
          moExcel.Range(Cell).Formula = sExcelFormFormula
       End If
   Else
      goSession.RaiseError "Error in mwSession.mwExcelForm.SetCell, No Excel Form is Open."
   End If
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwSession.mwExcelForm.SetCell, Cell= " & Cell & ": ", Err.Number, Err.Description

End Function

Public Function StuffNewFormProperties() As Boolean
   On Error GoTo FunctionError
   If Trim(moParent.FormProperties.subject) = "" Then
      moParent.FormProperties.subject = moParent.TemplateProperties.Description
   End If
   If moParent.TemplateProperties.IsProtectWorkbook Then
      UnprotectWorkbook
   End If
   moWorkbook.BuiltinDocumentProperties("Category") = moParent.EncodeMsOfficeFormInfo
   moWorkbook.BuiltinDocumentProperties("Subject") = moParent.FormProperties.subject
   moWorkbook.BuiltinDocumentProperties("Title") = moParent.FormProperties.FormID
   moWorkbook.BuiltinDocumentProperties("Author") = goSession.User.UserID
   On Error Resume Next
   moParent.FormProperties.Modified = moWorkbook.BuiltinDocumentProperties("Last Save Time")
   If Err Then
      moParent.FormProperties.Modified = Now()
   End If
   On Error GoTo FunctionError
   If moParent.TemplateProperties.IsProtectWorkbook Then
      'Unlock the workbook
      ProtectWorkbook
   End If
   moWorkbook.Save
   StuffNewFormProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwExcelForm.StuffNewFormProperties.", Err.Number, Err.Description
   StuffNewFormProperties = False
End Function



Public Function UpdateAsRecordset() As Boolean
   Dim strOperation As String
   Dim loRec As ConfigKeys
   Dim oKey As ConfigKey
   Dim i As Integer
   Dim loXapFile As ConfigGroups
   Dim strScriptFile As String
   Dim IsError As Boolean
   Dim IsAnyErrors As Boolean
   On Error GoTo FunctionError
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset, ExcelForm Not Open."
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' it must be an ExcelForm...
   '
   If moParent.FormType <> mwft_EXCEL And moParent.FormType <> mwft_EXCEL_XLSX And moParent.FormType <> mwft_EXCEL_XLSM Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset, Wrong Form Type."
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' and must be an Open Map XAP File...
   '
   If InStr(1, moParent.WorkflowProperties.MapScriptFile, "\") > 0 Then
      strScriptFile = moParent.WorkflowProperties.MapScriptFile
   Else
      strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
   End If
   Set loXapFile = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXapFile Is Nothing Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' XAP File must match...
   '
   If loXapFile.Count = 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset, XAP File Corrupt: "
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' and TemplateID of XAP file must match...
   '
'   If UCase(loXapFile(1).ConfigKeys.GetKeyValue("TEMPLATEID")) <> UCase(moParent.TemplateProperties.TemplateID) Then
'      goSession.RaiseError "Error in mwSession.mwExcelForm.UpdateAsRecordset, TemplateID of XAP File Incorrect: "
'      UpdateAsRecordset = False
'      Exit Function
'   End If


   ' Parent now manages data connection
   Set moCon = moParent.GetDataConnection
   If moCon Is Nothing Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset, unable to get a database connection."
      UpdateAsRecordset = False
      Exit Function
   End If
   '
   ' OK, Create a recordset object to use...
   '
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   '
   IsError = False
   IsAnyErrors = False
   moCon.BeginTrans
      '
      '
      ' For each record (except first moRecord, which is the Header) in the XAP file....
      ' 1-based collection
      '
      mwFormDetailCnt = 0
      For i = 2 To loXapFile.Count
         DoEvents
         Set loRec = loXapFile(i).ConfigKeys
         ' Need an Operation
         
         mCurrentSectionNumber = GetSectionNumber(loXapFile(i).ConfigGroup)
         
         strOperation = loRec.GetKeyValue("Operation")
         If strOperation = "A" Then
            IsError = Not UpdateRsTypeA(loRec)
         ElseIf strOperation = "SA" Then
            IsError = Not UpdateRsTypeSA(loRec)
         ElseIf strOperation = "LR" Then
            IsError = Not UpdateRsTypeLR(loRec)
         ElseIf strOperation = "SLR" Then
            IsError = Not UpdateRsTypeSLR(loRec)
         Else
            goSession.RaiseError "Error in mwExcelForm.UpdateAsRecordset, invalid Operation Type: " & strOperation
            IsError = True
         End If
         '
         ' Close the recordset for reuse...
         '
         If moRS.State = adStateOpen Then
            moRS.Close
         End If
         '
         ' Error Check...
         '
         If IsError Then
            If moParent.WorkflowProperties.TolerateOtherErrors Then
               IsError = False
               IsAnyErrors = True
'               goSession.RaisePublicError "Error in Record: " & str(i) & ". Tolerate Errors is set to True, continuing process..."
            Else
               goSession.RaisePublicError "Error in Record: " & str(i) & ". Tolerate Errors is set to False, cancelling process..."
               UpdateAsRecordset = False
               moCon.RollbackTrans
               Exit Function
            End If
         End If
      Next i
      '
      ' We be done...
      '
      
      If IsAnyErrors Then
'         goSession.RaisePublicError "Partially Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         UpdateAsRecordset = False
      Else
'         goSession.RaisePublicError "Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         UpdateAsRecordset = True
      End If
   moCon.CommitTrans
   Exit Function
FunctionError:
   If moParent.WorkflowProperties.TolerateDuplicateRecords And Err.Number = 1 Then
      'Future: Need to raise an event here to inform that a duplicate encountered...
      goSession.RaisePublicError "Error in UpdateAsRecordset: " & Err.Description
      Resume Next
   Else
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateAsRecordset.", Err.Number, Err.Description
      UpdateAsRecordset = False
   End If
   On Error Resume Next
   moCon.RollbackTrans
End Function

Private Function UpdateRsTypeA(xaprec As ConfigKeys) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      UpdateRsTypeA = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeA, Table name missing from Record. "
      Exit Function
   End If
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "select * from " & mCurrentTable & " where ID=-1"
   moRS.Open sSQL, moCon, adOpenStatic, adLockOptimistic
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      UpdateRsTypeA = False
      'goSession.CloseRecordset moRS
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   
   '
   ' Skip record Test
   '
   IsSkipRecord = False
   strNullTestCell = Trim(UCase(xaprec.GetKeyValue("NullTestCell")))
   If strNullTestCell <> "" And strNullTestCell <> "N/A" Then
      If IsNullCellValue(moExcel.Range(strNullTestCell)) Then
         IsSkipRecord = True
      End If
   End If
   If Not IsSkipRecord Then
   
      moRS.AddNew
      'If goSession.User.DebugMode Then MsgBox "after addnew"
      If Err Then
         UpdateRsTypeA = False
         goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeA," & " Unable to add new record, Table: " & mCurrentTable
         'goSession.CloseRecordset moRS
         Exit Function
      End If
      If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
         nKey = goSession.MakePK(mCurrentTable, moCon)
         'moRS.Fields(0).value = nKey
         moRS!ID = nKey
         '
         ' If Child, save key to the parent record...
         '
         If xaprec.GetKeyValue("IsChildKeyName") <> "" And _
              UCase(xaprec.GetKeyValue("IsChildKeyName")) <> "N/A" Then
            moRS.Fields(xaprec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
         End If
         ' Save initial record
         If moRS.Fields.Count < 75 Then
            moRS.Update
            If Err Then
               goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeA: ", Err.Number, Err.Description
               goSession.RaisePublicError "Failed to Update Table during Call to MakePK: " & mCurrentTable
               goSession.CloseRecordset moRS
               UpdateRsTypeA = False
               Exit Function
            End If
         End If
      End If
      If UpdateRSFields(xaprec) Then
         If goSession.User.DebugMode Then MsgBox "about to .update"
         moRS.Update
         If Err Then
            goSession.RaiseError "Error in mwSession.mwExcelForm.UpdateRsTypeA during .Update. ", _
              Err.Number, Err.Description
            UpdateRsTypeA = False
            moRS.CancelUpdate
            Exit Function
         End If
         '
         ' If Parent, save key to the parent record...
         '
         If xaprec.GetKeyValue("IsParentKeyName") <> "" Then
            mCurrentParentRecordID = moRS.Fields(xaprec.GetKeyValue("IsParentKeyName")).value
         End If
      Else
         moRS.CancelUpdate
         goSession.RaisePublicError "Error in mwDataForm.UpdateRsTypeA. CancelUpdate on Form " & moParent.FormProperties.FormID
         UpdateRsTypeA = False
         Exit Function
      End If
   End If
   UpdateRsTypeA = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.UpdateRsTypeA. ", Err.Number, Err.Description
   UpdateRsTypeA = False
End Function

Private Function UpdateRsTypeSA(xaprec As ConfigKeys) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   
   mCurrentTable = "mwFormDetail"
'   mCurrentTable = xaprec.GetKeyValue("table")
'
'   If Trim(mCurrentTable) = "" Then
'      '
'      UpdateRsTypeSA = False
'      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeSA, Table name missing from Record. "
'      Exit Function
'   End If
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "select * from " & mCurrentTable & " where ID=-1"
   moFormDataRs.Open sSQL, moCon, adOpenStatic, adLockOptimistic
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeSA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      UpdateRsTypeSA = False
      'goSession.CloseRecordset moRS
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   
   '
   ' Skip record Test
   '
   IsSkipRecord = False
   strNullTestCell = Trim(UCase(xaprec.GetKeyValue("NullTestCell")))
   If strNullTestCell <> "" And strNullTestCell <> "N/A" Then
      If IsNullCellValue(moExcel.Range(strNullTestCell)) Then
         IsSkipRecord = True
      End If
   End If
   If Not IsSkipRecord Then
      UpdateSimpleRSFields xaprec
   End If
   UpdateRsTypeSA = True
   CloseRecordset moFormDataRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.UpdateRsTypeSA. ", Err.Number, Err.Description
   UpdateRsTypeSA = False
   CloseRecordset moFormDataRs
End Function


Private Function UpdateRsTypeLR(xaprec As ConfigKeys) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      UpdateRsTypeLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.UpdateRsTypeLR, Worksheet Parameter Missing."
      UpdateRsTypeLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   ' First row a no brainer
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   '
   ' Open the recordset...
   '
   
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      UpdateRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeLR, Table name missing from Record. "
      Exit Function
   End If
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   '
   ' Open the recordset
   '
   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
      sSQL = "Select * from " & mCurrentTable & " where ID=-1"
      moRS.Open sSQL, moCon, adOpenStatic, adLockOptimistic
   Else
      moRS.Open mCurrentTable, moCon, adOpenStatic, adLockOptimistic, adCmdTable
   End If
   If Err Then
      UpdateRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moRS
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   iTotalRows = 0
   IsMoreRows = True
   iCurrentOffset = 0
   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))
   Do While IsMoreRows
      '
      ' Null Value Terminate Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, iCurrentOffset))) Then
'            goSession.RaisePublicError "NullCellTest: Finished adding rows, total count = " & iTotalRows
            Exit Do
         End If
      End If
      '
      ' Skip Row Check
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))) Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
         iTotalRows = iTotalRows + 1
         '
         ' Add new ADO record
         '
         On Error Resume Next
         moRS.AddNew
         'If goSession.User.DebugMode Then MsgBox "after addnew"
         If Err Then
            UpdateRsTypeLR = False
            goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeLR," & " Unable to add new record, Table: " & mCurrentTable
            'goSession.CloseRecordset moRS
            Exit Function
         End If
         If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
            moRS!ID = goSession.MakePK(mCurrentTable, moCon)
            moRS.Update
            If Err Then
               goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeLR: ", Err.Number, Err.Description
               goSession.RaisePublicError "Failed to Update Table during Call to MakePK: " & mCurrentTable
               goSession.CloseRecordset moRS
               UpdateRsTypeLR = False
               Exit Function
            End If
         End If
         '
         ' If Child, save key to the parent record...
         '
         On Error GoTo FunctionError
         If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
            moRS.Fields(xaprec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
            iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
         End If
         
         
         
         '
         ' OK, add the non repeating values first...
         '
         If mLrp.TotalNonLoopingRowFields > 0 Then
            For i = iTotalLrParameters To (mLrp.TotalNonLoopingRowFields + (iTotalLrParameters - 1))
               If Not StuffRsField(xaprec(i)) Then
                  ' Terminate
                  moRS.CancelUpdate
                  UpdateRsTypeLR = False
                  Exit Function
               End If
            Next i
         End If
         '
         ' Then, current row variables
         '
         For i = (mLrp.TotalNonLoopingRowFields + iTotalLrParameters) To (xaprec.Count)
            If Not StuffRsField(xaprec(i), iCurrentOffset) Then
               ' Terminate
               moRS.CancelUpdate
               UpdateRsTypeLR = False
               Exit Function
            End If
         Next i
         '
         ' Update record...
         '
         On Error Resume Next
         If goSession.User.DebugMode Then MsgBox "about to .update"
         moRS.Update
         If Err Then
            goSession.RaiseError "Error in mwSession.mwExcelForm.UpdateRsTypeLR during .Update. ", _
              Err.Number, Err.Description
            UpdateRsTypeLR = False
            moRS.CancelUpdate
            Exit Function
         End If
      End If
      '
      ' Another row to process ?
      '
      iCurrentOffset = iCurrentOffset + mLrp.OffsetIncrement
      If (iStartingRowOffset + iCurrentOffset) > mLrp.MaxRowOffset Then
         Exit Do
      End If
      IsSkipRow = False
   Loop
   UpdateRsTypeLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.UpdateRsTypeLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.UpdateRsTypeLR. ", Err.Number, Err.Description
   End If
   UpdateRsTypeLR = False
End Function

Private Function UpdateRsTypeSLR(xaprec As ConfigKeys) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      UpdateRsTypeSLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.UpdateRsTypeSLR, Worksheet Parameter Missing."
      UpdateRsTypeSLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   ' First row a no brainer
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   '
   ' Open the recordset...
   '
   
   mCurrentTable = "mwFormDetail"
'   mCurrentTable = xaprec.GetKeyValue("table")
'   If Trim(mCurrentTable) = "" Then
'      '
'      UpdateRsTypeSLR = False
'      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeSLR, Table name missing from Record. "
'      Exit Function
'   End If
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   '
   ' Open the recordset
   '
'   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
      sSQL = "Select * from " & mCurrentTable & " where ID=-1"
      moFormDataRs.Open sSQL, moCon, adOpenStatic, adLockOptimistic
'   Else
'      moRS.Open mCurrentTable, moCon, adOpenStatic, adLockOptimistic, adCmdTable
'   End If
   If Err Then
      UpdateRsTypeSLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.UpdateRsTypeSLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moRS
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   iTotalRows = 0
   IsMoreRows = True
   iCurrentOffset = 0
   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))
   Do While IsMoreRows
      '
      ' Null Value Terminate Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, iCurrentOffset))) Then
'            goSession.RaisePublicError "NullCellTest: Finished adding rows, total count = " & iTotalRows
            Exit Do
         End If
      End If
      '
      ' Skip Row Check
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))) Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
         iTotalRows = iTotalRows + 1
         On Error GoTo FunctionError
         If mLrp.TotalNonLoopingRowFields > 0 Then
            For i = iTotalLrParameters To (mLrp.TotalNonLoopingRowFields + (iTotalLrParameters - 1))
               If Not StuffSimpleRsField(xaprec(i)) Then
                  ' Terminate
                  UpdateRsTypeSLR = False
                  Exit Function
               End If
            Next i
         End If
         '
         ' Then, current row variables
         '
         For i = (mLrp.TotalNonLoopingRowFields + iTotalLrParameters) To (xaprec.Count)
            If Not StuffSimpleRsField(xaprec(i), iCurrentOffset) Then
               ' Terminate
'               moRS.CancelUpdate
               UpdateRsTypeSLR = False
               Exit Function
            End If
         Next i
      End If
      '
      ' Another row to process ?
      '
      iCurrentOffset = iCurrentOffset + mLrp.OffsetIncrement
      If (iStartingRowOffset + iCurrentOffset) > mLrp.MaxRowOffset Then
         Exit Do
      End If
      IsSkipRow = False
   Loop
   UpdateRsTypeSLR = True
   CloseRecordset moFormDataRs
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.UpdateRsTypeSLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.UpdateRsTypeSLR. ", Err.Number, Err.Description
   End If
   UpdateRsTypeSLR = False
   CloseRecordset moFormDataRs
End Function

Private Function UpdateRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   DoEvents
   If Not IsNumeric(oRec.GetKeyValue("WORKSHEET")) Then
      goSession.RaisePublicError "Error in Excel Script, non-numeric WORKSHEET specified (using 1)."
      moExcel.Worksheets(1).Activate
   Else
      moExcel.Worksheets(Int(oRec.GetKeyValue("WORKSHEET"))).Activate
   End If
   For Each loKey In oRec
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         Case Else
            'Debug.Print loKey.KeyName
            If Not StuffRsField(loKey) Then
               ' terminate...
               UpdateRSFields = False
               Exit Function
            End If
      End Select
   Next loKey
   UpdateRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.UpdateRSFields: ", _
     Err.Number, Err.Description
   UpdateRSFields = False
End Function

Private Function UpdateSimpleRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   DoEvents
   If Not IsNumeric(oRec.GetKeyValue("WORKSHEET")) Then
      goSession.RaisePublicError "Error in Excel Script, non-numeric WORKSHEET specified (using 1)."
      moExcel.Worksheets(1).Activate
   Else
      moExcel.Worksheets(Int(oRec.GetKeyValue("WORKSHEET"))).Activate
   End If
   For Each loKey In oRec
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         Case Else
            If Not StuffSimpleRsField(loKey) Then
               ' terminate...
               UpdateSimpleRSFields = False
               Exit Function
            End If
      End Select
   Next loKey
   UpdateSimpleRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.UpdateSimpleRSFields: ", _
     Err.Number, Err.Description
   UpdateSimpleRSFields = False
End Function


Public Function ConvertToDatagram() As mwForm
   Dim strType As String
   On Error GoTo FunctionError
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagram, ExcelForm Not Open."
      Exit Function
   End If
   '
   ' it must be an ExcelForm...
   '
   If moParent.FormType <> mwft_EXCEL And moParent.FormType <> mwft_EXCEL_XLSX And moParent.FormType <> mwft_EXCEL_XLSM Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagram, Wrong Form Type."
      Exit Function
   End If
   '
   ' Branch off, depending upon  Datagram Type...
   '
   strType = moParent.GetExtendedProperty("ConvertToDataGramType")
   If strType = "" Or strType = "1" Then
      Set ConvertToDatagram = ConvertToDatagramRS()
   ElseIf strType = "2" Then
      Set ConvertToDatagram = ConvertToDatagramMTML
   Else
      goSession.RaiseError "Error in mwExcelForm.ConvertToDatagram, invalid DatagramType: " & strType
      moParent.IsIncompleteForm = True
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwExcelForm.ConvertToDatagram: ", Err.Number, Err.Description
   moParent.IsIncompleteForm = True
End Function


Private Function ConvertToDatagramRS() As mwForm
   ' XAP File input
   Dim loXap As ConfigGroups
   Dim loXapRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim iSheet As Integer
   ' Output to datagram...
   Dim loDgForm As mwForm
   Dim loDatagram As ConfigGroups
   Dim loDatagramRec As ConfigKeys
   Dim loDatagramKey As ConfigKey
   ' working variables
   Dim strDgGeneralDateFormat As String
   Dim strDgShortDateFormat As String
   Dim strOperation As String
   Dim strScriptFile As String
   Dim IsError As Boolean
   Dim i As Integer
   
   On Error GoTo FunctionError
   '
   ' Open XAP File...
   '
   strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
   Set loXap = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXap Is Nothing Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      Exit Function
   End If
   '
   ' XAP File must be valid...
   '
   If loXap.Count = 0 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, XAP File Corrupt: "
      Exit Function
   End If
   '
   ' and TemplateID of Script must match...
   '
'   If UCase(loXap(1).ConfigKeys("TEMPLATEID").KeyValue) <> UCase(moParent.TemplateProperties.TemplateID) Then
'      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, TemplateID of XAP File Incorrect: " & _
'        moParent.TemplateProperties.TemplateID & "/" & UCase(loXap(1).ConfigKeys("TEMPLATEID").KeyValue)
'      Exit Function
'   End If
   
   
   '''''''''''''''''
   '
   ' Duck's Guts...
   '
   ' 1. Staging
   strDgGeneralDateFormat = DATAGRAM_DATE_FORMAT
   strDgShortDateFormat = DATAGRAM_SHORT_DATE_FORMAT
   
   '
   ' Make Transaction File
   '
   Set loDgForm = goSession.OpenNewForm()
   If Not loDgForm.MakeNewForm(moParent.TemplateProperties.DatagramTemplateID) Then
   'If Not loDgForm.MakeNewForm(moParent.TemplateProperties.DatagramTemplateID, , , moParent.FormProperties.FormID) Then
      goSession.KillObject loDgForm
      Exit Function
   End If
   '
   ' 2. Master Loop through Excel Mapping Script
   '
   '
   IsError = False
   '
   ' Outer Loop: For each record in Excel Map File
   '
   For i = 2 To loXap.Count
      DoEvents
      Set loXapRec = loXap(i).ConfigKeys
      ' Need an Operation
      strOperation = UCase(loXapRec.GetKeyValue("Operation"))
      If Trim(strOperation) = "" Then
         ' Must be "Operation" Code.
         goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, Operation Code missing from Excel Mapping Record: " & str(i)
         Exit Function
      End If
      '
      ' Need Worksheet for record
      '
      If Not IsNumeric(loXapRec("WORKSHEET").KeyValue) Then
         ' Must be "Worksheet" Code.
         goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, Worksheet Code missing or non-numeric from Excel Mapping Record: " & str(i)
         'moExcel.Worksheets(1).Activate
         Exit Function
      Else
         moExcel.Worksheets(Int(loXapRec("WORKSHEET").KeyValue)).Activate
      End If
      '
      ' Need a table name for rest of commands...
      '
      If Trim(loXapRec.GetKeyValue("table")) = "" Then
         '
         goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramRS, Table name missing from Excel Mapping  Record: " & str(i)
         Exit Function
      End If
      '
      ' 7/2002 - NullTestCell
      '
      If Trim(UCase(loXapRec.GetKeyValue("NullTestCell"))) <> "N/A" And _
        Trim(UCase(loXapRec.GetKeyValue("NullTestCell"))) <> "" Then
         If Trim(moExcel.Range(loXapRec.GetKeyValue("NullTestCell"))) = "" Then
            '
            ' Blank row encountered, we are finished
            '
            goSession.Logger.LogIt mwl_User_Defined, mwl_Information, "Skipping record based on NullTestCell: " & _
              loXapRec.GetKeyValue("NullTestCell")
            GoTo SkipRecord
         End If
      End If
      
      '
      ' OK build the Datagram record....
      '
      If strOperation = "A" Then
         If Not ConvertTypeA(loDgForm, loXapRec) Then
            ' Error Handler
            moParent.IsIncompleteForm = True
            If Not moParent.WorkflowProperties.TolerateOtherErrors Then
               Exit Function
            Else
'               goSession.RaisePublicError "Tolerate Errors turned on, will keep on trucking with: " _
'                 & moParent.TemplateProperties.TemplateID
            End If
         End If
      ElseIf strOperation = "LR" Then
         If Not ConvertTypeLR(loDgForm, loXapRec) Then
            ' Error Handler
            moParent.IsIncompleteForm = True
            If Not moParent.WorkflowProperties.TolerateOtherErrors Then
               Exit Function
            Else
'               goSession.RaisePublicError "Tolerate Errors turned on, will keep on trucking with: " _
'                 & moParent.TemplateProperties.TemplateID
            End If
         End If
      Else
         goSession.RaiseError "Error in mwExcelForm.ConvertToDatagramRS, invalid operation: " & strOperation & _
            ", in Excel Mapping File: " & moParent.WorkflowProperties.MapScriptFile
         If Not moParent.WorkflowProperties.TolerateOtherErrors Then
            Exit Function
         Else
'            goSession.RaisePublicError "Tolerate Errors turned on, will keep on trucking with: " _
'              & moParent.TemplateProperties.TemplateID
         End If
      End If
SkipRecord:
   Next i ' next record in script map file
   '
   ' Write the DataForm File
   '
   If Not loDgForm.DataForm.WriteDataForm() Then
      'ConvertToDatagramRS = ""
      goSession.KillObject loDgForm
      Exit Function
   End If
   '
   ' Return new Datagram Form...
   Set ConvertToDatagramRS = loDgForm
   '
   ' Cleanup Objects
   '
   goSession.KillObject loXapRec
   goSession.KillObject loXap
   goSession.KillObject loDgForm
   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwExcelForm.ConvertToDatagramRS: ", Err.Number, Err.Description
   moParent.IsIncompleteForm = True
   goSession.KillObject loXapRec
   goSession.KillObject loXap
   goSession.KillObject loDgForm
End Function

Private Function ConvertTypeA(ByRef DgForm As mwForm, xaprec As ConfigKeys) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim strKeyName As String
   Dim strKeyValue As String
   Dim iWorksheet As Integer
   Dim IsSkipRow As Boolean
   Dim strNullTestCell As String
   '
   'Loop through each key in script map, creating fields in datagram....
   '
   On Error GoTo FunctionError
   '
   ' Skip record Test
   '
   IsSkipRow = False
   strNullTestCell = Trim(UCase(xaprec.GetKeyValue("NullTestCell")))
   If strNullTestCell <> "" And strNullTestCell <> "N/A" Then
      If IsNullCellValue(moExcel.Range(strNullTestCell)) Then
         IsSkipRow = True
      End If
   End If
   If Not IsSkipRow Then
      Set loDgRec = DgForm.DataForm.AddRecord
      loDgRec.Add "Table", xaprec.GetKeyValue("table")
      loDgRec.Add "Operation", "A"
      If IsNumeric(xaprec.GetKeyValue("WORKSHEET")) Then
         iWorksheet = xaprec.GetKeyValue("WORKSHEET")
         If iWorksheet < 1 Then
            iWorksheet = 1
         End If
      Else
         iWorksheet = 1
      End If
      '
      ' Check if Parent or Child record flag nees to be written to datagram
      '
      If Trim(UCase(xaprec.GetKeyValue("IsParentKeyName"))) <> "" Then
         loDgRec.Add "IsParentKeyName", xaprec.GetKeyValue("IsParentKeyName")
      End If
      If Trim(UCase(xaprec.GetKeyValue("IsChildKeyName"))) <> "" Then
            loDgRec.Add "IsChildKeyName", xaprec.GetKeyValue("IsChildKeyName")
      End If
      moWorkbook.Worksheets(iWorksheet).Activate
   
      For Each loXapKey In xaprec
         '
         ' Real Ducks Guts - Make the Datagram line
         '
         
         Select Case UCase(loXapKey.KeyName)
            ' exclude from processing
            Case Is = "TABLE"
            Case Is = "OPERATION"
            Case Is = "WORKSHEET"
            Case Is = "NULLTESTCELL"
            Case Is = "ISPARENTKEYNAME"
            Case Is = "ISCHILDKEYNAME"
            Case Else
               strKeyName = loXapKey.KeyName
               strKeyValue = loXapKey.KeyValue
               If Not IsNullCellValue(moExcel.Range(strKeyValue)) Then
                  loDgRec.Add strKeyName, ReplaceVbCrLf(moExcel.Range(strKeyValue))
               End If
         End Select
         '
      Next loXapKey
   End If
   ConvertTypeA = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertTypeA, Invalid Cell ?: " & strKeyValue & ". ", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.ConvertTypeA, Cell: " & strKeyValue & " .", Err.Number, Err.Description
   End If
   ConvertTypeA = False
End Function

Private Function ConvertTypeLR(ByRef DgForm As mwForm, xaprec As ConfigKeys) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalmLrparameters As Integer
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      ConvertTypeLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.ConvertTypeLR, Worksheet Parameter Missing."
      ConvertTypeLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   '
   ' First row parameters
   '
   iTotalRows = 0
   IsMoreRows = True
   iCurrentOffset = 0
   '
   ' Set Total Parameters...
   '
   If Trim(UCase(xaprec.GetKeyValue("IsChildKeyName"))) <> "" And _
     UCase(xaprec.GetKeyValue("IsChildKeyName")) <> "N/A" Then
      iTotalmLrparameters = MW_TOTAL_LR_PARAMETERS + 1
   Else
      iTotalmLrparameters = MW_TOTAL_LR_PARAMETERS
   End If
   '
   ' Starting Row Offset - Used in MaxRowOffset Test
   '
   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(xaprec(mLrp.TotalNonLoopingRowFields + iTotalmLrparameters).KeyValue, 0))
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   
   Do While IsMoreRows
      '
      ' Null Value Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, iCurrentOffset))) Then
            '
            ' Blank row encountered, we are finished
            '
            Exit Do
         End If
      End If
      '
      ' Skip Row Test
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))) Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         'iTotalmLrparameters = MW_TOTAL_LR_PARAMETERS
         Set loDgRec = DgForm.DataForm.AddRecord()
         iTotalRows = iTotalRows + 1
         loDgRec.Add "Table", xaprec.GetKeyValue("table")
         loDgRec.Add "Operation", "A"
         '
         ' Check if Child record flag needs to be written to datagram
         '
         If Trim(UCase(xaprec.GetKeyValue("IsChildKeyName"))) <> "" And _
           UCase(xaprec.GetKeyValue("IsChildKeyName")) <> "N/A" Then
            loDgRec.Add "IsChildKeyName", xaprec.GetKeyValue("IsChildKeyName")
            'iTotalmLrparameters = MW_TOTAL_LR_PARAMETERS + 1
         End If
         '
         ' OK, add the non repeating values first...
         '
         If mLrp.TotalNonLoopingRowFields > 0 Then
            For i = iTotalmLrparameters To (mLrp.TotalNonLoopingRowFields + (iTotalmLrparameters - 1))
               loDgRec.Add xaprec(i).KeyName, ReplaceVbCrLf(moExcel.Range(xaprec(i).KeyValue))
            Next i
         End If
         '
         ' Then, current row variables
         '
         'For i = (mLrp.TotalNonLoopingRowFields + iTotalmLrparameters) To xaprec.Count
         For i = (mLrp.TotalNonLoopingRowFields + iTotalmLrparameters) To (xaprec.Count)
            If Not IsNullCellValue(moExcel.Range(GetLoopingRowCell(xaprec(i).KeyValue, iCurrentOffset))) Then
               loDgRec.Add xaprec(i).KeyName, ReplaceVbCrLf(moExcel.Range(GetLoopingRowCell(xaprec(i).KeyValue, iCurrentOffset)))
            End If
         Next i
      End If
      '
      ' Another row to process ?
      '
      iCurrentOffset = iCurrentOffset + mLrp.OffsetIncrement
      If (iStartingRowOffset + iCurrentOffset) > mLrp.MaxRowOffset Then
         Exit Do
      End If
      IsSkipRow = False
   Loop
   ConvertTypeLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertTypeLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.ConvertTypeLR. ", Err.Number, Err.Description
   End If
   ConvertTypeLR = False
End Function


Private Function CheckLrparameters(ByRef xaprec As ConfigKeys) As Boolean
   On Error GoTo FunctionError
   
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, Worksheet Parameter Missing."
      CheckLrparameters = False
      Exit Function
   End If
   If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
      mLrp.Worksheet = xaprec.GetKeyValue("Worksheet")
   Else
      mLrp.Worksheet = 1
   End If
   
   If Trim(xaprec.GetKeyValue("OffsetIncrement")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, OffsetIncrement Parameter Missing."
      CheckLrparameters = False
      Exit Function
   Else
      If IsNumeric(xaprec("OffsetIncrement").KeyValue) Then
         mLrp.OffsetIncrement = xaprec.GetKeyValue("OffsetIncrement")
      Else
         mLrp.OffsetIncrement = 1
      End If
   
   End If
   If Trim(xaprec.GetKeyValue("MaxRowOffset")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, MaxRowOffset Parameter Missing."
      CheckLrparameters = False
      Exit Function
   Else
      mLrp.MaxRowOffset = xaprec.GetKeyValue("MaxRowOffset")
   End If
   If Trim(xaprec.GetKeyValue("NullTestCell")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, NullTestCell Parameter Missing."
      CheckLrparameters = False
      Exit Function
   Else
      mLrp.NullTestCell = xaprec.GetKeyValue("NullTestCell")
   End If
   If Trim(xaprec.GetKeyValue("SkipRowTestCell")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, SkipRowTestCell Parameter Missing."
      CheckLrparameters = False
      Exit Function
   Else
      mLrp.SkipRowTestCell = xaprec.GetKeyValue("SkipRowTestCell")
   End If
   If Trim(xaprec.GetKeyValue("TotalNonLoopingRowFields")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.CheckLrparameters, TotalNonLoopingRowFields Parameter Missing."
      CheckLrparameters = False
      Exit Function
   Else
      mLrp.TotalNonLoopingRowFields = xaprec.GetKeyValue("TotalNonLoopingRowFields")
   End If
   CheckLrparameters = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwExcelForm.CheckLrparameters. ", Err.Number, Err.Description
   CheckLrparameters = False
End Function




' Returns the spreadsheet cell, taking into account the offset increment...
Public Function GetLoopingRowCell(BaselineCell As String, OffsetIncrement As Integer) As String
   Dim i As Integer
   Dim j As Integer
   Dim s As String
   Dim strPrefix As String
   Dim iNumberStart As Integer
   Dim StrCell As String
   
   On Error GoTo FunctionError
   
   StrCell = Trim(Replace(BaselineCell, vbTab, " "))

   If OffsetIncrement = 0 Then
      GetLoopingRowCell = Trim(StrCell)
   Else
      ' Find start of numeric position
      For i = 1 To Len(StrCell)
         s = mID$(StrCell, i, 1)
         ' Look for start of numeric portion of cell reference
         If Asc(s) >= Asc("0") And Asc(s) <= Asc("9") Then
            Exit For
         End If
      Next i
      'If i = Len(StrCell) Then
      '   goSession.RaiseError "Error in mwExcelForm.GetLoopingRowCell, invalid cell reference: " & StrCell
      '   GetLoopingRowCell = ""
      'Else
         '
         ' Recalculate new cell reference
         'grab Number...
         j = CInt(mID(StrCell, i)) + OffsetIncrement
         GetLoopingRowCell = Left(StrCell, (i - 1)) & LTrim(str(j))
      'End If
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwExcelForm.GetLoopingRowCell. ", Err.Number, Err.Description
   GetLoopingRowCell = ""
End Function

Public Function GetBaselineRowOffset(BaselineCell As String) As Integer
   Dim i As Integer
   Dim s As String
   Dim StrCell As String
   
   On Error GoTo FunctionError
   
   StrCell = Trim(Replace(BaselineCell, vbTab, " "))
   
   If UCase(StrCell) = "N/A" Then
      GetBaselineRowOffset = 0
   Else
   
      For i = 1 To Len(StrCell)
         s = mID$(StrCell, i, 1)
         ' Look for start of numeric portion of cell reference
         If Asc(s) >= Asc("0") And Asc(s) <= Asc("9") Then
            Exit For
         End If
      Next i
      If IsNumeric(mID(StrCell, i)) Then
         GetBaselineRowOffset = CInt(mID(StrCell, i))
      Else
         goSession.RaiseError "Error in mwExcelForm.GetBaselineRowOffset, bad cell reference: " & StrCell
         GetBaselineRowOffset = -1
      End If
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwExcelForm.GetBaselineRowOffset, Cell: " & StrCell, Err.Number, Err.Description
   GetBaselineRowOffset = -1

End Function


Private Sub Class_Initialize()
   MTML_QUOTE = Chr(34)
   mwrChangeTableKey = 0

   mLocalGeneralDateFormat = goSession.API.GetLocaleString(LOCALE_LONG_DATE)
   mLocalShortDateFormat = goSession.API.GetLocaleString(LOCALE_SHORT_DATE)
   mLocalDecimalSeparator = goSession.API.GetLocaleString(LOCALE_DECIMAL_SEPARATOR)
   
End Sub

Private Sub Class_Terminate()
   If Not moWorkbook Is Nothing Then Set moWorkbook = Nothing
   If Not moExcel Is Nothing Then Set moExcel = Nothing
   If Not moRS Is Nothing Then Set moRS = Nothing
   If Not moCon Is Nothing Then Set moCon = Nothing
   If Not moParent Is Nothing Then Set moParent = Nothing
End Sub


Public Function TearDown(Optional LeaveAppOpen As Boolean)
   On Error GoTo FunctionError
   If Not moWorkbook Is Nothing Then Set moWorkbook = Nothing
   
   
   If Not moExcel Is Nothing Then
      If Not LeaveAppOpen Then
         If Not moExcel.Application.ActiveWorkbook Is Nothing Then
            moExcel.Application.ActiveWorkbook.Close False
         End If
         moExcel.Quit
      End If
      Set moExcel = Nothing
   End If
   CloseRecordset moRS
   CloseRecordset moFormHeaderRs
   CloseRecordset moFormDataRs

   If Not moCon Is Nothing Then Set moCon = Nothing
   If Not moParent Is Nothing Then Set moParent = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.Teardown. ", Err.Number, Err.Description
End Function


Public Function SaveWorkbook() As Boolean
   On Error GoTo FunctionError
   moWorkbook.Save
   SaveWorkbook = True
   Exit Function
FunctionError:
   SaveWorkbook = False
End Function



Private Function ConvertToDatagramMTML() As mwForm

   ' XAP File input
   Dim loDgForm As mwForm
   Dim loXap As ConfigGroups
   Dim loXapKey As ConfigKey
   Dim iSheet As Integer
   Dim strScriptFile As String
   
   ' Local Variables
   Dim strReq As String
   Dim strTemp As String
   Dim iTotalLines As Integer
   
   ' Set Shipnet MTML Flag
   If goSession.Site.GetExtendedProperty("IsShipnetMtml") = "True" Then
      mIsShipnetMtml = True
   End If
   
   
   On Error GoTo FunctionError
   '
   ' Load XAP Template information
   '
   '
   ' Open XAP File...
   '
   strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
   Set loXap = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXap Is Nothing Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramMTML, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      Exit Function
   End If
   '
   ' XAP File must be valid...
   '
   If loXap.Count = 0 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramMTML, XAP File Corrupt: "
      Exit Function
   End If
   '
   ' and TemplateID of Script must match...
   '
'   If UCase(loXap(1).ConfigKeys("TEMPLATEID").KeyValue) <> UCase(moParent.TemplateProperties.TemplateID) Then
'      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramMTML, TemplateID of XAP File Incorrect: " & _
'        moParent.TemplateProperties.TemplateID & "/" & UCase(loXap(1).ConfigKeys("TEMPLATEID").KeyValue)
'      Exit Function
'   End If
   '
   ' Make Datagram form...
   '
   Set loDgForm = goSession.OpenNewForm()
   
   If mIsShipnetMtml Then
      If Not loDgForm.MakeNewForm(moParent.TemplateProperties.DatagramTemplateID) Then
         goSession.KillObject loDgForm
         Exit Function
      End If
   Else
      If Not loDgForm.MakeNewForm(moParent.TemplateProperties.DatagramTemplateID, , , _
        moParent.FormProperties.FormID) Then
         goSession.KillObject loDgForm
         Exit Function
      End If
   End If
   '
   ' Set Form ID to same as requisition....
   '
   loDgForm.FormProperties.FormID = moParent.FormProperties.FormID
   '
   ' Staging the XAP file...
   '
   Set moXapRecHeader = loXap(2).ConfigKeys
   Set moXapRecLines = loXap(3).ConfigKeys
   
   If Not IsNumeric(moXapRecHeader("WORKSHEET").KeyValue) Then
      ' Must be "Worksheet" Code.
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramMTML, Header Worksheet Code missing or non-numeric from Excel Mapping Record."
      Exit Function
   Else
       miHeaderWorksheet = Int(moXapRecHeader("WORKSHEET").KeyValue)
   End If
   
   If Not IsNumeric(moXapRecLines("WORKSHEET").KeyValue) Then
      ' Must be "Worksheet" Code.
      goSession.RaiseError "Error in mwSession.mwExcelForm.ConvertToDatagramMTML, Lines Worksheet Code missing or non-numeric from Excel Mapping Record."
      Exit Function
   Else
       mLrp.Worksheet = Int(moXapRecLines("WORKSHEET").KeyValue)
   End If
   '
   ' File Header
   '
   'Shipnet chokes on below...
   '
   'strReq = "<?xml version=" & MTML_QUOTE & "1.0" & MTML_QUOTE _
     & " encoding=" & MTML_QUOTE & "utf-8" & MTML_QUOTE & " ?>" & vbCrLf
   strReq = "<?xml version=" & MTML_QUOTE & "1.0" & MTML_QUOTE & " ?>" & vbCrLf
   strReq = strReq & "<!DOCTYPE MTML SYSTEM " & MTML_QUOTE & "MTML.DTD" & MTML_QUOTE & "> " & vbCrLf
   strReq = strReq & "<MTML>" & vbCrLf
   '
   ' Create Interchange Sender string
   '
   strReq = strReq & MTML_INDENT _
     & "<Interchange Sender=" & MTML_QUOTE & goSession.Site.GetExtendedProperty("mtmlShipImoNumber") & MTML_QUOTE _
     & " SenderCodeQualifier=" & MTML_QUOTE & "_11" & MTML_QUOTE _
     & " Recipient=" & MTML_QUOTE & goSession.ThisSite.GetExtendedProperty("mtmlBuyersImpaID") & MTML_QUOTE _
     & " RecipientCodeQualifier=" & MTML_QUOTE & "ZIM" & MTML_QUOTE _
     & " PreparationDate=" & MTML_QUOTE & Format(Now(), "YYMMDD") & MTML_QUOTE _
     & " PreparationTime=" & MTML_QUOTE & Format(Now(), "hhmm") & MTML_QUOTE _
     & " ControlReference=" & MTML_QUOTE & goSession.ThisSite.GetExtendedProperty("mtmlInterchangeReference") & MTML_QUOTE _
     & " Identifier=" & MTML_QUOTE & "UNOC" & MTML_QUOTE _
     & " VersionNumber=" & MTML_QUOTE & "2" & MTML_QUOTE & ">" & vbCrLf
   
   '
   ' Requisition Contents Qualifier
   '
   SetActiveWorksheet miHeaderWorksheet
   If Trim(GetMtmlHCell("Priority")) = "" Then
      moExcel.Range(moXapRecHeader("Priority").KeyValue) = "Low"
   End If
   strReq = strReq & MTML_INDENT & MTML_INDENT & "<Requisition ContentsQualifier=" & MTML_QUOTE _
     & GetMtmlHCell("ContentsQualifier") _
     & MTML_QUOTE & " CurrencyCode=" & MTML_QUOTE & GetMtmlHCell("CurrencyCode") & MTML_QUOTE _
     & " MessageNumber=" & MTML_QUOTE & moParent.FormProperties.FormID & MTML_QUOTE _
     & " MessageReferenceNumber=" & MTML_QUOTE & moParent.FormProperties.FormID & MTML_QUOTE _
     & " FunctionCode=" & MTML_QUOTE & "_9" & MTML_QUOTE _
     & " LineItemCount=" & MTML_QUOTE & "[LineItemCount]" & MTML_QUOTE _
     & " Priority=" & MTML_QUOTE & GetMtmlHCell("Priority") & MTML_QUOTE _
     & " VersionNumber=" & MTML_QUOTE & "D" & MTML_QUOTE _
     & " ReleaseNumber=" & MTML_QUOTE & "96A" & MTML_QUOTE _
     & " ControllingAgency= " & MTML_QUOTE & "UN" & MTML_QUOTE _
     & " AssociationAssignedCode=" & MTML_QUOTE & "MARL10" & MTML_QUOTE & ">" & vbCrLf
     
   '
   ' Date/Time Enumerations:
   '
   '  Format Qualifier:
   '  _203 - Date/Time
   '  _102 - Date
   '
   ' DateTime period Qualifiers:
   '
   '
   '  _137 - Document date/Time stamp
   '  _2   - Requested Delivery Date
   '  _132 - ETA
   '  _133 - ETD
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<DateTimePeriod Qualifier=" & MTML_QUOTE & "_137" & MTML_QUOTE _
     & " FormatQualifier=" & MTML_QUOTE & "_203" & MTML_QUOTE _
     & " Value= " & MTML_QUOTE & Format(Now(), "YYYYMMDDhhmm") & MTML_QUOTE & "/>" & vbCrLf
   
   If Trim(GetMtmlHCell("RequestedDeliveryDate")) <> "" Then
      If IsDate((GetMtmlHCell("RequestedDeliveryDate"))) Then
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<DateTimePeriod Qualifier=" & MTML_QUOTE & "_2" & MTML_QUOTE _
           & " FormatQualifier=" & MTML_QUOTE & "_102" & MTML_QUOTE _
           & " Value= " & MTML_QUOTE & Format(GetMtmlHCell("RequestedDeliveryDate"), "YYYYMMDD") & MTML_QUOTE & "/>" & vbCrLf
      End If
   End If
   
   If Trim(GetMtmlHCell("EtaPort")) <> "" Then
      If IsDate(GetMtmlHCell("EtaPort")) Then
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<DateTimePeriod Qualifier=" & MTML_QUOTE & "_132" & MTML_QUOTE _
           & " FormatQualifier=" & MTML_QUOTE & "_102" & MTML_QUOTE _
           & " Value= " & MTML_QUOTE & Format(GetMtmlHCell("EtaPort"), "YYYYMMDD") & MTML_QUOTE & "/>" & vbCrLf
      End If
   End If
   
   If Trim(GetMtmlHCell("EtdPort")) <> "" Then
      If IsDate(GetMtmlHCell("EtdPort")) Then
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<DateTimePeriod Qualifier=" & MTML_QUOTE & "_133" & MTML_QUOTE _
           & " FormatQualifier=" & MTML_QUOTE & "_102" & MTML_QUOTE _
           & " Value= " & MTML_QUOTE & Format(GetMtmlHCell("EtdPort"), "YYYYMMDD") & MTML_QUOTE & "/>" & vbCrLf
      End If
   End If
   
   
   
   '
   ' Comments Qualifier
   '  ZAT - Internal Comments
   '  PUR - Comments to Supplier
   '
   strTemp = GetMtmlHCell("InternalComments")
   If Trim(strTemp) <> "" Then
    strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Comments Qualifier=" & MTML_QUOTE & "ZAT" & MTML_QUOTE & ">" _
      & " <Value> " & strTemp & " </Value></Comments>" & vbCrLf
   End If
   strTemp = GetMtmlHCell("CommentsToSupplier")
   If Trim(strTemp) <> "" Then
      strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Comments Qualifier=" & MTML_QUOTE & "PUR" & MTML_QUOTE & ">" _
        & " <Value> " & strTemp & "</Value></Comments>" & vbCrLf
   End If
   '
   ' Reference Qualifier
   '  UC - Ship's Req Number
   '  ZYH - Shipyard Hull Number
   '  ZYS - Shipyard
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Reference Qualifier=" & MTML_QUOTE & "UC" & MTML_QUOTE _
     & " ReferenceNumber= " & MTML_QUOTE & GetMtmlHCell("RequisitionID") & MTML_QUOTE & " />" & vbCrLf
   
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Reference Qualifier=" & MTML_QUOTE & "ZYH" & MTML_QUOTE _
     & " ReferenceNumber= " & MTML_QUOTE & goSession.Site.GetExtendedProperty("mtmlHullNumber") & MTML_QUOTE & " />" & vbCrLf
   
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Reference Qualifier=" & MTML_QUOTE & "ZYS" & MTML_QUOTE _
     & " ReferenceNumber= " & MTML_QUOTE & goSession.Site.GetExtendedProperty("mtmlShipyardCode") & MTML_QUOTE & " />" & vbCrLf
   '
   ' Party Qualifier
   '  UD - Name of Vessel
   '  _11 - Ship's IMO Number
   '  BY - Buyer's Organization Buyer's IMPA Code
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Party Qualifier=" & MTML_QUOTE & "UD" & MTML_QUOTE _
     & " Name= " & MTML_QUOTE & goSession.Site.SiteName & MTML_QUOTE _
     & " CodeListResponsibleAgency= " & MTML_QUOTE & "_11" & MTML_QUOTE _
     & " Identification= " & MTML_QUOTE & goSession.Site.GetExtendedProperty("mtmlShipImoNumber") & MTML_QUOTE & ">" & vbCrLf
   '
   ' Port
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<PartyLocation Qualifier=" & MTML_QUOTE & "ZUC" & MTML_QUOTE & " Port=" & _
     MTML_QUOTE & GetMtmlHCell("RequestedDeliveryPort") & MTML_QUOTE & "/>" & vbCrLf
   '
   ' Contact/Email
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Contact FunctionCode=" & MTML_QUOTE & "CN" & MTML_QUOTE & " Name=" & _
     MTML_QUOTE & GetMtmlHCell("ContactName") & MTML_QUOTE & ">" & vbCrLf
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<CommunicationMethod Number=" & MTML_QUOTE & _
     GetMtmlHCell("EmailContact") & MTML_QUOTE & " Qualifier=" & MTML_QUOTE & "EM" & MTML_QUOTE & "/>" & vbCrLf
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & "</Contact>" & vbCrLf
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "</Party>" & vbCrLf
   
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Party Qualifier=" & MTML_QUOTE & "BY" & MTML_QUOTE _
     & " Name= " & MTML_QUOTE & goSession.ThisSite.GetExtendedProperty("mtmlBuyersName") & MTML_QUOTE _
     & " Identification= " & MTML_QUOTE & goSession.ThisSite.GetExtendedProperty("mtmlBuyersImpaID") & MTML_QUOTE _
     & " CodeListResponsibleAgency= " & MTML_QUOTE & "ZIM" & MTML_QUOTE & "/>" & vbCrLf
   
   '
   ' Equipment description
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & "<Equipment Description=" & MTML_QUOTE _
     & GetMtmlHCell("EqptDescription") & MTML_QUOTE _
     & " Manufacturer= " & MTML_QUOTE & GetMtmlHCell("Manufacturer") & MTML_QUOTE _
     & " ModelNumber= " & MTML_QUOTE & GetMtmlHCell("ModelNumber") & MTML_QUOTE _
     & " SerialNumber= " & MTML_QUOTE & GetMtmlHCell("SerialNumber") & MTML_QUOTE _
     & " Type= " & MTML_QUOTE & GetMtmlHCell("EqptType") & MTML_QUOTE _
     & " Name= " & MTML_QUOTE & GetMtmlHCell("EqptName") & MTML_QUOTE & "/>" & vbCrLf
     
     
   '
   ' Build  Line Items...
   '
   If Not ConvertTypeMTML(strReq, iTotalLines, GetMtmlHCell("AccountCode")) Then
      Exit Function
   End If
   '
   ' Closing stuff...
   '
   strReq = strReq & MTML_INDENT & MTML_INDENT & "</Requisition>" & vbCrLf
   strReq = strReq & MTML_INDENT & "</Interchange>" & vbCrLf
   strReq = strReq & "</MTML>" & vbCrLf
   strReq = Replace(strReq, "[LineItemCount]", str(iTotalLines))
   loDgForm.DataForm.MtmlReqBody = strReq
   '
   ' Write the DataForm File
   '
   If Not loDgForm.DataForm.WriteDataForm() Then
      'ConvertToDatagramRS = ""
      goSession.KillObject loDgForm
      Exit Function
   End If
   
   Set ConvertToDatagramMTML = loDgForm
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertToDatagramMTML, Invalid Cell Reference ?: " & loXapKey.KeyValue & "."
   ElseIf Err.Number = 91 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertToDatagramMTML, Parameter missing from Mapping File ?: " & loXapKey.KeyName
   Else
      goSession.RaiseError "General Error in mwExcelForm.ConvertToDatagramMTML. ", Err.Number, Err.Description
   End If

End Function



Private Function ConvertTypeMTML(ByRef strReq As String, _
  ByRef TotalRows As Integer, AccountCode As String) As Boolean
   Dim iTotalRows As Integer
   Dim iWorksheet As Integer
   Dim strTemp1 As String
   Dim strTemp2 As String
   Dim strTemp3 As String
   Dim iOffsetIncrement As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(moXapRecLines) Then
      ConvertTypeMTML = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   iTotalRows = 0
   IsMoreRows = True
   miCurrentOffset = 0
   '
   
   Do While IsMoreRows
      moWorkbook.Worksheets(mLrp.Worksheet).Activate
      '
      ' Null Value Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If Trim(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, miCurrentOffset))) = "" Then
            '
            ' Blank row encountered, we are finished
            '
            Exit Do
         End If
      End If
      '
      ' Skip Row Test
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If Trim(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, miCurrentOffset))) = "" Or _
          Trim(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, miCurrentOffset))) = "0" Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         iTotalRows = iTotalRows + 1
         '
         '------------Add Line Item------------------------------------
         '
         
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT _
           & "<LineItem Number=" & MTML_QUOTE & GetMtmlLCell("LineItemNumber") & MTML_QUOTE _
           & " Description= " & MTML_QUOTE & GetMtmlLCell("ItemDescription") & MTML_QUOTE _
           & " TypeCode= " & MTML_QUOTE & "MF" & MTML_QUOTE _
           & " Identification= " & MTML_QUOTE & GetMtmlLCell("MfgPartNo") & MTML_QUOTE _
           & " MeasureUnitQualifier= " & MTML_QUOTE & UCase(GetMtmlLCell("UnitOfMeasure")) & MTML_QUOTE _
           & " Quantity= " & MTML_QUOTE & GetMtmlLCell("Quantity") & MTML_QUOTE _
           & " RemainingOnBoard= " & MTML_QUOTE & GetMtmlLCell("PartRemainingOnboard") & MTML_QUOTE & ">" & vbCrLf
         '
         ' Comment ?
         '
         If Trim(GetMtmlLCell("AdditionalComments")) <> "" Then
            strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT _
              & "<Comments Qualifier=" & MTML_QUOTE & "LIN" & MTML_QUOTE & "> <Value> " _
              & GetMtmlLCell("AdditionalComments") _
              & " </Value>" & vbCrLf & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT _
              & " </Comments>" & vbCrLf
         End If
         '
         ' Section
         '
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT _
           & " <Section AccountNumber=" & MTML_QUOTE & AccountCode & MTML_QUOTE
         
         
         If Trim(GetMtmlLCell("PartEqptMfg")) <> "" Then
            'Add additional description...
            strReq = strReq _
              & " Manufacturer=" & MTML_QUOTE & GetMtmlLCell("PartEqptMfg") & MTML_QUOTE
         ElseIf mIsShipnetMtml Then
            If Trim(GetMtmlHCell("Manufacturer")) <> "" Then
               strReq = strReq _
                 & " Manufacturer=" & MTML_QUOTE & GetMtmlHCell("Manufacturer") & MTML_QUOTE
            End If
         End If
         If Trim(GetMtmlLCell("PartEqptName")) <> "" Then
            strReq = strReq _
              & " Name =" & MTML_QUOTE & GetMtmlLCell("PartEqptName") & MTML_QUOTE
         ElseIf mIsShipnetMtml Then
            If Trim(GetMtmlHCell("EqptName")) <> "" Then
               strReq = strReq _
                 & " Name=" & MTML_QUOTE & GetMtmlHCell("EqptName") & MTML_QUOTE
            End If
         End If
         '
         ' Other possible section stuff...
         '
         If mIsShipnetMtml Then
            'Description
            If Trim(GetMtmlHCell("EqptDescription")) <> "" Then
               strReq = strReq _
                 & " Description=" & MTML_QUOTE & GetMtmlHCell("EqptDescription") & MTML_QUOTE
            End If
            'Serial Number
            If Trim(GetMtmlHCell("SerialNumber")) <> "" Then
               strReq = strReq _
                 & " SerialNumber=" & MTML_QUOTE & GetMtmlHCell("SerialNumber") & MTML_QUOTE
            End If
            'Model Number
            If Trim(GetMtmlHCell("ModelNumber")) <> "" Then
               strReq = strReq _
                 & " ModelNumber=" & MTML_QUOTE & GetMtmlHCell("ModelNumber") & MTML_QUOTE
            End If
            'Type
            'If Trim(GetMtmlHCell("EqptType")) <> "" Then
            '   strReq = strReq _
            '     & " Type=" & MTML_QUOTE & GetMtmlHCell("EqptType") & MTML_QUOTE
            'End If
         
         
         End If
         '
         ' Finish Line Item...
         '
         strReq = strReq & " />" & vbCrLf
         strReq = strReq & MTML_INDENT & MTML_INDENT & MTML_INDENT & MTML_INDENT & " </LineItem>" & vbCrLf
      End If
      '
      '-------------------------------------------------------------
      '
      ' Another row to process ?
      '
      miCurrentOffset = miCurrentOffset + mLrp.OffsetIncrement
      If miCurrentOffset > mLrp.MaxRowOffset Then
         ' we be maxed out...
         Exit Do
      End If
      IsSkipRow = False
   Loop
   ' Finish line item
   TotalRows = iTotalRows
   ConvertTypeMTML = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertTypeMTML, probably Invalid Cell Reference in XAP mapping file. "
   ElseIf Err.Number = 91 Then
      goSession.RaiseError "Error in mwExcelForm.ConvertTypeMTML, probably parameter missing from the XAP mapping file. "
   Else
      goSession.RaiseError "General Error in mwExcelForm.ConvertTypeMTML. ", Err.Number, Err.Description
   End If
   ConvertTypeMTML = False

End Function

Private Function GetMtmlLCell(Cellname As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   moWorkbook.Worksheets(mLrp.Worksheet).Activate
   strTemp = moExcel.Range(GetLoopingRowCell(moXapRecLines(Cellname).KeyValue, miCurrentOffset))
   GetMtmlLCell = moParent.ReplaceXmlCharacters(strTemp)
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.GetMtmlLCell. ", Err.Number, Err.Description
End Function

Private Function GetMtmlHCell(Cellname As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   moWorkbook.Worksheets(miHeaderWorksheet).Activate
   strTemp = moExcel.Range(moXapRecHeader(Cellname).KeyValue)
   GetMtmlHCell = moParent.ReplaceXmlCharacters(strTemp)
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.GetMtmlLCell. ", Err.Number, Err.Description
End Function

Public Function UnprotectWorksheet(Optional SheetOffset As Integer)
   Dim strPassword As String
   On Error GoTo FunctionError
   If SheetOffset <> 0 Then
      moWorkbook.Worksheets(SheetOffset).Activate
   End If
   If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
      moWorkbook.ActiveSheet.Unprotect
      moWorkbook.Unprotect
   Else
      moWorkbook.ActiveSheet.Unprotect moParent.GetExtendedProperty("Protect")
      'moWorkbook.unProtect moParent.GetExtendedProperty("Protect")
   End If
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.UnprotectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwExcelForm.UnprotectWorkbook. ", Err.Number, Err.Description
   End If
End Function

Public Function UnprotectWorkbook() As Boolean
   Dim strPassword As String
   On Error GoTo FunctionError
   If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
      moWorkbook.Unprotect
   Else
      moWorkbook.Unprotect moParent.GetExtendedProperty("Protect")
   End If
   UnprotectWorkbook = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.UnprotectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwExcelForm.UnprotectWorkbook. ", Err.Number, Err.Description
   End If
   UnprotectWorkbook = False
End Function

Public Function UnprotectAllWorksheets()
   Dim strPassword As String
   Dim i As Integer
   On Error GoTo FunctionError
'   If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
'      Exit Function
'   End If
   For i = 1 To moWorkbook.Worksheets.Count
      moWorkbook.Worksheets(i).Activate
      If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
         moWorkbook.ActiveSheet.Unprotect
      Else
         moWorkbook.ActiveSheet.Unprotect moParent.GetExtendedProperty("Protect")
      End If
   Next i
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.UnprotectAllWorksheets, password in MW configuration does not match all worksheets."
   Else
      goSession.RaiseError "General Error in mwExcelForm.UnprotectAllWorksheets. ", Err.Number, Err.Description
   End If
End Function

Public Function ProtectWorksheet(Optional SheetOffset As Integer)
   Dim strPassword As String
   On Error GoTo FunctionError
   If SheetOffset <> 0 Then
      moWorkbook.Worksheets(SheetOffset).Activate
   End If
   If moParent.GetExtendedProperty("IsRetainExcelProtectOptions") = "True" Then
      SetAdvancedProtectOptions moWorkbook.ActiveSheet
   Else
      If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
         moWorkbook.ActiveSheet.Protect
         'moWorkbook.Protect
      Else
         moWorkbook.ActiveSheet.Protect moParent.GetExtendedProperty("Protect")
         'moWorkbook.Protect moParent.GetExtendedProperty("Protect")
      End If
   End If


   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ProtectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwExcelForm.ProtectWorkbook. ", Err.Number, Err.Description
   End If
End Function

Public Function ProtectWorkbook() As Boolean
   Dim strPassword As String
   On Error GoTo FunctionError
   If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
      moWorkbook.Protect
   Else
      moWorkbook.Protect moParent.GetExtendedProperty("Protect")
   End If
   ProtectWorkbook = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ProtectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwExcelForm.ProtectWorkbook. ", Err.Number, Err.Description
   End If
   ProtectWorkbook = False
End Function

Public Function ProtectAllWorksheets()
   Dim strPassword As String
   Dim i As Integer
   On Error GoTo FunctionError
'   If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
'      Exit Function
'   End If
   For i = 1 To moWorkbook.Worksheets.Count
      moWorkbook.Worksheets(i).Activate
      
      If moParent.GetExtendedProperty("IsRetainExcelProtectOptions") = "True" Then
         SetAdvancedProtectOptions moWorkbook.ActiveSheet
      Else
         If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
            moWorkbook.ActiveSheet.Protect
         Else
            moWorkbook.ActiveSheet.Protect moParent.GetExtendedProperty("Protect")
         End If
      End If
   Next i
   moWorkbook.Worksheets(1).Activate
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.ProtectAllWorksheets, password in MW configuration does not match all worksheets."
   Else
      goSession.RaiseError "General Error in mwExcelForm.ProtectAllWorksheets. ", Err.Number, Err.Description
   End If
End Function

Public Function SetActiveWorksheet(SheetOffset As Integer)
   On Error GoTo FunctionError
   If SheetOffset <> 0 Then
      moWorkbook.Worksheets(SheetOffset).Activate
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.SetActiveWorksheet. ", Err.Number, Err.Description
End Function


Friend Function StuffPocketPcProperties(ByRef loCG As ConfigGroup) As Boolean
   On Error GoTo FunctionError
   loCG.ConfigKeys.Add "Category", moWorkbook.BuiltinDocumentProperties("Category")
   If Trim(moParent.FormProperties.subject) = "" Then
      loCG.ConfigKeys.Add "Subject", moWorkbook.BuiltinDocumentProperties("Subject")
   Else
      loCG.ConfigKeys.Add "Subject", moParent.FormProperties.subject
   End If
   loCG.ConfigKeys.Add "Title", moWorkbook.BuiltinDocumentProperties("Title")
   loCG.ConfigKeys.Add "Author", moWorkbook.BuiltinDocumentProperties("Author")
   StuffPocketPcProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.StuffPocketPcProperties. ", Err.Number, Err.Description
   StuffPocketPcProperties = False
End Function


Friend Function FetchPocketPcProperties(ByRef loCG As ConfigGroup) As Boolean
   On Error GoTo FunctionError
   moWorkbook.BuiltinDocumentProperties("Category") = loCG.ConfigKeys("Category").KeyValue
   moWorkbook.BuiltinDocumentProperties("Subject") = loCG.ConfigKeys("Subject").KeyValue
   moParent.FormProperties.subject = loCG.ConfigKeys("Subject").KeyValue
   moWorkbook.BuiltinDocumentProperties("Title") = loCG.ConfigKeys("Title").KeyValue
   moWorkbook.BuiltinDocumentProperties("Author") = loCG.ConfigKeys("Author").KeyValue
   FetchPocketPcProperties = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.StuffPocketPcProperties. ", Err.Number, Err.Description
   FetchPocketPcProperties = False
End Function


Private Function OpenXAPFile() As Boolean
   Dim strScriptFile As String
      
   On Error GoTo FunctionError
   If mIsXapOpen Then
      OpenXAPFile = True
      Exit Function
   End If
   strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
   Set moXAP = goSession.LoadConfigGroupsFile(strScriptFile)
   If moXAP Is Nothing Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenXAPFile, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      OpenXAPFile = False
      Exit Function
   End If
   '
   ' XAP File must be valid...
   '
   If moXAP.Count = 0 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenXAPFile, XAP File Corrupt: "
      OpenXAPFile = False
      Exit Function
   End If
   '
   ' and TemplateID of Script must match...
   '
'   If UCase(moXAP(1).ConfigKeys("TEMPLATEID").KeyValue) <> UCase(moParent.TemplateProperties.TemplateID) Then
'      goSession.RaiseError "Error in mwSession.mwExcelForm.OpenXAPFile, TemplateID of XAP File Incorrect: " & _
'        moParent.TemplateProperties.TemplateID & "/" & UCase(moXAP(1).ConfigKeys("TEMPLATEID").KeyValue)
'      OpenXAPFile = False
'      Exit Function
'   End If
   mXapCurrentRecord = 2
   mIsXapOpen = True
   OpenXAPFile = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.StuffPocketPcProperties. ", Err.Number, Err.Description
   OpenXAPFile = False
End Function


Public Function GetMappedCell(ColumnName As String, Optional SetRecordOffset As Integer)
   Dim StrCell As String
   On Error GoTo FunctionError
   If Not mIsXapOpen Then
      If Not OpenXAPFile() Then
         GetMappedCell = ""
         Exit Function
      End If
   End If
   '
   If SetRecordOffset > 0 Then
      mXapCurrentRecord = SetRecordOffset
   End If
   GetMappedCell = moExcel.Range(moXAP(mXapCurrentRecord).ConfigKeys(ColumnName).KeyValue)
   Exit Function
FunctionError:
   If Err.Number = 91 Then
      goSession.RaiseError "Error in mwExcelForm.SetMappedCell, invalid ColumnName/Record Offset: " _
        & ColumnName & "/" & mXapCurrentRecord
   Else
      goSession.RaiseError "General Error in mwExcelForm.GetMappedCell. ", Err.Number, Err.Description
   End If
   GetMappedCell = ""
End Function

Public Function SetMappedCell(ColumnName As String, NewValue As String, Optional SetRecordOffset As Integer) As Boolean
   On Error GoTo FunctionError
   If Not mIsXapOpen Then
      If Not OpenXAPFile() Then
         SetMappedCell = False
         Exit Function
      End If
   End If
   '
   If SetRecordOffset > 0 Then
      mXapCurrentRecord = SetRecordOffset
   End If
   moExcel.Range(moXAP(mXapCurrentRecord).ConfigKeys(ColumnName).KeyValue) = NewValue
   SetMappedCell = True
   Exit Function
FunctionError:
   If Err.Number = 91 Then
      goSession.RaiseError "Error in mwExcelForm.SetMappedCell, invalid ColumnName/Record Offset: " _
        & ColumnName & "/" & mXapCurrentRecord
   Else
      goSession.RaiseError "General Error in mwExcelForm.SetMappedCell. ", Err.Number, Err.Description
   End If
   SetMappedCell = False
End Function


'
' Stuff an Excel cell value into an ADO record, trapping data type errors as we go
'
Private Function StuffRsField(oKey As ConfigKey, Optional OffsetIncrement As Integer) As Boolean
   Dim strValue As String
   Dim IsBooleanNumber As Boolean
   Dim IsDateConvert As Boolean
   On Error GoTo FunctionError
   If Not IsNull(oKey.KeyValue) Then
      If oKey.KeyName <> "" And oKey.KeyValue <> "" Then
         If moExcel.Range(oKey.KeyValue) = "" Then
            StuffRsField = True
            Exit Function
         End If
         ' Multi-step error avoidance in different data types.
         If moRS.Fields(oKey.KeyName).Type = adNumeric Or moRS.Fields(oKey.KeyName).Type = adDouble Then
            '
            ' Numeric Fields...
            '
            If IsNumeric(moExcel.Range(oKey.KeyValue)) Then
               ' Oracle does not have boolean data type, so string values of "True" and "False" error out...
               If UCase(moExcel.Range(oKey.KeyValue)) = "TRUE" Or UCase(moExcel.Range(oKey.KeyValue)) = "FALSE" Then
                  IsBooleanNumber = True
               End If
               StuffRsField = True
            Else
               goSession.RaisePublicError "Error in mwSession.StuffRsField, Bad value in numeric field: " & oKey.KeyName & "/" & oKey.KeyValue, 999, "MW Defined Error"
               StuffRsField = False
            End If
         ElseIf moRS.Fields(oKey.KeyName).Type = adDate Or moRS.Fields(oKey.KeyName).Type = adDBTimeStamp Then
            '
            ' Date Fields
            '
            If IsDate(moExcel.Range(oKey.KeyValue)) Then
               StuffRsField = True
            Else
               '
               ' Look for time only field...
               '
               If IsNumeric(moExcel.Range(oKey.KeyValue)) Then
                  If moExcel.Range(oKey.KeyValue) < 1 Then
                     'ok we have time...
                     StuffRsField = True
                     IsDateConvert = True
                  Else
                     goSession.RaisePublicError "Error in mwSession.StuffRsField, Bad value in Date field: " & oKey.KeyName & "/" & oKey.KeyValue, 999, "MW Defined Error"
                     StuffRsField = False
                  End If
               Else
                  goSession.RaisePublicError "Error in mwSession.StuffRsField, Bad value in Date field: " & oKey.KeyName & "/" & oKey.KeyValue, 999, "MW Defined Error"
                  StuffRsField = False
               End If
                  
            End If
         ElseIf moRS.Fields(oKey.KeyName).Type = adSmallInt Then
            If IsNumeric(moExcel.Range(oKey.KeyValue)) Then
               If moExcel.Range(oKey.KeyValue) > 32000 Then
                  goSession.RaisePublicError "Error in mwSession.StuffRsField, Value too large in SmallInt field: " & _
                    oKey.KeyName & "/" & moExcel.Range(oKey.KeyValue), 999, "MW Defined Error"
                  StuffRsField = False
               Else
                  StuffRsField = True
               End If
            Else
               goSession.RaisePublicError "Error in mwSession.StuffRsField, Bad value in SmallInt field: " & oKey.KeyName & "/" & oKey.KeyValue, 999, "MW Defined Error"
               StuffRsField = False
            End If
         Else
            StuffRsField = True
         End If
         If StuffRsField Then
            If OffsetIncrement > 0 Then
               If IsDateConvert Then
                  strValue = CDate(moExcel.Range(GetLoopingRowCell(oKey.KeyValue, OffsetIncrement)))
               Else
                  strValue = moExcel.Range(GetLoopingRowCell(oKey.KeyValue, OffsetIncrement))
               End If
            Else
               If IsDateConvert Then
                  strValue = CDate(moExcel.Range(oKey.KeyValue))
               Else
                  strValue = moExcel.Range(oKey.KeyValue)
               End If
            End If
            If IsBooleanNumber Or moRS.Fields(oKey.KeyName).Type = adBoolean Then
               If UCase(strValue) = "TRUE" Then
                  strValue = "1"
               Else
                  strValue = "0"
               End If
            End If
            
            If Trim(strValue) <> "" Then
               moRS.Fields(oKey.KeyName) = strValue
            End If
            
            
         End If
      End If
   End If
   Exit Function
FunctionError:
   If Err.Number = 3265 Then
      goSession.RaisePublicError "Error in mwExcelForm.StuffRsField, Column Missing: " & oKey.KeyName & "/" & oKey.KeyValue, 999, "MW Defined Error"
      If moParent.WorkflowProperties.TolerateMissingColumns Then
         StuffRsField = True
      Else
         StuffRsField = False
      End If
   Else
      'Assume invalid field, cancel operation...
      goSession.RaisePublicError "General Error in mwExcelForm.StuffRsField: ", _
        Err.Number, Err.Description
      goSession.RaisePublicError "Failed Recordset Update on: " & oKey.KeyName & _
        " - " & oKey.KeyValue
      StuffRsField = False
   End If

End Function

'
' Stuff an Excel cell value into an ADO record, trapping data type errors as we go
'
Private Function StuffSimpleRsField(oKey As ConfigKey, Optional OffsetIncrement As Integer) As Boolean
   Dim strValue As String
   Dim IsDateConvert As Boolean
   Dim ThisCell As String
   Dim strFormat As String
   Dim loVariant As Variant
   Dim strTypeName As String
   Dim nDataType As Long
   
   
   On Error GoTo FunctionError
   If Not IsNull(oKey.KeyValue) Then
      If oKey.KeyName <> "" And oKey.KeyValue <> "" Then
         ThisCell = GetLoopingRowCell(oKey.KeyValue, OffsetIncrement)
         
         strValue = moExcel.Range(ThisCell)
         
         If Trim(strValue) <> "" Then
            mwFormDetailCnt = mwFormDetailCnt + 1
            
            moFormDataRs.AddNew
            moFormDataRs!ID = goSession.MakePK(mCurrentTable)
            moFormDataRs!mwFormHeaderKey = mwFormHeaderKey
            moFormDataRs!XAPSection = mCurrentSectionNumber
            moFormDataRs!DetailSeq = mwFormDetailCnt
'            moFormDataRs!LRRowOffset = OffsetIncrement
            moFormDataRs!FieldName = ThisCell
            moFormDataRs!mwcSitesKey = mwcSitesKey

            loVariant = moExcel.Range(ThisCell)
            strTypeName = TypeName(loVariant)
            
            Select Case strTypeName
               Case "Byte", "Integer", "Long"
                  nDataType = adInteger
                  moFormDataRs!IsIntegerData = True
                  moFormDataRs!IntegerData = CLng(moExcel.Range(ThisCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!IntegerData)
               Case "Single", "Double", "Currency", "Decimal"
                  nDataType = adDouble
                  moFormDataRs!IsDoubleData = True
                  moFormDataRs!DoubleData = CDbl(moExcel.Range(ThisCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!DoubleData)
               Case "Date"
                  nDataType = adDate
                  moFormDataRs!IsDateData = True
                  moFormDataRs!DateData = CDate(moExcel.Range(ThisCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!DateData)
               Case "Boolean"
                  nDataType = adBoolean
                  moFormDataRs!IsBooleanData = True
                  moFormDataRs!BooleanData = CBool(moExcel.Range(ThisCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!BooleanData)
               Case Else
                  nDataType = adChar
                  If Len(strValue) > 250 Then
                     moFormDataRs!IsMemoData = True
                     moFormDataRs!MemoData = CStr(moExcel.Range(ThisCell))
                  Else
                     moFormDataRs!IsStringData = True
                     moFormDataRs!StringData = CStr(moExcel.Range(ThisCell)) ' was ShortData
                  End If
            End Select

            moFormDataRs.Update
         End If
      End If
   End If
   Exit Function
FunctionError:
      'Assume invalid field, cancel operation...
      goSession.RaisePublicError "General Error in mwExcelForm.StuffSimpleRsField: ", _
        Err.Number, Err.Description
      goSession.RaisePublicError "Failed Recordset Update on: " & oKey.KeyName & _
        " - " & oKey.KeyValue
      StuffSimpleRsField = False

End Function

Public Function IsNullCellValue(NullTestValue As String) As Boolean
   IsNullCellValue = False
   If Trim(NullTestValue) = "" Then
      '
      ' Blank row encountered, we are finished
      '
      IsNullCellValue = True
   ElseIf IsDate(NullTestValue) Then
      If NullTestValue = "12:00:00 AM" Then
         '
         ' Blank Date Encountered
         '
         IsNullCellValue = True
      End If
   ElseIf IsNumeric(NullTestValue) Then
      If NullTestValue = 0 Then
         '
         ' Test value on numeric cannot be "0"
         '
         IsNullCellValue = True
      End If
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.IsNullCellValue. ", Err.Number, Err.Description
End Function


Private Function ReplaceVbCrLf(strVal As String) As String
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
   
   ReplaceVbCrLf = ByteArray
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.ReplaceVbCrLf.", Err.Number, Err.Description
   ReplaceVbCrLf = strVal
End Function


Friend Function PopulateForm(Optional Worksheet As Integer, Optional IsRepopulate As Boolean) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   '
   ' Phase 1 - Intrinsic Properties
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwwfFormPopulate where TemplateID='" & moParent.TemplateProperties.TemplateID & "'"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   'MsgBox strSQL & loRS.RecordCount
   If loRs.RecordCount > 0 Then
      '
      'Set the worksheet
      If Worksheet > 0 Then
         SetActiveWorksheet Worksheet
      Else
         SetActiveWorksheet 1
      End If
      UnprotectWorksheet
      '
      ' Populate...
      '
      Do While Not loRs.EOF
         If Not IsNull(loRs!VesselCode) Then
            SetCell loRs!VesselCode, goSession.Site.SiteID
         End If
         If Not IsNull(loRs!VesselName) Then
            SetCell loRs!VesselName, goSession.Site.SiteName
         End If
         If Not IsNull(loRs!Voyage_NO) Then
            SetCell loRs!Voyage_NO, goSession.Site.CurrentVoyage
         End If
         If Not IsNull(loRs!CurrentDate) Then
            SetCell loRs!CurrentDate, Format(Now(), "Short Date")
         End If
         If Not IsNull(loRs!CurrentMaster) Then
            SetCell loRs!CurrentMaster, goSession.Site.CurrentMaster
         End If
         If Not IsNull(loRs!CurrentCheng) Then
            SetCell loRs!CurrentCheng, goSession.Site.GetExtendedProperty("CurrentChiefEngineer")
         End If
         If Not IsNull(loRs!CurrentPort) Then
            SetCell loRs!CurrentPort, goSession.Site.CurrentPort
         End If
         
         'MOL-143 Adding Ship IMO Number
         If Not IsNull(loRs!MtmlShipImoNumber) Then
            SetCell loRs!MtmlShipImoNumber, goSession.Site.ShipImoNumber
         End If
         
         '
         ' Populate Logo ?
         '
         If loRs!IsInsertLogo Then
            If Not IsRepopulate Then
               If Not (IsNull(loRs!LogoCellBookmark) Or goSession.Site.GetExtendedProperty("mwcCommercialOperatorKey") = "") Then
                  'MsgBox "about to insert logo"
                  moParent.InsertLogo loRs!LogoCellBookmark
               End If
            End If
         End If
         ' next one...
         loRs.MoveNext
      Loop
   End If
   goSession.CloseRecordset loRs
   
   
   '
   ' Close Up Shop
   '
   ProtectWorksheet
   SaveWorkbook
   PopulateForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.PopulateForm. ", Err.Number, Err.Description
   ProtectWorksheet
   SaveWorkbook
   PopulateForm = False
End Function

'
' Cheap Dog Tricks required in following function because Excel does not properly handle
' Regional settings and decimal separators
'
Friend Function PopulateEventForm(ByRef oRsColumns As Recordset, ByRef oRsData As Recordset) As Boolean
   Dim strField As String
   On Error GoTo FunctionError
   UnprotectAllWorksheets
   Do While Not oRsColumns.EOF
      'store to string - eliminate pointer...
      strField = oRsColumns!ColumnName
      Debug.Print BlankNull(oRsColumns!CellFieldReference)
      If Not IsNull(oRsData.Fields(strField).value) Then
         If IsNull(oRsColumns!FormatString) Then
            If oRsData.Fields(strField).Type < adVarChar Then
               If oRsData.Fields(strField).Type = adDouble Then
                  SetCell oRsColumns!CellFieldReference, CStr(oRsData.Fields(strField).value), oRsColumns!SheetDocOffset
               Else
                  SetCell oRsColumns!CellFieldReference, str(oRsData.Fields(strField).value), oRsColumns!SheetDocOffset
               End If
            Else
               SetCell oRsColumns!CellFieldReference, oRsData.Fields(strField).value, oRsColumns!SheetDocOffset
            End If
         Else
            SetCell oRsColumns!CellFieldReference, Format(oRsData.Fields(strField).value, oRsColumns!FormatString), oRsColumns!SheetDocOffset
         End If
      End If
      oRsColumns.MoveNext
   Loop
   PopulateEventForm = True
   ProtectAllWorksheets
   SaveWorkbook
   Exit Function
FunctionError:
   If Err.Number = 3265 Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwExcelForm.PopulateEventForm, column reference error: " & strField
      Resume Next
   Else
      goSession.RaiseError "General Error in mwExcelForm.PopulateEventForm. ", Err.Number, Err.Description
   End If
   PopulateEventForm = False
End Function


Public Function PrintForm() As Boolean
   On Error GoTo FunctionError
   moWorkbook.PrintOut
   PrintForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.PrintForm. ", Err.Number, Err.Description
   PrintForm = False

End Function

Public Function PreviewForm2(FileName As String) As Boolean
   Dim loExcel As Object
   Dim loWorkbook As Object
   Dim fso As FileSystemObject
   
   On Error GoTo FunctionError
   
   
   Set fso = New FileSystemObject
   
   If fso.FileExists(FileName) Then
      Set loExcel = CreateObject("Excel.Application")
      Set loWorkbook = loExcel.Workbooks.Open(FileName)
      loWorkbook.Application.Visible = True
      loWorkbook.PrintPreview
   End If
   
   KillObject fso
   PreviewForm2 = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.PreviewForm2. ", Err.Number, Err.Description
   KillObject fso
   PreviewForm2 = False
End Function

Public Function ResetRevisedFlag() As Boolean
   Dim loXap As ConfigGroups
   Dim loXapRec As ConfigKeys
   Dim i As Integer
   Dim strOperation As String
   Dim strScriptFile As String
   On Error GoTo FunctionError
   
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ResetRevisedFlag, ExcelForm Not Open."
      ResetRevisedFlag = False
      Exit Function
   End If
   '
   ' it must be an ExcelForm...
   '
   If moParent.FormType <> mwft_EXCEL And moParent.FormType <> mwft_EXCEL_XLSX And moParent.FormType <> mwft_EXCEL_XLSM Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ResetRevisedFlag, Wrong Form Type."
      ResetRevisedFlag = False
      Exit Function
   End If
   '
   ' Open XAP File...
   '
   strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
   Set loXap = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXap Is Nothing Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ResetRevisedFlag, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      Exit Function
   End If
   '
   ' XAP File must be valid...
   '
   If loXap.Count = 0 Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.ResetRevisedFlag, XAP File Corrupt: "
      Exit Function
   End If
   '
   ' Parse, looking for LR types...
   '
   For i = 2 To loXap.Count
      DoEvents
      Set loXapRec = loXap(i).ConfigKeys
      ' Need an Operation
      strOperation = UCase(loXapRec.GetKeyValue("Operation"))
      If Trim(strOperation) = "LR" Then
         'call the reset function
         If Not ResetLRRevisedFlag(loXapRec) Then
            goSession.KillObject loXap
            goSession.KillObject loXapRec
            ResetRevisedFlag = False
            Exit Function
         End If
      End If
   Next i
   goSession.KillObject loXap
   goSession.KillObject loXapRec
   ResetRevisedFlag = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwExcelForm.ResetRevisedFlag", Err.Number, Err.Description
   ResetRevisedFlag = False
  End Function

Private Function ResetLRRevisedFlag(loXapRec As ConfigKeys) As Boolean
   Dim IsMoreRows As Boolean
   Dim iCurrentOffset As Integer
   Dim iWorksheet As Integer
   Dim strCellOffset As String
   Dim strCellValue As String
   On Error GoTo FunctionError
   iWorksheet = loXapRec.GetKeyValue("Worksheet")
   SetActiveWorksheet iWorksheet
   UnprotectWorksheet iWorksheet
   IsMoreRows = True
   iCurrentOffset = 0
   '
   
   Do While IsMoreRows
      '
      ' NullTestCell skip out...
      '
      strCellOffset = GetLoopingRowCell(loXapRec.GetKeyValue("NullTestCell"), iCurrentOffset)
      strCellValue = GetCell(strCellOffset, iWorksheet)
      If IsNullCellValue(strCellValue) Then
         IsMoreRows = False
      Else
         strCellOffset = GetLoopingRowCell(loXapRec.GetKeyValue("SkipRowTestCell"), iCurrentOffset)
         SetCell strCellOffset, "", iWorksheet
      End If
      iCurrentOffset = iCurrentOffset + loXapRec.GetKeyValue("OffsetIncrement")
      If iCurrentOffset > loXapRec.GetKeyValue("MaxRowOffset") Then
         IsMoreRows = False
      End If
   Loop
   ProtectWorksheet iWorksheet
   SaveWorkbook
   ResetLRRevisedFlag = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwExcelForm.ResetLRRevisedFlag", Err.Number, Err.Description
   ResetLRRevisedFlag = False
   ProtectWorksheet iWorksheet
  End Function

Public Function LoadPicture(FileName As String, CellReference As String) As Boolean
   Dim loSheet As Object
   'Dim loSheet As Worksheet
   Dim i() As String
   If CellReference <> "" Then
      i = Split(CellReference, ",")
   End If
   On Error GoTo FunctionError
   Set loSheet = moWorkbook.ActiveSheet
   'If loSheet.Shapes.Count < 1 Then
   '
   ' Insert shape
   '
   If UBound(i) <> 3 Then
      loSheet.Shapes.AddPicture FileName, True, True, 1, 1, 50, 50
   Else
      loSheet.Shapes.AddPicture FileName, True, True, CInt(i(0)), CInt(i(1)), CInt(i(2)), CInt(i(3))
   End If
   'End If
   Set loSheet = Nothing
   
   LoadPicture = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.mwExcelForm.LoadPicture", Err.Number, Err.Description
   LoadPicture = False
End Function

' GetFormDetailData - Read data from the form and store it in the database

Public Function GetFormDetailData(FormHeaderKey As Long, SitesKey As Long) As Boolean
'   Dim loField As FormField
   Dim loField As Object
   Dim loCnt As Long
   Dim sSQL As String
   Dim strOperation As String
   Dim loRec As ConfigKeys
   Dim oKey As ConfigKey
   Dim i As Integer
   Dim loXapFile As ConfigGroups
   Dim strScriptFile As String
   Dim IsError As Boolean
   Dim IsAnyErrors As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   
   mwcSitesKey = SitesKey
   
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, ExcelForm Not Open."
      GetFormDetailData = False
      Exit Function
   End If
   '
   ' it must be an ExcelForm...
   '
   If moParent.FormType <> mwft_EXCEL And moParent.FormType <> mwft_EXCEL_XLSX And moParent.FormType <> mwft_EXCEL_XLSM Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, Wrong Form Type."
      GetFormDetailData = False
      Exit Function
   End If
   
   If FormHeaderKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, illegal FormHeaderKey " & FormHeaderKey
      GetFormDetailData = False
      Exit Function
   End If
   
   If goSession.Workflow.TemplatePlacesID <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, No Template Source PlaceID defined for the " & goSession.Workflow.WorkflowID & " Workflow Profile."
      GetFormDetailData = False
      Exit Function
   End If
   
   If Len(moParent.GetExtendedProperty("MapFName")) <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, No XAP File Name defined for the " & moParent.TemplateProperties.TemplateID & " Form Template."
      GetFormDetailData = False
      Exit Function
   End If
   
   If Len(moParent.GetExtendedProperty("MapFExt")) <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, No XAP File Extension defined for the " & moParent.TemplateProperties.TemplateID & " Form Template."
      GetFormDetailData = False
      Exit Function
   End If
   
   mwFormHeaderKey = FormHeaderKey
   
   If mwcSitesKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, illegal mwcSitesKey " & mwcSitesKey
      GetFormDetailData = False
      Exit Function
   End If
   mwcSitesKey = SitesKey
   
   Set moFormHeaderRs = New Recordset
   moFormHeaderRs.CursorLocation = adUseClient
   sSQL = "SELECT * FROM mwFormHeader WHERE ID= " & mwFormHeaderKey
   moFormHeaderRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not IsRecordLoaded(moFormHeaderRs) Then
      goSession.RaiseError "Error in mwSession.mwExcelForm.GetFormDetailData, could not find an mwFormHeader record for mwFormHeaderKey: " & mwFormHeaderKey
      GetFormDetailData = False
      CloseRecordset moFormHeaderRs
      KillObject fso
      Exit Function
   End If
   
   '
   ' and must be an Open Map XAP File...
   '
   Set fso = New FileSystemObject
   
   strScriptFile = GetTemplateScriptFileName(mwFormHeaderKey)
   
   If Not fso.FileExists(strScriptFile) Then
      If fso.FileExists(moParent.WorkflowProperties.MapScriptFile) Then
         strScriptFile = moParent.WorkflowProperties.MapScriptFile
      Else
         strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
      End If
   End If
   
   Set loXapFile = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXapFile Is Nothing Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      GetFormDetailData = False
      KillObject fso
      Exit Function
   End If
   '
   ' XAP File must match...
   '
   If loXapFile.Count = 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData, XAP File Corrupt: "
      GetFormDetailData = False
      KillObject fso
      Exit Function
   End If
   '
   ' Execute Submit Event for the Form
   '
   '
   ' Run external VB DLL - if it exists... OnSubmit form event...
   '
   If Not moParent.RunFormDll("OnSubmit") Then
      goSession.LogIt mwl_User_Defined, mwl_Information, "Submit has been cancelled by the OnSubmit DLL Code."
      GetFormDetailData = False
      Exit Function
   End If
   '
   ' OK, Create a recordset object to use...
   '
   '
   IsError = False
   IsAnyErrors = False
   goCon.BeginTrans
      
      ' First we have to delete any existing records for this form header
      
      For i = loXapFile.Count To 2 Step -1
      
      
         CloseRecordset moFormDataRs
         
         Set moFormDataRs = New Recordset
         moFormDataRs.CursorLocation = adUseClient
         
         Set loRec = loXapFile(i).ConfigKeys
         
         strOperation = loRec.GetKeyValue("Operation")
      
         If strOperation = "A" Or strOperation = "LR" Then
            mCurrentTable = loRec.GetKeyValue("table")
         
         ElseIf strOperation = "SA" Or strOperation = "SLR" Then
            mCurrentTable = "mwFormDetail"
         Else
            mCurrentTable = ""
         End If
      
         If mCurrentTable <> "" Then
            sSQL = "select * from " & mCurrentTable & " where mwFormHeaderKey= " & mwFormHeaderKey
            moFormDataRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      
            If IsRecordLoaded(moFormDataRs) Then
               Do While Not moFormDataRs.EOF
                  moFormDataRs.Delete
                  moFormDataRs.MoveNext
               Loop
            End If
         End If
      Next i
      
      '
      '
      ' For each record (except first moRecord, which is the Header) in the XAP file....
      ' 1-based collection
      '
      CloseRecordset moFormDataRs
      
      Set moFormDataRs = New Recordset
      moFormDataRs.CursorLocation = adUseClient
      
      For i = 2 To loXapFile.Count
         DoEvents
         Set loRec = loXapFile(i).ConfigKeys
         
         mCurrentSectionNumber = GetSectionNumber(loXapFile(i).ConfigGroup)
         ' Need an Operation
         strOperation = loRec.GetKeyValue("Operation")
         If strOperation = "A" Then
            IsError = Not GetFormDataRsTypeA(loRec, mwFormHeaderKey)
            If IsError = False Then
               If Len(mCurrentParentTable) > 0 And mCurrentParentRecordID > 0 Then
                  moFormHeaderRs!DataTableName = mCurrentParentTable
                  moFormHeaderRs!DataTableID = mCurrentParentRecordID
                  moFormHeaderRs.Update
               End If
            End If
         ElseIf strOperation = "SA" Then
            IsError = Not GetFormDataRsTypeSA(loRec, mwFormHeaderKey)
         ElseIf strOperation = "LR" Then
            IsError = Not GetFormDataRsTypeLR(loRec, mwFormHeaderKey)
         ElseIf strOperation = "SLR" Then
            IsError = Not GetFormDataRsTypeSLR(loRec, mwFormHeaderKey)
         Else
            goSession.RaiseError "Error in mwExcelForm.GetFormDetailData, invalid Operation Type: " & strOperation
            IsError = True
         End If
         '
         ' Close the recordset for reuse...
         '
         If moFormDataRs.State = adStateOpen Then
            moFormDataRs.Close
         End If
         '
         ' Error Check...
         '
         If IsError Then
            If moParent.WorkflowProperties.TolerateOtherErrors Then
               IsError = False
               IsAnyErrors = True
            Else
               goSession.RaisePublicError "Error in Record: " & str(i) & ". Tolerate Errors is set to False, cancelling process..."
               GetFormDetailData = False
               goCon.RollbackTrans
               KillObject fso
               Exit Function
            End If
         End If
      Next i
      '
      ' We be done...
      '
      
      If IsAnyErrors Then
'         goSession.RaisePublicError "Partially Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         GetFormDetailData = False
      Else
'         goSession.RaisePublicError "Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         GetFormDetailData = True
      End If
   goCon.CommitTrans
   KillObject fso
   Exit Function
FunctionError:
   If moParent.WorkflowProperties.TolerateDuplicateRecords And Err.Number = 1 Then
      'Future: Need to raise an event here to inform that a duplicate encountered...
      goSession.RaisePublicError "Error in GetFormDetailData: " & Err.Description
      Resume Next
   Else
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDetailData.", Err.Number, Err.Description
      GetFormDetailData = False
   End If
   On Error Resume Next
   goCon.RollbackTrans
   KillObject fso
End Function


Private Function GetFormDataRsTypeA(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA, Table name missing from Record. "
      GetFormDataRsTypeA = False
      Exit Function
   End If
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      GetFormDataRsTypeA = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "select * from " & mCurrentTable & " where ID=-1"
   'moFormDataRs.CursorLocation = adUseClient
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockOptimistic
'   moFormDataRs.Open mCurrentTable, goCon, adOpenStatic, adLockOptimistic, adCmdTable
'   moFormDataRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      GetFormDataRsTypeA = False
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   
   '
   ' Skip record Test
   '
   IsSkipRecord = False
   strNullTestCell = Trim(UCase(xaprec.GetKeyValue("NullTestCell")))
   If strNullTestCell <> "" And strNullTestCell <> "N/A" Then
      If IsNullCellValue(moExcel.Range(strNullTestCell)) Then
         IsSkipRecord = True
      End If
   End If
   If Not IsSkipRecord Then
   
      moFormDataRs.AddNew
      'If goSession.User.DebugMode Then MsgBox "after addnew"
      If Err Then
         GetFormDataRsTypeA = False
         goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA," & _
           " Unable to add new record, Table: " & mCurrentTable
         'goSession.CloseRecordset moFormDataRs
         Exit Function
      End If
      nKey = goSession.MakePK(mCurrentTable, goCon)
      'moFormDataRs.Fields(0).value = nKey
      moFormDataRs!ID = nKey
      moFormDataRs!mwcSitesKey = mwcSitesKey
      moFormDataRs!mwFormHeaderKey = FormHeaderKey
      moFormDataRs!XAPSection = mCurrentSectionNumber
      
      '
      ' If Child, save key to the parent record...
      '
      If xaprec.GetKeyValue("IsChildKeyName") <> "" And _
           UCase(xaprec.GetKeyValue("IsChildKeyName")) <> "N/A" Then
         moFormDataRs.Fields(xaprec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
      End If
'
' 3/2007 ms - adBoolean fields do not work if .UPDATE executed...
' -2147217864Row cannot be located for updating. Some values may have been changed since it was last read.|
'
'      If moFormDataRs.Fields.Count < 75 Then
'         moFormDataRs.Update
'         If err Then
'            goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA: ", err.Number, err.Description
'            goSession.RaisePublicError "Failed to Update Table during Call to MakePK: " & mCurrentTable
'            CloseRecordset moFormDataRs
'            GetFormDataRsTypeA = False
'            Exit Function
'         End If
'      End If
      If GetFormDataRSFields(xaprec) Then
         If goSession.User.DebugMode Then MsgBox "about to .update"
         moFormDataRs.Update
         If Err Then
            goSession.RaiseError "Error in mwSession.mwExcelForm.GetFormDataRsTypeA during .Update. ", _
              Err.Number, Err.Description
            GetFormDataRsTypeA = False
            moFormDataRs.CancelUpdate
            Exit Function
         End If
         '
         ' If Parent, save key to the parent record...
         '
         If xaprec.GetKeyValue("IsParentKeyName") <> "" Then
            mCurrentParentRecordID = moFormDataRs.Fields(xaprec.GetKeyValue("IsParentKeyName")).value
            mCurrentParentTable = mCurrentTable
         End If
      Else
         moFormDataRs.CancelUpdate
         goSession.RaisePublicError "Error in mwDataForm.GetFormDataRsTypeA. CancelUpdate on Form " & moParent.FormProperties.FormID
         GetFormDataRsTypeA = False
         Exit Function
      End If
   End If
   GetFormDataRsTypeA = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.GetFormDataRsTypeA. ", Err.Number, Err.Description
   GetFormDataRsTypeA = False
End Function

Private Function GetFormDataRsTypeSA(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   
   mCurrentTable = "mwFormDetail"
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeSA, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      GetFormDataRsTypeSA = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "select * from " & mCurrentTable & " where ID=-1"
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockOptimistic
   
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeSA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      GetFormDataRsTypeSA = False
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   
   '
   ' Skip record Test
   '
   IsSkipRecord = False
   strNullTestCell = Trim(UCase(xaprec.GetKeyValue("NullTestCell")))
   If strNullTestCell <> "" And strNullTestCell <> "N/A" Then
      If IsNullCellValue(moExcel.Range(strNullTestCell)) Then
         IsSkipRecord = True
      End If
   End If
   If Not IsSkipRecord Then
      GetFormDataRsTypeSA = GetFormDataSimpleRSFields(xaprec)
   Else
      GetFormDataRsTypeSA = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.GetFormDataRsTypeSA. ", Err.Number, Err.Description
   GetFormDataRsTypeSA = False
End Function


Private Function GetFormDataRsTypeLR(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      GetFormDataRsTypeLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.GetFormDataRsTypeLR, Worksheet Parameter Missing."
      GetFormDataRsTypeLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   ' First row a no brainer
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   '
   ' Open the recordset...
   '
   
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      GetFormDataRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR, Table name missing from Record. "
      Exit Function
   End If
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      GetFormDataRsTypeLR = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   '
   ' Open the recordset
   '
'   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
      sSQL = "Select * from " & mCurrentTable & " where ID=-1"
      moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockOptimistic
'   Else
'      moFormDataRs.Open mCurrentTable, moCon, adOpenStatic, adLockOptimistic, adCmdTable
'   End If
   If Err Then
      GetFormDataRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   iTotalRows = 0
   IsMoreRows = True
   iCurrentOffset = 0
'   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))
   iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
   If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
      iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
   End If
   iStartingRowOffset = GetBaselineRowOffset(xaprec(mLrp.TotalNonLoopingRowFields + iTotalLrParameters).KeyValue)
   Do While IsMoreRows
      '
      ' Null Value Terminate Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, iCurrentOffset))) Then
'            goSession.RaisePublicError "NullCellTest: Finished adding rows, total count = " & iTotalRows
            Exit Do
         End If
      End If
      '
      ' Skip Row Check
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))) Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
         iTotalRows = iTotalRows + 1
         '
         ' Add new ADO record
         '
         On Error Resume Next
         moFormDataRs.AddNew
         'If goSession.User.DebugMode Then MsgBox "after addnew"
         If Err Then
            GetFormDataRsTypeLR = False
            goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR," & _
              " Unable to add new record, Table: " & mCurrentTable
            'goSession.CloseRecordset moFormDataRs
            Exit Function
         End If
'         If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
            moFormDataRs!ID = goSession.MakePK(mCurrentTable, moCon)
            moFormDataRs!mwcSitesKey = mwcSitesKey
            moFormDataRs!mwFormHeaderKey = FormHeaderKey
            moFormDataRs!XAPSection = mCurrentSectionNumber
            moFormDataRs!LRRowOffset = iCurrentOffset
            
            moFormDataRs.Update
            If Err Then
               goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR: ", Err.Number, Err.Description
               goSession.RaisePublicError "Failed to Update Table during Call to MakePK: " & mCurrentTable
               goSession.CloseRecordset moFormDataRs
               GetFormDataRsTypeLR = False
               Exit Function
            End If
'         End If
         '
         ' If Child, save key to the parent record...
         '
         On Error GoTo FunctionError
         If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
            moFormDataRs.Fields(xaprec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
            iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
         End If
         
         '
         ' OK, add the non repeating values first...
         '
         If mLrp.TotalNonLoopingRowFields > 0 Then
            For i = iTotalLrParameters To (mLrp.TotalNonLoopingRowFields + (iTotalLrParameters - 1))
               If Not GetRsField(xaprec(i)) Then
                  ' Terminate
                  moFormDataRs.CancelUpdate
                  GetFormDataRsTypeLR = False
                  Exit Function
               End If
            Next i
         End If
         '
         ' Then, current row variables
         '
         For i = (mLrp.TotalNonLoopingRowFields + iTotalLrParameters) To (xaprec.Count)
            If Not GetRsField(xaprec(i), iCurrentOffset) Then
               ' Terminate
               moFormDataRs.CancelUpdate
               GetFormDataRsTypeLR = False
               Exit Function
            End If
         Next i
         '
         ' Update record...
         '
         On Error Resume Next
         If goSession.User.DebugMode Then MsgBox "about to .update"
         moFormDataRs.Update
         If Err Then
            goSession.RaiseError "Error in mwSession.mwExcelForm.GetFormDataRsTypeLR during .Update. ", _
              Err.Number, Err.Description
            GetFormDataRsTypeLR = False
            moFormDataRs.CancelUpdate
            Exit Function
         End If
      End If
      '
      ' Another row to process ?
      '
      iCurrentOffset = iCurrentOffset + mLrp.OffsetIncrement
      If (iStartingRowOffset + iCurrentOffset) > mLrp.MaxRowOffset Then
         Exit Do
      End If
      IsSkipRow = False
   Loop
   GetFormDataRsTypeLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.GetFormDataRsTypeLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.GetFormDataRsTypeLR. ", Err.Number, Err.Description
   End If
   GetFormDataRsTypeLR = False
End Function
Private Function GetFormDataRsTypeSLR(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      GetFormDataRsTypeSLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.GetFormDataRsTypeSLR, Worksheet Parameter Missing."
      GetFormDataRsTypeSLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   ' First row a no brainer
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   '
   ' Open the recordset...
   '
   
   mCurrentTable = "mwFormDetail"
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeSLR, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      GetFormDataRsTypeSLR = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   '
   ' Open the recordset
   '
'   If moParent.GetExtendedProperty("IsMakePkOnProcess") = "True" Then
      sSQL = "Select * from " & mCurrentTable & " where ID=-1"
      moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockOptimistic
'   Else
'      moFormDataRs.Open mCurrentTable, moCon, adOpenStatic, adLockOptimistic, adCmdTable
'   End If
   If Err Then
      GetFormDataRsTypeSLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.GetFormDataRsTypeSLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   iTotalRows = 0
   IsMoreRows = True
   iCurrentOffset = 0
'   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))
   iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
   If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
      iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
   End If
   iStartingRowOffset = GetBaselineRowOffset(xaprec(mLrp.TotalNonLoopingRowFields + iTotalLrParameters).KeyValue)
   
   Do While IsMoreRows
      '
      ' Null Value Terminate Check
      '
      If UCase(mLrp.NullTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.NullTestCell, iCurrentOffset))) Then
'            goSession.RaisePublicError "NullCellTest: Finished adding rows, total count = " & iTotalRows
            Exit Do
         End If
      End If
      '
      ' Skip Row Check
      '
      If UCase(mLrp.SkipRowTestCell) <> "N/A" Then
         If IsNullCellValue(moExcel.Range(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))) Then
            IsSkipRow = True
         End If
      End If
      If Not IsSkipRow Then
         iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
         iTotalRows = iTotalRows + 1
         '
         ' Add new ADO record
         '
         '
         ' If Child, save key to the parent record...
         '
         On Error GoTo FunctionError
         If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
'            moFormDataRs.Fields(xaprec.GetKeyValue("IsChildKeyName")).value = mCurrentParentRecordID
            iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
         End If
         
         '
         ' OK, add the non repeating values first...
         '
         If mLrp.TotalNonLoopingRowFields > 0 Then
            For i = iTotalLrParameters To (mLrp.TotalNonLoopingRowFields + (iTotalLrParameters - 1))
               If Not GetSimpleRsField(xaprec(i)) Then
                  ' Terminate
                  GetFormDataRsTypeSLR = False
                  Exit Function
               End If
            Next i
         End If
         '
         ' Then, current row variables
         '
         For i = (mLrp.TotalNonLoopingRowFields + iTotalLrParameters) To (xaprec.Count)
            If Not GetSimpleRsField(xaprec(i), iCurrentOffset) Then
               ' Terminate
'               moFormDataRs.CancelUpdate
               GetFormDataRsTypeSLR = False
               Exit Function
            End If
         Next i
         '
         ' Update record...
         '
      End If
      '
      ' Another row to process ?
      '
      iCurrentOffset = iCurrentOffset + mLrp.OffsetIncrement
      If (iStartingRowOffset + iCurrentOffset) > mLrp.MaxRowOffset Then
         Exit Do
      End If
      IsSkipRow = False
   Loop
   GetFormDataRsTypeSLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.GetFormDataRsTypeSLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.GetFormDataRsTypeSLR. ", Err.Number, Err.Description
   End If
   GetFormDataRsTypeSLR = False
End Function

Private Function GetFormDataRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   DoEvents
   If Not IsNumeric(oRec.GetKeyValue("WORKSHEET")) Then
      goSession.RaisePublicError "Error in Excel Script, non-numeric WORKSHEET specified (using 1)."
      moExcel.Worksheets(1).Activate
   Else
      moExcel.Worksheets(Int(oRec.GetKeyValue("WORKSHEET"))).Activate
   End If
   For Each loKey In oRec
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         
         Case Is = "OFFSETINCREMENT"
         Case Is = "MAXROWOFFSET"
         Case Is = "TOTALNONLOOPINGROWFIELDS"
   
         Case Else
            If Not GetRsField(loKey) Then
               ' terminate...
               GetFormDataRSFields = False
               Exit Function
            End If
      End Select
   Next loKey
   GetFormDataRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.GetFormDataRSFields: ", Err.Number, Err.Description
   GetFormDataRSFields = False
End Function
Private Function GetFormDataSimpleRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   DoEvents
   If Not IsNumeric(oRec.GetKeyValue("WORKSHEET")) Then
      goSession.RaisePublicError "Error in Excel Script, non-numeric WORKSHEET specified (using 1)."
      moExcel.Worksheets(1).Activate
   Else
      moExcel.Worksheets(Int(oRec.GetKeyValue("WORKSHEET"))).Activate
   End If
   For Each loKey In oRec
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         
         Case Is = "OFFSETINCREMENT"
         Case Is = "MAXROWOFFSET"
         Case Is = "TOTALNONLOOPINGROWFIELDS"
   
         Case Else
            If Not GetSimpleRsField(loKey) Then
               ' terminate...
               GetFormDataSimpleRSFields = False
               Exit Function
            End If
      End Select
   Next loKey
   GetFormDataSimpleRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.GetFormDataSimpleRSFields: ", Err.Number, Err.Description
   GetFormDataSimpleRSFields = False
End Function

'
' Stuff an Excel cell value into an ADO record, trapping data type errors as we go
'
Private Function GetRsField(oKey As ConfigKey, Optional OffsetIncrement As Integer) As Boolean
   Dim strValue As String
   Dim CurrentCell As String
   Dim IsDateConvert As Boolean
   Dim nLongValue As Long
   Dim nDoubleValue As Double
   
   On Error GoTo FunctionError
   
   If Not IsNull(oKey.KeyValue) Then
      If oKey.KeyName <> "" And oKey.KeyValue <> "" Then
         
         CurrentCell = GetLoopingRowCell(oKey.KeyValue, OffsetIncrement)
   
         If IsBlank(moExcel.Range(CurrentCell)) Then
            GetRsField = True
            Exit Function
         End If
         
         Select Case (moFormDataRs.Fields(oKey.KeyName).Type)
         ' Multi-step error avoidance in different data types.
            Case adSingle, adDouble, adDecimal, adNumeric
            
               ' Oracle stores booleans in a Numeric(1) type.
               If (moFormDataRs.Fields(oKey.KeyName).Type = adNumeric And moFormDataRs.Fields(oKey.KeyName).Precision = 1 And goSession.IsOracle) Then
                  If TranslateBoolean(moExcel.Range(CurrentCell)) = "1" Then
                     moFormDataRs.Fields(oKey.KeyName).value = 1
                  Else
                     moFormDataRs.Fields(oKey.KeyName).value = 0
                  End If
                  GetRsField = True
               Else
                  '
                  ' Numeric Fields...
                  '
                  If IsNumeric(moExcel.Range(CurrentCell)) Then
                     
                     moFormDataRs.Fields(oKey.KeyName) = CDbl(moExcel.Range(CurrentCell))
                     GetRsField = True
                  Else
                     goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in numeric field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                     GetRsField = False
                  End If
               End If
            Case adDate, adDBTimeStamp
               '
               ' Date Fields
               '
               If IsDate(moExcel.Range(CurrentCell)) Then
                  moFormDataRs.Fields(oKey.KeyName) = CDate(moExcel.Range(CurrentCell))
                  GetRsField = True
               Else
                  '
                  ' Look for time only field...
                  '
                  If IsNumeric(moExcel.Range(CurrentCell)) Then
                     If moExcel.Range(CurrentCell) < 1 Then
                        'ok we have time...
                        moFormDataRs.Fields(oKey.KeyName) = CDate(moExcel.Range(CurrentCell))
                        GetRsField = True
                        IsDateConvert = True
                     Else
                        goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in Date field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                        GetRsField = False
                     End If
                  Else
                     goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in Date field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                     GetRsField = False
                  End If
                     
               End If
            Case adSmallInt
               If IsNumeric(moExcel.Range(CurrentCell)) Then
                  nLongValue = moExcel.Range(CurrentCell)
                  If nLongValue > 32000 Or nLongValue < -32000 Then
                     goSession.RaisePublicError "Error in mwSession.GetRsField, Value too large in SmallInt field: " & _
                       oKey.KeyName & "/" & moExcel.Range(CurrentCell), 999, "MW Defined Error"
                     GetRsField = False
                  Else
                     moFormDataRs.Fields(oKey.KeyName) = CLng(moExcel.Range(CurrentCell))
                     GetRsField = True
                  End If
               Else
                  goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in SmallInt field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                  GetRsField = False
               End If
            Case adInteger
               If IsNumeric(moExcel.Range(CurrentCell)) Then
                  nDoubleValue = moExcel.Range(CurrentCell)
                  If nDoubleValue > 2147483647 Or nDoubleValue < -2147483647 Then
                     goSession.RaisePublicError "Error in mwSession.GetRsField, Value too large in LongInt field: " & _
                       oKey.KeyName & "/" & moExcel.Range(CurrentCell), 999, "MW Defined Error"
                     GetRsField = False
                  Else
                     moFormDataRs.Fields(oKey.KeyName) = CLng(moExcel.Range(CurrentCell))
                     GetRsField = True
                  End If
               Else
                  goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in LongInt field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                  GetRsField = False
               End If
            Case adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt
               If IsNumeric(moExcel.Range(CurrentCell)) Then
                  moFormDataRs.Fields(oKey.KeyName) = CLng(moExcel.Range(CurrentCell))
                  GetRsField = True
               Else
                  goSession.RaisePublicError "Error in mwSession.GetRsField, Bad value in Numeric field: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
                  GetRsField = False
               End If
            Case adBoolean
               If TranslateBoolean(moExcel.Range(CurrentCell)) = "1" Then
                  moFormDataRs.Fields(oKey.KeyName).value = 1
               Else
                  moFormDataRs.Fields(oKey.KeyName).value = 0
               End If
               GetRsField = True
            Case Else
               strValue = moExcel.Range(CurrentCell)
               If Trim(strValue) <> "" Then
                  moFormDataRs.Fields(oKey.KeyName) = strValue
               End If
               GetRsField = True
         End Select
         
      End If
   End If
   Exit Function
FunctionError:
'Resume Next
   If Err.Number = 3265 Then
      goSession.RaisePublicError "Error in mwExcelForm.GetRsField, Column Missing: " & oKey.KeyName & "/" & CurrentCell, 999, "MW Defined Error"
      If moParent.WorkflowProperties.TolerateMissingColumns Then
         GetRsField = True
      Else
         GetRsField = False
      End If
   Else
      'Assume invalid field, cancel operation...
      goSession.RaisePublicError "General Error in mwExcelForm.GetRsField: ", Err.Number, Err.Description
      goSession.RaisePublicError "Failed Recordset Update on: " & oKey.KeyName & " - " & CurrentCell
      GetRsField = False
   End If

End Function

Private Function GetSimpleRsField(oKey As ConfigKey, Optional OffsetIncrement As Integer) As Boolean
   Dim strValue As String
   Dim CurrentCell As String
   Dim IsDateConvert As Boolean
   Dim strFormat As String
   Dim strLocalFormat As String
   Dim loVariant As Variant
   Dim strTypeName As String
   Dim nDataType As Long
   
   On Error GoTo FunctionError
   
   If Not IsNull(oKey.KeyValue) Then
      If oKey.KeyName <> "" And oKey.KeyValue <> "" Then
         
         CurrentCell = GetLoopingRowCell(oKey.KeyValue, OffsetIncrement)
   
         strValue = moExcel.Range(CurrentCell)
         strFormat = moExcel.Range(CurrentCell).NumberFormat
         strLocalFormat = moExcel.Range(CurrentCell).NumberFormatLocal
         
         
         If Trim(strValue) <> "" Then
         
            loVariant = moExcel.Range(CurrentCell)
            strTypeName = TypeName(loVariant)
         
            mwFormDetailCnt = mwFormDetailCnt + 1
            
            moFormDataRs.AddNew
            moFormDataRs!ID = goSession.MakePK(mCurrentTable)
            moFormDataRs!mwFormHeaderKey = mwFormHeaderKey
            moFormDataRs!XAPSection = mCurrentSectionNumber
            moFormDataRs!DetailSeq = mwFormDetailCnt
'            moFormDataRs!LRRowOffset = OffsetIncrement
            moFormDataRs!FieldName = CurrentCell
            moFormDataRs!mwcSitesKey = mwcSitesKey

            
            Select Case strTypeName
               Case "Byte", "Integer", "Long"
                  nDataType = adInteger
                  moFormDataRs!IsIntegerData = True
                  moFormDataRs!IntegerData = CLng(moExcel.Range(CurrentCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!IntegerData)
               Case "Single", "Double", "Currency", "Decimal"
                  nDataType = adDouble
                  moFormDataRs!IsDoubleData = True
                  moFormDataRs!DoubleData = CDbl(moExcel.Range(CurrentCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!DoubleData)
               Case "Date"
                  nDataType = adDate
                  moFormDataRs!IsDateData = True
                  moFormDataRs!DateData = CDate(moExcel.Range(CurrentCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!DateData)
               Case "Boolean"
                  nDataType = adBoolean
                  moFormDataRs!IsBooleanData = True
                  moFormDataRs!BooleanData = CBool(moExcel.Range(CurrentCell))
                  moFormDataRs!StringData = CStr(moFormDataRs!BooleanData)
               Case Else
                  nDataType = adChar
                  If Len(strValue) > 250 Then
                     moFormDataRs!IsMemoData = True
                     moFormDataRs!MemoData = CStr(moExcel.Range(CurrentCell))
                  Else
                     moFormDataRs!IsStringData = True
                     moFormDataRs!StringData = CStr(moExcel.Range(CurrentCell)) ' was ShortData
                  End If
            End Select

            moFormDataRs.Update
         End If
      End If
   End If
   GetSimpleRsField = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwExcelForm.GetSimpleRsField: ", Err.Number, Err.Description
   goSession.RaisePublicError "Failed Recordset Update on: " & oKey.KeyName & " - " & CurrentCell
   GetSimpleRsField = False

End Function


Private Function GetmwrChangeTableKey(TableName As String) As Long
   Dim loRs As Recordset
   Dim sSQL As String

   Dim UpperTableName As String
   
   On Error GoTo FunctionError
   
   'validate
   If Len(TableName) < 1 Then
      goSession.RaisePublicError "Error In GetmwrChangeTableKey, no TableName Value", vbInformation, "Check If Child References Exist"
      Exit Function
   End If
   
   UpperTableName = UCase(TableName)
   
   If goSession.IsAccess Then
      sSQL = "SELECT ID FROM mwrChangeTable WHERE UCase(TableName)  = '" & UpperTableName & "'"
   ElseIf goSession.IsSqlServer Then
      sSQL = "SELECT ID FROM mwrChangeTable WHERE Upper(TableName)  = '" & UpperTableName & "'"
   ElseIf goSession.IsOracle Then
      sSQL = "SELECT ID FROM mwrChangeTable WHERE UPPER(TableName)  = '" & UpperTableName & "'"
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      GetmwrChangeTableKey = loRs!ID
   Else
      GetmwrChangeTableKey = 0
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwGUI.GetmwrChangeTableKey. ", Err.Number, Err.Description
   CloseRecordset loRs
   GetmwrChangeTableKey = 0
End Function

Private Sub moFormHeaderRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.moFormHeaderRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub

Private Sub moFormDataRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   Static loCurrentTableName As String
   On Error GoTo SubError


   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
   End If

   If mCurrentTable <> loCurrentTableName Then
      loCurrentTableName = mCurrentTable
      If Not loWork.Initialize(mCurrentTable) Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If

   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.moFormDataRs_WillChangeRecord. ", Err.Number, Err.Description

End Sub

' SetFormDetailData - Read data from the database and store it in the form


Public Function SetFormDetailData(FormHeaderKey As Long) As Boolean
'   Dim loField As FormField
   Dim loField As Object
   Dim loCnt As Long
   Dim sSQL As String
   Dim strOperation As String
   Dim loRec As ConfigKeys
   Dim oKey As ConfigKey
   Dim i As Integer
   Dim loXapFile As ConfigGroups
   Dim strScriptFile As String
   Dim IsError As Boolean
   Dim IsAnyErrors As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   '
   ' Must be an open form...
   '
   If Not moParent.IsOpen Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, ExcelForm Not Open."
      SetFormDetailData = False
      Exit Function
   End If
   '
   ' it must be an ExcelForm...
   '
   If moParent.FormType <> mwft_EXCEL And moParent.FormType <> mwft_EXCEL_XLSX And moParent.FormType <> mwft_EXCEL_XLSM Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, Wrong Form Type."
      SetFormDetailData = False
      Exit Function
   End If
   
   If goSession.Workflow.TemplatePlacesID <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, No Template Source PlaceID defined for the " & goSession.Workflow.WorkflowID & " Workflow Profile."
      SetFormDetailData = False
      Exit Function
   End If
   
   If Len(moParent.GetExtendedProperty("MapFName")) <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, No XAP File Name defined for the " & moParent.TemplateProperties.TemplateID & " Form Template."
      SetFormDetailData = False
      Exit Function
   End If
   
   If Len(moParent.GetExtendedProperty("MapFExt")) <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, No XAP File Extension defined for the " & moParent.TemplateProperties.TemplateID & " Form Template."
      SetFormDetailData = False
      Exit Function
   End If
   
   If FormHeaderKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, illegal FormHeaderKey " & FormHeaderKey
      SetFormDetailData = False
      Exit Function
   End If
   mwFormHeaderKey = FormHeaderKey
'   If mwcSitesKey <= 0 Then
'      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, illegal mwcSitesKey " & mwcSitesKey
'      SetFormDetailData = False
'      Exit Function
'   End If
   
   '
   ' and must be an Open Map XAP File...
   '
   Set fso = New FileSystemObject
   
   strScriptFile = GetTemplateScriptFileName(mwFormHeaderKey)
   
   If Not fso.FileExists(strScriptFile) Then
      If fso.FileExists(moParent.WorkflowProperties.MapScriptFile) Then
         strScriptFile = moParent.WorkflowProperties.MapScriptFile
      Else
         strScriptFile = goSession.Workflow.ScriptFolder & "\" & moParent.WorkflowProperties.MapScriptFile
      End If
   End If
   
   Set loXapFile = goSession.LoadConfigGroupsFile(strScriptFile)
   If loXapFile Is Nothing Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, Open XAP File Error: " & _
        moParent.WorkflowProperties.MapScriptFile
      SetFormDetailData = False
      KillObject fso
      Exit Function
   End If
   '
   ' XAP File must match...
   '
   If loXapFile.Count = 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData, XAP File Corrupt: "
      SetFormDetailData = False
      KillObject fso
      Exit Function
   End If
   
   '
   ' OK, Create a recordset object to use...
   '
   CloseRecordset moFormDataRs
   
   Set moFormDataRs = New Recordset
   moFormDataRs.CursorLocation = adUseClient
   '
   IsError = False
   IsAnyErrors = False
   
   UnprotectAllWorksheets
   
      '
      '
      ' For each record (except first moRecord, which is the Header) in the XAP file....
      ' 1-based collection
      '
      For i = 2 To loXapFile.Count
         DoEvents
         Set loRec = loXapFile(i).ConfigKeys
         
         mCurrentSectionNumber = GetSectionNumber(loXapFile(i).ConfigGroup)
         ' Need an Operation
         strOperation = loRec.GetKeyValue("Operation")
         If strOperation = "A" Then
            IsError = Not SetFormDataRsTypeA(loRec, mwFormHeaderKey)
         ElseIf strOperation = "SA" Then
            IsError = Not SetFormDataRsTypeSA(loRec, mwFormHeaderKey)
         ElseIf strOperation = "LR" Then
            IsError = Not SetFormDataRsTypeLR(loRec, mwFormHeaderKey)
         ElseIf strOperation = "SLR" Then
            IsError = Not SetFormDataRsTypeSLR(loRec, mwFormHeaderKey)
         Else
            goSession.RaiseError "Error in mwExcelForm.SetFormDetailData, invalid Operation Type: " & strOperation
            IsError = True
         End If
         '
         ' Close the recordset for reuse...
         '
         If moFormDataRs.State = adStateOpen Then
            moFormDataRs.Close
         End If
         '
         ' Error Check...
         '
         If IsError Then
            If moParent.WorkflowProperties.TolerateOtherErrors Then
               IsError = False
               IsAnyErrors = True
'               goSession.RaisePublicError "Error in Record: " & str(i) & ". Tolerate Errors is set to True, continuing process..."
            Else
               goSession.RaisePublicError "Error in Record: " & str(i) & ". Tolerate Errors is set to False, cancelling process..."
               SetFormDetailData = False
               KillObject fso
               Exit Function
            End If
         End If
      Next i
      '
      ' We be done...
      '
      
      If IsAnyErrors Then
'         goSession.RaisePublicError "Partially Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         SetFormDetailData = False
      Else
'         goSession.RaisePublicError "Processed Data Form: " & moParent.FormProperties.FormID & " at: " & Now()
         SetFormDetailData = True
      End If
   ProtectAllWorksheets
   SaveWorkbook
      
   KillObject fso
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDetailData.", Err.Number, Err.Description
   SetFormDetailData = False
   KillObject fso
   ProtectAllWorksheets
End Function

Private Function GetSectionNumber(SectionLabel As String) As Long
   Dim WorkingStr As String
   Dim ThisChar As String
   Dim xx As Integer
   
   On Error GoTo FunctionError
   
   For xx = 1 To Len(SectionLabel)
      ThisChar = mID(SectionLabel, xx, 1)
      If IsNumeric(ThisChar) Then
         WorkingStr = WorkingStr & ThisChar
      End If
   Next xx
   
   If Len(WorkingStr) > 0 Then
      GetSectionNumber = CStr(WorkingStr)
   Else
      goSession.RaisePublicError "General Error in mwExcelForm.GetSectionNumber, XAP section header " & SectionLabel & " must contain a numeric value", Err.Number, Err.Description
      GetSectionNumber = -1
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwExcelForm.GetSectionNumber. ", Err.Number, Err.Description
   GetSectionNumber = -1
End Function

Private Function SetFormDataRsTypeA(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      SetFormDataRsTypeA = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeA, Table name missing from Record. "
      Exit Function
   End If
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeA, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      SetFormDataRsTypeA = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "SELECT * from " & mCurrentTable & " WHERE mwFormHeaderKey= " & FormHeaderKey & " AND XAPSection = " & mCurrentSectionNumber
   
   '
   ' If Child, use key to the parent record...
   '
   If xaprec.GetKeyValue("IsChildKeyName") <> "" And mCurrentParentRecordID > 0 Then
      sSQL = sSQL & " AND " & xaprec.GetKeyValue("IsChildKeyName") & " = " & mCurrentParentRecordID
   End If
   
   sSQL = sSQL & " ORDER BY ID"
   
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockReadOnly
   
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      SetFormDataRsTypeA = False
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   If SetFormDataRSFields(xaprec) Then
      '
      ' If Parent, save key to the parent record...
      '
      If xaprec.GetKeyValue("IsParentKeyName") <> "" Then
         mCurrentParentRecordID = moFormDataRs!ID
         mCurrentParentTable = mCurrentTable
      End If
   Else
      goSession.RaisePublicError "Error in mwDataForm.SetFormDataRsTypeA. CancelUpdate on Form " & moParent.FormProperties.FormID
      SetFormDataRsTypeA = False
      Exit Function
   End If
   
   SetFormDataRsTypeA = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.SetFormDataRsTypeA. ", Err.Number, Err.Description
   SetFormDataRsTypeA = False
End Function

Private Function SetFormDataRsTypeSA(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim sSQL As String
   Dim nKey As Long
   Dim IsSkipRecord As Boolean
   Dim iWorksheet As Integer
   Dim strNullTestCell As String
   On Error GoTo FunctionError
   
   mCurrentTable = "mwFormDetail"
'   mCurrentTable = xaprec.GetKeyValue("table")
'   If Trim(mCurrentTable) = "" Then
'      '
'      SetFormDataRsTypeSA = False
'      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeSA, Table name missing from Record. "
'      Exit Function
'   End If
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeSA, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      SetFormDataRsTypeSA = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.SetFormDataRsTypeSA, Worksheet Parameter Missing."
      SetFormDataRsTypeSA = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   
   moWorkbook.Worksheets(iWorksheet).Activate
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   sSQL = "SELECT * from " & mCurrentTable & " WHERE mwFormHeaderKey= " & FormHeaderKey & " AND XAPSection = " & mCurrentSectionNumber
   
   '
   ' If Child, use key to the parent record...
   '
   
   sSQL = sSQL & " ORDER BY DetailSeq"
   
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockReadOnly
   
   If Err Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeSA, Unable to open Table: " & _
         mCurrentTable & "Error: " & Err.Description
      SetFormDataRsTypeSA = False
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   
   SetFormDataSimpleRSFields xaprec
'   If SetFormDataSimpleRSFields(xaprec) Then
      '
      ' If Parent, save key to the parent record...
      '
'      If xaprec.GetKeyValue("IsParentKeyName") <> "" Then
'         mCurrentParentRecordID = moFormDataRs!ID
'         mCurrentParentTable = mCurrentTable
'      End If
'   Else
'      goSession.RaisePublicError "Error in mwDataForm.SetFormDataRsTypeSA. CancelUpdate on Form " & moParent.FormProperties.FormID
'      SetFormDataRsTypeSA = False
'      Exit Function
'   End If
   
   SetFormDataRsTypeSA = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwExcelForm.SetFormDataRsTypeSA. ", Err.Number, Err.Description
   SetFormDataRsTypeSA = False
End Function


Private Function SetFormDataRsTypeLR(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      SetFormDataRsTypeLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.SetFormDataRsTypeLR, Worksheet Parameter Missing."
      SetFormDataRsTypeLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   ' First row a no brainer
   '
   moWorkbook.Worksheets(iWorksheet).Activate
   '
   ' Open the recordset...
   '
   
   mCurrentTable = xaprec.GetKeyValue("table")
   If Trim(mCurrentTable) = "" Then
      '
      SetFormDataRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeLR, Table name missing from Record. "
      Exit Function
   End If
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeLR, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      SetFormDataRsTypeLR = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   
'            moFormDataRs!mwFormHeaderKey = FormHeaderKey
'            moFormDataRs!XAPSection = mCurrentSectionNumber
'            moFormDataRs!LRRowOffset = iCurrentOffset
   
   '
   ' Open the recordset
   '
   
   sSQL = "SELECT * from " & mCurrentTable & " WHERE mwFormHeaderKey= " & FormHeaderKey & " AND XAPSection = " & mCurrentSectionNumber
   
   '
   ' If Child, use key to the parent record...
   '
   If xaprec.GetKeyValue("IsChildKeyName") <> "" And mCurrentParentRecordID > 0 Then
      sSQL = sSQL & " AND " & xaprec.GetKeyValue("IsChildKeyName") & " = " & mCurrentParentRecordID
   End If
   
   sSQL = sSQL & " ORDER BY LRRowOffset"
   
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockReadOnly
   
   If Err Then
      SetFormDataRsTypeLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   iTotalRows = 0
   IsMoreRows = True
'   iStartingRowOffset = GetBaselineRowOffset(GetLoopingRowCell(mLrp.SkipRowTestCell, iCurrentOffset))
   Do While moFormDataRs.EOF = False
      
      iCurrentOffset = moFormDataRs!LRRowOffset
      
      iTotalLrParameters = MW_TOTAL_LR_PARAMETERS
      iTotalRows = iTotalRows + 1
      
      If xaprec.GetKeyValue("IsChildKeyName") <> "" Then
         iTotalLrParameters = MW_TOTAL_LR_PARAMETERS + 1
      End If
      
      '
      ' OK, add the non repeating values first...
      '
      If mLrp.TotalNonLoopingRowFields > 0 Then
         For i = iTotalLrParameters To (mLrp.TotalNonLoopingRowFields + (iTotalLrParameters - 1))
            If Not SetRsField(xaprec(i)) Then
               SetFormDataRsTypeLR = False
               Exit Function
            End If
         Next i
      End If
      '
      ' Then, current row variables
      '
      For i = (mLrp.TotalNonLoopingRowFields + iTotalLrParameters) To (xaprec.Count)
         If Not SetRsField(xaprec(i), iCurrentOffset) Then
            SetFormDataRsTypeLR = False
            Exit Function
         End If
      Next i
      moFormDataRs.MoveNext
   Loop
   SetFormDataRsTypeLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.SetFormDataRsTypeLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.SetFormDataRsTypeLR. ", Err.Number, Err.Description
   End If
   SetFormDataRsTypeLR = False
End Function

Private Function SetFormDataRsTypeSLR(xaprec As ConfigKeys, FormHeaderKey As Long) As Boolean
   Dim loDgRec As ConfigKeys
   Dim loXapKey As ConfigKey
   Dim sSQL As String
   Dim loKey As ConfigKey
   Dim lrp As LrParameters
   Dim iCurrentOffset As Integer
   Dim iStartingRowOffset As Integer
   Dim iTotalRows As Integer
   Dim iTotalNonLoopingRowFields As Integer
   Dim iWorksheet As Integer
   Dim i As Integer
   Dim ii As Integer
   Dim IsMoreRows As Boolean
   Dim IsSkipRow As Boolean
   Dim iTotalLrParameters As Integer
   Dim strNullTestValue As String
   
   On Error GoTo FunctionError
   '
   ' Step 1 - Required LR Parameters
   '
   If Not CheckLrparameters(xaprec) Then
      SetFormDataRsTypeSLR = False
      Exit Function
   End If
   
   '
   ' Set the worksheet
   '
   If Trim(xaprec.GetKeyValue("Worksheet")) = "" Then
      goSession.RaiseError "Error in mwExcelForm.SetFormDataRsTypeSLR, Worksheet Parameter Missing."
      SetFormDataRsTypeSLR = False
      Exit Function
   Else
      If IsNumeric(xaprec.GetKeyValue("Worksheet")) Then
         iWorksheet = xaprec.GetKeyValue("Worksheet")
      Else
         iWorksheet = 1
      End If
   End If
   
   moWorkbook.Worksheets(iWorksheet).Activate
   
   mCurrentTable = "mwFormDetail"
   
   mwrChangeTableKey = GetmwrChangeTableKey(mCurrentTable)
   If mwrChangeTableKey <= 0 Then
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeSLR, Table has not been declared as a replication table in mwrChangeTable: " & mCurrentTable
      SetFormDataRsTypeSLR = False
      Exit Function
   End If
   
   '
   ' Turn on manual error checking
   '
   On Error Resume Next
   If goSession.User.DebugMode Then MsgBox "before open table: " & mCurrentTable
   
   '
   ' Open the recordset
   '
   
   sSQL = "SELECT * from " & mCurrentTable & " WHERE mwFormHeaderKey= " & FormHeaderKey & " AND XAPSection = " & mCurrentSectionNumber
   
   '
   ' If Child, use key to the parent record...
   '
'   If xaprec.GetKeyValue("IsChildKeyName") <> "" And mCurrentParentRecordID > 0 Then
'      sSQL = sSQL & " AND " & xaprec.GetKeyValue("IsChildKeyName") & " = " & mCurrentParentRecordID
'   End If
   
   sSQL = sSQL & " ORDER BY DetailSeq"
   
   moFormDataRs.Open sSQL, goCon, adOpenStatic, adLockReadOnly
   
   If Err Then
      SetFormDataRsTypeSLR = False
      goSession.RaisePublicError "Error in mwSession.mwExcelForm.SetFormDataRsTypeSLR, Unable to open Table: " & mCurrentTable
      'goSession.CloseRecordset moFormDataRs
      Exit Function
   End If
   If goSession.User.DebugMode Then MsgBox "After open table: " & mCurrentTable
   On Error GoTo FunctionError
   
   SetFormDataSimpleRSFields xaprec
   
   SetFormDataRsTypeSLR = True
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwExcelForm.SetFormDataRsTypeSLR, Invalid Cell ?: " & loXapKey.KeyValue & ".", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwExcelForm.SetFormDataRsTypeSLR. ", Err.Number, Err.Description
   End If
   SetFormDataRsTypeSLR = False
End Function

Private Function SetFormDataRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   DoEvents
   If Not IsNumeric(oRec.GetKeyValue("WORKSHEET")) Then
      goSession.RaisePublicError "Error in Excel Script, non-numeric WORKSHEET specified (using 1)."
      moExcel.Worksheets(1).Activate
   Else
      moExcel.Worksheets(Int(oRec.GetKeyValue("WORKSHEET"))).Activate
   End If
   For Each loKey In oRec
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         
         Case Is = "OFFSETINCREMENT"
         Case Is = "MAXROWOFFSET"
         Case Is = "TOTALNONLOOPINGROWFIELDS"
   
         Case Else
            If Not SetRsField(loKey) Then
               ' terminate...
               SetFormDataRSFields = False
               Exit Function
            End If
      End Select
   Next loKey
   SetFormDataRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.SetFormDataRSFields: ", Err.Number, Err.Description
   SetFormDataRSFields = False
End Function
Private Function SetFormDataSimpleRSFields(ByRef oRec As ConfigKeys) As Boolean
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   Dim sExcelFormFormula As String
   On Error GoTo FunctionError
   DoEvents
   
   Do While moFormDataRs.EOF = False
   
      If Not IsNull(moFormDataRs!FieldName) Then
         If moExcel.Range(moFormDataRs!FieldName).Locked = False Then
            sExcelFormFormula = moExcel.Range(moFormDataRs!FieldName).Formula
            If CBool(moFormDataRs!IsIntegerData) = True Then
               If IsNull(moFormDataRs!IntegerData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CLng(moFormDataRs!IntegerData)
               End If
            ElseIf CBool(moFormDataRs!IsDoubleData) = True Then
               If IsNull(moFormDataRs!DoubleData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CDbl(moFormDataRs!DoubleData)
               End If
            ElseIf CBool(moFormDataRs!IsDateData) = True Then
               If IsNull(moFormDataRs!DateData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CDate(moFormDataRs!DateData)
               End If
            ElseIf CBool(moFormDataRs!IsBooleanData) = True Then
               If IsNull(moFormDataRs!BooleanData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CBool(moFormDataRs!BooleanData)
               End If
            ElseIf CBool(moFormDataRs!IsStringData) = True Then
               If IsNull(moFormDataRs!StringData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CStr(moFormDataRs!StringData)
               End If
            ElseIf CBool(moFormDataRs!IsMemoData) = True Then
               If IsNull(moFormDataRs!MemoData) Then
                  moExcel.Range(moFormDataRs!FieldName) = ""
               Else
                  moExcel.Range(moFormDataRs!FieldName) = CStr(moFormDataRs!MemoData)
               End If
            End If
            If sExcelFormFormula <> "" Then
               moExcel.Range(moFormDataRs!FieldName).Formula = sExcelFormFormula
            End If
         End If
         
      End If
      moFormDataRs.MoveNext
   Loop
   
   SetFormDataSimpleRSFields = True
   Exit Function
FunctionError:
   'Assume invalid field, cancel operation...
   goSession.RaisePublicError "General Error in mwDataForm.SetFormDataSimpleRSFields: ", Err.Number, Err.Description
   SetFormDataSimpleRSFields = False
End Function

'
' Stuff an Excel cell value into an ADO record, trapping data type errors as we go
'
Private Function SetRsField(oKey As ConfigKey, Optional OffsetIncrement As Integer) As Boolean
   Dim strValue As String
   Dim CurrentCell As String
   Dim IsDateConvert As Boolean
   Dim sExcelFormFormula As String
   On Error GoTo FunctionError
   
   If Not IsNull(oKey.KeyValue) Then
      If oKey.KeyName <> "" And oKey.KeyValue <> "" Then
         
         CurrentCell = GetLoopingRowCell(oKey.KeyValue, OffsetIncrement)
         sExcelFormFormula = moExcel.Range(CurrentCell).Formula
         If IsNull(moFormDataRs.Fields(oKey.KeyName)) Then
            moExcel.Range(CurrentCell) = ""
         Else
            Select Case (moFormDataRs.Fields(oKey.KeyName).Type)
            ' Multi-step error avoidance in different data types.
               Case adSingle, adDouble, adDecimal, adNumeric
                     
                  moExcel.Range(CurrentCell) = CDbl(moFormDataRs.Fields(oKey.KeyName))
               
               Case adDate, adDBTimeStamp
               
                  moExcel.Range(CurrentCell) = CDate(moFormDataRs.Fields(oKey.KeyName))
               
               Case adSmallInt, adInteger, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt
               
                  moExcel.Range(CurrentCell) = CLng(moFormDataRs.Fields(oKey.KeyName))
               
               Case Else
               
                  moExcel.Range(CurrentCell) = CStr(moFormDataRs.Fields(oKey.KeyName))
                  
            End Select

         End If
         If sExcelFormFormula <> "" Then
            moExcel.Range(CurrentCell).Formula = sExcelFormFormula
         End If
      End If
   End If
   
   SetRsField = True
   Exit Function
FunctionError:
      'Assume invalid field, cancel operation...
      goSession.RaisePublicError "General Error in mwExcelForm.SetRsField: ", Err.Number, Err.Description
      goSession.RaisePublicError "Failed Update on: " & oKey.KeyName & " - " & CurrentCell
      SetRsField = False
End Function


Private Function CalcFirstDataRecord(ByRef oRec As ConfigKeys) As Integer
   Dim loKey As ConfigKey
   Dim iErr As Long
   Dim sErr As String
   Dim iWS As Integer
   On Error GoTo FunctionError
   CalcFirstDataRecord = 0
   For Each loKey In oRec
      CalcFirstDataRecord = CalcFirstDataRecord + 1
      Select Case UCase(loKey.KeyName)
         Case Is = "TABLE"
         Case Is = "OPERATION"
         Case Is = "WORKSHEET"
         Case Is = "ISPARENTKEYNAME"
         Case Is = "ISCHILDKEYNAME"
         Case Is = "NULLTESTCELL"
         Case Is = "SKIPROWTESTCELL"
         
         Case Is = "OFFSETINCREMENT"
         Case Is = "MAXROWOFFSET"
         Case Is = "TOTALNONLOOPINGROWFIELDS"
   
         Case Else
            Exit Function
      End Select
   Next loKey
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.CalcFirstDataRecord: ", Err.Number, Err.Description
   CalcFirstDataRecord = 0
End Function
Private Function GetTemplateScriptFileName(mwFormHeaderKey As Long) As String
   Dim RevNo As Long
   On Error GoTo FunctionError
   
   If mwFormHeaderKey > 0 Then
      RevNo = GetFormHeadersRevNo(mwFormHeaderKey)
   Else
      RevNo = 1
   End If
   
   If RevNo > 1 Then
      GetTemplateScriptFileName = BuildPlacesPath(goSession.Workflow.TemplatePlacesID) & moParent.GetExtendedProperty("MapFName") & "_R" & Format(RevNo, "000") & "." & moParent.GetExtendedProperty("MapFExt")
   Else
      GetTemplateScriptFileName = BuildPlacesPath(goSession.Workflow.TemplatePlacesID) & moParent.GetExtendedProperty("MapFName") & "." & moParent.GetExtendedProperty("MapFExt")
   End If
      
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwExcelForm.GetTemplateScriptFileName.", Err.Number, Err.Description
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
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.GetFormHeadersRevNo ", Err.Number, Err.Description
   CloseRecordset loRs
   GetFormHeadersRevNo = 1
End Function

Private Function BuildPlacesPath(Optional mwcPlacesKey As Long, Optional PlaceID As String) As String
   Dim oRs As Recordset
   On Error GoTo FunctionError
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   
   If mwcPlacesKey > 0 Then
      oRs.Open "select * from mwcPlaces where ID= " & mwcPlacesKey, goCon, adOpenForwardOnly, adLockReadOnly
   ElseIf Len(PlaceID) > 0 Then
      oRs.Open "select * from mwcPlaces where PlaceID='" & PlaceID & "'", goCon, adOpenForwardOnly, adLockReadOnly
   Else
      goSession.RaisePublicError "General Error in mwSession.mwExcelForm.BuildPlacesPath: function called with no mwcPlacesID or PlaceID parms."
   End If
   
   If oRs.RecordCount < 1 Then
      If mwcPlacesKey > 0 Then
         goSession.RaisePublicError "General Error in mwSession.mwExcelForm.BuildPlacesPath: Could not find a record in mwcPlaces with ID = " & mwcPlacesKey
      Else
         goSession.RaisePublicError "General Error in mwSession.mwExcelForm.BuildPlacesPath: Could not find a record in mwcPlaces with PlaceID = " & PlaceID
      End If
   Else
      BuildPlacesPath = ""
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
   
   CloseRecordset oRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.BuildPlacesPath. ", Err.Number, Err.Description
   BuildPlacesPath = ""
   CloseRecordset oRs
End Function


Private Function CopyExcelProtectOptions() As Boolean
   Dim i As Integer
   Dim nSheetIndex As Integer
   Dim loProtection As Object
   On Error GoTo FunctionError
   
   ReDim arrXLProtect(0) As XLProtect
   
   For i = 1 To moWorkbook.Worksheets.Count
      Set loProtection = moWorkbook.Worksheets(i).Protection
      ReDim Preserve arrXLProtect(i) As XLProtect
      If Not loProtection Is Nothing Then
         With arrXLProtect(nSheetIndex)
            .IsAllowDeletingColumns = loProtection.AllowDeletingColumns
            .IsAllowDeletingRows = loProtection.AllowDeletingRows
            .IsAllowFiltering = loProtection.AllowFiltering
            .IsAllowFormattingCells = loProtection.AllowFormattingCells
            .IsAllowFormattingColumns = loProtection.AllowFormattingColumns
            .IsAllowFormattingRows = loProtection.AllowFormattingRows
            .IsAllowInsertingColumns = loProtection.AllowInsertingColumns
            .IsAllowInsertingHyperlinks = loProtection.AllowInsertingHyperlinks
            .IsAllowInsertingRows = loProtection.AllowInsertingRows
            .IsAllowSorting = loProtection.AllowSorting
            .IsAllowUsingPivotTables = loProtection.AllowUsingPivotTables
         End With
      End If
      nSheetIndex = nSheetIndex + 1
   Next i
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.CopyExcelProtectOptions. ", Err.Number, Err.Description
End Function

Private Function SetAdvancedProtectOptions(loSheet As Object) As Boolean
   Dim nArrIndex As Integer
   On Error GoTo FunctionError

   If loSheet.Index <= UBound(arrXLProtect) Then
      nArrIndex = loSheet.Index - 1
      If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
         loSheet.Protect , , , , , arrXLProtect(nArrIndex).IsAllowFormattingCells, arrXLProtect(nArrIndex).IsAllowFormattingColumns _
            , arrXLProtect(nArrIndex).IsAllowFormattingRows, arrXLProtect(nArrIndex).IsAllowInsertingColumns, arrXLProtect(nArrIndex).IsAllowInsertingRows, _
            arrXLProtect(nArrIndex).IsAllowInsertingHyperlinks, arrXLProtect(nArrIndex).IsAllowDeletingColumns, arrXLProtect(nArrIndex).IsAllowDeletingRows, _
            arrXLProtect(nArrIndex).IsAllowSorting, arrXLProtect(nArrIndex).IsAllowFiltering, arrXLProtect(nArrIndex).IsAllowUsingPivotTables
      Else
         loSheet.Protect moParent.GetExtendedProperty("Protect"), , , , , arrXLProtect(nArrIndex).IsAllowFormattingCells, arrXLProtect(nArrIndex).IsAllowFormattingColumns _
            , arrXLProtect(nArrIndex).IsAllowFormattingRows, arrXLProtect(nArrIndex).IsAllowInsertingColumns, arrXLProtect(nArrIndex).IsAllowInsertingRows, _
            arrXLProtect(nArrIndex).IsAllowInsertingHyperlinks, arrXLProtect(nArrIndex).IsAllowDeletingColumns, arrXLProtect(nArrIndex).IsAllowDeletingRows, _
            arrXLProtect(nArrIndex).IsAllowSorting, arrXLProtect(nArrIndex).IsAllowFiltering, arrXLProtect(nArrIndex).IsAllowUsingPivotTables
      End If
   Else
      If Trim(moParent.GetExtendedProperty("Protect")) = "" Then
         loSheet.Protect
      Else
         loSheet.Protect moParent.GetExtendedProperty("Protect")
      End If
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwExcelForm.SetAdvancedProtectOptions. ", Err.Number, Err.Description
End Function
