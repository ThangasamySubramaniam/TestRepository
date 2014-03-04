VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwCrystal8Work"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwCrystal8Work - Maritime Systems Wrapper for Crystal Report Functions
' 5/6/2000 ms
' 12/5/2000 ms
   Option Explicit
   
   ' Working Crystal Objects...
   Private moApp As CRPEAuto.Application
   Private moRep As CRPEAuto.Report
   Private moExp As CRPEAuto.ExportOptions
   Private moView As CRPEAuto.View
   Private WithEvents moWin As CRPEAuto.Window
Attribute moWin.VB_VarHelpID = -1
   Private moEventInfo As New CRPEAuto.EventInfo
   
   
   Private moPWO As CRPEAuto.PrintWindowOptions
   
   ' Private Variables
   Private mDestination As Integer
   Private mWindowTitle As String
   Private mSourceFile As String
   Private mTargetFile As String
   Private mTitle As String
   Private mAuthor As String
   Private mSummary As String
   Private mName As String
   Private mPrintDate As String
   Private mPrintFileType As Integer
   Private mPrinterName As String
   Private mRefreshData As Boolean
   Private mIsReport As Boolean
   Private mIsError As Boolean
   Private mCurrentError As mwError
   Private mIsOpen As Boolean
   Private mAutoPrintAction As Boolean
   Private mParameterPromptingEnabled As Boolean
   
   ' 7/2003 - Crystal Login Info...
   Private mServer As String
   Private mDatabase As String
   Private mUserID As String
   Private mPassword As String
   
   
   Private Const CRYSTAL_OUT_WINDOW = 0
   Private Const CRYSTAL_OUT_PRINTER = 1
   Private Const CRYSTAL_OUT_FILE = 2
   
   Public Enum mwCrystalValueTypes
      mw_crCurrencyField = 8
      mw_crDateField = 10
      mw_crDateTimeField = 16
      mw_crStringField = 12
      mw_crNumberField = 7
   End Enum
   
   Public Enum mwCrystalDestination
      mw_crWindow = 0
      mw_crPrinter = 1
      mw_crFile = 2
   End Enum
   
      
   
   Public Event ReportWindowOpen()
    
Private Sub Class_Initialize()
    ' Set default values for preview...
   mIsReport = False
   Set moApp = New CRPEAuto.Application
End Sub

Private Sub Class_Terminate()
    If Not moApp Is Nothing Then Set moApp = Nothing
    If Not moRep Is Nothing Then Set moRep = Nothing
    If Not moView Is Nothing Then Set moView = Nothing
    If Not moPWO Is Nothing Then Set moPWO = Nothing
End Sub

Public Property Get TargetFile() As String
    TargetFile = mTargetFile
End Property

Public Property Let TargetFile(ByVal sNewValue As String)
    mTargetFile = sNewValue
End Property

Public Property Get SourceFile() As String
    SourceFile = mSourceFile
End Property

Public Property Let SourceFile(ByVal sNewValue As String)
    mSourceFile = sNewValue
End Property

Public Property Get WindowTitle() As String
    WindowTitle = mWindowTitle
End Property

Public Property Let WindowTitle(ByVal sNewValue As String)
    mWindowTitle = sNewValue
End Property


Public Property Get Destination() As mwCrystalDestination
    Destination = mDestination
End Property

Public Property Let Destination(ByVal sNewValue As mwCrystalDestination)
    mDestination = sNewValue
End Property

    
Public Property Get RefreshData() As Boolean
    RefreshData = mRefreshData
End Property

Public Property Let RefreshData(ByVal bNewValue As Boolean)
    mRefreshData = bNewValue
End Property
    
Public Property Get Title() As String
    Title = mTitle
End Property
    
Public Property Get PrintDate() As String
    PrintDate = mPrintDate
End Property
    
Public Property Get IsReport() As Boolean
   IsReport = mIsReport
End Property
   
Public Property Get IsError() As Boolean
   IsError = mIsError
End Property
    
Public Property Get IsOpen() As Boolean
   IsOpen = mIsOpen
End Property
    
Public Property Get Author() As String
   Author = mAuthor
End Property
    
Public Property Get Summary() As String
   Summary = mSummary
End Property
    
Public Property Get Name() As String
   Name = mName
End Property
    
Public Property Get AutoPrintAction() As Boolean
    AutoPrintAction = mAutoPrintAction
End Property

Public Property Let AutoPrintAction(ByVal vNewValue As Boolean)
    mAutoPrintAction = vNewValue
End Property

Public Property Let ParameterPromptingEnabled(ByVal Data As Boolean)
   If mIsOpen Then
      moRep.ParameterPromptingEnabled = Data
      If Data Then
         moRep.DiscardSavedData
      End If
   End If
End Property

Public Property Let PrinterName(ByVal Data As String)
    mPrinterName = Data
End Property

Public Property Get PrinterName() As String
    PrinterName = mPrinterName
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
    
Public Function OpenReport(file As String, Optional DbConnectID As String) As Boolean
   On Error GoTo OpenReport_error
   Dim loData As mwDataWork
   mSourceFile = file
   Set moRep = moApp.OpenReport(file, 1)
   If DbConnectID <> "" Then
      Set loData = goSession.OpenNewDataWork
      If Not loData.GetCrystalLogin(DbConnectID, mServer, mDatabase, mUserID, mPassword) Then
         mIsError = True
         OpenReport = False
         mIsOpen = False
         mIsReport = False
         KillObject loData
         Exit Function
      End If
   End If
   OpenReport = True
   mIsOpen = True
   mIsReport = True
   KillObject loData
   Exit Function
OpenReport_error:
   goSession.RaiseError "General Error in mwCrystal8Work.OpenReport. ", err.Number, err.Description
   mIsError = True
   OpenReport = False
   mIsOpen = False
   mIsReport = False
   KillObject moRep
End Function
    
Public Function SetParameter(ItemOffset As Integer, value As String, ParamType As mwCrystalValueTypes) As Boolean
   Dim dDate As Date
   Dim nLong As Long
   On Error GoTo SetParameter_error
   If mIsOpen = False Then
      SetParameter = False
      Exit Function
   End If
   Select Case ParamType
      Case mw_crDateField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CDate(value)
      Case mw_crCurrencyField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CCur(value)
      Case mw_crNumberField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CDbl(value)
      Case mw_crStringField
         moRep.ParameterFields(ItemOffset).SetCurrentValue value
      Case mw_crDateTimeField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CDate(value)
      Case Else
         goSession.RaiseError "Error in mwCrystal8Work.SetParameter, unsupported parameter value encountered."
         SetParameter = False
         Exit Function
   End Select
   SetParameter = True
   Exit Function
SetParameter_error:
   goSession.RaiseError "General Error in mwCrystal8Work.SetParameter. ", err.Number, err.Description
   SetParameter = False
End Function
    
Public Function SetLocation(Item As Integer, LocationString As String) As Boolean
   Dim oTables As CRPEAuto.DatabaseTables
   On Error GoTo SetLocation_error
   If mIsOpen = False Then
      SetLocation = False
      Exit Function
   End If
   Set oTables = moRep.Database.Tables
   oTables(Item).Location = LocationString
   SetLocation = True
   Exit Function
SetLocation_error:
   goSession.RaiseError "General Error in mwCrystal8Work.SetLocation. ", err.Number, err.Description
   SetLocation = False
End Function
    
Public Function GetReportInfo()
   Dim loSum As ReportSummaryInfo
   Dim iLen As Integer
   On Error GoTo OpenReport_error
   mIsReport = True
   Set moRep = moApp.OpenReport(mSourceFile, 1)
   On Error GoTo GetTitle_error
   Set loSum = moRep.ReportSummaryInfo
   mName = Trim(loSum.Name)
   mTitle = Trim(loSum.Title)
   mAuthor = Trim(loSum.Author)
   mSummary = Trim(loSum.subject)
   'mTitle = moRep.FormulaFields("title").Text
   'iLen = Trim(Len(mTitle))
   ' Crystal strings seem to have single quotes in there...
   'If Left(mTitle, 1) = "'" Then
   '    mTitle = Mid$(mTitle, 2, iLen - 4)
   'End If
   mPrintDate = moRep.PrintDate
   Set loSum = Nothing
   Set moRep = Nothing
   Exit Function
OpenReport_error:
   mIsReport = False
GetTitle_error:
   On Error Resume Next
   mTitle = ""
   mPrintDate = ""
   If Not moRep Is Nothing Then Set moRep = Nothing
   If Not loSum Is Nothing Then Set loSum = Nothing
End Function
    
'End Function
    
Public Function ProcessReport()
   Dim strError As String
   Dim i As Integer
   Dim loPrintInfo As CRPEAuto.PrinterInfo
   
   On Error GoTo ProcessReport_error
   
   If Len(goSession.CrystalDLL) > 0 Then
      moApp.LogOnServer goSession.CrystalDLL, goSession.CrystalServer, goSession.CrystalDatabaseName, goSession.CrystalUser, goSession.CrystalPswd
   End If
   
   ' Note: the ",1" opens the report as a temporary file(making an
   ' extra copy in the "temp" folder. The original version is left
   ' untouched. Not well documented in Crystal manual...
   
   If moRep Is Nothing Then
      Set moRep = moApp.OpenReport(mSourceFile, 1)
   End If
   'moRep.Database.Tables(1).SetLogOnInfo mServer, mDatabase, mUserID, mPassword
   'moRep.Database.Tables(1).SetLogOnInfo "MW", "mwDataVela", "sa", ""
   
   If mServer <> "" Then
      For i = 1 To moRep.Database.Tables.Count
         moRep.Database.Tables(i).SetLogOnInfo mServer, mDatabase, mUserID, mPassword
      Next i
   End If
   
   
   Set moEventInfo = moRep.EventInfo
   moEventInfo.ActivatePrintWindowEventEnabled = True
   If mRefreshData Then
      ' Reorder following property set's to accomodate suppresion of parameters...
      'moRep.ParameterPromptingEnabled = mParameterPromptingEnabled
      'moRep.DiscardSavedData = mDiscardSavedData
      Set moExp = moRep.ExportOptions
      moExp.DiskFileName = mTargetFile
      moExp.DestinationType = crEDTDiskFile
      moExp.FormatType = crEFTCrystalReport
      moRep.Export False
   End If
   Set moPWO = moRep.PrintWindowOptions
   moPWO.CanDrillDown = True
   moPWO.HasCloseButton = True
   moPWO.HasCancelButton = True
   moPWO.HasExportButton = True
   moPWO.HasRefreshButton = False
   moPWO.HasNavigationControls = True
   moPWO.HasPrintButton = True
   moPWO.HasPrintSetupButton = True
   moPWO.HasSearchButton = True
   moPWO.HasZoomControl = True
   moPWO.HasGroupTree = True
   If mAutoPrintAction Then
      '
      ' Special case - auto print report, no visual printing...
      '
      Set loPrintInfo = moRep.PrinterInfo
      'loPrintInfo.PrinterName = mPrinterName
      moRep.PrintOut False
   Else
      Set moView = moRep.Preview
      Set moWin = moView.Parent
   End If
   Exit Function
ProcessReport_error:
   goSession.RaiseError "General Error in mwCrystal8Work.ProcessReport: ", err.Number, err.Description
   mIsError = True
End Function



Public Function DiscardSavedData()
   If mIsOpen Then
      moRep.DiscardSavedData
   End If
End Function

Public Function SetReportField(FieldName As String, FieldValue As String, valueType As mwCrystalValueTypes) As Boolean
   On Error GoTo SetReportField_error
   Select Case valueType
      Case Is = crStringField
         moRep.FormulaFields(FieldName).Text = Chr(34) & FieldValue & Chr(34)
      Case Is = crNumberField
         moRep.FormulaFields(FieldName).Text = FieldValue
   End Select
   SetReportField = True
   Exit Function
SetReportField_error:
   goSession.RaiseError "General Error in mwCrystal8Work.SetReportField. ", err.Number, err.Description
   SetReportField = False

End Function

'Public Function ClearParameterValues() As Boolean
'   Dim i As Long
'   On Error GoTo ClearParameterValues_error
'   If mIsOpen = False Then
'      ClearParameterValues = False
'      Exit Function
'   ElseIf moRep.ParameterFields.Count < 1 Then
'      ClearParameterValues = True
'      Exit Function
'   End If
'   For i = 1 To moRep.ParameterFields.Count
'      moRep.ParameterFields(i).SetCurrentValue
'   Next i
'   ClearParameterValues = True
'   Exit Function
'ClearParameterValues_error:
'   goSession.RaiseError "General Error in mwCrystal8Work.ClearParameterValues. ", err.Number, err.Description
'   ClearParameterValues = False
'End Function


