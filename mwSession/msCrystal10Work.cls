VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwCrystal10Work"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwCrystal10Work - Maritime Systems Wrapper for Crystal Report Functions
' 5/6/2000 ms
' 12/5/2000 ms
   Option Explicit
   
   ' Working Crystal Objects...
   Private moApp As CRAXDRT.Application
   Private moRep As CRAXDRT.Report
   Private moExp As CRAXDRT.ExportOptions
   Private moDB As CRAXDRT.Database
   'Private moView As CRAXDRT.View
   'Private WithEvents moWin As CRAXDRT.Window
   'Private moEventInfo As New CRAXDRT.EventInfo
   
   
   'Private moPWO As CRAXDRT.PrintWindowOptions
   
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
   
   Private mIsMwsReport As Boolean
   Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"
   
   Private Const CRYSTAL_OUT_WINDOW = 0
   Private Const CRYSTAL_OUT_PRINTER = 1
   Private Const CRYSTAL_OUT_FILE = 2
   
   Public Enum mwCrystalValueTypes
      mw_crCurrencyField = 8
      mw_crDateField = 10
      mw_crDateTimeField = 16
      mw_crStringField = 12
      mw_crNumberField = 7
      mw_crBooleanField = 1
   End Enum
   
   Public Enum mwCrystalDestination
      mw_crWindow = 0
      mw_crPrinter = 1
      mw_crFile = 2
   End Enum
   
   Dim mo
   
   Public Event ReportWindowOpen()
    
Private Sub Class_Initialize()
   On Error GoTo FunctionError
    ' Set default values for preview...
   mIsReport = False
   Set moApp = New CRAXDRT.Application
   Exit Sub
FunctionError:
   goSession.RaiseError "General Error in mwCrystal10Work.Class_Initialize. ", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
    If Not moApp Is Nothing Then Set moApp = Nothing
    If Not moRep Is Nothing Then Set moRep = Nothing
    'If Not moView Is Nothing Then Set moView = Nothing
    'If Not moPWO Is Nothing Then Set moPWO = Nothing
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
      moRep.EnableParameterPrompting = Data
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
   On Error GoTo FunctionError
   Dim loData As mwDataWork
   mSourceFile = file
   Set moRep = moApp.OpenReport(file, 1)
   '
   ' DbConnectID used in msWorkstation.frmCrystalMain to distinguish between "MWS" reports and
   ' non-MWS reports where we do not mess with the data connection string at all
   '
   If DbConnectID <> "NON_MWS" Then
      mIsMwsReport = True
      'Set loData = goSession.OpenNewDataWork
      'If Not loData.GetCrystalLogin("MW", mServer, mDatabase, mUserID, mPassword) Then
      '   mIsError = True
      '   OpenReport = False
      '   mIsOpen = False
      '   mIsReport = False
      '   KillObject loData
      '   Exit Function
      'End If
   Else
      mIsMwsReport = False
   End If
   OpenReport = True
   mIsOpen = True
   mIsReport = True
   KillObject loData
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwCrystal10Work.OpenReport. ", Err.Number, Err.Description
   mIsError = True
   OpenReport = False
   mIsOpen = False
   mIsReport = False
   KillObject moRep
End Function

Public Function GetReportParameterName(ParamIndex As Long) As String
   On Error GoTo FunctionError

   If ParamIndex <= moRep.ParameterFields.Count Then
      GetReportParameterName = moRep.ParameterFields(ParamIndex).ParameterFieldName
   End If

   Exit Function
FunctionError:
   GetReportParameterName = ""
   goSession.RaiseError "General Error in mwCrystal10Work.GetReportParameterName. ", Err.Number, Err.Description
End Function
    
Public Function GetReportParameterType(ParamIndex As Long) As Integer
   On Error GoTo FunctionError

   If ParamIndex <= moRep.ParameterFields.Count Then
      GetReportParameterType = moRep.ParameterFields(ParamIndex).valueType
   End If

   Exit Function
FunctionError:
   GetReportParameterType = 0
   goSession.RaiseError "General Error in mwCrystal10Work.GetReportParameterType. ", Err.Number, Err.Description
End Function

Public Function GetReportParameterCount() As Long
   On Error GoTo FunctionError

   GetReportParameterCount = moRep.ParameterFields.Count

   Exit Function
FunctionError:
   GetReportParameterCount = 0
   goSession.RaiseError "General Error in mwCrystal10Work.GetReportParameterCount. ", Err.Number, Err.Description
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
         'nLong = 1166017860
         'moRep.ParameterFields(ItemOffset).AddCurrentValue nLong
      Case mw_crStringField
         moRep.ParameterFields(ItemOffset).SetCurrentValue value
      Case mw_crDateTimeField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CDate(value)
      Case mw_crBooleanField
         moRep.ParameterFields(ItemOffset).SetCurrentValue CBool(value)
      Case Else
         goSession.RaiseError "Error in mwCrystal10Work.SetParameter, unsupported parameter value encountered."
         SetParameter = False
         Exit Function
   End Select
   SetParameter = True
   Exit Function
SetParameter_error:
   goSession.RaiseError "General Error in mwCrystal10Work.SetParameter. ", Err.Number, Err.Description
   SetParameter = False
End Function
    
Public Function SetLocation(Item As Integer, LocationString As String) As Boolean
   Dim oTables As CRAXDRT.DatabaseTables
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
   goSession.RaiseError "General Error in mwCrystal10Work.SetLocation. ", Err.Number, Err.Description
   SetLocation = False
End Function
    

Public Function GetReportInfo()
   Dim iLen As Integer
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   Set moRep = moApp.OpenReport(mSourceFile, 1)
   mName = fso.GetFileName(mSourceFile)
   mTitle = moRep.ReportTitle
   mAuthor = moRep.ReportAuthor
   mSummary = moRep.ReportSubject
   mPrintDate = moRep.PrintDate
   Set moRep = Nothing
   mIsReport = True
   Exit Function
FunctionError:
   mIsReport = False
   mTitle = ""
   mPrintDate = ""
   mAuthor = ""
   mName = ""
   mSummary = ""
   Set moRep = Nothing
End Function
    
    
Public Function PreProcessReport()
'   Private moRep As CRAXDRT.Report
   Dim loCryCon As CRAXDRT.ConnectionProperties
   Dim strError As String
   Dim i As Integer
   Dim sConn As String
   Dim nPos As Long
   Dim sPwd As String
   Dim sTemp As String
   
   On Error GoTo FunctionError
   If moRep Is Nothing Then
      Set moRep = moApp.OpenReport(mSourceFile, 1)
   End If
   moRep.ReportComments = "Shipnet Fleet Version " & App.Major & "." & App.Minor & "." & Left(Format(App.Revision, "0000"), 1) & " Build " & CStr(Val(mID(Format(App.Revision, "0000"), 2, 2)))
   FixReportTables moRep
   
   ' handy dandy debugging code...
   'moRep.Database.Verify
   
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwCrystal10Work.PreProcessReport: ", Err.Number, Err.Description
   mIsError = True

End Function

Private Sub FixReportTables(loRep As CRAXDRT.Report)
   Dim loCryCon As CRAXDRT.ConnectionProperties
   Dim loSection As CRAXDRT.Section
   Dim loSubReport As CRAXDRT.SubreportObject
   Dim loRepObj As CRAXDRT.Report
   Dim loObject As Object
   Dim strError As String
   Dim i As Integer
   Dim X As Integer
   Dim sConn As String
   Dim nPos As Long
   Dim sPwd As String
   Dim sTemp As String
   
   On Error GoTo FunctionError
   If loRep Is Nothing Then
      Exit Sub
   End If
   
   If mIsMwsReport Then
      On Error Resume Next
      If loRep.Database.Tables.Count > 0 Then
         loRep.Database.Tables(1).DllName = "crdb_ado.dll"
      End If
      On Error GoTo FunctionError
      For i = 1 To loRep.Database.Tables.Count
         With loRep.Database.Tables(i).ConnectionProperties
            If goSession.IsOracle Then
               ' OLEDB
               .DeleteAll
               
               .Add "Provider", goCon.Provider
               sTemp = goCon.Properties("Data Source")
               .Add "Data Source", sTemp
               sTemp = goCon.Properties("User ID")
               .Add "User ID", sTemp
               sTemp = goCon.Properties("Password")
               .Add "Password", sTemp
               'sTemp = goCon.Properties("Server Name")
               '.Add "Server Name", sTemp
'               .Add "Persist Security Info", goCon.Properties("Persist Security Info")
               loRep.Database.Tables(i).Location = UCase(loRep.Database.Tables(i).Location)
'               loRep.Database.Tables(i).Name = UCase(loRep.Database.Tables(i).Name)
               
               '
               ' ODBC
               '.Item("Database") = "mwDB"
               '.Item("Database") = mServer
               '.Item("User ID") = mPassword
               '.Item("Password") = mUserID
               '.Item("UseDSNProperties") = 1
            ElseIf goSession.IsAccess Then
               .DeleteAll
               .Add "Provider", UCase(goCon.Provider)
               .Add "Data Source", UCase(goCon.Properties("Data Source"))
               sPwd = GetAccessDbPassword
               If Trim(sPwd) <> "" Then
                  .Add "Jet Database Password", sPwd
               End If
               '.Add "Password", ""
               'loRep.Database.Tables(1).DllName = "crdb_ado.dll"
               
            Else 'SQL Server
               ' OLEDB
               .DeleteAll
'               .Add "Provider", "SQLOLEDB"
               .Add "Provider", UCase(goCon.Provider)
               sTemp = goCon.Properties("Data Source")
               .Add "Data Source", sTemp
               '.Add "Data Source", "MS_LAP"
               '.Add "Data Source", goCon.Properties("Data Source")
               '.Add "Initial Catalog", goCon.Properties("Initial Catalog")
               sTemp = goCon.Properties("Initial Catalog")
               '.Add "Initial Catalog", goCon.Properties("Initial Catalog")
               .Add "Initial Catalog", sTemp
               .Add "Integrated Security", False
               sTemp = goCon.Properties("User ID")
               '.Add "User ID", goCon.Properties("User ID")
               .Add "User ID", sTemp
               '.Add "User ID", "sa"
               
               ' SQL Server doesn't keep the Password so we need to find it in the connection string
               ' then parse it out and strip off junk
               
               sConn = goSession.GetDecryptedDBConnectString(ENCRYPT_PSWD)
               nPos = InStr(1, sConn, "Password=", vbTextCompare)
               If nPos > 1 Then
                  sPwd = mID$(sConn, nPos + 9)
                  nPos = InStr(1, sPwd, ";", vbTextCompare)
                  If nPos > 1 Then
                     sPwd = Left$(sPwd, nPos - 1)
                  End If
                  
                  sPwd = Trim$(sPwd)
                  
                  .Add "Password", sPwd
                  '.Add "Password", "sa"
               End If
               On Error Resume Next
               loRep.Database.Tables(1).DllName = "crdb_ado.dll"
               On Error GoTo FunctionError

'               .Add "Server", mServer
               ' ODBC
               '.Item("User ID") = mUserID
               '.Item("Password") = mPassword
               '.DeleteAll
               '.Add "DSN", mServer
               '.Add "User ID", "sa"
               '.Add "Password", "sa"
            End If
         End With
      Next i
      
      For Each loSection In loRep.Sections
         For Each loObject In loSection.ReportObjects
            If loObject.Kind = crSubreportObject Then
               Set loSubReport = loObject
               Set loRepObj = loSubReport.OpenSubreport
               FixReportTables loRepObj
            End If
         Next loObject
      Next loSection
      
   End If
   
   ' handy dandy debugging code...
   loRep.Database.Verify
   
   
   Exit Sub
FunctionError:
   Resume Next
   goSession.RaiseError "General Error in mwCrystal10Work.PreProcessReport: ", Err.Number, Err.Description
   mIsError = True

End Sub
    
    
    
Public Function ProcessReport(Optional IsModal As Boolean)
   Dim loCryCon As CRAXDRT.ConnectionProperties
   Dim strError As String
   Dim i As Integer
   Dim sConn As String
   Dim nPos As Long
   Dim sPwd As String
   Dim sTemp As String
   
   On Error GoTo FunctionError
   
   
   If moRep Is Nothing Then
      goSession.RaiseError "Error in mwCrystal10Work.ProcessReport, report not opened and pre-processed."
      Exit Function
   End If
   If goRepForm Is Nothing Then
      Set goRepForm = New frmCrystalViewer
   End If
   
   
   If mAutoPrintAction Then
      '
      ' Special case - auto print report, no visual printing...
      '
      'moRep.PrintOut True, 1
      moRep.PrintOut False, 1
   Else
      'Set goRepForm = New frmCrystalViewer
      '
      ' Debug code
      '
      'moRep.ExportOptions.DestinationType = crEDTDiskFile
      'moRep.ExportOptions.DiskFileName = "c:\test.rpt"
      'moRep.ExportOptions.FormatType = crEFTCrystalReport
      'moRep.Export True
      'moRep.ParameterFields(1).AddCurrentValue 511706507
      
      'goRepForm.ShowReport moRep.ReportSubject, moRep
      
      goRepForm.ShowReport WindowTitle, moRep
      
      If IsModal Then
         goRepForm.Show vbModal
      Else
         goRepForm.Show
      End If
      
      
      'Set moView = moRep.Preview
      'Set moWin = moView.Parent
   End If
   Exit Function
FunctionError:
   If Err.Number = 400 Or Err.Number = 401 Then
      goSession.RaiseError "Can not display this report because another report window is already open." _
      & vbCrLf & "Please close the other report window then try again.", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwCrystal10Work.ProcessReport: ", Err.Number, Err.Description
   End If
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
   goSession.RaiseError "General Error in mwCrystal10Work.SetReportField. ", Err.Number, Err.Description
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
'   goSession.RaiseError "General Error in mwCrystal10Work.ClearParameterValues. ", err.Number, err.Description
'   ClearParameterValues = False
'End Function


Private Function GetAccessDbPassword() As String
   Dim sPwd As String
   Dim nPos As Long
   Dim sCon As String
   On Error GoTo FunctionError
   sCon = goSession.GetDecryptedDBConnectString(ENCRYPT_PSWD)
   nPos = InStr(1, sCon, "Jet OLEDB:Database Password=", vbTextCompare)
   If nPos > 1 Then
      sPwd = mID$(sCon, nPos + 28)
      nPos = InStr(1, sPwd, ";", vbTextCompare)
      If nPos > 1 Then
         sPwd = Left$(sPwd, nPos - 1)
      End If
      GetAccessDbPassword = Trim$(sPwd)
   Else
      GetAccessDbPassword = ""
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwCrystal10Work.ClearParameterValues. ", Err.Number, Err.Description
End Function

Public Function SetRecordset(oRs As Recordset) As Boolean
   Dim oTables As CRAXDRT.DatabaseTables
   On Error GoTo FunctionError
   If mIsOpen = False Then
      SetRecordset = False
      Exit Function
   End If
   
   Set moDB = moRep.Database
   moRep.DiscardSavedData
   moDB.SetDataSource oRs, 3, 1

   SetRecordset = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwCrystal10Work.SetRecordset. ", Err.Number, Err.Description
   SetRecordset = False
End Function
    
Public Function SetSubReportLocation() As Boolean
   Dim loSection As CRAXDRT.Section
   Dim loSubReport As CRAXDRT.SubreportObject
   Dim loRepObj As CRAXDRT.Report
   Dim loObject As Object
   On Error GoTo FunctionError

   For Each loSection In moRep.Sections
      For Each loObject In loSection.ReportObjects
         If loObject.Kind = crSubreportObject Then
            Set loSubReport = loObject
            Set loRepObj = loSubReport.OpenSubreport
            FixReportTables loRepObj
         End If
      Next loObject
   Next loSection
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwCrystal10Work.SetSubReportLocation. ", Err.Number, Err.Description
   SetSubReportLocation = False
End Function
