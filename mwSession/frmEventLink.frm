VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEventLink 
   Caption         =   "Edit Event Link"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Tag             =   "moRS.Fields(RST_LinkTitle).value"
   Begin VB.TextBox txtDescription 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2160
      Width           =   5535
   End
   Begin VB.CommandButton cmdPasteHyperlink 
      Caption         =   "Paste Clipboard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox chkMoveToEventFolder 
      Caption         =   "Copy To Event Folder"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      Picture         =   "frmEventLink.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdViewChm 
      Caption         =   "View Link"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Picture         =   "frmEventLink.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFormHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Picture         =   "frmEventLink.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7200
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Link"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Picture         =   "frmEventLink.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Height          =   975
      Left            =   6240
      Picture         =   "frmEventLink.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   1
      Top             =   1080
      Width           =   5535
   End
   Begin PVCOMBOLibCtl.PVComboBox pvcboHelpLink 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   5535
      _Version        =   524288
      _cx             =   9763
      _cy             =   661
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   2
      Sorted          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowPictures    =   0   'False
      ColumnHeaders   =   0   'False
      PrimaryColumn   =   0
      VisibleItems    =   10
      ColumnHeaderHeight=   20
      ListMember      =   ""
      ColumnHeaderForeColor=   0
      ColumnHeaderBackColor=   13160660
      SelectedForeColor=   16777215
      SelectedBackColor=   6956042
      AlternateBackColor=   16777215
      ItemLabelStyle  =   1
      ItemLabelType   =   0
      ItemLabelWidth  =   40
      ItemLabelForeColor=   0
      ItemLabelBackColor=   13160660
      ColumnHeaderStyle=   1
      VerticalGridLines=   0   'False
      HorizontalGridLines=   -1  'True
      ColumnResize    =   0   'False
      ItemLabelResize =   0   'False
      AllowDBAutoConfig=   -1  'True
      GridLineColor   =   13421772
      List            =   ""
      NullString      =   "[NULL]"
      DropShadow      =   -1  'True
      Text            =   ""
      SortOnColumnHeaderClick=   0   'False
      DropEffect      =   0
      ColumnCount     =   2
      Column0.Heading =   "Help Context"
      Column0.Width   =   100
      Column0.Alignment=   0
      Column0.Hidden  =   0   'False
      Column0.Name    =   ""
      Column0.Format  =   ""
      Column0.Bound   =   0   'False
      Column0.Locked  =   0   'False
      Column0.HeaderAlignment=   0
      Column1.Heading =   "Context ID"
      Column1.Width   =   40
      Column1.Alignment=   0
      Column1.Hidden  =   0   'False
      Column1.Name    =   ""
      Column1.Format  =   ""
      Column1.Bound   =   0   'False
      Column1.Locked  =   0   'False
      Column1.HeaderAlignment=   0
      SortKey1.Column =   -1
      SortKey1.Ascending=   -1  'True
      SortKey1.CaseInsensitive=   -1  'True
      SortKey2.Column =   -1
      SortKey2.Ascending=   -1  'True
      SortKey2.CaseInsensitive=   -1  'True
      SortKey3.Column =   -1
      SortKey3.Ascending=   -1  'True
      SortKey3.CaseInsensitive=   -1  'True
      BoundColumn     =   ""
      Border          =   -1  'True
      VertAlign       =   1
      Format          =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblLinkType 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Link Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblHelpLink 
      Caption         =   "Help Link"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEqptID 
      Caption         =   "Link to File:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmEventLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmEventLink - Manage External Link
' 8/30/02 ms
Option Explicit
Dim mIsCancelled As Boolean
Dim mInitDir As String
Dim mIsDeleteLink As Boolean
Dim mCurrentContextLink As String
Dim mStartingFullFilename As String
Dim mEventType As Long
Dim mIsButtonTitle As Boolean
Dim moRS As Recordset

Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventLinkTypeKey = 3
Const RS_LinkTitle = 4
Const RS_IsCreated = 5
Const RS_IsSubmit = 6
Const RS_FullFilename = 7
Const RS_ContextID = 8
Const RS_DateTimeCreated = 9
Const RS_BriefDescription = 10
Const RS_UserId = 11
Const RST_ID = 12
Const RST_mwEventTypeKey = 13
Const RST_LinkTitle = 14
Const RST_DisplaySequence = 15
Const RST_DefaultDescription = 16
Const RST_DisplayIcon = 17
Const RST_IsMandatory = 18
Const RST_IsSuggested = 19
Const RST_IsSubmitAllowed = 20
Const RST_SourceLocationPath = 21
Const RST_SourceFilePattern = 22
Const RST_IsOverrideSourceFolderAllowed = 23
Const RST_IsMustCopyToEventFolder = 24


' loRsLink - must be set to correct record...
Public Function InitForm(EventType As Long, loRsLink As Recordset) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   mEventType = EventType
   Set moRS = loRsLink
   If Not IsNull(moRS.Fields(RS_FullFilename).value) Then
   '
   ' Set InitDir
   '
      Set fso = New FileSystemObject
      txtFile.Text = moRS.Fields(RS_FullFilename).value
      mStartingFullFilename = moRS.Fields(RS_FullFilename).value
      mInitDir = fso.GetParentFolderName(moRS.Fields(RS_FullFilename).value)
      '
      If UCase(fso.GetExtensionName(txtFile.Text)) = "CHM" Then
         If LoadHHFile() Then
            pvcboHelpLink.Visible = True
            lblHelpLink.Visible = True
         Else
            pvcboHelpLink.Visible = False
            lblHelpLink.Visible = False
         End If
      Else
         pvcboHelpLink.Visible = False
         lblHelpLink.Visible = False
      End If
      Set fso = Nothing
   Else
      If Not IsNull(moRS.Fields(RST_SourceLocationPath).value) Then
         mInitDir = moRS.Fields(RST_SourceLocationPath).value
      End If
   End If
   '
   ' If I only had a brain...
   '
   If moRS.Fields(RST_IsMustCopyToEventFolder).value Then
      chkMoveToEventFolder.value = 1
      chkMoveToEventFolder.Enabled = False
   End If
   If Not IsNull(moRS.Fields(RST_LinkTitle).value) Then
      lblLinkType.Caption = moRS.Fields(RST_LinkTitle).value
   End If
   If Not IsNull(moRS.Fields(RS_BriefDescription).value) Then
      txtDescription.Text = moRS.Fields(RS_BriefDescription).value
   End If
   
   goSession.SetDotNetTheme Me
   
   InitForm = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in frmEventLink.SetCurrentFile. ", err.Number, err.Description
   InitForm = False
End Function


Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdDelete_Click()
   mIsDeleteLink = True
   Me.Hide
End Sub

Private Sub cmdFormHelp_Click()
   'goSession.API.ShowMwHelp SWEQ_HELP_MANUAL, SWEQ_External_Links

End Sub

Private Sub cmdOK_Click()
   Dim fso As FileSystemObject
   Dim loEvWork As mwEventWork
   Dim strPath As String
   Dim strFile As String
   Dim i As Integer
   Dim strNewFilename As String
   On Error GoTo SubError
   '
   ' Save screen values to Recordset...
   '
   If pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1) > 1 Then
      moRS.Fields(RS_ContextID).value = pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1)
   Else
      moRS.Fields(RS_ContextID).value = Null
   End If
   ' Reset the IsCreated flag
   If Trim(txtFile.Text) = "" Then
      moRS.Fields(RS_FullFilename).value = Null
      moRS.Fields(RS_IsCreated).value = False
      mIsDeleteLink = True
      Me.Hide
      Exit Sub
   Else
      moRS.Fields(RS_FullFilename).value = txtFile.Text
      moRS.Fields(RS_IsCreated).value = True
   End If
   If Trim(txtDescription.Text) = "" Then
      moRS.Fields(RS_BriefDescription).value = Null
   Else
      moRS.Fields(RS_BriefDescription).value = txtDescription.Text
   End If
   
   '
   ' Further action required ?
   '
   If chkMoveToEventFolder.value <> 1 Then
      Me.Hide
      Exit Sub
   End If
   If UCase(mStartingFullFilename) = UCase(txtFile.Text) Then
      Me.Hide
      Exit Sub
   End If
   Set fso = New FileSystemObject
   ' Is source a file or a pasted link ?
   If fso.FileExists(txtFile.Text) Then
      '
      ' Copy to Voyage Folder...
      '
      Set loEvWork = New mwEventWork
      strPath = loEvWork.GetVoyageFolder(mEventType)
      Set loEvWork = Nothing
      If strPath = "" Then
         goSession.RaiseError "Error in frmEventLink.CmdOK_Click, Event Form Folder missing. "
         mIsCancelled = True
         Exit Sub
      End If
      strFile = strPath & "\" & fso.GetFileName(txtFile.Text)
      If fso.FileExists(strFile) Then
         '
         ' File already exists...
         '
         
         MsgBox "File already exists in Event Folder, a new name will be assigned.", vbInformation, "Copy Linked File to Voyage Folder"
         i = 0
         Do While fso.FileExists(strFile)
            i = i + 1
            If i > 100 Then
               Exit Do
            End If
            strFile = strPath & "\" & fso.GetBaseName(txtFile.Text) & LTrim(str(i)) & "." & fso.GetExtensionName(txtFile.Text)
         Loop
      End If
      fso.CopyFile txtFile.Text, strFile
      txtFile.Text = strFile
   End If
   moRS.Fields(RS_FullFilename).value = txtFile.Text
   Set fso = Nothing
   Set loEvWork = Nothing
   
   ' That's all folks...
   Me.Hide
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmEventLink.CmdOK_Click. ", err.Number, err.Description
   Me.Hide
End Sub

Private Sub cmdPasteHyperlink_Click()
   Dim strHyperlink As String
   On Error GoTo SubError
   strHyperlink = Clipboard.GetText
   If strHyperlink <> "" Then
      txtFile.Text = strHyperlink
   Else
      MsgBox "Nothing in paste buffer to use.", vbInformation
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmEventLink.cmdPasteHyperlink_Click. ", err.Number, err.Description
End Sub

Private Sub cmdViewChm_Click()
   If txtFile.Text = "" Then
      Beep
      Exit Sub
   End If
   '
   ' Launch External Link...
   '
   If Trim(pvcboHelpLink.Text) = "" Or pvcboHelpLink.Text = "<No Context Help>" Then
      goSession.API.LaunchExternalLink txtFile.Text
   Else
      goSession.API.LaunchExternalLink txtFile.Text, pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1)
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmEventLink.cmdViewChm_Click. ", err.Number, err.Description
End Sub


Private Sub txtFile_GotFocus()
   Dim fso As FileSystemObject
   On Error GoTo SubError
   If txtFile.Text = "" Then
      GetLinkToFile
      Exit Sub
   End If
   Set fso = New FileSystemObject
   If fso.FileExists(txtFile.Text) Then
      GetLinkToFile
   End If
   Set fso = Nothing
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmEventLink.txtFile_GotFocus. ", err.Number, err.Description
End Sub


Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function IsDeleteLink() As Boolean
   IsDeleteLink = mIsDeleteLink
End Function





Public Function GetCurrentFile() As String
   If Trim(txtFile.Text) = "" Then
      GetCurrentFile = ""
   Else
      GetCurrentFile = txtFile.Text
   End If
End Function

Public Function SetFilter(FileFilter As String)
   On Error GoTo FunctionError
   cd1.Filter = FileFilter
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in frmEventLink.SetFilter. ", err.Number, err.Description
End Function

Public Function SetCurrentContextLink(ContextLink As String)
   mCurrentContextLink = ContextLink
End Function


Public Function GetCurrentContextLink() As Long
   If pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1) > 1 Then
      GetCurrentContextLink = pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1)
   Else
      GetCurrentContextLink = -1
   End If
End Function




'
' Open the ".hh" help file (from Robohelp) and loads CBO...
'
Private Function LoadHHFile() As Boolean
   On Error GoTo FunctionError
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim strHeaderFilename As String
   Dim strHeader() As String
   Dim strTemp
   Dim iRow As Long
   Dim strFolder As String
   
   '  Select help Context from Header file found on same
   '  folder as the help file...
   '
   If Trim(txtFile.Text) = "" Then
      LoadHHFile = False
      Exit Function
   End If
   
   Set fso = New FileSystemObject
   ' Need current, column 0's value...xxx
   strHeaderFilename = fso.GetParentFolderName(txtFile.Text) _
     & "\" & fso.GetBaseName(txtFile.Text) _
     & ".hh"
   If Not fso.FileExists(strHeaderFilename) Then
      LoadHHFile = False
   Else
      '
      ' Load file into frmGetComboxBox... cbo field
      Set ts = fso.OpenTextFile(strHeaderFilename, ForReading)
      
      'PopulateHelpContext
      With Me.pvcboHelpLink
         Do While Not ts.AtEndOfStream
            strTemp = mID(ts.ReadLine, 9)
            strHeader = Split(strTemp, Chr(9))
            If UBound(strHeader) = 1 Then
               .AddItem strHeader(0)
               iRow = .NewIndex
               ' HelpContext Description...
               .SubItem(iRow, 0) = strHeader(0)
               ' HelpContext Number
               .SubItem(iRow, 1) = strHeader(1)
               If strHeader(1) = mCurrentContextLink Then
                  .Text = strHeader(0)
               End If
            End If
         Loop
         .AddItem "<No Context Help>"
         .SubItem(.NewIndex, 0) = "<No Context Help>"
         ' HelpContext Number
         .SubItem(.NewIndex, 1) = " "
         ts.Close
      End With
      LoadHHFile = True
   End If
   KillObject fso
   KillObject ts
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in frmEventLink.LoadHHFile. ", err.Number, err.Description
   LoadHHFile = False
End Function


Private Function GetLinkToFile()
   Dim strFile As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   strFile = txtFile.Text
   If Trim(txtFile.Text) <> "" Then
      cd1.DialogTitle = "Link File to Event"
      cd1.InitDir = fso.GetParentFolderName(txtFile.Text)
      cd1.FileName = fso.GetFileName(txtFile.Text)
   Else
      If Trim(mInitDir) <> "" Then
         cd1.InitDir = mInitDir
      End If
   End If
   cd1.CancelError = True
   cd1.ShowOpen
   pvcboHelpLink.ClearItems
   txtFile.Text = cd1.FileName
   ' is it a chm file ?
   If UCase(fso.GetExtensionName(txtFile.Text)) = "CHM" Then
      If LoadHHFile() Then
         pvcboHelpLink.Visible = True
         lblHelpLink.Visible = True
      Else
         pvcboHelpLink.Visible = False
         lblHelpLink.Visible = False
      End If
   Else
      pvcboHelpLink.Visible = False
      lblHelpLink.Visible = False
   End If
   'txtFile.Text = cd1.FileName
Exit Function
FunctionError:
   If err.Number = 32755 Then
      ' Cancel button pressed on Common Dialog...
      txtFile.Text = strFile
   Else
      goSession.RaisePublicError "Error in frmEventLink.GetLinkToFile. ", err.Number, err.Description
   End If
   KillObject fso
End Function

