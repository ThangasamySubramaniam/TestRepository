VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmEditExternalLink 
   Caption         =   "Edit External Link"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
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
      Left            =   240
      Picture         =   "frmEditExternalLink.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtButtonTitle 
      DataField       =   "EqptID"
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
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "EqptID"
      Top             =   1800
      Width           =   1815
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
      Left            =   3000
      Picture         =   "frmEditExternalLink.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
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
      Left            =   4680
      Picture         =   "frmEditExternalLink.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7080
      Top             =   240
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
      Left            =   1560
      Picture         =   "frmEditExternalLink.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Height          =   975
      Left            =   6240
      Picture         =   "frmEditExternalLink.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      DataField       =   "EqptID"
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
      Left            =   2040
      TabIndex        =   0
      Tag             =   "EqptID"
      Top             =   240
      Width           =   4815
   End
   Begin PVCOMBOLibCtl.PVComboBox pvcboHelpLink 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   4815
      _Version        =   524288
      _cx             =   8493
      _cy             =   661
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   2
      Sorted          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      DropEffect      =   1
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
   Begin VB.Label Label2 
      Caption         =   "Button Title"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   4
      Top             =   1080
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
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditExternalLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmExternalLink - Manage External Link
' 8/30/02 ms
Option Explicit
Dim mIsCancelled As Boolean
Dim mInitDir As String
Dim mIsDeleteLink As Boolean
Dim mCurrentLink As String


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

Private Sub CmdOK_Click()
   Me.Hide
End Sub

Private Sub cmdViewChm_Click()
   If txtFile.Text = "" Then
      Beep
      Exit Sub
   End If
   '
   ' Launch External Link...
   '
   If Trim(pvcboHelpLink.Text) <> "" And pvcboHelpLink.Text <> "<No Context Help>" Then
      goSession.API.LaunchExternalLink txtFile.Text, pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1)
   Else
      goSession.API.LaunchExternalLink txtFile.Text
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwWorks5.frmExternalLink.cmdViewChm_Click. ", err.Number, err.Description
End Sub

Private Sub txtFile_GotFocus()
   GetLinkToFile
End Sub

Public Function SetWhichButton(WhichButton As String)
   Me.Caption = Me.Caption & " - " & WhichButton
End Function


Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function IsDeleteLink() As Boolean
   IsDeleteLink = mIsDeleteLink
End Function

Public Function SetCurrentFile(CurrentFile As String)
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   txtFile.Text = CurrentFile
   cd1.InitDir = fso.GetParentFolderName(CurrentFile)
   Set fso = Nothing
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwWorks5.frmExternalLink.SetCurrentFile. ", err.Number, err.Description
End Function

Public Function SetInitDir(InitDir As String)
   mInitDir = InitDir
End Function

Public Function SetButtonTitle(ButtonTitle As String)
   txtButtonTitle.Text = ButtonTitle
End Function


Public Function GetCurrentFile() As String
   If Trim(txtFile.Text) = "" Then
      GetCurrentFile = ""
   Else
      GetCurrentFile = txtFile.Text
   End If
End Function

Public Function SetCurrentLink(CurrentLink As Long)
   mCurrentLink = CurrentLink
End Function

Public Function GetCurrentLink() As Long
   If pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1) > 1 Then
      GetCurrentLink = pvcboHelpLink.SubItem(pvcboHelpLink.ListIndex, 1)
   Else
      GetCurrentLink = -1
   End If
End Function

Public Function GetButtonTitle() As String
   If Trim(txtButtonTitle.Text) <> "" Then
      GetButtonTitle = txtButtonTitle.Text
   Else
      GetButtonTitle = ""
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
      MsgBox "Help Link Header file is missing: " & strHeaderFilename, vbExclamation
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
               If strHeader(1) = mCurrentLink Then
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
   End If
   KillObject fso
   KillObject ts
   LoadHHFile = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwWorks5.frmExternalLink.LoadHHFile. ", err.Number, err.Description
   LoadHHFile = False
End Function


Private Function GetLinkToFile()
   Dim strFile As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   strFile = txtFile.Text
   If Trim(txtFile.Text) <> "" Then
      cd1.DialogTitle = "Link File to System View"
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
      Else
         pvcboHelpLink.Visible = False
      End If
   Else
      pvcboHelpLink.Visible = False
   End If
   'txtFile.Text = cd1.FileName
Exit Function
FunctionError:
   If err.Number = 32755 Then
      ' Cancel button pressed on Common Dialog...
      txtFile.Text = strFile
   Else
      goSession.RaisePublicError "Error in frmExternalLink.GetLinkToFile. ", err.Number, err.Description
   End If
   KillObject fso
End Function

