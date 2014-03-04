VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmCfgBrowser 
   Caption         =   "ShipNet Fleet Configuration Utility"
   ClientHeight    =   7245
   ClientLeft      =   1440
   ClientTop       =   1365
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCfgBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSortBy 
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
      Left            =   8880
      TabIndex        =   18
      Top             =   1800
      Width           =   3075
   End
   Begin UltraGrid.SSUltraGrid ug1 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7223
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   68157440
      Override        =   "frmCfgBrowser.frx":08CA
      Caption         =   "Configuration Table"
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap DD/MM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Picture         =   "frmCfgBrowser.frx":0920
      TabIndex        =   5
      ToolTipText     =   "Print"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdApplyFilter 
      Caption         =   "Apply Filter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6348
      TabIndex        =   3
      Top             =   1728
      Width           =   1455
   End
   Begin VB.TextBox txtFilter 
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
      Left            =   900
      TabIndex        =   2
      Top             =   1800
      Width           =   5355
   End
   Begin VB.TextBox txtRecordCount 
      Enabled         =   0   'False
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
      Left            =   1380
      TabIndex        =   15
      Top             =   6540
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Row"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   6480
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "frmCfgBrowser.frx":0D62
      ScaleHeight     =   1335
      ScaleWidth      =   1335
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin PVCOMBOLibCtl.PVComboBox pvc1 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1260
      Width           =   7335
      _Version        =   524288
      _cx             =   12938
      _cy             =   661
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   0
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
      VerticalGridLines=   -1  'True
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
      Column0.Heading =   "Table Description"
      Column0.Width   =   40
      Column0.Alignment=   0
      Column0.Hidden  =   0   'False
      Column0.Name    =   ""
      Column0.Format  =   ""
      Column0.Bound   =   0   'False
      Column0.Locked  =   0   'False
      Column0.HeaderAlignment=   0
      Column1.Heading =   "Table Name"
      Column1.Width   =   40
      Column1.Alignment=   0
      Column1.Hidden  =   -1  'True
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
   Begin VB.CommandButton cmdOpenTable 
      Caption         =   "Open"
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
      Left            =   10440
      TabIndex        =   9
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtTableName 
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
      Left            =   7920
      TabIndex        =   8
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   6480
      Width           =   1215
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   10920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   12030
      DesignHeight    =   7245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Sort By"
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
      Left            =   7920
      TabIndex        =   19
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Filter"
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
      Left            =   60
      TabIndex        =   17
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Records"
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
      Left            =   180
      TabIndex        =   16
      Top             =   6540
      Width           =   1095
   End
   Begin VB.Label lblDbConnectionString 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Connection"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   840
      Width           =   2595
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Common Config Tables"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   1260
      Width           =   2595
   End
   Begin VB.Label Label2 
      Caption         =   "Table Name"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ShipNet Fleet Configuration Browser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1823
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmCfgBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moParent As mwSession.Session
Attribute moParent.VB_VarHelpID = -1
Dim moCon As ADODB.Connection
'dim moConShape as ADODB.da
Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1

' Replication variables
   Dim mMWRT As Long
   Dim mIsFleetNonReplicate As Boolean
   Dim mCurrentTableName As String

Dim mLastActiveCellIndex As Long

Private Sub cmdApplyFilter_Click()
   On Error GoTo SubError
   If IsRecordLoaded(moRS) Then
      moRS.Filter = txtFilter.Text
      txtRecordCount.Text = moRS.RecordCount
   End If
   Exit Sub
SubError:
   'moParent.RaiseError "General Error in frmCfgBrowser.cmdApplyFilter, Bad Filter Expression...Resetting to none ", Err.Number, Err.Description
   MsgBox "Invalid Filter Expression", vbInformation, "Configuration Table"
   moRS.Filter = adFilterNone
End Sub

Private Sub cmdInsert_Click()
   On Error GoTo SubError
   moRS.AddNew
   Exit Sub
SubError:
   moParent.RaiseError "General Error in frmCfgBrowser, unable to Add New Record: ", Err.Number, Err.Description
End Sub

Private Sub cmdOpenTable_Click()
   Dim sSQL As String
   Dim nIndexKey As Long
   On Error GoTo SubError
   
   If Trim(txtTableName.Text) = "" Then
      Exit Sub
   End If
   ug1.Caption = "Configuration Table"
   ' clear pvc1 display for nonReplication tables
   pvc1.Text = ""
   moParent.CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   
   mCurrentTableName = txtTableName.Text
   
   If UCase(txtTableName.Text) = "SMFILECABINET" Then
      sSQL = "SELECT ID, ParentKey, smBlobFileKey, mwcSitesKey, smFileCabFleetIndexKey,  Description, Abbreviation,  FullFilename, IsReplicate, " & _
      " IsFastTrack, mwEventTypeKey, mwEventDetailKey, DateMod, IsSignedOut, FileExtension, smDocTypeKey, IsSendNow,  mwrBatchLogOutboundKey " & _
      " FROM smFileCabinet"
   ElseIf UCase(txtTableName.Text) = "MWBLOBFILE" Then
      sSQL = "SELECT ID, mwBlobFileTypeKey, BriefDescription, FullFileName), BlobCreated, BlobUpdated, FileTypeDescription, " & _
      " mwEventTypeKey, mwEventDetailKey, VersionNumber " & _
      " FROM mwBlobFile"
   Else
      sSQL = "SELECT * FROM " & txtTableName.Text
   End If
   
   If txtSortBy.Text <> "" Then
      sSQL = sSQL & " ORDER BY " & txtSortBy
   End If
   
   moRS.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   Set ug1.DataSource = moRS
   txtRecordCount.Text = moRS.RecordCount
   
   ' log & display TableDescription
   goSession.LogIt mwl_User_Defined, mwl_Information, "Opening Configuration Table: " & mCurrentTableName
   ug1.Caption = mCurrentTableName
   
'   '
'   ' get the pvcbo index and replication ID key
'   nIndexKey = pvc1.Search(Trim(txtTableName.Text), PV_CT_TableName, -1)
'   If nIndexKey >= 0 Then
'      mMWRT = pvc1.SubItem(nIndexKey, PV_CT_ID)
'   Else
'      mMWRT = -1
'   End If
'
'   If Not (UCase(txtTableName.Text) = "SMFILECABINET" _
'   Or UCase(txtTableName.Text) = "MWBLOBFILE") Then
'      '  is Fleet Table
'      If pvc1.SubItem(nIndexKey, PV_CT_mwrBatchTypeKey) = 1 Then
'         mIsFleetNonReplicate = True
'      Else
'         mIsFleetNonReplicate = False
'      End If
'   Else
'      ' check for Fleet table
'      If pvc1.SubItem(nIndexKey, PV_CT_mwrBatchTypeKey) = 1 Then
'         mIsFleetNonReplicate = True
'      End If
'   End If
   
   ' lock it down shipboy
   If SITE_TYPE_SHIP = goSession.Site.SiteType Then
      ug1.Override.AllowDelete = ssAllowDeleteNo
      ug1.Override.AllowUpdate = ssAllowUpdateNo
      cmdInsert.Enabled = False
   End If
      
   Exit Sub
SubError:
   If Err.Number = -2147217900 Or Err.Number = -2147217904 Then
      MsgBox "Invalid Column name in 'Sort By' Field", vbInformation, "Configuration Table"
      Exit Sub
   ElseIf Err.Number = -2147217865 Then
      MsgBox "Invalid Table name", vbInformation, "Configuration Table"
      Exit Sub
   ElseIf Err.Number = -2147467259 Then
      MsgBox "Invalid Table or Column name", vbInformation, "Configuration Table"
      Exit Sub
   End If
   moParent.RaiseError "General Error in frmCfgBrowser.cmdOpenTable_click: ", Err.Number, Err.Description
End Sub

Private Sub cmdSave_Click()
   On Error GoTo SubError
   If Trim(mCurrentTableName) <> "" Then
      moRS.Update
   End If
   Exit Sub
SubError:
   moParent.RaiseError "General Error in frmCfgBrowser.cmdSave_Click. ", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
   '
   ' Configuration Tables
   '
   Dim loRs As Recordset
   Dim sSQL As String
   goSession.LogIt mwl_Workstation, mwl_Information, "Entering Configuration Browser, UserID: " & goSession.User.UserID
   
'   sSQL = "SELECT * FROM mwrChangeTable WHERE IsActive <> 0  ORDER BY TableName"
   
   If goSession.IsAccess Then
      sSQL = "SELECT * FROM mwrChangeTable WHERE IsActive <> 0  ORDER BY UCase(TableName)"
   ElseIf goSession.IsSqlServer Then
      sSQL = "SELECT * FROM mwrChangeTable WHERE IsActive <> 0  ORDER BY Upper(TableName)"
   ElseIf goSession.IsOracle Then
      sSQL = "SELECT * FROM mwrChangeTable WHERE IsActive <> 0  ORDER BY UPPER(TableName)"
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, moCon, adOpenForwardOnly, adLockReadOnly
   
   If loRs.RecordCount > 0 Then
      pvc1.BoundColumn = "ID"
      pvc1.PrimaryColumn = 1
      Set pvc1.ListSource = loRs
      pvc1.ColumnWidth(PV_CT_ID) = 25
      pvc1.ColumnWidth(PV_CT_TableName) = 90
      pvc1.ColumnWidth(4) = 100
      pvc1.ColumnHidden(PV_CT_ID) = True
      pvc1.ColumnHidden(PV_CT_IsActive) = True
      pvc1.ColumnHidden(PV_CT_mwrBatchTypeKey) = True
      pvc1.ColumnHidden(PV_CT_SaveAuditLogs) = True
   Else
      MsgBox "No Replication Tables to Display", vbInformation, "Replication tables"
      Exit Sub
   End If
   CloseRecordset loRs
   
   mCurrentTableName = ""
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   moParent.RaiseError "General Error in frmCfgBrowser.Form_load. ", Err.Number, Err.Description
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   moParent.CloseRecordset moRS
   goSession.LogIt mwl_Workstation, mwl_Information, "Exiting Configuration Browser, UserID: " & goSession.User.UserID
End Sub


Private Sub pvc1_Change()
   On Error GoTo SubError
   
   If pvc1.Text = "" Then
      Exit Sub
   End If
   txtTableName.Text = ""
   ug1.Caption = "Configuration Table"

   moParent.CloseRecordset moRS
   Set moRS = New Recordset
   Dim strSQL As String
   moCon.CursorLocation = adUseClient
   moRS.CursorLocation = adUseClient
   
   mCurrentTableName = pvc1.SubItem(pvc1.ListIndex, PV_CT_TableName)
   If Trim(mCurrentTableName) <> "" Then
      If UCase(pvc1.SubItem(pvc1.ListIndex, PV_CT_TableName)) = "SMFILECABINET" Then
         strSQL = "SELECT ID, ParentKey, smBlobFileKey, mwcSitesKey, smFileCabFleetIndexKey,  Description, Abbreviation,  FullFilename, IsReplicate, " & _
         " IsFastTrack, mwEventTypeKey, mwEventDetailKey, DateMod, IsSignedOut, FileExtension, smDocTypeKey, IsSendNow,  mwrBatchLogOutboundKey " & _
         " FROM smFileCabinet"
      ElseIf UCase(pvc1.SubItem(pvc1.ListIndex, PV_CT_TableName)) = "MWBLOBFILE" Then
         strSQL = "SELECT ID, mwBlobFileTypeKey, BriefDescription, FullFileName), BlobCreated, BlobUpdated, FileTypeDescription, " & _
         " mwEventTypeKey, mwEventDetailKey, VersionNumber " & _
         " FROM mwBlobFile"
      Else
         strSQL = "SELECT * FROM " & pvc1.SubItem(pvc1.ListIndex, PV_CT_TableName)
      End If
      If txtSortBy.Text <> "" Then
         strSQL = strSQL & " ORDER BY " & txtSortBy
      End If
      moRS.Open strSQL, moCon, adOpenDynamic, adLockOptimistic
      Set ug1.DataSource = moRS
      
      txtRecordCount.Text = moRS.RecordCount
      ' log & display TableDescription
      goSession.LogIt mwl_User_Defined, mwl_Information, "Opening Configuration Table: " & pvc1.SubItem(pvc1.ListIndex, 4)
      ug1.Caption = pvc1.SubItem(pvc1.ListIndex, 4)
   Else
      MsgBox "Please select the valid Table Name", vbInformation, "Configuration Table"
   End If
   Exit Sub
SubError:
   If Err.Number = -2147217900 Or Err.Number = -2147217904 Then
      MsgBox "Invalid Column name in 'Sort By' Field", vbInformation, "Configuration Table"
      Exit Sub
   ElseIf Err.Number = -2147217865 Then
      MsgBox "Invalid Table name", vbInformation, "Configuration Table"
      Exit Sub
   ElseIf Err.Number = -2147467259 Then
      MsgBox "Invalid Table or Column name", vbInformation, "Configuration Table"
      Exit Sub
   End If
   moParent.RaiseError "General Error in frmCfgBrowser.pvc1_change: ", Err.Number, Err.Description
End Sub



Public Function SetParentSession(ByRef ses As Session)
   On Error GoTo FunctionError
   Set moParent = ses
   Set moCon = moParent.DBConnection
   lblDbConnectionString.Caption = moParent.DBConnectString
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmCfgBrowser.SetParentSession. ", Err.Number, Err.Description
End Function




Private Sub ug1_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
   'Debug.Print "KeyDown Event: KeyCode: " & KeyCode & ", Shift: " & Shift
   On Error GoTo SubError
   If KeyCode = 38 Then
      KeyCode = 0
      ug1.PerformAction ssKeyActionExitEditMode
      ug1.PerformAction ssKeyActionAboveCell
      ug1.PerformAction ssKeyActionEnterEditMode
      ug1.PerformAction ssKeyActionToggleCellSel
      If Not IsNull(ug1.ActiveCell.value) Then
         ug1.ActiveCell.SelStart = Len(ug1.ActiveCell.value)
      End If
   ElseIf KeyCode = 40 Then
      ' Down Arrow
      KeyCode = 0
      ug1.PerformAction ssKeyActionExitEditMode
      ug1.PerformAction ssKeyActionBelowCell
      ug1.PerformAction ssKeyActionEnterEditMode
      ug1.PerformAction ssKeyActionToggleCellSel
      If Not IsNull(ug1.ActiveCell.value) Then
         ug1.ActiveCell.SelStart = Len(ug1.ActiveCell.value)
      End If
   ElseIf KeyCode = 9 And Shift = 0 Then
      KeyCode = 0
      ug1.PerformAction ssKeyActionExitEditMode
      ug1.PerformAction ssKeyActionNextCellByTab
      ug1.PerformAction ssKeyActionToggleCellSel
      If ug1.ActiveCell Is Nothing Then
         Exit Sub
      ElseIf Not IsNull(ug1.ActiveCell.value) Then
         ug1.ActiveCell.SelStart = Len(ug1.ActiveCell.value)
      End If
   ElseIf KeyCode = 9 And Shift = 1 Then
      KeyCode = 0
      ug1.PerformAction ssKeyActionExitEditMode
      ug1.PerformAction ssKeyActionPrevCellByTab
      ug1.PerformAction ssKeyActionToggleCellSel
      If Not IsNull(ug1.ActiveCell.value) Then
         ug1.ActiveCell.SelStart = Len(ug1.ActiveCell.value)
      End If
   ElseIf KeyCode = 13 Then
      KeyCode = 0
      ug1.ActiveRow.ExpandAll
   
   End If
   Exit Sub
SubError:
   If Err.Number <> 380 And Err.Number <> 429 Then
      goSession.RaisePublicError "General error in frmCfgBrowser.ug1_Keydown. ", Err.Number, Err.Description
   End If

End Sub

Private Sub ug1_LostFocus()
   If ug1.ActiveCell Is Nothing Then
      mLastActiveCellIndex = 0
   Else
      mLastActiveCellIndex = ug1.ActiveCell.Column.Index
   End If
   ug1.PerformAction ssKeyActionDeactivateCell
   ug1.PerformAction ssKeyActionExitEditMode
End Sub

Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   Static loCurrentTableName As String
   On Error GoTo SubError


   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
'      TableName = moRsChangeTable!TableName
'      If Not loWork.Initialize(TableName) Then
'         Set loWork = Nothing
'         Exit Sub
'      End If
   End If

   If mCurrentTableName <> loCurrentTableName Then
      loCurrentTableName = mCurrentTableName
      If Not loWork.Initialize(mCurrentTableName) Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If

   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmCfgBrowser.moRs_WillChangeRecord. ", Err.Number, Err.Description
End Sub

Private Sub cmdSwap_Click()
   On Error GoTo SubError
   Dim sTempDate As String
   Dim dNewDate As Date
   Dim bChangeSeparator As Boolean
   
   If mLastActiveCellIndex > 0 Then
      ug1.ActiveRow.Cells(mLastActiveCellIndex).Activation = ssActivationAllowEdit
      ug1.ActiveRow.Cells(mLastActiveCellIndex).Selected = True
      ug1.ActiveCell = ug1.ActiveRow.Cells(mLastActiveCellIndex)
      ug1.PerformAction ssKeyActionEnterEditMode
   End If

   If ug1.ActiveCell Is Nothing Then
      MsgBox "Please select the Field Date", vbInformation, "Date Swap"
      Exit Sub
   End If
   
   ' ssDataTypeDate for Access, ssDataTypeDBTimeStamp for sql/Oracle
   If Not (ug1.ActiveCell.Column.Datatype = ssDataTypeDate Or ug1.ActiveCell.Column.Datatype = ssDataTypeDBTimeStamp) Then
      MsgBox "Date columns are only editible for swap", vbInformation, "Swap DD/MM"
      Exit Sub
   End If
   
   If Not IsDate(ug1.ActiveCell.value) Then
      MsgBox "Please select a Field Value with a valid date value. If the field data is indeed a date you may need to change the date separators." & vbCrLf & "For example: change 02.03.04 to 02/03/04"
      Exit Sub
   End If
   
   If Year(ug1.ActiveCell.value) = "1899" Then
      MsgBox "The field data selected contains an unrecognized regional date separator. Please change the date separator manually and try again." & vbCrLf & "For example: change 02.03.04 to 02/03/04"
      Exit Sub
   End If

   'Regional date work around
   sTempDate = ug1.ActiveCell.value
   sTempDate = Format(sTempDate, "yyyy") & "-" & Format(sTempDate, "dd") & "-" & Format(sTempDate, "mm")
      
   If IsDate(sTempDate) Then
      dNewDate = sTempDate
      ug1.ActiveCell.value = dNewDate
      ug1.Update
      ug1.ActiveCell.Appearance.BackColor = &HFFC0FF
   Else
      If CStr(ug1.ActiveCell.value) = CStr(Format(ug1.ActiveCell.value, "Short Date")) Then
         MsgBox "Could not swap day and month. You are most likely requesting a month which exceeds the number 12"
      Else
         'adjust to regional setting
         ug1.ActiveCell.value = Format(ug1.ActiveCell.value, "Short Date")
      End If
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.frmCfgBrowser.cmdSwap_Click", Err.Number, Err.Description
End Sub

Private Sub txtSortBy_LostFocus()
   pvc1_Change
End Sub


