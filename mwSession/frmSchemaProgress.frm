VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmSchemaProgress 
   Caption         =   "Schema Updates Processing"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "frmSchemaProgress.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Stop processing Schemas and return to the Windows Desktop"
      Top             =   4440
      Width           =   1155
   End
   Begin VB.CommandButton cmdRunOne 
      Caption         =   "Run One"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Start Processing Schemas"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdRunAll 
      Caption         =   "Run All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Start Processing Schemas"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Stop processing Schemas and return to the Windows Desktop"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      Picture         =   "frmSchemaProgress.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Start Processing Schemas"
      Top             =   4440
      Width           =   1155
   End
   Begin UltraGrid.SSUltraGrid ugSchemas 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5847
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68158484
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColScrollRegions=   "frmSchemaProgress.frx":0FD4
      Override        =   "frmSchemaProgress.frx":1012
      Caption         =   "Schemas To Be Processed"
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   6060
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   9390
      DesignHeight    =   5370
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3780
      Width           =   9135
   End
End
Attribute VB_Name = "frmSchemaProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim moRS As Recordset
Dim mIsCancelled As Boolean
Dim moMenuItem As mwMenuItem
Dim moSchemaUpdateWork As mwSchemaUpdateWork

Dim mCurrentID As Long
Dim mIsPaused As Boolean

Const ICON_GREEN_DOT = "GreenCheck16.ico"
Const ICON_RED_DOT = "Red_X_16.ico"
Const ICON_BLUE_DOT = "Dot_Blue_16.ico"
Const ICON_CLEAR_DOT = "Dot_Clear_16.ico"
Const ICON_YELLOW_DOT = "HelpSmall.ico"

' frmOffice UG Recordset mapping...
Const UG_ID = 0
Const UG_Status = 1
Const UG_Path = 2
Const UG_Name = 3
Const UG_ChangeID = 4
Const UG_Description = 5
Const UG_Message = 6
'   loRs.Fields.Append "ID", adInteger, 4
'   loRs.Fields.Append "Status", adVarWChar, 50
'   loRs.Fields.Append "Path", adVarWChar, 200
'   loRs.Fields.Append "Name", adVarWChar, 56
'   loRs.Fields.Append "ChangeID", adVarWChar, 50
'   loRs.Fields.Append "Description", adVarWChar, 255
'   loRs.Fields.Append "Message", adVarWChar, 255

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function


Public Function InitForm(InRs As Recordset) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   
   Set moRS = InRs
   
   If moRS.EOF = False Then
      mCurrentID = moRS!ID
   Else
      mCurrentID = -1
   End If
   
   cmdOK.Visible = False
   cmdPause.Visible = False
   mIsPaused = False
   
   If ugSchemas.Images.Count <= 0 Then
      Dim ThePath As String
   
      ThePath = goSession.GetAppPath() & "\icons\16x16\"
      ugSchemas.Images.Add , "ICON_Clear", LoadPicture(ThePath & ICON_CLEAR_DOT)
      ugSchemas.Images.Add , "ICON_Green", LoadPicture(ThePath & ICON_GREEN_DOT)
      ugSchemas.Images.Add , "ICON_Red", LoadPicture(ThePath & ICON_RED_DOT)
      ugSchemas.Images.Add , "ICON_Yellow", LoadPicture(ThePath & ICON_YELLOW_DOT)
   
   End If
   
   Set ugSchemas.DataSource = moRS
   
   lblStatus.Caption = ""
   HideUltragridColumns ugSchemas, 0
   ugSchemas.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   
   ugSchemas.Bands(0).Columns(UG_Status).Hidden = False
   ugSchemas.Bands(0).Columns(UG_Status).Width = 1500
   ugSchemas.Bands(0).Columns(UG_Status).Header.Caption = "Status"
   ugSchemas.Bands(0).Columns(UG_Status).CellAppearance.PictureVAlign = ssVAlignTop
   
   ugSchemas.Bands(0).Columns(UG_Name).Hidden = False
   ugSchemas.Bands(0).Columns(UG_Name).Width = 2000
   ugSchemas.Bands(0).Columns(UG_Name).Header.Caption = "File Name"
   
   ugSchemas.Bands(0).Columns(UG_ChangeID).Hidden = False
   ugSchemas.Bands(0).Columns(UG_ChangeID).Width = 1500
   ugSchemas.Bands(0).Columns(UG_ChangeID).Header.Caption = "Change ID"
   
   ugSchemas.Bands(0).Columns(UG_Description).Hidden = False
   ugSchemas.Bands(0).Columns(UG_Description).Width = 5000
   ugSchemas.Bands(0).Columns(UG_Description).Header.Caption = "Description"
   
'   ugSchemas.Bands(0).Columns(UG_Message).Hidden = False
'   ugSchemas.Bands(0).Columns(UG_Message).Width = 15000
'   ugSchemas.Bands(0).Columns(UG_Message).Header.Caption = "Error Raised"
   
   On Error GoTo FunctionError
   Set moSchemaUpdateWork = New mwSchemaUpdateWork
   
   goSession.SetDotNetTheme Me
   
   InitForm = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in frmSchemaProgress.InitForm.", Err.Number, Err.Description
   InitForm = False
End Function


Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()

   Me.Hide
   
End Sub
Private Sub CmdPause_Click()

   mIsPaused = True
   cmdPause.Visible = False
   lblStatus.Caption = "Pause button pressed, Run All halted. Press Run All or Run One to continue, Cancel to exit."
   
End Sub

Private Sub CmdRunAll_Click()

   cmdRunAll.Enabled = False
   cmdRunOne.Enabled = False
   
   cmdPause.Visible = True
   mIsPaused = False
   
   moRS.MoveFirst
   mIsPaused = False
   
   Do While moRS.EOF = False And mIsPaused = False
   
      If moRS!Status = "" Then
         If moSchemaUpdateWork.ProcessSchema(moRS, False) Then
            lblStatus.Caption = "Error encountered, can not continue processing. Please press the Ok button to exit."
            cmdRunAll.Visible = False
            cmdRunOne.Visible = False
            cmdCancel.Visible = False
            cmdPause.Visible = False
            
            cmdOK.Visible = True
            cmdRunAll.Enabled = True
            cmdRunOne.Enabled = True
            Exit Sub
         End If
         DoEvents
         Me.Refresh
         
      End If
      moRS.MoveNext
   Loop
   
   If mIsPaused = False Then
   
      lblStatus.Caption = "All schemas processed. Please press the Ok button to exit."
      cmdRunAll.Visible = False
      cmdRunOne.Visible = False
      cmdCancel.Visible = False
      cmdPause.Visible = False
      
      cmdOK.Visible = True
   End If
   
   cmdRunAll.Enabled = True
   cmdRunOne.Enabled = True
End Sub
Private Sub CmdRunOne_Click()

   cmdRunOne.Enabled = False
   cmdRunAll.Enabled = False
      
   cmdPause.Visible = False
   mIsPaused = False
   
   moRS.MoveFirst
   mIsPaused = False
   
   Do While moRS.EOF = False
   
      If moRS!Status = "" Then
         If moSchemaUpdateWork.ProcessSchema(moRS, False) Then
            lblStatus.Caption = "Error encountered, can not continue processing. Please press the Ok button to exit."
            cmdRunAll.Visible = False
            cmdRunOne.Visible = False
            cmdCancel.Visible = False
            cmdPause.Visible = False
            
            cmdOK.Visible = True
            
            cmdRunOne.Enabled = True
            cmdRunAll.Enabled = True
            Exit Sub
         Else
            moRS.MoveNext
            If moRS.EOF = True Then
               lblStatus.Caption = "All schemas processed. Please press the Ok button to exit."
               cmdRunAll.Visible = False
               cmdRunOne.Visible = False
               cmdCancel.Visible = False
               cmdPause.Visible = False
               
               cmdOK.Visible = True
            End If
            cmdRunOne.Enabled = True
            cmdRunAll.Enabled = True

            Exit Sub
         End If
      End If
      moRS.MoveNext
   Loop
   
   lblStatus.Caption = "All schemas processed. Please press the Ok button to exit."
   cmdRunAll.Visible = False
   cmdRunOne.Visible = False
   cmdCancel.Visible = False
   cmdPause.Visible = False
   
   cmdOK.Visible = True
   cmdRunOne.Enabled = True
   cmdRunAll.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set moSchemaUpdateWork = Nothing
End Sub

Private Sub ugSchemas_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
   
   If ugSchemas.Selected.Rows.Count > 0 Then
      If Len(BlankNull(ugSchemas.Selected.Rows(0).Cells(UG_Message).value)) > 0 Then
         goSession.GUI.ImprovedMsgBox ugSchemas.Selected.Rows(0).Cells(UG_Message).value, vbOKOnly, "Error Processing Schema"
      End If
   End If
   
End Sub

Private Sub ugSchemas_DblClick()
'   Me.Hide
End Sub
Private Sub ChooseIconColor(ByVal Row As UltraGrid.SSRow)
   If Not IsNull(Row.Cells(UG_Status).value) Then
      
      If Row.Cells(UG_Status).value = "" Then
         Row.Cells(UG_Status).Appearance.Picture = ugSchemas.Images("ICON_Clear").Picture
      ElseIf Row.Cells(UG_Status).value = "Complete" Then
         Row.Cells(UG_Status).Appearance.Picture = ugSchemas.Images("ICON_Green").Picture
      ElseIf Row.Cells(UG_Status).value = "Error" Then
         Row.Cells(UG_Status).Appearance.Picture = ugSchemas.Images("ICON_Red").Picture
      ElseIf Row.Cells(UG_Status).value = "Info" Then
         Row.Cells(UG_Status).Appearance.Picture = ugSchemas.Images("ICON_Yellow").Picture
      End If
      
   End If
End Sub
Private Sub ugSchemas_AfterRowUpdate(ByVal Row As UltraGrid.SSRow)
   
   ChooseIconColor Row
   
End Sub

Private Sub ugSchemas_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
   
   ChooseIconColor Row
   
End Sub

