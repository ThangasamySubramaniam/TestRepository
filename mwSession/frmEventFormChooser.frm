VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmEventFormChooser 
   Caption         =   "Select Event Submitted Form"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPreviewForm 
      Caption         =   "Preview Form"
      Height          =   855
      Left            =   3600
      Picture         =   "frmEventFormChooser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Preview the highlighted form"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   855
      Left            =   7080
      Picture         =   "frmEventFormChooser.frx":0498
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Selected highlighted record"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   855
      Left            =   960
      Picture         =   "frmEventFormChooser.frx":1162
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   855
      Left            =   2310
      Picture         =   "frmEventFormChooser.frx":146C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   5760
      Picture         =   "frmEventFormChooser.frx":1776
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancel Select "
      Top             =   5280
      Width           =   1095
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   9240
      DesignHeight    =   6285
   End
   Begin UltraGrid.SSUltraGrid ug1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8281
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
      ColScrollRegions=   "frmEventFormChooser.frx":1A80
      Override        =   "frmEventFormChooser.frx":1ABE
      Caption         =   "Select Event Submitted Form"
   End
End
Attribute VB_Name = "frmEventFormChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmEventFormChooser - Choose Submitted forms for same event type...
Option Explicit


Dim moRS As Recordset
Dim mIsCancelled As Boolean
Dim mSearch As String

Const UG_FORM_ID = 0
Const UG_FORM_FormName = 1
Const UG_FORM_Subject = 2
Const UG_FORM_SubmittedDateTime = 3
Const UG_FORM_FormID = 4

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function FormInitChooser(EventType As Long, TemplateID As String, SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loVfWork As mwEventFormWork
   On Error GoTo FunctionError
   Me.Caption = "Select Event Submitted Form"
   ug1.Caption = "Select Event Submitted Form"
   Set loVfWork = New mwEventFormWork
   Set moRS = loVfWork.FetchSubmittedFormsRS(EventType, TemplateID, SiteKey)
   If moRS.RecordCount < 1 Then
      MsgBox "No submitted forms are available for EventType: " & goSession.EventTypes.Item(EventType).Description, vbInformation
      mIsCancelled = True
      CloseRecordset moRS
      FormInitChooser = False
      Exit Function
   End If
   Set ug1.DataSource = moRS
   HideUltragridColumns ug1, 0
   ug1.Override.SelectTypeRow = ssSelectTypeSingle
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Template Id
   ug1.Bands(0).Columns(UG_FORM_FormName).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_FormName).Width = 1400
   ug1.Bands(0).Columns(UG_FORM_FormName).Header.Caption = "Template ID"
   ' Description
   ug1.Bands(0).Columns(UG_FORM_Subject).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_Subject).Width = 3000
   ug1.Bands(0).Columns(UG_FORM_Subject).Header.Caption = "Description"
   ' Submitted
   ug1.Bands(0).Columns(UG_FORM_SubmittedDateTime).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_SubmittedDateTime).Width = 2400
   ug1.Bands(0).Columns(UG_FORM_SubmittedDateTime).Header.Caption = "Submitted"
   ' Form Id
   ug1.Bands(0).Columns(UG_FORM_FormID).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_FormID).Width = 1800
   ug1.Bands(0).Columns(UG_FORM_FormID).Header.Caption = "Form ID"
   
   goSession.SetDotNetTheme Me
   
   FormInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventFormChooser.FormInitChooser.", Err.Number, Err.Description
   FormInitChooser = False
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdFind_Click()
   On Error GoTo SubError

   mSearch = InputBox("Enter Search Phrase", "Search for Record")
   If Len(mSearch) > 0 Then
      goSession.GUI.TraverseUgSearch mSearch, ug1, True
   End If
   
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmEventFormChooser.cmdFind_Click. ", Err.Number, Err.Description
End Sub

Private Sub cmdFindNext_Click()
   On Error GoTo SubError
   
   If Len(mSearch) > 0 Then
      goSession.GUI.TraverseUgSearch mSearch, ug1, False
   End If
   
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmEventFormChooser.cmdFindNext_Click. ", Err.Number, Err.Description
End Sub

Private Sub cmdOK_Click()
   On Error GoTo SubError

   If ug1.Selected.Rows.Count < 1 Then
      If Not ug1.ActiveRow Is Nothing Then
         ug1.ActiveRow.Selected = True
      End If
   End If
   mIsCancelled = False
   Me.Hide
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmEventFormChooser.cmdOK_Click. ", Err.Number, Err.Description
End Sub

Private Sub cmdPreviewForm_Click()
   Dim moFormMaintenanceWork As Object
   On Error GoTo SubError
   If ug1.ActiveRow Is Nothing Then
      Beep
      Exit Sub
   End If
   ' open the form
   If ZeroNull(ug1.ActiveRow.Cells(UG_FORM_ID).value) > 0 Then
      Set moFormMaintenanceWork = CreateObject("mwManuals.mwFormMaintenanceWork")
      moFormMaintenanceWork.InitSession goSession
      moFormMaintenanceWork.PreviewForm (ug1.ActiveRow.Cells(UG_FORM_ID).value)
      KillObject moFormMaintenanceWork
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmEventFormChooser.cmdPreviewForm_Click: ", Err.Number, Err.Description
   KillObject moFormMaintenanceWork
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   CloseRecordset moRS
End Sub

Private Sub ug1_DblClick()
   Me.Hide
End Sub

Public Function IsSelectedRows() As Boolean
   If ug1.Selected.Rows.Count > 0 Then
      IsSelectedRows = True
   Else
      IsSelectedRows = False
   End If
End Function

Public Function FetchSelected() As Long
   Dim loRow As SSRow
   On Error GoTo FunctionError
   If ug1.Selected.Rows.Count < 1 Then
      FetchSelected = -1
      Exit Function
   End If
   Set loRow = ug1.Selected.Rows(0)
   FetchSelected = loRow.Cells(UG_FORM_ID).value
   loRow.Selected = False
   Set loRow = Nothing
      
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchSelected.", Err.Number, Err.Description
   FetchSelected = -1
End Function

