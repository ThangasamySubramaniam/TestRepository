VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form frmEmailTemplate 
   Caption         =   "Select Email"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   975
      Left            =   240
      Picture         =   "frmEmailTemplate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   975
      Left            =   4440
      Picture         =   "frmEmailTemplate.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin UltraGrid.SSUltraGrid ug 
      Height          =   2655
      Left            =   41
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4683
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   1048596
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "frmEmailTemplate.frx":0FD4
   End
End
Attribute VB_Name = "frmEmailTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moRS As Recordset
Public mIsCancelled As Boolean
Public mEmailTemplatePath As String
Public mEmailAddress As String

' constants from mwSession.mwDataForm
   Const MWWF_EMAILTEMPLATETITLE = 0
   Const MWWF_EMAILTEMPLATEPATH = 1
   Const MWWF_TEMPLATEID = 2
   Const MWWF_EmailAddress = 3


Public Function InitForm(Rs As Recordset) As Boolean
   On Error GoTo FunctionError
   
   If IsNull(Rs) Then
      Exit Function
   End If
   
   CloseRecordset moRS
   Set moRS = Rs
   
   ' refreshUgColumns
   Set ug.DataSource = moRS
   
   If moRS.RecordCount = 1 Then
      Set ug.ActiveRow = ug.GetRow(ssChildRowFirst)
      ug.ActiveRow.Selected = True
   End If
   
   HideUltragridColumns ug, 0
   ug.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug.Bands(0).Columns(MWWF_EMAILTEMPLATETITLE).Hidden = False
   ug.Bands(0).Columns(MWWF_EMAILTEMPLATETITLE).Width = 4000
   ug.Bands(0).Columns(MWWF_EMAILTEMPLATETITLE).Header.Caption = "Title"
   ' Description
'   ug.Bands(0).Columns(MWWF_TEMPLATEID).Hidden = False
'   ug.Bands(0).Columns(MWWF_TEMPLATEID).Width = 1400
'   ug.Bands(0).Columns(MWWF_TEMPLATEID).Header.Caption = "Template ID"
   
   goSession.SetDotNetTheme Me
   
   InitForm = True

   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.InitForm: ", Err.Number, Err.Description
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   mEmailTemplatePath = ""
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   If Not ug.ActiveRow Is Nothing Then
      mEmailTemplatePath = ug.ActiveRow.Cells(MWWF_EMAILTEMPLATEPATH).value
      mEmailAddress = BlankNull(ug.ActiveRow.Cells(MWWF_EmailAddress).value)
   End If
   
   Me.Hide
   mIsCancelled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo SubError
   CloseRecordset moRS
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.frmEmailTemplate: ", Err.Number, Err.Description
   CloseRecordset moRS
End Sub


Public Function GetEmailTemplateRs() As Recordset
   Dim loRs As Recordset
   Dim loRow As SSRow
   On Error GoTo FunctionError
   
   If Not ug.ActiveRow Is Nothing Then
      If ug.ActiveRow.Selected = False Then
         ug.ActiveRow.Selected = True
      End If
   End If
   
   If ug.Selected.Rows.Count < 1 Then
      Exit Function
   End If
   
       
   Set loRs = New Recordset
   loRs.Fields.Append "EmailTemplatePath", adVarChar, 200, adFldIsNullable And adFldUpdatable And adFldMayBeNull
   loRs.Fields.Append "EmailAddress", adVarChar, 100, adFldIsNullable And adFldUpdatable And adFldMayBeNull
   
   loRs.Open
   For Each loRow In ug.Selected.Rows
      loRs.AddNew
      If Not IsNull(loRow.Cells(MWWF_EMAILTEMPLATEPATH).value) Then
         loRs!EmailTemplatePath = loRow.Cells(MWWF_EMAILTEMPLATEPATH).value
      End If
      If Not IsNull(loRow.Cells(MWWF_EmailAddress).value) Then
         loRs!EmailAddress = loRow.Cells(MWWF_EmailAddress).value
      End If
      loRs.Update
   Next loRow
   
   Set GetEmailTemplateRs = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwSession.frmEmailTemplate.GetEmailTemplateRs. ", Err.Number, Err.Description
End Function
