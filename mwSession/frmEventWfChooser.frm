VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form frmEventWfChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Position in Workflow"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin UltraGrid.SSUltraGrid ug 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "frmEventWfChooser.frx":0000
      Caption         =   "ug"
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   855
      Left            =   360
      Picture         =   "frmEventWfChooser.frx":0056
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   2160
      Picture         =   "frmEventWfChooser.frx":0360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   855
      Left            =   3960
      Picture         =   "frmEventWfChooser.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Caption         =   "Select Workflow Stage"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmEventWfChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mIsCancelled As Boolean
Dim mHelpFileID As String
Dim mHelpContextID As String

' mwEventWfStage
   Const UG_ID = 0
   Const UG_DisplaySequence = 1
   Const UG_mwEventGroupKey = 2
   Const UG_WfStageName = 3
   
Public RecordCount As Long


Public Function InitForm(EventGroup As Long, Optional UpdateStagesOnly As Boolean) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError:
      
   If EventGroup < 1 Then
      Exit Function
   End If
   
   If UpdateStagesOnly Then
      'sSQL = "SELECT DISTINCT mwEventWfStage.ID, mwEventWfStage.DisplaySequence, mwEventWfStage.mwEventGroupKey, " _
       '  & " mwEventWfStage.WfStageName, mwEventWfStagePermissions.IsUpdateAllowed FROM mwEventWfStage " _
       '  & " INNER JOIN mwEventWfStagePermissions ON mwEventWfStage.ID = mwEventWfStagePermissions.mwEventWfStageKey " _
       '  & " WHERE mwEventWfStage.mwEventGroupKey = " & EventGroup & "  And mwEventWfStagePermissions.IsUpdateAllowed = True " _
       '  & "  AND mwEventWfStagePermissions.mwcRoleTypeKey = " & goSession.User.RoleTypeKey _
       '  & " ORDER BY mwEventWfStage.DisplaySequence "
      sSQL = "SELECT DISTINCT mwEventWfStage.ID, mwEventWfStage.DisplaySequence, mwEventWfStage.mwEventGroupKey,  " _
        & " mwEventWfStage.WfStageName, mwEventWfStagePermissions.IsUpdateAllowed FROM mwEventWfStage " _
        & " INNER JOIN mwEventWfStagePermissions ON mwEventWfStage.ID = mwEventWfStagePermissions.mwEventWfStageKey " _
        & " Where mwEventWfStage.mwEventGroupKey = " & EventGroup _
        & " And mwEventWfStagePermissions.IsUpdateAllowed Is Not Null And mwEventWfStagePermissions.IsUpdateAllowed <> 0" _
        & " And mwEventWfStagePermissions.mwcRoleTypeKey = " & goSession.User.RoleTypeKey _
        & " ORDER BY mwEventWfStage.DisplaySequence "
 
   Else
      sSQL = "SELECT ID, DisplaySequence, mwEventGroupKey, WfStageName " _
         & " FROM mwEventWfStage WHERE mwEventGroupKey = " & EventGroup _
         & " ORDER BY DisplaySequence"
   End If
      
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRs.RecordCount > 0 Then
      RecordCount = loRs.RecordCount
   End If
   Set ug.DataSource = loRs
   HideUltragridColumns ug, 0
   ' Expose selected columns...
   If ug.Bands.Count = 0 Then
      Exit Function
   End If
   HideUltragridColumns ug, 0
   ug.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   
   ug.Bands(0).Columns(UG_WfStageName).Hidden = False
   ug.Bands(0).Columns(UG_WfStageName).Width = 4100
   ug.Bands(0).Columns(UG_WfStageName).Header.Caption = "Workflow Stage"
   
   goSession.SetDotNetTheme Me
   
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaisePublicError "General Error in frmMwEventWfChooser.InitForm. ", Err.Number, Err.Description
   mIsCancelled = True
   Me.Hide
End Function




Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdHelp_Click()
    goSession.API.ShowVbFormHelp Me.Name
End Sub

Private Sub cmdOK_Click()
   If ug.ActiveRow Is Nothing Then
      MsgBox "You must select from the list, or click cancel.", vbExclamation, "Select Event Workflow Stage"
      Exit Sub
   End If
   Me.Hide
End Sub


Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function GetEventWfStageKey() As Long
   On Error GoTo FunctionError
   If Not ug.ActiveRow Is Nothing Then
      GetEventWfStageKey = ug.ActiveRow.Cells(UG_ID).value
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in frmMwEventWfChooser.GetEventWfStageKey. ", Err.Number, Err.Description
End Function

Private Sub ug_DblClick()
   Me.Hide
End Sub

'Public Function SetHelpFileID(HelpFileID As String, Optional HelpContextID As String)
'   On Error GoTo FunctionError
'   mHelpFileID = HelpFileID
'   If Not HelpContextID = "" Then
'      mHelpContextID = HelpContextID
'   End If
'   Exit Function
'FunctionError:
'   goSession.RaisePublicError "General Error in frmMwEventWfChooser.SetHelpFileID. ", Err.Number, Err.Description
'End Function
