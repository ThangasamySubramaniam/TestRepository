VERSION 5.00
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frmAlertNotifyInterval 
   Caption         =   "Change Alert Notify Interval (Minutes) "
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin PVNumericLib.PVNumeric pvnInterval 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2535
      _Version        =   524288
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   253
      Text            =   "0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Alignment       =   2
      SpinButtons     =   0
      LimitValue      =   -1  'True
      ValueMin        =   0
      ValueMax        =   1440
      LimitValueByType=   5
      DecimalMax      =   0
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Picture         =   "frmAlertNotifyInterval.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Picture         =   "frmAlertNotifyInterval.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "(Range: 0 - 1440 Minutes)  "
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
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblUserID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Admin"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Current User ID:"
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
      Index           =   3
      Left            =   300
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Interval Minutes:"
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
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   2655
   End
End
Attribute VB_Name = "frmAlertNotifyInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mIsCancelled As Boolean

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdChange_Click()
'
   On Error GoTo SubError
   
   If UpdateAlertNotifyInterval() Then
      goSession.GUI.ImprovedMsgBox "Alert Notify Interval has been updated.", vbInformation, "Change Alert Notify Interval (Minutes)"
   Else
      mIsCancelled = True
      goSession.GUI.ImprovedMsgBox "Alert Notify Interval has not been updated.", vbInformation, "Change Alert Notify Interval (Minutes)"
   End If
   Me.Hide
   
   Exit Sub
SubError:
   goSession.RaiseError "General Error in msWorkstation.frmAlertNotifyInterval.cmdChange_Click ", Err.Number, Err.Description
End Sub

Private Function UpdateAlertNotifyInterval() As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo ErrorHandler

   strSQL = "SELECT * FROM mwcUsers WHERE ID=" & goSession.User.UserKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic

   loRs!AlertNotifyIntervalMinutes = pvnInterval.ValueInteger
   loRs.Update
   
   CloseRecordset loRs
   
   goSession.User.SetExtendedProperty "AlertNotifyIntervalMinutes", pvnInterval.ValueInteger
   UpdateAlertNotifyInterval = True
   
   Exit Function
ErrorHandler:
   UpdateAlertNotifyInterval = False
   goSession.RaisePublicError "General Error in msWorkstation.frmAlertNotifyInterval.UpdateAlertNotifyInterval.", Err.Number, Err.Description
End Function

Private Sub Form_Load()
   On Error GoTo SubError
   
   mIsCancelled = False
   lblUserID.Caption = goSession.User.UserID
   pvnInterval.Text = ZeroNull(goSession.User.GetExtendedProperty("AlertNotifyIntervalMinutes"))
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   MsgBox "Error in mwSession.frmAlertNotifyInterval.Form_Load ", Err.Number, Err.Description
End Sub
