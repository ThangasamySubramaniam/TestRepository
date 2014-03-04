VERSION 5.00
Object = "{22FCC75D-5FDD-4E46-8C0F-E178216EB1B0}#5.0#0"; "mwEventReports.ocx"
Begin VB.Form frmEventReports 
   Caption         =   "Available Reports"
   ClientHeight    =   5964
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8088
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5964
   ScaleWidth      =   8088
   StartUpPosition =   1  'CenterOwner
   Begin mwEventReports.mwEventReports_ocx mwER 
      Height          =   4572
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   8065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Picture         =   "frmEventReports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmEventReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' frmEventReports - Wrapper for Event Reports OCX
' 10/28/2003 ms
'
Option Explicit

Dim mIsModal As Boolean


Private Sub cmdOK_Click()
   Me.Hide
End Sub


Public Function InitForm(FormTitle As String, EventType As Long, _
 SiteKey As Long, Optional EventSubType As Long, _
 Optional IsModal As Boolean, Optional DetailKey As Long, Optional DetailKey2 As Long, Optional DetailKey3 As Long) As Boolean

   On Error GoTo FunctionError
   '
   ' Initialize the OCX..
   '
   goSession.SetDotNetTheme Me
   mIsModal = IsModal
   Me.Caption = FormTitle
   If Not mwER.InitControl(FormTitle, goSession, EventType, SiteKey, EventSubType, mIsModal) Then
      MsgBox "These reports can not be displayed because an error occurred.", vbExclamation, "Report Initialization Error"
      Me.Hide
      InitForm = False
   Else
      mwER.SetPK DetailKey, DetailKey2, DetailKey3
      InitForm = True
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventReports.InitControl: ", Err.Number, Err.Description
   InitForm = False
End Function

