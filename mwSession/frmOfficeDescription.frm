VERSION 5.00
Begin VB.Form frmOfficeDescription 
   Caption         =   "New Form Subject"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSubject 
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
      Left            =   2040
      MaxLength       =   40
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Picture         =   "frmOfficeDescription.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Subject:"
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
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmOfficeDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   txtSubject.Text = ""
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   Me.Hide
End Sub


Private Sub txtSubject_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   On Error GoTo SubError
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSecurity.frmOfficeDescription.Form_Load", Err.Number, Err.Description
End Sub