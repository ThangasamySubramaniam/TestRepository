VERSION 5.00
Begin VB.Form frmImprovedMsgBox 
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton1 
      Caption         =   "OK"
      Height          =   975
      Left            =   7920
      Picture         =   "frmImprovedMsgBox.frx":0000
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "OK"
      Height          =   975
      Left            =   6540
      Picture         =   "frmImprovedMsgBox.frx":0CCA
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton3 
      Caption         =   "OK"
      Height          =   975
      Left            =   5160
      Picture         =   "frmImprovedMsgBox.frx":1994
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image FormIcon 
      Height          =   600
      Left            =   120
      Picture         =   "frmImprovedMsgBox.frx":265E
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Width           =   8115
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmImprovedMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRetode As Integer

Public Function WhichButton() As Integer
   WhichButton = mRetode
End Function

Private Sub cmdButton3_Click()
   mRetode = cmdButton3.Tag
   Me.Hide
End Sub
Private Sub cmdButton2_Click()
   mRetode = cmdButton2.Tag
   Me.Hide
End Sub
Private Sub cmdButton1_Click()
   mRetode = cmdButton1.Tag
   Me.Hide
End Sub

'      Buttons
'      vbOKOnly                     0 Display OK button only.
'      vbOKCancel                   1 Display OK and Cancel buttons.
'      vbAbortRetryIgnore           2 Display Abort, Retry, and Ignore buttons.
'      vbYesNoCancel                3 Display Yes, No, and Cancel buttons.
'      vbYesNo                      4 Display Yes and No buttons.
'      vbRetryCancel                5 Display Retry and Cancel buttons.

'      vbCritical                  16 Display Critical Message icon.
'      vbQuestion                  32 Display Warning Query icon.
'      vbExclamation               48 Display Warning Message icon.
'      vbInformation               64 Display Information Message icon.

'      vbDefaultButton1             0 First button is default.
'      vbDefaultButton2           256 Second button is default.
'      vbDefaultButton3           512 Third button is default.
'      vbDefaultButton4           768 Fourth button is default.

'      vbApplicationModal           0 Application modal; the user must respond to the message box before continuing work in the current application.
'      vbSystemModal             4096 System modal; all applications are suspended until the user responds to the message box.
'      vbMsgBoxHelpButton       16384 Adds Help button to the message box
'      VbMsgBoxSetForeground    65536 Specifies the message box window as the foreground window
'      vbMsgBoxRight           524288 Text is right aligned
'      vbMsgBoxRtlReading     1048576 Specifies text should appear as right-to-left reading on Hebrew and Arabic systems

'      Return Values
'
'      vbOK       1 OK
'      vbCancel   2 Cancel
'      vbAbort    3 Abort
'      vbRetry    4 Retry
'      vbIgnore   5 Ignore
'      vbYes      6 Yes
'      vbNo       7 No

Public Function InitForm(Prompt As String, Buttons As Integer, Title As String) As Boolean
   Dim ButtonChoice As Integer
   Dim ImageChoice As Integer
   Dim ButtonBottom As Long
   
   On Error GoTo FunctionError
   
   ButtonChoice = Buttons Mod 16
   
   Select Case ButtonChoice
      Case vbOKOnly
         cmdButton3.Visible = False
         cmdButton2.Visible = False
         cmdButton1.Visible = True
         cmdButton1.Caption = "OK"
         cmdButton1.Tag = vbOK
         
      Case vbOKCancel
         cmdButton3.Visible = False
         cmdButton2.Visible = True
         cmdButton1.Visible = True
         cmdButton2.Caption = "OK"
         cmdButton2.Tag = vbOK
         
         cmdButton1.Caption = "Cancel"
         cmdButton1.Tag = vbCancel
      Case vbAbortRetryIgnore
         cmdButton3.Visible = True
         cmdButton2.Visible = True
         cmdButton1.Visible = True
         
         cmdButton3.Caption = "Abort"
         cmdButton3.Tag = vbAbort
         
         cmdButton2.Caption = "Retry"
         cmdButton2.Tag = vbRetry
         
         cmdButton1.Caption = "Ignore"
         cmdButton1.Tag = vbIgnore
      Case vbYesNoCancel
         cmdButton3.Visible = True
         cmdButton2.Visible = True
         cmdButton1.Visible = True
         
         cmdButton3.Caption = "Yes"
         cmdButton3.Tag = vbYes
         
         cmdButton2.Caption = "No"
         cmdButton2.Tag = vbNo
         
         cmdButton1.Caption = "Cancel"
         cmdButton1.Tag = vbCancel
      Case vbYesNo
         cmdButton3.Visible = False
         cmdButton2.Visible = True
         cmdButton1.Visible = True
         
         cmdButton2.Caption = "Yes"
         cmdButton2.Tag = vbYes
         
         cmdButton1.Caption = "No"
         cmdButton1.Tag = vbNo
      Case vbRetryCancel
         cmdButton3.Visible = False
         cmdButton2.Visible = True
         cmdButton1.Visible = True
         
         cmdButton2.Caption = "Retry"
         cmdButton2.Tag = vbRetry
         
         cmdButton1.Caption = "Cancel"
         cmdButton1.Tag = vbCancel
   End Select
   
'      vbCritical                  16 Display Critical Message icon.
'      vbQuestion                  32 Display Warning Query icon.
'      vbExclamation               48 Display Warning Message icon.
'      vbInformation               64 Display Information Message icon.
   
   ImageChoice = (Buttons - ButtonChoice) Mod 256
   
   If ImageChoice = vbCritical Then
      GetIcon goSession.GetAppPath() & "\icons\32x32\Critical_Bubble.ico"
   ElseIf ImageChoice = vbQuestion Then
      GetIcon goSession.GetAppPath() & "\icons\32x32\Question_Bubble.ico"
   ElseIf ImageChoice = vbExclamation Then
      GetIcon goSession.GetAppPath() & "\icons\32x32\Exclamation_Bubble.ico"
   Else
      GetIcon goSession.GetAppPath() & "\icons\32x32\Info_Bubble.ico"
   End If
   
      
   lblPrompt.Caption = Prompt
   Me.Caption = Title
   
   If Title = "" Then
      Me.Caption = "ShipNet Fleet"
   End If
   
   ButtonBottom = cmdButton1.Top + cmdButton1.Height
   
   cmdButton3.Top = lblPrompt.Top + lblPrompt.Height + 180
   cmdButton2.Top = cmdButton3.Top
   cmdButton1.Top = cmdButton3.Top
   
   ButtonBottom = (cmdButton1.Top + cmdButton1.Height) - ButtonBottom
   Me.Height = Me.Height + ButtonBottom
'   Me.Height = (cmdButton1.Top + cmdButton1.Height + 180) + ButtonBottom
   
   goSession.SetDotNetTheme Me
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in frmImprovedMsgBox.InitForm. ", Err.Number, Err.Description
End Function
   
Private Sub GetIcon(IconFile As String)
   Dim fso As FileSystemObject
   Set fso = New FileSystemObject
   
   On Error Resume Next
   
   If fso.FileExists(IconFile) Then
      FormIcon.Picture = LoadPicture(IconFile)
   End If
   
   KillObject fso
End Sub
