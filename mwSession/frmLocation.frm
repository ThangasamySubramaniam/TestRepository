VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmLocation 
   Caption         =   "Select Location"
   ClientHeight    =   6216
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7248
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6216
   ScaleWidth      =   7248
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   696
      Left            =   4277
      Picture         =   "frmLocation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5436
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   696
      Left            =   1997
      Picture         =   "frmLocation.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5436
      Width           =   1095
   End
   Begin VB.DirListBox Folder1 
      Height          =   2184
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   6972
   End
   Begin VB.DriveListBox drive1 
      Height          =   336
      Left            =   120
      TabIndex        =   0
      Top             =   1992
      Width           =   6972
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   132
      Top             =   5724
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   12
      DesignWidth     =   7248
      DesignHeight    =   6216
   End
   Begin VB.Label lbltarget 
      BorderStyle     =   1  'Fixed Single
      Height          =   408
      Left            =   2508
      TabIndex        =   7
      Top             =   4776
      Width           =   4584
   End
   Begin VB.Label Label1 
      Caption         =   "Selected Location:"
      Height          =   408
      Left            =   120
      TabIndex        =   6
      Top             =   4776
      Width           =   2268
   End
   Begin VB.Label lblOptCaption 
      Height          =   1404
      Left            =   120
      TabIndex        =   3
      Top             =   456
      Visible         =   0   'False
      Width           =   6972
   End
   Begin VB.Label lblDefaultCaption 
      Alignment       =   2  'Center
      Caption         =   "Choose Location"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2478
      TabIndex        =   2
      Top             =   72
      Width           =   2292
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmLocation - Form to Select a Folder Location
' 10/26/2001 ms  Maritime Systems Inc
'

Option Explicit
Dim mDrive As String
Dim mPath As String
Dim moParent As Session
Dim mIsCancelled As Boolean
Dim mIsNotSendByMedia As Boolean

Public Sub SetOptCaption(Caption As String)
   lblOptCaption.Caption = Caption
   lblOptCaption.Visible = True
End Sub
Public Function SetParent(ByRef ses As Session)
   Set moParent = ses
End Function

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function SetDrive(Drive As String)
   On Error Resume Next
   drive1.Drive = Drive
   mDrive = Drive
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLocation.SetDrive. ", err.Number, err.Description
   mIsCancelled = True
   Me.Hide
End Function

Public Function SetFolder(DefaultFolder As String)
   On Error Resume Next
   Folder1.Path = DefaultFolder
End Function

Public Function GetPath() As String
   GetPath = mPath
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   On Error GoTo SubError
'   mPath = Folder1.Path
   Me.Hide
   Exit Sub
SubError:
   moParent.RaiseError "General Error in frmLocation.cmdOK_Click. ", err.Number, err.Description
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub drive1_Change()
   On Error GoTo SubError
   Folder1.Path = drive1.Drive
   mPath = drive1.Drive
   mDrive = drive1.Drive
   lbltarget.Caption = mPath
   Exit Sub
SubError:
   If err.Number = 68 Then
      On Error GoTo SubError2
      MsgBox "Drive " & drive1.Drive & " is Unavailable: ", vbExclamation
      'MsgBox mPath
      Folder1.Path = mPath
      drive1.Drive = Left(mPath, 2)
      Exit Sub
   Else
      moParent.RaiseError "General Error in frmLocation.SetDrive. ", err.Number, err.Description
   End If
   mIsCancelled = True
   Me.Hide
SubError2:
   MsgBox "sub error 2" & err.Number & err.Description
   moParent.RaiseError "General Error in frmLocation.drive1_change_SubError2. ", err.Number, err.Description
   mIsCancelled = True
   Me.Hide

End Sub

Private Sub Folder1_Change()
   mPath = Folder1.Path
   lbltarget.Caption = mPath
End Sub


Private Sub Form_Activate()
   Dim strFolder As String
   On Error GoTo SubError
   If mIsNotSendByMedia Then
      Exit Sub
   End If
   strFolder = moParent.Workflow.SendByMediaFolder
   If Trim(strFolder) <> "" Then
      mPath = strFolder
      drive1.Drive = strFolder
      Folder1.Path = strFolder
   Else
      mPath = ""
   End If
   
   lbltarget.Caption = mPath
   
   moParent.SetDotNetTheme Me
   
   Exit Sub
SubError:
   moParent.RaiseError "General error in frmLocation. ", err.Number, err.Description
   mIsCancelled = True
   Me.Hide
End Sub

Public Function SetIsNotSendByMedia() As Boolean
   mIsNotSendByMedia = True
End Function


Private Sub Form_Load()
   On Error GoTo SubError
   moParent.SetDotNetTheme Me
   Exit Sub
SubError:
   moParent.RaiseError "General error in frmLocation. ", Err.Number, Err.Description
End Sub

