VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipnet Fleet  Login"
   ClientHeight    =   4776
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6768
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4776
   ScaleWidth      =   6768
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pvmPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox pvmLoginID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1950
      MaxLength       =   16
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   5280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   6768
      DesignHeight    =   4776
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1695
      Picture         =   "frmLogin.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      Picture         =   "frmLogin.frx":0BD4
      ScaleHeight     =   1092
      ScaleWidth      =   2412
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3735
      Picture         =   "frmLogin.frx":1B77
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin PVCOMBOLibCtl.PVComboBox pvCbo 
      Height          =   405
      Left            =   1950
      TabIndex        =   2
      Top             =   2880
      Width           =   4335
      _Version        =   524288
      _cx             =   7646
      _cy             =   714
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   0
      Sorted          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.6
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
      ColumnCount     =   1
      Column0.Heading =   ""
      Column0.Width   =   40
      Column0.Alignment=   0
      Column0.Hidden  =   0   'False
      Column0.Name    =   ""
      Column0.Format  =   ""
      Column0.Bound   =   0   'False
      Column0.Locked  =   0   'False
      Column0.HeaderAlignment=   0
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
   Begin VB.Label lblDatabase 
      Alignment       =   1  'Right Justify
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   225
      TabIndex        =   8
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   225
      TabIndex        =   6
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Login ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   1680
      Width           =   1395
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ShipNet Fleet Login Screen
'
' 9/21/2001 ms  Maritime Systems Inc
'
Option Explicit
Dim mIsCancelled As Boolean
Dim mIsLoggedIn As Boolean
Dim mDbConnection As Connection
Dim mIsMultiDB As Boolean
Dim mIsID As Boolean
Dim mIsPW As Boolean
Dim mIsDB As Boolean
Dim mIsDBS As Boolean
Dim mIsEMP As Boolean
Dim mIsDBOpen As Boolean
Dim mGotNetworkLogin As Boolean
Dim mNetworkLoginID As String
Dim mIsIntegratedWinLogin As Boolean
Dim mRootFolderSemaphore As String

Dim moIni As ConfigGroups
Dim moParent As Session
Dim moDbConnection As Connection

Const EVENTTYPE_MARINE_ASSURANCE_APP = 7000

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
            (ByVal lpBuffer As String, nSize As Long) As Long

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Const BASE_REGISTRY As String = "Software\Maritime Systems Inc"
Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"

Private Function CheckIfFieldExists(ByVal loRs As Recordset, ByVal sFieldName As String) As Boolean
   Dim vTemp As Variant
   On Error GoTo ErrorHandler

   'DEV-1782 Assign User to specific Site
   'Added By N.Angelakis On 07 Jan 2010

   vTemp = loRs(sFieldName)
   CheckIfFieldExists = True
   
Exit Function
ErrorHandler:
   CheckIfFieldExists = False
End Function

Public Function SetParent(ByRef ses As Session)
   Set moParent = ses
End Function

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function IsLoggedIn() As Boolean
   IsLoggedIn = mIsLoggedIn
End Function

Public Function GetDbConnection() As Connection
   If Not mDbConnection Is Nothing Then
      Set GetDbConnection = mDbConnection
   End If
End Function

Public Function PreActivateForm() As Boolean
   On Error GoTo FunctionError
   If mIsDBOpen Then
      pvCbo.Visible = False
      lblDatabase.Visible = False
      
   Else
      If Not LoadDbList() Then
         moParent.RaiseError "Error Logging into ShipNet Fleet, No Database Connections Defined !"
         mIsLoggedIn = False
         Me.Hide
      End If
   End If
   PreActivateForm = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLogin.PreActivateForm.", Err.Number, Err.Description
   PreActivateForm = False
End Function



Public Function LoadDbList() As Boolean
   '
   ' Load Dropdown box
   '
   Dim i As Integer
   On Error GoTo FunctionError
   For i = 2 To moIni.Count
      pvCbo.AddItem moIni(i).ConfigKeys("NAME").KeyValue
   Next i
   If pvCbo.ListCount > 0 Then
      LoadDbList = True
   Else
      LoadDbList = False
   End If
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLogin.LoadDbList.", Err.Number, Err.Description
   LoadDbList = False
End Function


Public Function DetermineLoginParameters() As Boolean

   Dim loRsThisSite As Recordset
   Dim strSQL As String
   Dim loReg As Registry
   Dim strFile As String
   Dim fso As FileSystemObject
   Dim strLoginID As String
   
   On Error GoTo FunctionError
   
   mGotNetworkLogin = False
   
   If mIsLoggedIn Then
      DetermineLoginParameters = True
      Exit Function
   End If
   '
   ' Find the staging parameters...
   '
   ' 1. Via Parameters...
   If Trim(pvmLoginID.Text) <> "" Then
      mIsID = True
   End If
   If Trim(pvmPassword.Text) <> "" Then
      mIsPW = True
   End If
   If Trim(gDBConnectString) <> "" Then
      mIsDB = True
   End If
   If Trim(gDbShapeConnectString) <> "" Then
      mIsDBS = True
   End If
   '
   If Not (mIsID And mIsPW And mIsDB And mIsDBS And mIsEMP) Then
      '
      ' Not everything passed in, so Load the Registry
      '
      Set loReg = New Registry
      loReg.BaseRegistry = BASE_REGISTRY
      If Not mIsID Then
         pvmLoginID.Text = loReg.GetReg("MwUserID")
         If Trim(pvmLoginID.Text) <> "" Then mIsID = True
      End If
      If Not mIsPW Then
         pvmPassword.Text = loReg.GetReg("MwPassword")
         If Trim(pvmPassword.Text) <> "" Then mIsPW = True
      End If
      If Not mIsDB Then
         gDBConnectString = loReg.GetReg("DBConnectString")
         If Trim(gDBConnectString) <> "" Then mIsDB = True
      End If
      If Not mIsDBS Then
         gDbShapeConnectString = loReg.GetReg("DBShapeConnectString")
         If Trim(gDbShapeConnectString) <> "" Then mIsDBS = True
      End If
      '
      If Not (mIsDB And mIsDBS) Then
         '
         ' Last - read the INI file
         '
         strFile = goSession.GetAppPath() & "\" & App.EXEName & ".ini"
         '
         ' File must exist...
         '
         Set fso = New FileSystemObject
         If Not fso.FileExists(strFile) Then
            ' msg box as required error objects do not exist...
            MsgBox "Error in mwSession.frmLogin.DetermineLoginParameters, " & App.EXEName & ".ini" & " missing." & _
            vbCrLf & "Has the baseline database been installed ?", vbCritical
            DetermineLoginParameters = False
            moParent.KillObject fso
            Exit Function
         End If
         moParent.KillObject fso
         '
         Set moIni = moParent.LoadConfigGroupsFile(strFile)
         If moIni Is Nothing Then
            DetermineLoginParameters = False
            Exit Function
         End If
         '
         ' Is there a login section ?
         '
         On Error Resume Next
         strFile = moIni("login").ConfigKeys.Count
         If Err Then
            MsgBox "Unlogged Error in mwSession.frmLogin, [LOGIN] section is missing from " & App.EXEName & ".ini"
            DetermineLoginParameters = False
            Exit Function
         End If
            
         '
         ' Load from INI...
         '
         If Not mIsID Then
            pvmLoginID.Text = moIni("login").ConfigKeys.GetKeyValue("MwUserID")
            If Trim(pvmLoginID.Text) <> "" Then mIsID = True
         End If
         If Not mIsPW Then
            pvmPassword.Text = moIni("login").ConfigKeys.GetKeyValue("MwPassword")
            If Trim(pvmPassword.Text) <> "" Then mIsPW = True
         End If
         If Not mIsDB Then
            gDBConnectString = moIni("login").ConfigKeys.GetKeyValue("DBConnectString")
            If Trim(gDBConnectString) <> "" Then mIsDB = True
         End If
         If Not mIsDBS Then
            gDbShapeConnectString = moIni("login").ConfigKeys.GetKeyValue("DBShapeConnectString")
            If Trim(gDbShapeConnectString) <> "" Then mIsDBS = True
         End If
         ' Are there multiple database connections in the INI file ?
         If moIni.Count > 2 Then
            mIsMultiDB = True
         End If
      End If
   End If
   '
   DetermineLoginParameters = True
   '
   ' Set Database Connection strings in the Session Object...
   '
   moParent.DBConnectString = gDBConnectString
   moParent.DbShapeConnectString = gDbShapeConnectString
   
   
   Exit Function
FunctionError:
'Resume Next

   moParent.RaiseError "General Error in frmLogin.DetermineLoginParameters: ", Err.Number, Err.Description
   DetermineLoginParameters = False
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdLogin_Click()
   AttemptLogin
End Sub

Private Function AttemptLogin()
   ' At this point, let the checks begin...
   Dim fso As FileSystemObject
   Dim iLogin As Integer
   Dim sLogId As String
   mIsID = True
   mIsPW = True
   mIsDB = True
   mIsDBS = True
   mIsEMP = True
   If ValidateUser() Then
      
      Set fso = New FileSystemObject
      
      If mRootFolderSemaphore <> "" Then
         If Not fso.FileExists(mRootFolderSemaphore) Then
            iLogin = MsgBox("Warning: Semaphore File name: """ & mRootFolderSemaphore & """ does not exist in the root folder." & vbCrLf & _
               "Please make sure that you are connected to the correct Root Folder Tree before running ShipNet Fleet." & vbCrLf & _
               "If you are connected to the wrong Root Folder Tree, your session will be using files and making files in the wrong place, which can cause unpredicted results." & vbCrLf & vbCrLf & _
               "Do you still want to continue?", vbYesNo + vbQuestion, Me.Caption)
            If iLogin = vbNo Then
               mIsCancelled = True
               mIsID = False
               mIsLoggedIn = False
               moParent.KillObject fso
               Me.Hide
               Exit Function
            End If
         End If
      End If
      moParent.KillObject fso
      
      mIsCancelled = False
      Me.Hide
   
      ' check LoginID registrykey to avoid multiple logins
      sLogId = UCase(Trim(pvmLoginID.Text))
      
' This code replaced with similar functionality in msWorkstation.frmWorkflowAgent
' For SF-9999
      
'      If sLogId = "AGENT" Or sLogId = "VESSELAGENT" Or sLogId = "MASTER" Then
'
'         If IsWorkflowUserActive(sLogId) Then
'            iLogin = MsgBox("Warning: There is already a user logged in running the Workflow Agent..." & vbCrLf & _
'               "Running multiple copies of the Workflow Agent can create an unstable condition." & vbCrLf & vbCrLf & _
'               "Do you still want to login as " & Trim(pvmLoginID.Text) & "?", vbYesNo, "Multiple Workflow Agent Users warning")
'            If iLogin = vbNo Then
'               mIsCancelled = True
'               mIsID = False
'               mIsLoggedIn = False
'               Me.Hide
'               Exit Function
'            End If
'         End If
'      End If

      SetActiveUser sLogId
      
      'By N.Angelakis On 22 April 2009
      'DEV-1174 Advance Password Settings
      If moParent.ThisSite.IsPasswordStrong = True Then 'strong passwords enabled
         Dim intCalcDays As Integer
         If moParent.User.IsShoreUser = True And moParent.ThisSite.PasswordExpireNoDays > 0 And DateDiff("d", 0, moParent.User.PasswordLastChangedDate) > 0 Then
            'calculate days between now and (add number of days till expire to date original changed password)
            intCalcDays = DateDiff("d", Now, DateAdd("d", moParent.ThisSite.PasswordExpireNoDays, moParent.User.PasswordLastChangedDate))
            Select Case intCalcDays
               Case 1 To 10
                  'date is within last 10 days show alert or change password
                  Select Case goSession.GUI.ImprovedMsgBox("Your password will Expire in  " & intCalcDays & " days." & vbCrLf & vbCrLf & "Would you like to change it now?", vbYesNoCancel + vbInformation, Me.Caption)
                     Case vbNo, vbCancel
                        'password will expire within next 10 days, send alert to remind user
                        Call SendPasswordAlert(intCalcDays, moParent.User.UserID)

                     Case vbYes
                        'show change password screen
                        If Not goSession.Encrypt.ChangePasswordX(False) Then
                           'user cancelled to change password so send alert to remind expiry
                           Call SendPasswordAlert(intCalcDays, moParent.User.UserID)
                        End If
                  End Select
                  
               Case Is < 0
                  'date has passed , show change password screen
                  goSession.Encrypt.ChangePasswordX True
            End Select
         ElseIf DateDiff("d", 0, moParent.User.PasswordLastChangedDate) = 0 Then
            'advanced security settings enabled, but no date found for password last changed date for
            'current user. Maybe first time using so force change password to include advanced options
            goSession.Encrypt.ChangePasswordX True
         End If
      ElseIf DateDiff("d", 0, moParent.User.PasswordLastChangedDate) = 0 Then
         'forced change password at next logon was set
         If moParent.ThisSite.IsPasswordStrong = True Then 'strong passwords enabled
            goSession.Encrypt.ChangePasswordX True
         End If
      End If
   Else
      If mIsDBOpen Then
         MsgBox "Connected to database, but user login failed."
         pvCbo.Visible = False
         lblDatabase.Visible = False
      Else
         MsgBox "Login Failed. Database Connection Not established."
         mIsDB = False
         mIsDBS = False
         gDBConnectString = ""
         gDbShapeConnectString = ""
         moParent.DBConnectString = ""
         moParent.DbShapeConnectString = ""
      End If
   End If

End Function
'Private Function IsWorkflowUserActive(sUserLogin As String) As Boolean
'   Dim loRs As Recordset
'   Dim sSQL As String
'   Dim sLogin_Info As String
'   Dim xx As Integer
'   Dim IsSqlServer As Boolean
'   Dim sChar As Byte
'
'   On Error GoTo FunctionError
'
'   If goSession.IsOracle Then
'
'      sSQL = "select client_info as LOGIN_INFO from v$session"
'
'      IsSqlServer = False
'
'   ElseIf goSession.IsSqlServer Then
'
'      sSQL = "select context_info as LOGIN_INFO from master.dbo.sysprocesses"
'
'      IsSqlServer = True
'
'   Else
'      IsWorkflowUserActive = False
'      Exit Function
'
'   End If
'
'   Set loRs = New Recordset
'   loRs.CursorLocation = adUseClient
'
'   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
'
'   If loRs.RecordCount < 1 Then
'      IsWorkflowUserActive = False
'   Else
'      Do While Not loRs.EOF
'
'         If IsSqlServer = True Then
'
'            sLogin_Info = ""
'
'            For xx = 0 To 31
'               sChar = loRs!LOGIN_INFO.value(xx)
'               sLogin_Info = sLogin_Info & Chr(sChar)
'            Next
'            sLogin_Info = Trim(sLogin_Info)
'
'         Else
'            sLogin_Info = Trim(BlankNull(loRs!LOGIN_INFO))
'         End If
'
'
'
'         If sLogin_Info = "AGENT" Or sLogin_Info = "VESSELAGENT" Or sLogin_Info = "MASTER" Then
'            IsWorkflowUserActive = True
'            CloseRecordset loRs
'            Exit Function
'         End If
'
'         loRs.MoveNext
'      Loop
'   End If
'
'   CloseRecordset loRs
'
'   IsWorkflowUserActive = False
'
'   Exit Function
'FunctionError:
'   moParent.RaiseError "General Error in frmLogin.IsWorkflowUserActive: ", Err.Number, Err.Description
'
'End Function
'
Private Function SetActiveUser(sUserLogin As String) As Boolean

   Dim sClient_Info As String
   Dim sSQL As String
   Dim xx As Integer

   On Error GoTo FunctionError

   sClient_Info = sUserLogin & "                                                                                "

   sClient_Info = Left(sClient_Info, 32)

   If goSession.IsOracle Then

      sSQL = "DBMS_APPLICATION_INFO.SET_CLIENT_INFO('" & sClient_Info & "')"
      goCon.Execute sSQL

   ElseIf goSession.IsSqlServer Then

      sSQL = "SET CONTEXT_INFO 0X"

      For xx = 1 To 32
         sSQL = sSQL & Hex(Asc(mID(sClient_Info, xx, 1)))
      Next

      sSQL = sSQL & "00"

      goCon.Execute sSQL

   Else
      SetActiveUser = True
      Exit Function

   End If

   SetActiveUser = True

   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLogin.SetActiveUser: ", Err.Number, Err.Description

End Function

Public Function SetLoginParameters( _
  Optional MwUserID As String, _
  Optional MwPassword As String, _
  Optional DBConnectString As String, _
  Optional DbShapeConnectString As String, _
  Optional EmployeeID As String) As Boolean
   pvmLoginID.Text = MwUserID
   pvmPassword.Text = MwPassword
   gDBConnectString = DBConnectString
   gDbShapeConnectString = DbShapeConnectString
   
   moParent.SetDotNetTheme Me
   
   SetLoginParameters = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLogin.SetLoginParameters: ", Err.Number, Err.Description
   SetLoginParameters = False
   Exit Function
End Function

Private Function GetUser() As String
   Dim username         As String
   Dim slength          As Long
   Dim retval           As Long

   On Error GoTo FunctionError
   
   username = Space$(255)
   slength = 255

   retval = GetUserName(username, slength)
   username = Left$(username, slength - 1)
   GetUser = username
   mGotNetworkLogin = True
   mNetworkLoginID = username
   
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in frmLogin.GetUser: ", Err.Number, Err.Description
   Exit Function

End Function
'
' Return True if user is validated...
'
Public Function ValidateUser() As Boolean
   Dim loRsThisSite As Recordset
   Dim IsUserIntLogin As Boolean
   Dim IsPasswordEncrypted As Boolean
   Dim strSQL As String
   Dim oRsUser As Recordset
   Dim UserIsActive As Boolean
   
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim loIntNumberOfLoginAttempts As Integer
   
   On Error GoTo FunctionError
   ' Database Prespecified, load it...
   If (mIsDB And mIsDBS) And Not mIsDBOpen Then
      If Not OpenDatabase() Then
         ValidateUser = False
         Exit Function
      End If
End If
   If mIsIntegratedWinLogin And pvmLoginID.Text = "" Then
'      pvmLoginID.Text = Environ("USERNAME")
      pvmLoginID.Text = GetUser()
      mIsID = True
   End If
   
   ' No Login ID
   If Not mIsID Then
      ValidateUser = False
      Exit Function
   End If
   
   If Not mIsDBOpen Then
      ValidateUser = False
      Exit Function
   End If
   
'   '
'   ' mwcThisSite determines Authentication Requirements...
'   '
'   If Not moParent.ThisSite.LoadConfiguration() Then
'      ValidateUser = False
'      Exit Function
'   End If

   '
   '
   ' Load the User Record...
   '
   Set oRsUser = New Recordset
   oRsUser.CursorLocation = adUseClient

   If goSession.IsOracle Then
      strSQL = "select * from mwcUsers where Upper(UserID)='" & UCase(Trim(pvmLoginID.Text)) & "'"
   Else
      strSQL = "select * from mwcUsers where UserID='" & Trim(pvmLoginID.Text) & "'"
   End If
   
   
   oRsUser.Open strSQL, moDbConnection, adOpenDynamic, adLockPessimistic

   If oRsUser.RecordCount < 1 Then
      '
      ValidateUser = False
      moParent.CloseRecordset oRsUser
      Exit Function
   End If
   '
   ' Windows Integrated login OR Force Password Valaidation - no bypass !
   '
   On Error Resume Next
   IsUserIntLogin = oRsUser!IsIntegratedWinLogin
   If Err Then
      IsUserIntLogin = False
   End If
   On Error GoTo FunctionError
   
   On Error Resume Next
   UserIsActive = oRsUser!IsActive
   If Err Then
      UserIsActive = True
   End If
   On Error GoTo FunctionError
   
   If UserIsActive <> True Then
      ValidateUser = False
      moParent.CloseRecordset oRsUser
      Exit Function
   End If
   
   If Len(Trim(mNetworkLoginID)) > 0 Then
      If UCase(Trim(mNetworkLoginID)) = UCase(Trim(pvmLoginID.Text)) Then
         mGotNetworkLogin = True
      Else
         mGotNetworkLogin = False
      End If
   End If

   
   'DEV-1846 Show/Hide Site Specific users column
   'By N.Angelakis 28 APril 2010
   If moParent.ThisSite.IsUserSiteSpecific = True Then
      'DEV-1782 Assign User to specific Site
      'Added By N.Angelakis On 07 Jan 2010
      If CheckIfFieldExists(oRsUser, "Siteskey") Then
         If ZeroNull(oRsUser!SitesKey) > 0 Then
            If moParent.Site.SiteKey <> ZeroNull(oRsUser!SitesKey) Then
               'user cannot logon to this site as user has only been registered for a particular business unit
               MsgBox "You cannot logon." & vbCrLf & vbCrLf & "You are trying to logon to '" & moParent.Site.SiteName & "' but you are registered to logon to '" & moParent.Site.GetSiteName(ZeroNull(oRsUser!SitesKey)) & "'" & vbCrLf & vbCrLf & "Please contact your office.", vbCritical, Me.Caption
               
               ValidateUser = False
               moParent.CloseRecordset oRsUser
               Exit Function
            End If
         End If
      End If
   End If

   If Not (mIsIntegratedWinLogin And IsUserIntLogin <> 0) Then
      On Error Resume Next
      IsPasswordEncrypted = oRsUser!PasswordEncrypted
      If Err Then
         IsPasswordEncrypted = False
      End If
      On Error GoTo FunctionError




      'By N.Angelakis On 22 April 2009
      'DEV-1174 Advance Password Settings
      If moParent.ThisSite.IsPasswordStrong = True Then
         loIntNumberOfLoginAttempts = GetLoginAttempt(pvmLoginID.Text)
         
         'Added By N.Angelakis On 12th October 2009
         'VEL-117-User is not locking after entering allowed number of wrong password
         'If user has exceeded the number of tries and alert (+1) already sent to designated then lock system (+2)
         If ((loIntNumberOfLoginAttempts + 2 > moParent.ThisSite.PasswordFailedAttempts) And (moParent.ThisSite.PasswordFailedAttempts > 0)) Then
            MsgBox "You have exceeded the number of login attempts allowed." & vbCrLf & "Please contact the " & goSession.RoleType.GetRoleTypeName(moParent.ThisSite.LoginFailNotifyRoleTypeID), vbCritical, Me.Caption
            pvmLoginID.Enabled = False
            pvmPassword.Enabled = False
            cmdLogin.Enabled = False
            
            ValidateUser = False
            moParent.CloseRecordset oRsUser
            Exit Function
         End If
      End If
      
      
      
      
      If IsPasswordEncrypted Then
         Dim loEncrypt As New mwEncrypt
         Dim MwPassword As String
         
         'By N.Angelakis On 25th may 2009/ 16th September 2009 /moved here from IF below
         'if user failed to logon then we need to assign value to userkey, needed for sending alert to designated role
         If Not goSession.User.UserKey > 0 Then
            'if user failed to logon and we need to send alert to designated role
            goSession.User.UserKey = oRsUser!ID
         End If
         
         
         If loEncrypt.EnableEncryption(ENCRYPT_PSWD) Then
            If IsNull(oRsUser!MwPassword) Then
               If Trim(pvmPassword.Text) <> Trim(BlankNull(oRsUser!MwPassword)) Then
                  'By N.Angelakis On 22 April 2009
                  'DEV-1174 Advance Password Settings

                  Call SetLoginAttemptCounter(False, pvmLoginID.Text, loIntNumberOfLoginAttempts, oRsUser!ID, oRsUser!mwcRoleTypekey)
                  
                  ValidateUser = False
                  moParent.CloseRecordset oRsUser
                  Exit Function
               End If
            Else
               MwPassword = loEncrypt.DecryptString(oRsUser!MwPassword)
               If Trim(pvmPassword.Text) <> Trim(MwPassword) Then

                  'By N.Angelakis On 22 April 2009
                  'DEV-1174 Advance Password Settings
                  Call SetLoginAttemptCounter(False, pvmLoginID.Text, loIntNumberOfLoginAttempts, oRsUser!ID, oRsUser!mwcRoleTypekey)
                  
                  ValidateUser = False
                  moParent.CloseRecordset oRsUser
                  Exit Function
               End If
               
               'By N.Angelakis On 22 April 2009
               'DEV-1174 Advance Password Settings
               Call SetLoginAttemptCounter(True, pvmLoginID.Text, 0, oRsUser!ID, oRsUser!mwcRoleTypekey)
            End If
         End If
      ElseIf Trim(pvmPassword.Text) <> Trim(BlankNull(oRsUser!MwPassword)) Then
         'By N.Angelakis On 22 April 2009
         'DEV-1174 Advance Password Settings
         Call SetLoginAttemptCounter(False, pvmLoginID.Text, loIntNumberOfLoginAttempts, oRsUser!ID, oRsUser!mwcRoleTypekey)
      
         ValidateUser = False
         moParent.CloseRecordset oRsUser
         Exit Function
      End If

      'By N.Angelakis On 22 April 2009
      'DEV-1174 Advance Password Settings
      Call SetLoginAttemptCounter(True, pvmLoginID.Text, 0, oRsUser!ID, oRsUser!mwcRoleTypekey)
      
   ElseIf IsUserIntLogin = True And mGotNetworkLogin = False Then
      ValidateUser = False
      moParent.CloseRecordset oRsUser
      Exit Function
   End If
   
   moParent.UserKey = oRsUser!ID
   moParent.User.LoadUser oRsUser
   moParent.CloseRecordset oRsUser
   ValidateUser = True
   mIsLoggedIn = True
   
   Exit Function
FunctionError:
   'Resume Next
   moParent.RaiseError "General Error in frmLogin.ValidateUser: ", Err.Number, Err.Description
   ValidateUser = False
End Function

Private Function OpenDatabase() As Boolean
   Dim iDbOffset As Integer
   Dim loDW As mwDataWork
   
   '
   ' Open Database Connection
   '
   On Error GoTo FunctionError
   If mIsDBOpen Then
      OpenDatabase = True
      Exit Function
   End If
   '
   ' Get Connection String
   '
   If Trim(gDBConnectString) = "" Then
      If pvCbo.ListIndex < 0 Then
         MsgBox "Error: You must select a database to log into."
         OpenDatabase = False
         Exit Function
      End If
      iDbOffset = pvCbo.ListIndex + 2
      gDBConnectString = moIni(iDbOffset).ConfigKeys.GetKeyValue("DBConnectString")
      moParent.DBConnectString = gDBConnectString
      gDBConnectString = moParent.GetDecryptedDBConnectString(ENCRYPT_PSWD)
      
      gDbShapeConnectString = moIni(iDbOffset).ConfigKeys.GetKeyValue("DBShapeConnectString")
      moParent.DbShapeConnectString = gDbShapeConnectString
      gDbShapeConnectString = moParent.GetDecryptedDbShapeConnectString(ENCRYPT_PSWD)
      
      mRootFolderSemaphore = moIni(iDbOffset).ConfigKeys.GetKeyValue("RootFolderSemaphore")
      
   End If
   'MsgBox "about to create ado connection"
   If moDbConnection Is Nothing Then
      Set moDbConnection = New ADODB.Connection
   End If
   'MsgBox "created ado connection"
   moDbConnection.CursorLocation = adUseClient
   moDbConnection.Open gDBConnectString
   Set goCon = moDbConnection
   mIsDBOpen = True
   OpenDatabase = True
   
   '
   ' mwcThisSite determines Authentication Requirements...
   '
   If Not moParent.ThisSite.LoadConfiguration() Then
      mIsDBOpen = False
      OpenDatabase = False
      Exit Function
   End If
   
   moParent.FinishDBConnection
   
'   Set loDW = New mwDataWork
'   gIsSqlServer = loDW.IsSqlServer()
'   gIsOracle = loDW.IsOracle()
'   gIsAccess = loDW.IsAccess()
'   Set loDW = Nothing
   
   If gIsSqlServer = True Then
      goCon.CommandTimeout = 3600
   End If
   
   '
   ' Integrated Windows Login ?
   '
   mIsIntegratedWinLogin = CheckIntegratedWinLogin
   Exit Function
FunctionError:
   ' Can't raise error - objects not created...
   MsgBox "Unlogged Error in frmLogin.OpenDatabase, Is Baseline Database installed ?" & vbCrLf & "Connect String: " & _
     vbCrLf & "Error: " & Err.Number & "-" & Err.Description
   mIsDBOpen = False
   OpenDatabase = False

End Function

Private Sub pvmLoginID_GotFocus()
   pvmLoginID.SelStart = 0
   pvmLoginID.SelLength = Len(pvmLoginID.Text)
End Sub
Private Sub pvmPassword_GotFocus()
   pvmPassword.SelStart = 0
   pvmPassword.SelLength = Len(pvmLoginID.Text)
End Sub

Private Sub pvmLoginID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      AttemptLogin
   ElseIf KeyAscii = 27 Then
      mIsCancelled = True
      Me.Hide
   End If
End Sub

Private Sub pvmPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      AttemptLogin
   ElseIf KeyAscii = 27 Then
      mIsCancelled = True
      Me.Hide
   End If

End Sub


Private Function CheckIntegratedWinLogin() As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   sSQL = "select IsIntegratedWinLogin from mwcSites, mwcThisSite where " & _
    "mwcSites.SiteID=mwcThisSite.ThisSite"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   On Error Resume Next
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If Err Then
      CheckIntegratedWinLogin = False
      Set loRs = Nothing
      Exit Function
   End If
   On Error GoTo FunctionError
   If loRs.RecordCount < 1 Then
      CheckIntegratedWinLogin = False
   ElseIf loRs.Fields(0) <> 0 Then
      CheckIntegratedWinLogin = True
   Else
      CheckIntegratedWinLogin = False
   End If
   loRs.Close
   Set loRs = Nothing
   Exit Function
FunctionError:
   MsgBox "Unlogged Error in frmLogin.CheckIntegratedWinLogin. " & Err.Number & " - " & Err.Description
End Function


'Private Function SendPasswordAlert(ByVal loIntValue As Integer, _
'                           Optional ByVal loStrLogonUser As String, _
'                           Optional ByVal loLngLogonUserID As Long, _
'                           Optional ByVal loLngLogonRoleTypeID As Long, _
'                           Optional ByVal loLngTargetRoleTypeID As Long _
'                           ) As Boolean
'   'By N.Angelakis On 22 April 2009
'   'DEV-1174 Advance Password Settings
'
'   Dim loAlertDetails As String
'   Dim loAlertTitle As String
'   Dim sSQL As String
'   Dim loNewID As Long
'   Dim loSiteKey As Long
'   Dim loTargetSiteKey As Long
'   Dim loSiteSeedNo As Long
'   Dim loRs As Recordset
'   Dim loRsSite As Recordset
'
'   On Error GoTo ErrorHandler
'
'   loSiteKey = goSession.Site.SiteKey
'
'   sSQL = "SELECT * FROM mwAlertLog WHERE ID = -1"
'   Set loRs = New Recordset
'   loRs.CursorLocation = adUseClient
'   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
'
'   loRs.AddNew
'
'   loNewID = goSession.MakePK("mwAlertLog")
'   loRs!ID = loNewID
'   loRs!mwcSitesKeySource = loSiteKey 'From Site
'
'
'   If loIntValue = 0 Then
'      If moParent.ThisSite.LoginFailNotifyRoleTypeID = 0 Then Exit Function 'no designated role specified so do not send alert
'      'user exceeded login attempt limit send notification to designated admin
'      loAlertDetails = loAlertDetails & " Site:  " & moParent.ThisSite.ThisSite & "." & _
'      "  Roletype:  " & goSession.RoleType.GetRoleTypeName(loLngTargetRoleTypeID) & "." & vbCrLf
'      loAlertDetails = "Alert notification(s) sent to the following." & vbCrLf & loAlertDetails & _
'      vbCrLf & vbCrLf & "User (" & loStrLogonUser & ") reached maximum logon attempts. Please contact them to reset their password"
'
'      'it can only be office role, as the list of available to notify is office role populated
'      sSQL = "SELECT st.id FROM mwcRoleType rt INNER JOIN mwcSites st on rt.mwcSiteType=st.SiteType WHERE rt.id=" & loLngTargetRoleTypeID
'      Set loRsSite = New Recordset
'      loRsSite.CursorLocation = adUseClient
'      loRsSite.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
'      If Not (loRsSite.BOF And loRsSite.EOF) Then
'         loTargetSiteKey = loRsSite!ID
'      Else
'         SendPasswordAlert = False
'         CloseRecordset loRsSite
'         Exit Function
'      End If
'      CloseRecordset loRsSite
'
'      'user exceeded login attempt limit send notification to designated admin
'      loRs!mwcRoleTypeKeySource = loLngLogonRoleTypeID 'mwcRoleType:id of logging on user
'      loRs!mwcRoleTypeKeyTarget = moParent.ThisSite.LoginFailNotifyRoleTypeID 'mwcRoleType:id target
'      loRs!mwcUsersKeySource = loLngLogonUserID 'mwcUsers:id of logging on user
'      loRs!mwcSitesKeyTarget = loTargetSiteKey 'target site is where designated role is
'   Else
'      'user within last 10 days before password expires
'
'      loAlertDetails = loAlertDetails & " Site:  " & moParent.ThisSite.ThisSite & "." & _
'      "  Roletype:  " & goSession.RoleType.GetRoleTypeName(goSession.User.RoleTypeKey) & "." & vbCrLf
''''      loAlertDetails = "Alert notification(s) sent to the following." & vbCrLf & loAlertDetails & vbCrLf & vbCrLf & "You password will Expire in " & 10 - loIntValue & " days. Please change your password!"
''''
''''      loRs!mwcRoleTypeKeySource = goSession.User.RoleTypeKey
''''      loRs!mwcRoleTypeKeyTarget = loRs!mwcRoleTypeKeySource 'mwcRoleType:id singing in role sends alert/reminder to themselves target is current user
''''      loRs!mwcUsersKeySource = goSession.User.userkey 'mwcUsers:id
''''      loRs!mwcSitesKeyTarget = loSiteKey 'target is always current site
''''   End If
''''
''''
''''   loAlertTitle = "Password Security Alert"
''''   loRs!Title = loAlertTitle
''''   loRs!AlertDetails = loAlertDetails
''''   loRs!AlertDateTime = Now()
''''
''''
''''   loRs!mwAlertLogStatusKey = MW_ALERT_STATUS_SENT
''''   loRs!mwAlertTypeKey = MW_ALERT_TYPE_USER
''''
''''   loRs!mwAlertLogKeyFirst = loRs!ID
''''   loRs!mwAlertEventsKey = Null
''''   loRs!mwEventTypeKey = Null
''''   loRs!mwEventDetailKey = Null
''''   loRs!ReceiverNotes = Null
''''   loRs!ExternalData = Null
''''   loRs!mwAlertLogKeyPrev = Null
''''   loRs!ReceivedDateTime = Null
''''
''''   loRs.Update
''''   CloseRecordset loRs
''''
''''   SendPasswordAlert = True
''''   Exit Function
''''ErrorHandler:
''''   CloseRecordset loRs
''''   CloseRecordset loRsSite
''''   SendPasswordAlert = False
''''End Function
''''
Private Function SendPasswordAlert(ByVal loIntValue As Integer, _
                           Optional ByVal loStrLogonUser As String, _
                           Optional ByVal loLngLogonUserID As Long, _
                           Optional ByVal loLngLogonRoleTypeID As Long, _
                           Optional ByVal loLngTargetRoleTypeID As Long _
                           ) As Boolean
   
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings

   Dim sSQL As String
   Dim loRsSite As Recordset
   Dim lomwAlertWork As Object
   Dim AlertTitle As String
   Dim AlertDetails As String
   Dim AlertTargetSiteKey As Long
   Dim AlertRoleTypeKeyTarget As Long
   Dim AlertRoleTypeKeySource As Long
   Dim AlertUsersKeySource As Long
   Dim AlertUsersKeyTarget As Long
   ' 11/2011 ms - Since emails go to a Role... kill the option as an annoyance
   Dim loKeys As ConfigKeys
   
   On Error GoTo ErrorHandler
   Set loKeys = goSession.GetEventSecurityKeys(EVENTTYPE_MARINE_ASSURANCE_APP)
   If loKeys.GetBoolKeyValue("DisablePasswordExpireAlerts") Then
      SendPasswordAlert = True
      KillObject loKeys
      Exit Function
   End If

   If loIntValue = 0 Then
      If moParent.ThisSite.LoginFailNotifyRoleTypeID = 0 Then Exit Function 'no designated role specified so do not send alert
      
      'user exceeded login attempt limit send notification to designated role
      AlertDetails = AlertDetails & " Site:  " & moParent.ThisSite.ThisSite & "." & _
      "  Roletype:  " & goSession.RoleType.GetRoleTypeName(loLngTargetRoleTypeID) & "." & vbCrLf
      AlertDetails = "Alert notification(s) sent to the following." & vbCrLf & AlertDetails & _
      vbCrLf & vbCrLf & "User (" & loStrLogonUser & ") reached maximum logon attempts. Please contact them to reset their password"
   
      'it can only be office role, as the list of available to notify is office role populated
      sSQL = "SELECT st.id FROM mwcRoleType rt INNER JOIN mwcSites st on rt.mwcSiteType=st.SiteType WHERE rt.id=" & loLngTargetRoleTypeID
      Set loRsSite = New Recordset
      loRsSite.CursorLocation = adUseClient
      loRsSite.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      If Not (loRsSite.BOF And loRsSite.EOF) Then
         'target site for designated role
         AlertTargetSiteKey = loRsSite!ID
      Else
         SendPasswordAlert = False
         CloseRecordset loRsSite
         Exit Function
      End If
      CloseRecordset loRsSite
   
      AlertRoleTypeKeyTarget = moParent.ThisSite.LoginFailNotifyRoleTypeID
      AlertRoleTypeKeySource = loLngLogonRoleTypeID
      AlertUsersKeySource = loLngLogonUserID
      AlertUsersKeyTarget = 0

   Else
      'user within last 10 days before password expires
      'currect site & role is sending alert to themselves
      
'      AlertDetails = AlertDetails & " Site:  " & moParent.ThisSite.ThisSite & "." & _
'      "  Roletype:  " & goSession.RoleType.GetRoleTypeName(goSession.User.RoleTypeKey) & "." & vbCrLf
'      AlertDetails = "Alert notification(s) sent to the following." & vbCrLf & AlertDetails & vbCrLf & vbCrLf & "Your password will Expire in " & 10 - loIntValue & " days. Please change your password!"
      AlertDetails = "The password for User (" & loStrLogonUser & ") will Expire in " & 10 - loIntValue & " days."
      
      AlertTargetSiteKey = goSession.Site.SiteKey
      AlertRoleTypeKeyTarget = 0
      AlertRoleTypeKeySource = goSession.User.RoleTypeKey
      AlertRoleTypeKeyTarget = goSession.User.RoleTypeKey
      AlertUsersKeySource = goSession.User.UserKey
      AlertUsersKeyTarget = goSession.User.UserKey
   End If
   
   AlertTitle = "Password Security Alert"
   
   Set lomwAlertWork = CreateObject("mwSession.mwAlertWork")
   SendPasswordAlert = lomwAlertWork.CreateAlert(AlertTargetSiteKey, _
                              AlertTitle, _
                              AlertRoleTypeKeyTarget, _
                              AlertDetails, _
                              MW_ALERT_STATUS_SENT, _
                              MW_ALERT_TYPE_USER, _
                              AlertRoleTypeKeySource, _
                              "", _
                              AlertUsersKeySource, _
                              AlertUsersKeyTarget)
   KillObject lomwAlertWork

   Exit Function
ErrorHandler:
   CloseRecordset loRsSite
   KillObject lomwAlertWork
   SendPasswordAlert = False
End Function


Private Function SetLoginAttemptCounter(ByVal loBlnLoginAttempt, _
                                        ByVal loStrLogonUsername As String, _
                                        ByVal loIntLoginAttempt As Integer, _
                                        ByVal loLngLogonUserID As Long, _
                                        ByVal loLngLogonRoleTypeID As Long) As Boolean
   
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim loStrSQL As String
   Dim loRs As Recordset
   
   On Error GoTo ErrorHandler
   
   If moParent.ThisSite.IsPasswordStrong = False Then
       SetLoginAttemptCounter = True
       Exit Function
   End If
      
      
   If loBlnLoginAttempt = False Then
      If ((loIntLoginAttempt + 1 > moParent.ThisSite.PasswordFailedAttempts) And (moParent.ThisSite.PasswordFailedAttempts > 0)) Then
         If loIntLoginAttempt + 1 = moParent.ThisSite.PasswordFailedAttempts + 1 Then
            'send alert to designated admin only once when maximum tries+1 have been reached before encrypting
            Call SendPasswordAlert(0, pvmLoginID.Text, loLngLogonUserID, _
            loLngLogonRoleTypeID, moParent.ThisSite.LoginFailNotifyRoleTypeID)
         End If
         
         MsgBox "You have exceeded the number of login attempts allowed." & vbCrLf & "Please contact the " & goSession.RoleType.GetRoleTypeName(moParent.ThisSite.LoginFailNotifyRoleTypeID), vbCritical, Me.Caption
         pvmLoginID.Enabled = False
         pvmPassword.Enabled = False
         cmdLogin.Enabled = False
      End If


      loIntLoginAttempt = loIntLoginAttempt + 1
      If goSession.IsOracle Then
         loStrSQL = "UPDATE mwcUsers SET LoginFailedAttempts=" & loIntLoginAttempt & " WHERE Upper(UserID)='" & UCase(Trim(loStrLogonUsername)) & "'"
      Else
         loStrSQL = "UPDATE mwcUsers SET LoginFailedAttempts=" & loIntLoginAttempt & " WHERE UserID='" & Trim(loStrLogonUsername) & "'"
      End If
         
      goCon.Execute loStrSQL

   ElseIf loBlnLoginAttempt = True Then
      'user login ok so reset counter to 0
      If goSession.IsOracle Then
         loStrSQL = "UPDATE mwcUsers SET LoginFailedAttempts=0 WHERE Upper(UserID)='" & UCase(Trim(loStrLogonUsername)) & "'"
      Else
         loStrSQL = "UPDATE mwcUsers SET LoginFailedAttempts=0 WHERE UserID='" & Trim(loStrLogonUsername) & "'"
      End If
      goCon.Execute loStrSQL
   End If
   
   SetLoginAttemptCounter = True
   
   Exit Function
ErrorHandler:
   CloseRecordset loRs
   SetLoginAttemptCounter = False
End Function

Private Function GetLoginAttempt(ByVal loStrUsername As String) As Integer
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   Dim loStrSQL As String
   Dim loRs As Recordset
   Dim loIntLoginFailedAttempts As Integer

   On Error GoTo ErrorHandler
   
   GetLoginAttempt = 0
   If moParent.ThisSite.IsPasswordStrong = False Then
       Exit Function
   End If

   'user login attempt failed so increment counter max 10 then lock
   If goSession.IsOracle Then
      loStrSQL = "SELECT id, LoginFailedAttempts FROM mwcUsers WHERE Upper(UserID)='" & UCase(Trim(loStrUsername)) & "'"
   Else
      loStrSQL = "SELECT id, LoginFailedAttempts FROM mwcUsers WHERE UserID='" & Trim(loStrUsername) & "'"
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open loStrSQL, goCon, adOpenDynamic, adLockOptimistic
      
   If Not (loRs.BOF And loRs.EOF) Then
      loIntLoginFailedAttempts = ZeroNull(loRs("LoginFailedAttempts"))
   End If
   CloseRecordset loRs
      
   GetLoginAttempt = loIntLoginFailedAttempts
      
   Exit Function
ErrorHandler:
   CloseRecordset loRs
   GetLoginAttempt = 0
End Function



