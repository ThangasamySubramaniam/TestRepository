VERSION 5.00
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmLicensing 
   Caption         =   "ShipNet Fleet Licensing"
   ClientHeight    =   6432
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   9852
   Icon            =   "frmLicensing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   9852
   StartUpPosition =   2  'CenterScreen
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   7440
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   9852
      DesignHeight    =   6432
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Close"
      Height          =   1095
      Left            =   8160
      Picture         =   "frmLicensing.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdFormHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      Picture         =   "frmLicensing.frx":1594
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtSiteKey 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtSiteCode 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton cmdCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Copy Site  Code To Clipboard To Send To ShipNet Fleet"
      Top             =   960
      Width           =   1812
   End
   Begin VB.CommandButton cmdUpdateLicense 
      Caption         =   "Update License Key"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      ToolTipText     =   "Enter new Site Key Received From ShipNet Fleet"
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Caption         =   "License Transfer Service"
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Transfer licenses between computers running ShipNet Fleet"
      Top             =   4080
      Width           =   6855
      Begin VB.TextBox txtLicenseCount 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdExportLicense 
         Caption         =   "Transfer License Out"
         Height          =   975
         Left            =   2280
         Picture         =   "frmLicensing.frx":189E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Using the License Request file from another computer, create a License Response file for the computer requesting the license."
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdRequestDiskette 
         Caption         =   "Request Transfer"
         Height          =   975
         Left            =   240
         Picture         =   "frmLicensing.frx":1BA8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generate License Request file to get a a license from another computer running ShipNet Fleet"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdImportLicense 
         Caption         =   "Transfer License In"
         Height          =   975
         Left            =   4440
         Picture         =   "frmLicensing.frx":1EB2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Process the License Response File generated by the computer issuing the license, completing the license transfer."
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "No. of Licenses on this computer"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdRequest 
      Caption         =   "EMail Request"
      Height          =   975
      Left            =   8280
      Picture         =   "frmLicensing.frx":21BC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Send License Request Email to ShipNet Fleet"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ShipNet Fleet Licensing Service"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label18 
      Caption         =   "Site Code:"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "License Key:"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "License Terms"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblLicenseTerms 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1092
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   7092
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEmail 
         Caption         =   "Email Settings"
      End
   End
End
Attribute VB_Name = "frmLicensing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Licensing Service Form
'
' 9/21/2001 ms
'

Option Explicit
Dim moCryp As mwCrypkey
Dim moParent As mwSession.Session
Dim mIsLicensed As Boolean
Dim mIsLicenseKeyDisabled As Boolean

Private Const CK_AUTH_OK = 0


Const HELP_MANUAL = "mwUser_810_ManageLicenses.chm"
Const Troubleshooting = 1
Const License_Manager = 2
Const Licensing_Overview = 3


Private Sub cmdFormHelp_Click()
   moParent.API.ShowMwHelp "mwUser810ManageLicenses.chm"
End Sub

Private Sub Form_Load()
   Dim iCrypInitError As Integer
   On Error GoTo SubError
   '
   ' Won't Work if not logged in
   '
   Set moCryp = New mwCrypkey
   iCrypInitError = moCryp.InitCrypkeyGE
   ' Initialize Crypkey
   If iCrypInitError = -102 Then
      'Unload Me
      txtSiteCode.Text = "<Licensing Service Is Not Operating>"
      txtSiteCode.FontBold = True
      txtSiteCode.BackColor = vbRed
      Exit Sub
   ElseIf iCrypInitError = -101 Then
      moParent.RaiseError "License Initialization Failed. Error Code is: " & str(iCrypInitError) & vbCrLf & _
       "Most likely Windows User does not have full permissions to the [APPLICATION] folder."
   ElseIf iCrypInitError <> 0 Then
      moParent.RaiseError "License Initialization Failed. Error Code is: " & str(iCrypInitError)
      Me.Hide
      Exit Sub
   End If
   txtSiteCode.Text = moCryp.GetSiteCodeGE
   UpdateLicenseTerms
   'License Distribution...
   'If moParent.IsLoggedIn Then
      If Not moParent.User.Security.UserConfigLicenseDistribution Then
         Frame1.Visible = False
      End If
   'End If
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   moParent.RaiseError "General Error in mwSession.frmLicensing.Load.", err.Number, err.Description
   Me.Hide
End Sub


Private Sub cmdRequestDiskette_Click()
   Dim strTarget As String
   Dim oForm As frmLocation
   On Error GoTo SubError

   '
   ' Allow user to select where to put license request files...
   '
   Set oForm = New frmLocation
   oForm.SetParent moParent
   oForm.Show vbModal
   '
   If oForm.IsCancelled Then
      Unload oForm
      Set oForm = Nothing
      Exit Sub
   End If
   strTarget = oForm.GetPath
   Unload oForm
   Set oForm = Nothing
   If moCryp.TransferRequestMS(strTarget) Then
      'MsgBox "License request has been placed in the specified folder or diskette.", vbInformation
   End If
   Exit Sub
SubError:
   moParent.RaiseError "General error in frmLicensing.cmdExportLicense. ", err.Number, err.Description

End Sub

Private Sub cmdSave_Click()
   Me.Hide
End Sub

Private Sub cmdUpdateLicense_Click()
   Dim iCryp As Integer
   Dim sLicense As String
   On Error GoTo SubError
   If Trim(txtSiteKey.Text) <> "" Then
      moCryp.SaveSiteKeyGE txtSiteKey.Text & Chr$(0)
   Else
      MsgBox "You must enter a valid License Key.", vbInformation, _
        "ShipNet Fleet Licensing"
      Exit Sub
   End If
   '
   ' Reconfirm License is ok...
   '
   iCryp = moCryp.GetAuthorizationGE
   If iCryp <> CK_AUTH_OK Then
      MsgBox "License is still not authorized for this site.", vbCritical, _
        "ShipNet Fleet Licensing"
      moParent.Logger.LogIt mwl_User_Defined, mwl_Critical, " License is still not authorized for this site."
   Else
      
      MsgBox "Workstation License Has Been Updated, you are authorized for " & GetLicenseTerms, vbInformation, _
        "ShipNet Fleet Licensing"
      moParent.Logger.LogIt mwl_User_Defined, mwl_Information, " Workstation License Has Been Updated."
      If mIsLicensed = False Then
         mIsLicensed = True
         Me.Hide
      Else
         txtSiteKey.Text = ""
         UpdateLicenseTerms
      End If
   End If
   Exit Sub
SubError:
   moParent.RaiseError "General Error in mwSession.frmLicensing. ", err.Number, err.Description
End Sub

Private Sub cmdCopyToClipboard_Click()
   Clipboard.Clear
   Clipboard.SetText txtSiteCode.Text
End Sub

Private Sub cmdExportLicense_Click()
   Dim strTarget As String
   Dim oForm As frmLocation
   On Error GoTo SubError

   '
   ' Allow user to select where to put license request files...
   '
   Set oForm = New frmLocation
   oForm.SetParent moParent
   oForm.Show vbModal
   '
   If oForm.IsCancelled Then
      Unload oForm
      Set oForm = Nothing
      Exit Sub
   End If
   strTarget = oForm.GetPath
   Unload oForm
   Set oForm = Nothing
   If moCryp.TransferOutMS(strTarget) Then
      UpdateLicenseTerms
   End If
   Exit Sub
SubError:
   moParent.RaiseError "General error in frmLicensing.cmdExportLicense. ", err.Number, err.Description
End Sub

Private Sub cmdImportLicense_Click()
   Dim strTarget As String
   Dim oForm As frmLocation
   On Error GoTo SubError

   '
   ' Allow user to select where to put license request files...
   '
   Set oForm = New frmLocation
   oForm.SetParent moParent
   oForm.Show vbModal
   '
   If oForm.IsCancelled Then
      Unload oForm
      Set oForm = Nothing
      Exit Sub
   End If
   strTarget = oForm.GetPath
   Unload oForm
   Set oForm = Nothing
   If moCryp.TransferInMS(strTarget) Then
      UpdateLicenseTerms
   End If
   Exit Sub
SubError:
   moParent.RaiseError "General error in frmLicensing.cmdImportLicense. ", err.Number, err.Description
End Sub

Private Sub cmdRequest_Click()
   Dim strAddress As String
   Dim strSubjectText As String
   Dim strBodyText As String
   
   On Error GoTo SubError
   If Trim(txtSiteCode.Text) = "" Then
      MsgBox "Licensing Service is not working, cannot send request.", vbInformation, "Email License Request"
      Exit Sub
   End If
   If goSession.User.DefaultEmailCarrier = mw_AMOS_MAIL Then
      strAddress = "<SMTP:license@shipnetfleet.com>"
   Else
      strAddress = "license@shipnetfleet.com"
   End If
   
   strSubjectText = "License Key Request From: " & goSession.Site.SiteName
   
   strBodyText = "SiteKey=" & txtSiteCode.Text & vbCrLf & vbCrLf
   strBodyText = strBodyText & "CompanyID = " & goSession.ThisSite.CompanyID & vbCrLf
   strBodyText = strBodyText & "CompanyCode = " & goSession.Site.GetExtendedProperty("CompanyCode") & vbCrLf
   strBodyText = strBodyText & "WorkflowSendToAddress = " & goSession.Site.WorkflowSendToAddress & vbCrLf
   strBodyText = strBodyText & "SiteRoot = " & goSession.SiteRoot & vbCrLf
   
   If Not goSession.SendNotification(strSubjectText, strBodyText, strAddress) Then
      MsgBox "Email Integration is not configured or available. Use File option, or Copy To Clipboard to manually request license.", vbExclamation, "Email Request Error"
   Else
      MsgBox "License Request has been sent out; you will receive your license key by email. You can close ShipNet Fleet now.", vbInformation, "Email Request Successful"
   End If
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.frmLicensing.mnuEmail_Click.", err.Number, err.Description
End Sub

Private Function UpdateLicenseTerms()
   Dim sLicense As String
   On Error GoTo FunctionError
   If mIsLicenseKeyDisabled Then
      lblLicenseTerms.Caption = "License Key has been disabled, you must request a new License Key."
      lblLicenseTerms.ForeColor = &HFF&
   ElseIf moCryp.IsUnlimited Then
      sLicense = "Unlimited Time. Licensed for "
      sLicense = sLicense & GetLicenseTerms()
      lblLicenseTerms.Caption = sLicense
      lblLicenseTerms.ForeColor = &HFF0000
      
   ElseIf moCryp.IsDaysRestricted Then
      sLicense = str(moCryp.NumberOfDaysRunsAllowed) & " day License issued. " & _
        moCryp.NumberOfDaysRunsRemaining & " days remaining. Licensed for "
      sLicense = sLicense & GetLicenseTerms()
      lblLicenseTerms.Caption = sLicense
      If moCryp.NumberOfDaysRunsAllowed - moCryp.NumberOfDaysRunsUsed < 30 Then
         lblLicenseTerms.ForeColor = &HFF&
      Else
         lblLicenseTerms.ForeColor = &HFF0000
      End If
   ElseIf moCryp.IsRunsRestricted Then
      sLicense = "Run Evaluation License issued. Licensed for "
      sLicense = sLicense & GetLicenseTerms()
      lblLicenseTerms.Caption = sLicense
      lblLicenseTerms.ForeColor = &HFF&
   Else
      'MsgBox "Critical Error Loading ShipNet Fleet: " & _
        "TERMS" & vbCrLf & "Please Contact  Maritime Systems for Assistance", vbCritical, _
        "ShipNet Fleet Initialization Failure"
   End If
   txtLicenseCount.Text = moCryp.GetNumCopiesGE
FunctionError:
   'MsgBox "General Error infrmLicensing.UpdateLicenseTerms: " & err.Number & " - " & err.Description
End Function



Private Sub Form_Unload(Cancel As Integer)
   Set moCryp = Nothing
End Sub

Public Function SetParentSession(ByRef ses As Session)
   Set moParent = ses
End Function

Public Function SetIsLicenseKeyDisabled()
   mIsLicenseKeyDisabled = True
End Function

Private Sub mnuEmail_Click()
   goSession.User.ConfigureEmail
End Sub


Private Function GetLicenseTerms() As String
   Dim sLicense As String
   On Error GoTo FunctionError
   If goSession.IsFeatureLicensed(LIC_01_Crewing) Then
      sLicense = sLicense & "Crewing, "
   End If
   If goSession.IsFeatureLicensed(LIC_02_Vessel_Reporting) Then
      sLicense = sLicense & "Vessel Reporting, "
   End If
   If goSession.IsFeatureLicensed(LIC_03_Safety_Management) Then
      sLicense = sLicense & "Safety Management, "
   End If
   If goSession.IsFeatureLicensed(LIC_04_Document_Control) Then
      sLicense = sLicense & "Document Control, "
   End If
   If goSession.IsFeatureLicensed(LIC_05_Warranty_Claims) Then
      sLicense = sLicense & "Warranty Claims, "
   End If
   If goSession.IsFeatureLicensed(LIC_06_ShipWorks_Equipment) Then
      sLicense = sLicense & "Equipment Management, "
   End If
   If goSession.IsFeatureLicensed(LIC_07_ShipWorks_Drydock) Then
      sLicense = sLicense & "Dry Docking, "
   End If
   If goSession.IsFeatureLicensed(LIC_08_ShipWorks_Requisitioning) Then
      sLicense = sLicense & "Requisitioning, "
   End If
   If goSession.IsFeatureLicensed(LIC_09_Maintenance) Then
      sLicense = sLicense & "Maintenance, "
   End If
   GetLicenseTerms = Left(sLicense, Len(sLicense) - 2) & "."
   Exit Function
FunctionError:
   moParent.RaiseError "General error in frmLicensing.GetLicenseTerms. ", err.Number, err.Description
End Function

