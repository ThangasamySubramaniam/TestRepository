VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form frmAlertCreate 
   Caption         =   "Create Alert"
   ClientHeight    =   6585
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5485.417
   ScaleMode       =   0  'User
   ScaleWidth      =   10680.28
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAlertAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      Picture         =   "frmAlertCreate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Add Recipient"
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdAlertDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      Picture         =   "frmAlertCreate.frx":0272
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Delete Recipient"
      Top             =   1800
      Width           =   975
   End
   Begin UltraGrid.SSUltraGrid ug 
      Height          =   2115
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3731
      _Version        =   131072
      GridFlags       =   17040384
      Images          =   "frmAlertCreate.frx":057C
      LayoutFlags     =   72351764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "frmAlertCreate.frx":0EDC
      CaptionAppearance=   "frmAlertCreate.frx":0F32
      Caption         =   "To"
   End
   Begin VB.CheckBox chkShoreUsersOnly 
      Caption         =   "Shore Users Only"
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
      Left            =   9480
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7121
      Picture         =   "frmAlertCreate.frx":0F6E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Select and Exit"
      Top             =   5535
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3593
      Picture         =   "frmAlertCreate.frx":1C38
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancel and Exit"
      Top             =   5535
      Width           =   975
   End
   Begin VB.CheckBox chkShowUsers 
      Caption         =   "Show Users"
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
      Left            =   9480
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdLookupSites 
      Height          =   408
      Left            =   4320
      Picture         =   "frmAlertCreate.frx":1F42
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtTargetSites 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   1020
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtTargetRoleTypes 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   5520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton cmdLookupRoleTypes 
      Height          =   408
      Left            =   10920
      Picture         =   "frmAlertCreate.frx":224C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtAlertDetails 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3480
      Width           =   11295
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   900
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2880
      Width           =   9555
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Shows the progress when sending the Alert"
      Top             =   4995
      Width           =   11295
   End
   Begin VB.Label lblmwcRoleTypeKeySource 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "mwcRoleTypeKeySource"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   10
      Top             =   5655
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblmwcSitesKeySource 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "mwcSitesKeySource"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "From:  Site, Role, User"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3015
   End
End
Attribute VB_Name = "frmAlertCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents moRS  As Recordset
Attribute moRS.VB_VarHelpID = -1

Dim moRsTargetRTKeys As Recordset
Dim moRsTargetSiteKeys As Recordset
Dim moRsUserKeys As Recordset
Dim moRsAlertToList As Recordset

Dim sTargetRtNames As String
Dim sTargetSiteNames As String
Dim sUserNames As String

Const MW_ALERT_STATUS_SENT = 1
Const MW_ALERT_STATUS_READ = 2
Const MW_ALERT_STATUS_REPLIED = 3
Const MW_ALERT_STATUS_CLOSED = 4

Const MW_ALERT_TYPE_USER = 1
Const MW_ALERT_TYPE_SYSTEM = 2

Dim m_mwcSitesKeySource As Long
Dim m_mwcRoleTypeKeySource As Long

Dim mIsCancelled As Boolean
Dim mNewID As Long
Dim mUGNewID As Long
Dim mAlertDetails As String
Dim mAlertTitle As String
Dim mAlertTarget As String

   Dim mEventTypeKey As Long
   Dim mEventDetailKey As Long
   
Dim mThisSiteKey As Long
Dim mThisSiteType As Integer

' UserSiteSpecific
   Dim mIsUserSiteSpecific As Boolean
   Dim mSSiteKeys As String
'PBT 1692
Dim moReg As Registry
Dim InRefresh As Boolean
Const UG_ID = 0
Const UG_mwcSitesKeyDest = 1
Const UG_mwcSitesName = 2
Const UG_mwcRoleTypeKeyDest = 3
Const UG_mwcRoleTypeName = 4
Const UG_mwcUsersKeyTarget = 5
Const UG_mwcUsersName = 6
Const UG_IsRoleType = 7

Private Function CreateAlertToStructure() As Boolean
   On Error GoTo FunctionError
   
   CloseRecordset moRsAlertToList
   
   Set moRsAlertToList = New Recordset
   
   moRsAlertToList.CursorLocation = adUseClient
   moRsAlertToList.Fields.Append "ID", adInteger, , adFldLong And adFldIsNullable And adFldUpdatable
   moRsAlertToList.Fields.Append "mwcSitesKeyTarget", adInteger, , adFldLong And adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "SiteName", adVarChar, 50, adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "mwcRoleTypeKeyTarget", adInteger, , adFldLong And adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "RoleName", adVarChar, 50, adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "mwcUsersKeyTarget", adInteger, , adFldLong And adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "UserName", adVarChar, 50, adFldIsNullable And adFldUpdatable And adFldMayBeNull
   moRsAlertToList.Fields.Append "IsRoleType", adInteger, , adFldLong And adFldIsNullable And adFldUpdatable And adFldMayBeNull
   
   moRsAlertToList.Open
   
   CreateAlertToStructure = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.CreateAlertToStructure ", Err.Number, Err.Description
End Function
'Public Const SITE_TYPE_SHIP = 1
'Public Const SITE_TYPE_SHORE = 2

Public Function NewID() As Long
   NewID = mNewID
End Function
Public Function GetAlertDetails() As String
   GetAlertDetails = mAlertDetails
End Function
Public Function GetAlertTitle() As String
   GetAlertTitle = mAlertTitle
End Function
Public Function GetAlertTarget() As String
   GetAlertTarget = mAlertTarget
End Function

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Property Let AlertTitle(sAlertTitle As String)
   mAlertTitle = Left(sAlertTitle, 50)
   If mAlertTitle <> "" Then
      txtTitle.Text = mAlertTitle
   End If
End Property
Public Property Let AlertDescription(sAlertDetails As String)
   mAlertDetails = sAlertDetails
   If mAlertDetails <> "" Then
      txtAlertDetails.Text = mAlertDetails
   End If
End Property
Public Property Let AlertTarget(sAlertTarget As String)
   mAlertTarget = sAlertTarget
End Property

Public Property Let EventTypeKey(nEventTypeKey As Long)
   mEventTypeKey = nEventTypeKey
End Property

Public Property Let EventDetailKey(nEventDetailKey As Long)
   mEventDetailKey = nEventDetailKey
End Property

Private Sub cmdAlertDelete_Click()
   Dim i As Integer
   On Error GoTo SubError
   '
   ' validate
   If ug.ActiveRow Is Nothing Then
      Beep
      Exit Sub
   ElseIf ug.Selected.Rows.Count < 1 Then
      Beep
      Exit Sub
   End If
   
   i = MsgBox("Do you want to delete these Recipient(s) ?", vbYesNo, "Delete Recipient")
   If i = vbYes Then
      ug.DeleteSelectedRows
   End If
   
   RefreshUgColumns
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.cmdAlertDelete_Click. ", Err.Number, Err.Description
End Sub

Private Sub cmdCancel_Click()
   mIsCancelled = True
   mAlertDetails = ""
   mAlertTitle = ""
   Me.Hide
End Sub
Private Sub Form_Load()
   Dim sSQL As String
   Dim loRs As Recordset
   Dim nShoreSiteKey As Long
   On Error GoTo FunctionError
   
   If moReg Is Nothing Then
      Set moReg = New Registry
   End If
   moReg.BaseRegistry = BASE_REG & "mwSession." & Me.Name
   InRefresh = False
   
   mIsCancelled = False
   '
   ' mwcSitesKeySource
   m_mwcSitesKeySource = goSession.Site.SiteKey
   If m_mwcSitesKeySource > 0 Then
      lblmwcSitesKeySource = goSession.Site.GetSiteName(m_mwcSitesKeySource) & ", "
   Else
      lblmwcSitesKeySource = ", "
   End If
   
   ' DEV-2144 display only Assigned UserToSite
      If goSession.ThisSite.IsUserSiteSpecific Then
         mIsUserSiteSpecific = True
      Else
         mIsUserSiteSpecific = False
      End If


   '
   ' list limitation variables
   mThisSiteKey = goSession.Site.SiteKey
   mThisSiteType = goSession.Site.SiteType
   
   ' mwcRoleTypeKeyTarget (provide backbard compatability)
   m_mwcRoleTypeKeySource = goSession.User.RoleTypeKey
   If m_mwcRoleTypeKeySource > 0 Then
      lblmwcRoleTypeKeySource = goSession.RoleType.GetRoleTypeName(m_mwcRoleTypeKeySource) & ", "
      lblmwcSitesKeySource = lblmwcSitesKeySource & lblmwcRoleTypeKeySource
   Else
      lblmwcSitesKeySource = lblmwcSitesKeySource & ", "
   End If
   
   lblmwcSitesKeySource = lblmwcSitesKeySource & goSession.User.GetExtendedProperty("UserName")
   Set loRs = Nothing
   
   CreateAlertToStructure 'PBT 1692
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in frmAlertCreate ", Err.Number, Err.Description
End Sub
Private Sub cmdOK_Click()
   Dim sSQL As String
   Dim nRoleTypeKey As Long
   Dim sInvalidMsg As String
   Dim IsListTruncated As Boolean
   Dim nRecipientCount As Long
   Dim sUserName As String
   Dim sRTName As String
   Dim nRTKey As Long
   Dim loAlertWork As mwAlertWork
   On Error GoTo SubError
   
   sSQL = "SELECT * FROM mwAlertLog WHERE ID = -1"
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   ' set memo description with sites & roletypes
   mAlertDetails = ""
   If IsRecordLoaded(moRsAlertToList) Then
      moRsAlertToList.MoveFirst
      IsListTruncated = False
      Set loAlertWork = New mwAlertWork
      
      If moRsAlertToList.RecordCount > 0 Then
         nRecipientCount = moRsAlertToList.RecordCount
      Else
         nRecipientCount = 0
      End If
      
      If nRecipientCount > 1 Then
         mAlertDetails = "Alert notification(s) sent to the following " & nRecipientCount & " recipients." & vbCrLf & mAlertDetails & vbCrLf & vbCrLf & txtAlertDetails.Text
      Else
         mAlertDetails = "Alert notification(s) sent to the following." & vbCrLf & mAlertDetails & vbCrLf & vbCrLf & txtAlertDetails.Text
      End If
      
      moRsAlertToList.MoveFirst
      mNewID = goSession.MakePK("mwAlertLog")
      
      ' loop for each TargetSiteKey selected
      Do While Not moRsAlertToList.EOF
         
         If ZeroNull(moRsAlertToList!mwcUsersKeyTarget) > 0 Then
            sUserName = GetUserNameRoleName(moRsAlertToList!mwcUsersKeyTarget, sRTName, nRTKey)
         
            lblStatus.Caption = "Sending Alert to Site: " & goSession.Site.GetSiteName(moRsAlertToList!mwcSitesKeyTarget) & _
            "  User/Roletype:  " & sUserName & "/" & sRTName
            lblStatus.Refresh
         Else
            lblStatus.Caption = "Sending Alert to Site: " & goSession.Site.GetSiteName(moRsAlertToList!mwcSitesKeyTarget) & _
                "  Roletype:  " & goSession.RoleType.GetRoleTypeName(moRsAlertToList!mwcRoleTypeKeyTarget)
                lblStatus.Refresh
         End If
         
         moRS.AddNew
         moRS!ID = mNewID
         
         moRS!mwcSitesKeySource = m_mwcSitesKeySource
         moRS!mwcRoleTypeKeySource = m_mwcRoleTypeKeySource
         moRS!mwcSitesKeyTarget = moRsAlertToList!mwcSitesKeyTarget
         
         'mAlertTarget = moRsTargetRTKeys!ID
         If ZeroNull(moRsAlertToList!mwcUsersKeyTarget) > 0 Then
            mAlertTarget = nRTKey
         Else
            mAlertTarget = moRsAlertToList!mwcRoleTypeKeyTarget
         End If
         
         mAlertTitle = "" & txtTitle.Text
         moRS!mwcRoleTypeKeyTarget = mAlertTarget
         moRS!Title = mAlertTitle
         moRS!AlertDetails = mAlertDetails
         moRS!AlertDateTime = Now()
         moRS!mwAlertLogKeyFirst = moRS!ID
         moRS!mwAlertLogKeyPrev = Null
         moRS!ReceivedDateTime = Null
         moRS!mwAlertLogStatusKey = MW_ALERT_STATUS_SENT
         moRS!mwAlertTypeKey = MW_ALERT_TYPE_USER
         moRS!mwAlertEventsKey = Null
         moRS!mwEventTypeKey = Null
         moRS!mwEventDetailKey = Null
         moRS!ReceiverNotes = Null
         moRS!ExternalData = Null
         
         If mEventTypeKey > 0 Then
            moRS!mwEventTypeKey = mEventTypeKey
         Else
            moRS!mwEventTypeKey = Null
         End If
         If mEventDetailKey > 0 Then
            moRS!mwEventDetailKey = mEventDetailKey
         Else
            moRS!mwEventDetailKey = Null
         End If
         
         moRS!mwcUsersKeySource = goSession.User.UserKey
         
         If ZeroNull(moRsAlertToList!mwcUsersKeyTarget) > 0 Then
            moRS!mwcUsersKeyTarget = moRsAlertToList!mwcUsersKeyTarget
         End If
         moRS.Update
         ' send notification?
         'loAlertWork.CheckRoleEmail (moRS!mwcUsersKeyTarget)
         
         ' send notification?
         loAlertWork.SetAlertNotification moRS, 0, ZeroNull(moRS!mwcUsersKeyTarget)
         
         mNewID = mNewID + 1
         DoEvents
         moRsAlertToList.MoveNext
      Loop
      If mNewID > 0 Then
         goSession.UpdatePrimaryKeySequence "mwAlertLog", mNewID
      End If

      CloseRecordset moRS
   End If
      
   lblStatus.Caption = " "
   lblStatus.Refresh
   
   Me.Hide
   KillObject loAlertWork
Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.cmdOK_Click: ", Err.Number, Err.Description
   CloseRecordset moRS
   Me.Hide
   KillObject loAlertWork
End Sub

Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static IsBeginAdd As Boolean
   Static IsBeginDelete As Boolean
   On Error GoTo SubError
   
   ' mwAlertLog must use this special WillChangeRecord handler because it is a
   ' Site Specific table that does NOT have an mwcSitesKey.
   
   ' DO NOT change this to the new-style mwReplicateWillChange class!
   
   If adReason = adRsnAddNew Then
      IsBeginAdd = True
   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
      IsBeginDelete = False
   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRS!ID, moRS!mwcSitesKeyTarget, moRS.Fields
      Else
         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRS!ID, goSession.Site.SiteKey, moRS.Fields
      End If
      IsBeginAdd = False
   ElseIf adReason = adRsnDelete Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRS!ID, moRS!mwcSitesKeyTarget
      Else
         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRS!ID, goSession.Site.SiteKey
      End If
      IsBeginDelete = True
   ElseIf adReason <> adRsnFirstChange Then
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRS.Fields, moRS!mwcSitesKeyTarget
      Else
         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRS.Fields, goSession.Site.SiteKey
      End If
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in frmAlertCreate.moRs_WillChangeRecord. ", Err.Number, Err.Description
End Sub

Private Sub cmdLookupRoleTypes_Click()
   Dim loWork As Object
   Dim sSQL As String
   Dim loRs As Recordset
   Dim sTargetRT As String
   Dim sUserNameTarget As String
   Dim sRoleName As String
   Dim sInList As String
   On Error GoTo SubError

   ' OPTIONs 1.lookup roletypes  OR 2.lookup roletypes/Users OR 3.(UserSiteSpecific OR RT/Users)
   ' Ship List or Shore List or RT List
   
   ' possible future option UserSite Specific Only
   Set loWork = CreateObject("mwUtility.mwLookUpTool")
   loWork.InitSession goSession
   
   ' IN List
   If Len(mSSiteKeys) > 3 Then
      sInList = " And mwcUsers.Siteskey IN " & mSSiteKeys
   Else
      sInList = ""
   End If
   
   ' Test 2-data sets by options      DEV-2144
   '3a. Send Alert by Roletype/User designated ShoreUser
   '3b. Send Alert by User assigned sitekey and sitetype=shore
   '1. Send Alert by Roletype ONLY
   '2a. Send Alert by Roletype/User designated Non-shoreuser
   '2b. Send Alert by User assigned SiteKey and SiteType=ship
   
  If mIsUserSiteSpecific Then
   
      If chkShoreUsersOnly.value = vbChecked Then
      
         sInList = sInList & " And mwcSites.SiteType = 2  "  ' ShoreType
         
         ' UserSiteSpecific OR IsShoreUser requirement  DEV-2144
         sSQL = "SELECT mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName, mwcUsers.Siteskey, " _
            & " mwcSites.SiteType, 'Rs2' AS recordtype " _
            & " FROM mwcRoleType INNER JOIN (mwcSites RIGHT JOIN mwcUsers ON mwcSites.ID = mwcUsers.Siteskey) " _
            & " ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
            & " WHERE (mwcUsers.IsActive is not null AND mwcUsers.IsActive <> 0) AND Not (mwcRoleType.RoleTypeName Is Null) AND " _
            & "    ((mwcUsers.Siteskey <> 0 And mwcUsers.Siteskey Is Not Null) " & sInList _
            & "      OR (mwcUsers.IsShoreUser <> 0 And mwcUsers.IsShoreUser Is Not Null) ) "
   
      Else 'IsUserSiteSpecific & IsShipUser IsShipType
         
         sInList = sInList & " And mwcSites.SiteType = 1  "     ' ShipType
      
         ' UserSiteSpecific OR Non-ShoreUser requirement  DEV-2144
         sSQL = "SELECT  mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName, mwcUsers.Siteskey, " _
            & " mwcSites.SiteType, 'Rs1' as recordtype " _
            & " FROM mwcRoleType INNER JOIN (mwcSites RIGHT JOIN mwcUsers ON mwcSites.ID = mwcUsers.Siteskey) " _
            & " ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
            & " WHERE (mwcUsers.IsActive is not null AND mwcUsers.IsActive <> 0) AND Not (mwcRoleType.RoleTypeName Is Null) AND " _
            & "   (((mwcUsers.Siteskey <> 0 And mwcUsers.Siteskey Is Not Null) " & sInList _
            & "     ) Or ((mwcUsers.IsShoreUser = 0 Or mwcUsers.IsShoreUser Is Null)  and mwcUsers.Siteskey Is Null )) "

         
      End If
      sSQL = sSQL & " ORDER BY mwcRoleType.RoleTypeName, mwcUsers.UserName "

      ' clear multiline text list & mRsTargetRoleTypes
'      If goSession.IsAccess Or goSession.IsSqlServer Then
         loWork.FilterField = "RoleTypeName"
'      Else
'         loWork.FilterField = "mwcRoleType.RoleTypeName"
'      End If
      Set loRs = loWork.GetKeys(sSQL, "User Name", 2000, "Role", 3000, , , True, True, True)
      
      If loWork.IsCancelled Then
         KillObject loWork
         Exit Sub
      ElseIf loWork.IsDeleted Then
         txtTargetRoleTypes.Text = ""
         sTargetRtNames = ""
         sUserNames = ""
         CloseRecordset moRsUserKeys
         CloseRecordset moRsTargetRTKeys
         Exit Sub
      End If
      
      ' initalize variables
      CloseRecordset moRsUserKeys
      sUserNameTarget = ""
      Set moRsUserKeys = loRs
      
      moRsUserKeys.MoveFirst
      If moRsUserKeys.RecordCount > 0 Then
      
         ' either RoleTypekeys or UserKeys
         CloseRecordset moRsTargetRTKeys
         txtTargetRoleTypes.Text = ""
         sTargetRtNames = ""
      
         Do While Not moRsUserKeys.EOF
            sUserNameTarget = GetUserNameRoleName(moRsUserKeys!ID, sRoleName)
            If sRoleName <> "" Then
               sUserNameTarget = sRoleName & ", " & sUserNameTarget
            End If
            sUserNames = sUserNameTarget
            'txtUserTarget.Text = txtUserTarget.Text & sUserNameTarget & vbCrLf
            txtTargetRoleTypes.Text = txtTargetRoleTypes.Text & sUserNameTarget & vbCrLf
            moRsUserKeys.MoveNext
         Loop
         moRsUserKeys.MoveFirst
      End If


  ElseIf chkShowUsers.value = vbUnchecked Then
      
      ' lookup RoletypeName
      sSQL = "SELECT ID, RoleTypeName FROM mwcRoleType ORDER BY RoleTypeName"
      
      ' clear multiline text list & mRsTargetRoleTypes
      loWork.FilterField = "RoleTypeName"
      Set loRs = loWork.GetKeys(sSQL, "Role Type", 3000, , , , , True, True)
      
      If loWork.IsCancelled Then
         KillObject loWork
         Exit Sub
      ElseIf loWork.IsDeleted Then
         txtTargetRoleTypes.Text = ""
         sTargetRtNames = ""
         sUserNames = ""
         CloseRecordset moRsUserKeys
         CloseRecordset moRsTargetRTKeys
         Exit Sub
      End If
      
      ' initalize variables
      CloseRecordset moRsTargetRTKeys
      sTargetRtNames = ""
      txtTargetRoleTypes.Text = ""
      Set moRsTargetRTKeys = loRs
      
      moRsTargetRTKeys.MoveFirst
      If moRsTargetRTKeys.RecordCount > 0 Then
      
         ' either RoleTypekeys or UserKeys
         CloseRecordset moRsUserKeys
         sUserNames = ""
      
         Do While Not moRsTargetRTKeys.EOF
            sTargetRT = goSession.RoleType.GetRoleTypeName(moRsTargetRTKeys!ID)
            sTargetRtNames = sTargetRtNames & " " & sTargetRT
            'txtTargetRoleTypes.Text = txtTargetRoleTypes.Text & sTargetRT & vbCrLf
            'txtUserTarget.Text = txtUserTarget.Text & sTargetRT & vbCrLf
            txtTargetRoleTypes.Text = txtTargetRoleTypes.Text & sTargetRT & vbCrLf
            moRsTargetRTKeys.MoveNext
         Loop
         moRsTargetRTKeys.MoveFirst
      End If
   
   Else


      If chkShoreUsersOnly.value = vbChecked Then
         sSQL = "SELECT mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName " _
            & " FROM mwcRoleType INNER JOIN mwcUsers ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
            & " Where (Not RoleTypeName Is Null) And (mwcUsers.IsShoreUser <> 0 And mwcUsers.IsShoreUser Is NOT Null) AND (mwcUsers.IsActive is not null AND mwcUsers.IsActive <> 0) "
      Else
         sSQL = "SELECT mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName " _
            & " FROM mwcRoleType INNER JOIN mwcUsers ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
            & " Where (Not RoleTypeName Is Null) And (mwcUsers.IsShoreUser = 0 Or mwcUsers.IsShoreUser Is Null) AND (mwcUsers.IsActive is not null AND mwcUsers.IsActive <> 0) "
      End If
      sSQL = sSQL & " ORDER BY mwcRoleType.RoleTypeName, mwcUsers.UserName "

      ' clear multiline text list & mRsTargetRoleTypes
'      If goSession.IsAccess Or goSession.IsSqlServer Then
         loWork.FilterField = "RoleTypeName"
'      Else
'         loWork.FilterField = "mwcRoleType.RoleTypeName"
'      End If
      Set loRs = loWork.GetKeys(sSQL, "User Name", 2000, "Role", 3000, , , True, True, True)
      
      If loWork.IsCancelled Then
         KillObject loWork
         Exit Sub
      ElseIf loWork.IsDeleted Then
         txtTargetRoleTypes.Text = ""
         sTargetRtNames = ""
         sUserNames = ""
         CloseRecordset moRsUserKeys
         CloseRecordset moRsTargetRTKeys
         Exit Sub
      End If
      
      ' initalize variables
      CloseRecordset moRsUserKeys
      sUserNameTarget = ""
      Set moRsUserKeys = loRs
      
      moRsUserKeys.MoveFirst
      If moRsUserKeys.RecordCount > 0 Then
      
         ' either RoleTypekeys or UserKeys
         CloseRecordset moRsTargetRTKeys
         txtTargetRoleTypes.Text = ""
         sTargetRtNames = ""
      
         Do While Not moRsUserKeys.EOF
            sUserNameTarget = GetUserNameRoleName(moRsUserKeys!ID, sRoleName)
            If sRoleName <> "" Then
               sUserNameTarget = sRoleName & ", " & sUserNameTarget
            End If
            sUserNames = sUserNameTarget
            'txtUserTarget.Text = txtUserTarget.Text & sUserNameTarget & vbCrLf
            txtTargetRoleTypes.Text = txtTargetRoleTypes.Text & sUserNameTarget & vbCrLf
            moRsUserKeys.MoveNext
         Loop
         moRsUserKeys.MoveFirst
      End If
   
   End If
   
   Set loRs = Nothing
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.frmAlertCreate.cmdLookupRoleTypes_Click. ", Err.Number, Err.Description
   CloseRecordset loRs
End Sub


Private Sub cmdLookupSites_Click()
   Dim loWork As Object
   Dim sSQL As String
   Dim loRs As Recordset
   Dim sTargetSite As String
   Dim nShoreSiteKey As Long
   Dim IsShipType As Boolean
   Dim IsFirstItem As Boolean
   On Error GoTo SubError
   
'   ' lookup roletypes  ORIGINAL
'   Set loWork = CreateObject("mwUtility.mwLookUpTool")
'   loWork.InitSession goSession
'
'   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
'      nShoreSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
'      sSQL = "SELECT ID, SiteName FROM mwcSites where ID=" & _
'             goSession.Site.SiteKey & " or ID=" & nShoreSiteKey & " ORDER BY SiteName"
'   Else
'      sSQL = "SELECT ID, SiteName FROM mwcSites WHERE IsReplicateSite Is Not Null and IsReplicateSite<>0 ORDER BY SiteName"
'   End If
'
'   ' clear multiline text list & mRsTargetSiteKeys
'   Set loRs = loWork.GetKeys(sSQL, "Site", 3500, , , , , True, False)
   
   
   ' use alternate form for mtml & limit of sites
   If mThisSiteType = SITE_TYPE_SHIP Then
      IsShipType = True
   End If
   nShoreSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
   Set loWork = CreateObject("mwUtility.mwFleetWork")
   loWork.InitSession goSession

   loWork.SetAlertVariables IsShipType, mThisSiteKey, nShoreSiteKey

   Set loRs = loWork.GetSites_KeysMultiSites(0, 0, "")
   
   If loWork.IsCancelled Then
      KillObject loWork
      Exit Sub
   ElseIf loWork.IsDeleted Then
      sTargetSite = ""
      txtTargetSites.Text = ""
      mSSiteKeys = ""
      KillObject loWork
      Exit Sub
   End If
   
   ' when IsUserSiteSpecific=1
   mSSiteKeys = "( "                            ' Set IN (list)
   IsFirstItem = True
   
   ' initalize variables
   CloseRecordset moRsTargetSiteKeys
   sTargetSite = ""
   txtTargetSites.Text = ""
   Set moRsTargetSiteKeys = loRs
   
   moRsTargetSiteKeys.MoveFirst
   If moRsTargetSiteKeys.RecordCount > 0 Then
      Do While Not moRsTargetSiteKeys.EOF
         sTargetSite = goSession.Site.GetSiteName(moRsTargetSiteKeys!ID)
         sTargetSiteNames = sTargetSiteNames & " " & sTargetSite
         txtTargetSites.Text = txtTargetSites.Text & sTargetSite & vbCrLf
         
         ' keys list for UserSiteSpecific RT list
         If IsFirstItem = True Then
            mSSiteKeys = mSSiteKeys & moRsTargetSiteKeys!ID
            IsFirstItem = False
         Else
            mSSiteKeys = mSSiteKeys & ", " & moRsTargetSiteKeys!ID
         End If
         
         moRsTargetSiteKeys.MoveNext
      Loop
      mSSiteKeys = mSSiteKeys & ")"             ' closeup IN (list)
   End If
   Set loRs = Nothing
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.frmAlertCreate.cmdLookupSites_Click. ", Err.Number, Err.Description
   CloseRecordset loRs
End Sub


Private Function GetUserNameRoleName(mwcUsersKey As Long, Optional RoleName As String, Optional RoleTypeKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   If mwcUsersKey < 1 Then
      Exit Function
   End If
   
   'sSQL = "SELECT ID, UserName FROM mwcUsers WHERE ID = " & mwcUsersKey
   'sSQL = "SELECT mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName, mwcUsers.mwcRoleTypeKey " _
   '   & " FROM mwcRoleType INNER JOIN mwcUsers ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
   '   & " Where mwcUsers.ID = " & mwcUsersKey
   
   ' User with Site limitation -- DEV-2144 must be allowed Ship-->shore & shore-->Ship selection
   sSQL = "SELECT mwcUsers.ID, mwcUsers.UserName, mwcRoleType.RoleTypeName, mwcUsers.mwcRoleTypeKey " _
      & " FROM mwcRoleType INNER JOIN mwcUsers ON mwcRoleType.ID = mwcUsers.mwcRoleTypeKey " _
      & " Where mwcUsers.ID = " & mwcUsersKey
'   If mThisSiteType = SITE_TYPE_SHIP Then
'      sSQL = sSQL & " And (mwcUsers.IsShoreUser Is Null Or mwcUsers.IsShoreUser = 0) "         ' IsShipUser List Required"
'   Else
'      sSQL = sSQL & " And Not (mwcUsers.IsShoreUser Is Null Or mwcUsers.IsShoreUser = 0) "    ' IsShoreUser List Required
'   End If

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If BlankNull(loRs.Fields(1)) <> "" Then
         GetUserNameRoleName = loRs.Fields(1)
      End If
      If BlankNull(loRs.Fields(2)) <> "" Then
         RoleName = loRs.Fields(2)
      End If
      If ZeroNull(loRs.Fields(3)) <> 0 Then
         RoleTypeKey = loRs.Fields(3)
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.frmAlertCreate.GetUserNameRoleName. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Private Sub chkShowUsers_Click()
   On Error GoTo SubError
   
   If chkShowUsers.value = vbChecked Then
      chkShoreUsersOnly.Visible = True
   Else
      chkShoreUsersOnly.Visible = False
   End If
   
   cmdLookupRoleTypes_Click
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwSession.frmAlertCreate.chkShowUsers_Click. ", Err.Number, Err.Description
End Sub


Private Sub chkShoreUsersOnly_Click()
   cmdLookupRoleTypes_Click
End Sub

Private Sub RefreshUgColumns()
   Dim VisiblePosition(250) As Long   ' Used to store Column Positions
   Dim xx As Long
   Dim yy As Long
   Dim ColPos As Long
   On Error GoTo SubError
   
   ' Initialize the array elements to a value larger than the array extent
   InRefresh = True
   For xx = 0 To UBound(VisiblePosition)
      VisiblePosition(xx) = 999
   Next xx
   ColPos = 0   ' initialize the starting column number.
   

   HideUltragridColumns ug, 0
   ug.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ug.Override.FetchRows = ssFetchRowsPreloadWithParent
   
   ug.Bands(0).ColHeadersVisible = True
           
   ug.RowConnectorStyle = ssConnectorStyleRaised
   ug.Caption = ""
   
   ug.Bands(0).Columns(UG_mwcSitesName).Hidden = False
   ug.Bands(0).Columns(UG_mwcSitesName).Header.Caption = " From Site "
   ug.Bands(0).Columns(UG_mwcSitesName).Activation = ssActivationActivateNoEdit
   ug.Bands(0).Columns(UG_mwcSitesName).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ug.Bands(0).Columns(UG_mwcSitesName).Width = moReg.ugGetWidth("Ug", UG_mwcSitesName, 3500)
   VisiblePosition(UG_mwcSitesName) = moReg.ugGetPosition("Ug", UG_mwcSitesName, IncrCounter(ColPos))

   ug.Bands(0).Columns(UG_mwcRoleTypeName).Hidden = False
   ug.Bands(0).Columns(UG_mwcRoleTypeName).Header.Caption = "To Role"
   ug.Bands(0).Columns(UG_mwcRoleTypeName).Activation = ssActivationActivateNoEdit
   ug.Bands(0).Columns(UG_mwcRoleTypeName).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ug.Bands(0).Columns(UG_mwcRoleTypeName).Width = moReg.ugGetWidth("Ug", UG_mwcRoleTypeName, 3500)
   VisiblePosition(UG_mwcRoleTypeName) = moReg.ugGetPosition("Ug", UG_mwcRoleTypeName, IncrCounter(ColPos))

   ug.Bands(0).Columns(UG_mwcUsersName).Hidden = False
   ug.Bands(0).Columns(UG_mwcUsersName).Header.Caption = "To User"
   ug.Bands(0).Columns(UG_mwcUsersName).Activation = ssActivationActivateNoEdit
   ug.Bands(0).Columns(UG_mwcUsersName).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ug.Bands(0).Columns(UG_mwcUsersName).Width = moReg.ugGetWidth("Ug", UG_mwcUsersName, 3732)
   VisiblePosition(UG_mwcUsersName) = moReg.ugGetPosition("Ug", UG_mwcUsersName, IncrCounter(ColPos))
   
   
   For xx = 0 To UBound(VisiblePosition)
      For yy = 0 To UBound(VisiblePosition)
         If VisiblePosition(yy) = xx Then
            ug.Bands(0).Columns(yy).Header.VisiblePosition = xx
         End If
      Next yy
   Next xx
   InRefresh = False   ' Remember to turn the flag off when we're finishedUse
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.RefreshUgColumns ", Err.Number, Err.Description
End Sub
Private Sub cmdAlertAdd_Click()
   Dim loRs As Recordset
   Dim loWork As Object
   On Error GoTo SubError
   Set loWork = CreateObject("mwUtility.mwLookUpTool")
   loWork.InitSession goSession

   Set loRs = loWork.GetKeysAlert()
    
   If loWork.IsCancelled Then
      KillObject loWork
      Exit Sub
   ElseIf loWork.IsDeleted Then
      Exit Sub
   End If
   If IsRecordLoaded(loRs) Then
      loRs.MoveFirst
      If loRs.RecordCount > 0 Then
         If mUGNewID < 1 Then
            mUGNewID = 1
         End If
         Do While Not loRs.EOF
            If ZeroNull(loRs!mwcUsersKeyTarget) = 0 Then
               moRsAlertToList.Filter = "mwcSitesKeyTarget = " & loRs!mwcSitesKeyTarget & " AND mwcRoleTypeKeyTarget = " & loRs!mwcRoleTypeKeyTarget
            Else
               moRsAlertToList.Filter = "mwcSitesKeyTarget = " & loRs!mwcSitesKeyTarget & " AND mwcUsersKeyTarget=" & loRs!mwcUsersKeyTarget
            End If
            If moRsAlertToList.RecordCount < 1 Then
                  moRsAlertToList.AddNew
                  moRsAlertToList!ID = mUGNewID
                  moRsAlertToList!mwcSitesKeyTarget = loRs!mwcSitesKeyTarget
                  moRsAlertToList!SiteName = goSession.Site.GetSiteName(loRs!mwcSitesKeyTarget)
                  moRsAlertToList!mwcRoleTypeKeyTarget = loRs!mwcRoleTypeKeyTarget
                  moRsAlertToList!RoleName = goSession.RoleType.GetRoleTypeName(loRs!mwcRoleTypeKeyTarget)
                  moRsAlertToList!mwcUsersKeyTarget = loRs!mwcUsersKeyTarget
                  moRsAlertToList!username = GetUserNameRoleName(loRs!mwcUsersKeyTarget)
                  moRsAlertToList!IsRoleType = 0
                  moRsAlertToList.Update
                  mUGNewID = mUGNewID + 1
            End If
            moRsAlertToList.Filter = adFilterNone
            loRs.MoveNext
         Loop
      End If
      Set ug.DataSource = moRsAlertToList
      RefreshUgColumns
   End If
   
   Exit Sub
SubError:
    goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.cmdAlertAdd_Click ", Err.Number, Err.Description
    CloseRecordset moRsAlertToList
End Sub

Private Sub ug_AfterColPosChanged(ByVal Action As UltraGrid.Constants_PosChanged, ByVal Columns As UltraGrid.SSSelectedCols)
   Dim loCol As SSColumn
   On Error GoTo SubError
   ' Only save the settings if the InRefresh flag is False.
   If InRefresh = False Then
      For Each loCol In ug.Bands(0).Columns
         moReg.ugSetWidth "ug", loCol.Index, loCol.Width
         moReg.ugSetPosition "ug", loCol.Index, loCol.Header.VisiblePosition
      Next loCol
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertCreate.ug_AfterColPosChanged ", Err.Number, Err.Description
End Sub
