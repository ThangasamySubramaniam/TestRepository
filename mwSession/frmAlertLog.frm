VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmAlertLog 
   Caption         =   "Alert Log"
   ClientHeight    =   6885
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   12450
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
   ScaleHeight     =   6885
   ScaleWidth      =   12450
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLookupMyUserSites 
      Height          =   408
      Left            =   11880
      Picture         =   "frmAlertLog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   360
      Width           =   375
   End
   Begin VB.CheckBox chkMySites 
      Caption         =   "Limit By Sites"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   60
      Width           =   2715
   End
   Begin VB.TextBox txtMyUserSiteNames 
      BackColor       =   &H80000000&
      Height          =   435
      Left            =   9000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdViewEventDetails 
      Caption         =   "View Event"
      Height          =   735
      Left            =   8640
      Picture         =   "frmAlertLog.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "View the Event Details associated with this Alert"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Exit"
      Height          =   735
      Left            =   10680
      Picture         =   "frmAlertLog.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exit this Form"
      Top             =   3336
      Width           =   1215
   End
   Begin VB.CommandButton cmdClosed 
      Caption         =   "Close Alert"
      Height          =   375
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdMarkRead 
      Caption         =   "Mark As Read"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alert Details"
      Height          =   2712
      Left            =   60
      TabIndex        =   6
      Top             =   4080
      Width           =   12375
      Begin VB.TextBox txAlertDetails 
         DataField       =   "AlertDetails"
         Height          =   2415
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   12135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Include the following Alerts (blank for all)"
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   6195
      Begin VB.CheckBox chkSentByMe 
         Caption         =   "Sent"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         ToolTipText     =   "Show Alerts sent"
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox chkMyAlerts 
         Caption         =   "My Alerts"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         ToolTipText     =   "Show User Alerts Only"
         Top             =   300
         Width           =   1455
      End
      Begin VB.CheckBox chkClosed 
         Caption         =   "Closed"
         Height          =   255
         Left            =   2355
         TabIndex        =   5
         Top             =   300
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read"
         Height          =   255
         Left            =   1245
         TabIndex        =   4
         Top             =   300
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkSent 
         Caption         =   "New"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Value           =   1  'Checked
         Width           =   795
      End
   End
   Begin UltraGrid.SSUltraGrid ugAlertLog 
      Height          =   2475
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   4366
      _Version        =   131072
      GridFlags       =   17040384
      Images          =   "frmAlertLog.frx":12DE
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
      Override        =   "frmAlertLog.frx":1C3E
      CaptionAppearance=   "frmAlertLog.frx":1C94
      Caption         =   "Alert Log"
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   12450
      DesignHeight    =   6885
   End
   Begin VB.Label Label1 
      Caption         =   "Alert Log"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   300
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   180
      Picture         =   "frmAlertLog.frx":1CD0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmAlertLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' remember column width and position changes
   Dim moReg As Registry
   Dim InRefresh As Boolean
   
' PBT#110 MyUserSiteKeys, MyUserSiteNames
   Dim mMyUserSiteKeys As String
   Dim mMyUserSiteNames As String
   Dim mMyUser As String
   Dim moRsTargetSiteKeys As Recordset
   Dim mIsLoading As Boolean

Dim m_mwRoleTypeKey As Long
Dim m_ShowNewAlerts As Boolean

Dim WithEvents moRsAlertLog As Recordset
Attribute moRsAlertLog.VB_VarHelpID = -1
Dim FilterString As String
Dim IsLoading As Boolean

Const UGAlertLog_ID = 0
Const UGAlertLog_Title = 1
Const UGAlertLog_AlertDateTime = 2
Const UGAlertLog_mwcSitesKeySource = 3
Const UGAlertLog_mwcRoleTypeKeySource = 4
Const UGAlertLog_mwcSitesKeyDest = 5
Const UGAlertLog_mwcRoleTypeKeyDest = 6
Const UGAlertLog_AlertDetails = 7
Const UGAlertLog_mwAlertLogKeyFirst = 8
Const UGAlertLog_mwAlertLogKeyPrev = 9
Const UGAlertLog_ReceivedDateTime = 10
Const UGAlertLog_mwAlertLogStatusKey = 11
Const UGAlertLog_mwAlertTypeKey = 12
Const UGAlertLog_mwAlertEventsKey = 13
Const UGAlertLog_mwEventTypeKey = 14
Const UGAlertLog_mwEventDetailKey = 15
Const UGAlertLog_ReceiverNotes = 16
Const UGAlertLog_ExternalData = 17
Const UGAlertLog_mwcUsersKeySource = 18
Const UGAlertLog_mwcUsersKeyTarget = 19

Private Enum EnumEventViewerType
   EVT_NO_VIEWER = 0
   EVT_CERT_SHIP = 1
   EVT_OCCURRENCE = 2
   EVT_REQ = 3
   EVT_WO = 4
   EVT_PMS_CHANGE_REQUEST = 5
End Enum

Dim ShortDateFormat As String
Const SHORT_DATE = True

Const ALERT_SENT_ICON = "NewShort.ico"
Const ALERT_READ_ICON = "read.ico"
Const ALERT_CLOSED_ICON = "Closed2.ico"


Private Sub chkSentByMe_Click()
   On Error GoTo SubError
   
   RefreshAlertLogView
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.chkSentByMe_Click ", Err.Number, Err.Description
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   On Error GoTo SubError
   
   IsLoading = True
   
   If moReg Is Nothing Then
      Set moReg = New Registry
   End If
   moReg.BaseRegistry = BASE_REG & "mwSession." & Me.Name
   InRefresh = False
   
   
   If goSession.API.UserSessionSettingGet("frmAlertLog.Sent") = "1" Then
      chkSent.value = 1
   Else
      chkSent.value = 0
   End If

   If goSession.API.UserSessionSettingGet("frmAlertLog.Read") = "1" Then
      chkRead.value = 1
   Else
      chkRead.value = 0
   End If
   If goSession.API.UserSessionSettingGet("frmAlertLog.Closed") = "1" Then
      chkClosed.value = 1
   Else
      chkClosed.value = 0
   End If
   

   mMyUser = "frmAlertLog.MyAlerts." & goSession.User.UserKey           'MyAlerts.106
   mMyUserSiteKeys = goSession.API.UserSessionSettingGet(mMyUser & ".SiteKeys")
   mMyUserSiteNames = goSession.API.UserSessionSettingGet(mMyUser & ".SiteNames")
   txtMyUserSiteNames.Text = mMyUserSiteNames
   txtMyUserSiteNames.ToolTipText = Replace(mMyUserSiteNames, vbCrLf, ",")
   
   ' toggle viewing MyAlerts with MySites
   If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'      If chkMyAlerts.value = 1 Then
         chkMySites.Visible = True                 ' shore can view selected sites
         cmdLookupMyUserSites.Visible = True
         txtMyUserSiteNames.Visible = True
'      Else
'         chkMySites.Visible = False                ' hide until MyAlerts ticked
'         cmdLookupMyUserSites.Visible = False
'         txtMyUserSiteNames.Visible = False
'      End If
   Else  ' ship site
         chkMySites.Visible = False                ' ship does not need
         cmdLookupMyUserSites.Visible = False
         txtMyUserSiteNames.Visible = False
   End If
   If goSession.API.UserSessionSettingGet(mMyUser) = "1" Then     ' change fires refresh
      chkMySites.value = 1
   Else
      chkMySites.value = 0
   End If

   If goSession.API.UserSessionSettingGet("frmAlertLog.MyAlerts") = "1" Then
      chkMyAlerts.value = 1
   Else
      chkMyAlerts.value = 0
   End If


   FilterString = ""
   
   ShortDateFormat = goSession.API.GetDisplayDateFormat(SHORT_DATE)
'   ShortDateFormat = GetDisplayDateFormat(SHORT_DATE, NO_TIME)

   IsLoading = False
   RefreshAlertLogView
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.Form_Load ", Err.Number, Err.Description
   IsLoading = False
End Sub

Private Function RefreshAlertLogView()
   Dim CheckedStatuses As String
   Dim sSQL As String
   Dim OptionalString As String
   On Error GoTo FunctionError
   
   If IsLoading = True Then
      Exit Function
   End If
   
   CloseRecordset moRsAlertLog
   Set ugAlertLog.DataSource = Nothing
   
   If m_mwRoleTypeKey < 1 Then
      m_mwRoleTypeKey = goSession.User.RoleTypeKey
   End If
   
   CheckedStatuses = ""
   
   If chkSent = 1 Or m_ShowNewAlerts = True Then
      CheckedStatuses = "1"
   End If
   If chkRead = 1 Then
      If Len(CheckedStatuses) > 0 Then
         CheckedStatuses = CheckedStatuses & ", "
      End If
      CheckedStatuses = CheckedStatuses & "2"
   End If
   If chkClosed = 1 Then
      If Len(CheckedStatuses) > 0 Then
         CheckedStatuses = CheckedStatuses & ", "
      End If
      CheckedStatuses = CheckedStatuses & "4"
   End If
   
   Set moRsAlertLog = New Recordset
   
'   sSQL = "Select ID , Title , AlertDateTime, mwcSitesKeySource , mwcRoleTypeKeySource , mwcSitesKeyTarget , mwcRoleTypeKeyTarget , " & _
'          " AlertDetails , mwAlertLogKeyFirst , mwAlertLogKeyPrev , ReceivedDateTime , mwAlertLogStatusKey , " & _
'          " mwAlertTypeKey , mwAlertEventsKey , mwEventTypeKey , mwEventDetailKey , ReceiverNotes , ExternalData , " & _
'          " mwcUsersKeySource , mwcUsersKeyTarget " & _
'          " FROM mwAlertLog WHERE mwcRoleTypeKeyTarget = " & m_mwRoleTypeKey & " and mwcSitesKeyTarget=" & goSession.Site.SiteKey

   sSQL = "Select ID , Title , AlertDateTime, mwcSitesKeySource , mwcRoleTypeKeySource , mwcSitesKeyTarget , mwcRoleTypeKeyTarget , " & _
       " AlertDetails , mwAlertLogKeyFirst , mwAlertLogKeyPrev , ReceivedDateTime , mwAlertLogStatusKey , " & _
       " mwAlertTypeKey , mwAlertEventsKey , mwEventTypeKey , mwEventDetailKey , ReceiverNotes , ExternalData , " & _
       " mwcUsersKeySource , mwcUsersKeyTarget  FROM mwAlertLog "

   If chkSentByMe.value = vbChecked Then
      sSQL = sSQL & " WHERE mwcSitesKeySource=" & goSession.Site.SiteKey
      If chkMyAlerts.value = vbUnchecked Then
          sSQL = sSQL & " AND (mwcUsersKeySource = " & goSession.User.UserKey & _
         "  OR (mwcRoleTypeKeySource = " & goSession.User.RoleTypeKey & " AND mwcUsersKeySource Is Null )) "
      Else
         sSQL = sSQL & " AND (mwcUsersKeySource = " & goSession.User.UserKey & " ) "
      End If
      
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         If chkMySites.value = 1 Then
            If Len(mMyUserSiteKeys) > 4 Then                            ' selected sites Len > (b)
               sSQL = sSQL & "  AND mwcSitesKeyTarget IN " & mMyUserSiteKeys  ' limit to My selected sites
            End If
         End If
      End If
   Else
      sSQL = sSQL & " WHERE mwcSitesKeyTarget=" & goSession.Site.SiteKey
      If chkMyAlerts.value = vbUnchecked Then
         sSQL = sSQL & " AND (mwcUsersKeyTarget = " & goSession.User.UserKey & _
          "  OR (mwcRoleTypeKeyTarget = " & goSession.User.RoleTypeKey & " AND mwcUsersKeyTarget Is Null )) "
      Else
         sSQL = sSQL & " AND (mwcUsersKeyTarget = " & goSession.User.UserKey & " ) "
      End If
      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
         If chkMySites.value = 1 Then
            If Len(mMyUserSiteKeys) > 4 Then                            ' selected sites Len > (b)
               sSQL = sSQL & "  AND mwcSitesKeySource IN " & mMyUserSiteKeys  ' limit to My selected sites
            End If
         End If
      End If
   End If
'   If chkMyAlerts.value = vbChecked Then
'      OptionalString = " "
'      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'         If Me.Visible = False And chkMySites.Visible = False And chkMySites.value = 1 Then
'            ' load form has not been displayed yet
'            If Len(mMyUserSiteKeys) > 4 Then                            ' selected sites Len > (b)
'               OptionalString = "  AND mwcSitesKeyTarget IN " & mMyUserSiteKeys  ' limit to My selected sites
'            End If
'         ElseIf chkMySites.Visible = True And chkMySites.value = 1 Then
'            If Len(mMyUserSiteKeys) > 4 Then                            ' selected sites Len > (b)
'               OptionalString = "  AND mwcSitesKeyTarget IN " & mMyUserSiteKeys  ' limit to My selected sites
'            End If
'         End If
'      End If
'      sSQL = sSQL & " OR (mwcSitesKeySource=" & goSession.Site.SiteKey & _
'          " AND (mwcUsersKeySource = " & goSession.User.UserKey & OptionalString & _
'          "           OR (mwcRoleTypeKeySource = " & goSession.User.RoleTypeKey & _
'          " AND mwcUsersKeySource Is Null " & OptionalString & " ))) "
'   End If
   
   
                        
   If Len(CheckedStatuses) > 0 Then
      FilterString = " AND mwAlertLogStatusKey IN( " & CheckedStatuses & " )"
      sSQL = sSQL & FilterString
   End If
   
   
   sSQL = sSQL & " ORDER BY AlertDateTime DESC"
   
   moRsAlertLog.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Set ugAlertLog.DataSource = moRsAlertLog
   If ugAlertLog.HasRows = True Then
      Set ugAlertLog.ActiveRow = ugAlertLog.GetRow(ssChildRowFirst)
   End If
   Set txAlertDetails.DataSource = moRsAlertLog
'   ugAlertLog.CollapseAll
'   ' temporary show record count in MySites Tooltiptext
'   If IsRecordLoaded(moRsAlertLog) Then
'      chkMySites.ToolTipText = moRsAlertLog.RecordCount
'   End If
   
   RefreshUgAlertLogColumns
      
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in frmAlertLog.RefreshAlertLogView ", Err.Number, Err.Description
   CloseRecordset moRsAlertLog
End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo SubError
   goSession.API.UserSessionSettingSet "frmAlertLog.Sent", Trim$(str(chkSent.value))
   goSession.API.UserSessionSettingSet "frmAlertLog.Read", Trim$(str(chkRead.value))
   goSession.API.UserSessionSettingSet "frmAlertLog.Closed", Trim$(str(chkClosed.value))
   goSession.API.UserSessionSettingSet "frmAlertLog.MyAlerts", Trim$(str(chkMyAlerts.value))
   
   goSession.API.UserSessionSettingSet mMyUser, Trim$(str(chkMySites.value))
   goSession.API.UserSessionSettingSet mMyUser & ".SiteKeys", mMyUserSiteKeys    '   'MyAlerts.106"
   goSession.API.UserSessionSettingSet mMyUser & ".SiteNames", mMyUserSiteNames  '   'MyAlerts.106"
      
   CloseRecordset moRsAlertLog
   
   KillObject moReg
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.Form_QueryUnload ", Err.Number, Err.Description
End Sub



Private Sub moRsAlertLog_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Dim mwcSitesKeyTarget As Long
   Dim mwcRoleTypeKeyTarget As Long
   Dim mwcUsersKeyTarget As Long
   
   On Error GoTo SubError
   
   If IsRecordLoaded(moRsAlertLog) Then
   
      If ViewableEventType() > EVT_NO_VIEWER And ZeroNull(moRsAlertLog!mwEventDetailKey) > 0 Then
         cmdViewEventDetails.Visible = True
      Else
         cmdViewEventDetails.Visible = False
      End If
      
      ' MT-56 If the alert was sent to your user/role then the buttons should be enabled, else disabled.
      
      mwcSitesKeyTarget = ZeroNull(moRsAlertLog!mwcSitesKeyTarget)
      mwcRoleTypeKeyTarget = ZeroNull(moRsAlertLog!mwcRoleTypeKeyTarget)
      mwcUsersKeyTarget = ZeroNull(moRsAlertLog!mwcUsersKeyTarget)
      
      ' If this alert is for my site and either my role or my user then it is for me.
      ' Otherwise it must be FROM me (or my roletype)
      
      If mwcSitesKeyTarget = goSession.Site.SiteKey Then
         If (mwcRoleTypeKeyTarget = goSession.User.RoleTypeKey And mwcUsersKeyTarget = 0) Then
            ' This Alert is for me
            cmdMarkRead.Enabled = True
            cmdClosed.Enabled = True
         ElseIf (mwcUsersKeyTarget = goSession.User.UserKey) Then
            ' This Alert is for me
            cmdMarkRead.Enabled = True
            cmdClosed.Enabled = True
         Else
            cmdMarkRead.Enabled = False
            cmdClosed.Enabled = False
         End If
         
      Else ' This Alert was sent by me and not TO me
         cmdMarkRead.Enabled = False
         cmdClosed.Enabled = False
      End If
      
      Select Case ZeroNull(moRsAlertLog!mwAlertLogStatusKey)
         
         Case MW_ALERT_STATUS_READ
            cmdMarkRead.Enabled = False
         
         Case MW_ALERT_STATUS_CLOSED
            cmdMarkRead.Enabled = False
            cmdClosed.Enabled = False
      End Select
   End If
      
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.moRsAlertLog_MoveComplete ", Err.Number, Err.Description
End Sub


Private Sub ugAlertLog_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
'   HandleUG_KeyDown ugAlertLog, KeyCode, Shift
   Exit Sub
End Sub
Private Sub ugAlertLog_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
   Cancel = True
End Sub

Private Function RefreshUgAlertLogColumns()
   Dim VisiblePosition(250) As Long   ' Used to store Column Positions
   Dim xx As Long
   Dim yy As Long
   Dim ColPos As Long
   On Error GoTo FunctionError
   
   ' Initialize the array elements to a value larger than the array extent
   InRefresh = True
   For xx = 0 To UBound(VisiblePosition)
      VisiblePosition(xx) = 999
   Next xx
   ColPos = 0   ' initialize the starting column number.
   
   
'Const UGAlertLog_ID = 0
'Const UGAlertLog_Title = 1
'Const UGAlertLog_AlertDateTime = 2
'Const UGAlertLog_mwcSitesKeySource = 3
'Const UGAlertLog_mwcRoleTypeKeySource = 4
'Const UGAlertLog_mwcSitesKeyDest = 5
'Const UGAlertLog_mwcRoleTypeKeyDest = 6
'Const UGAlertLog_AlertDetails = 7
'Const UGAlertLog_mwAlertLogKeyFirst = 8
'Const UGAlertLog_mwAlertLogKeyPrev = 9
'Const UGAlertLog_ReceivedDateTime = 10
'Const UGAlertLog_mwAlertLogStatusKey = 11
'Const UGAlertLog_mwAlertTypeKey = 12
'Const UGAlertLog_mwAlertEventsKey = 13
'Const UGAlertLog_mwEventTypeKey = 14
'Const UGAlertLog_mwEventDetailKey = 15
'Const UGAlertLog_ReceiverNotes = 16
'Const UGAlertLog_ExternalData = 17, 18, 19
   
   HideUltragridColumns ugAlertLog, 0
'   HideUltragridColumns UgAlertLog, 1

   'ugAlertLog.Override.HeaderClickAction = ssHeaderClickActionSelect
   ugAlertLog.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ugAlertLog.Override.FetchRows = ssFetchRowsPreloadWithParent
   
   ugAlertLog.Bands(0).ColHeadersVisible = True
'   ugAlertLog.Bands(1).ColHeadersVisible = False
'   UgAlertLog.Bands(0).HeaderVisible = False
'   UgAlertLog.Bands(1).HeaderVisible = False
           
   ugAlertLog.RowConnectorStyle = ssConnectorStyleRaised
   ugAlertLog.Caption = ""
   
   
   'ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Hidden = False
   'ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Width = 1000
   'ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Header.Caption = "Status"
   'ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Activation = ssActivationActivateNoEdit
   'ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor

   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Header.Caption = " Created "
   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Format = goSession.API.GetDisplayDateFormat
   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_AlertDateTime, 2400)
   VisiblePosition(UGAlertLog_AlertDateTime) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_AlertDateTime, IncrCounter(ColPos))

   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Header.Caption = "From Site"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcSitesKeySource, 1800)
   VisiblePosition(UGAlertLog_mwcSitesKeySource) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcSitesKeySource, IncrCounter(ColPos))

   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Header.Caption = "From Role"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcRoleTypeKeySource, 1800)
   VisiblePosition(UGAlertLog_mwcRoleTypeKeySource) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcRoleTypeKeySource, IncrCounter(ColPos))
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).Header.Caption = "Title"
   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_Title, 3000)
   VisiblePosition(UGAlertLog_Title) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_Title, IncrCounter(ColPos))

   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).Header.Caption = "Read"
   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).Format = goSession.API.GetDisplayDateFormat
   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_ReceivedDateTime).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_ReceivedDateTime, 2500)
   VisiblePosition(UGAlertLog_ReceivedDateTime) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_ReceivedDateTime, IncrCounter(ColPos))


   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Header.Caption = "Sent to Role"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcRoleTypeKeyDest, 1800)
   VisiblePosition(UGAlertLog_mwcRoleTypeKeyDest) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcRoleTypeKeyDest, IncrCounter(ColPos))

   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Header.Caption = "From User"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcUsersKeySource, 1800)
   VisiblePosition(UGAlertLog_mwcUsersKeySource) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcUsersKeySource, IncrCounter(ColPos))

   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).Hidden = False
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).Header.Caption = "Sent To User"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).Activation = ssActivationActivateNoEdit
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcUsersKeyTarget, 1800)
   VisiblePosition(UGAlertLog_mwcUsersKeyTarget) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcUsersKeyTarget, IncrCounter(ColPos))
      
   If chkSentByMe.value = vbChecked Then
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Hidden = False
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Header.Caption = "Target Site"
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Activation = ssActivationActivateNoEdit
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Width = moReg.ugGetWidth("ugAlertLog", UGAlertLog_mwcSitesKeyDest, 1800)
      VisiblePosition(UGAlertLog_mwcSitesKeyDest) = moReg.ugGetPosition("ugAlertLog", UGAlertLog_mwcSitesKeyDest, IncrCounter(ColPos))
   Else
      ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Hidden = True
   End If
   
   For xx = 0 To UBound(VisiblePosition)
      For yy = 0 To UBound(VisiblePosition)
         If VisiblePosition(yy) = xx Then
            ugAlertLog.Bands(0).Columns(yy).Header.VisiblePosition = xx
         End If
      Next yy
   Next xx
   InRefresh = False   ' Remember to turn the flag off when we're finishedUse
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwCrew.frmPayroll.RefreshUgPeoplePayActColumns ", Err.Number, Err.Description
End Function

Private Sub UgAlertLog_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
   
   Dim loRs As Recordset
   Dim sSQL As String
   Dim nID As Long
   Dim sValue As String
   On Error GoTo SubError
   ugAlertLog.Bands(0).ColHeaderLines = 1
      
   On Error GoTo SubError
   
'    ugAlertLog.Override.RowSpacingAfter = 45
   
   ' Source SITE
   Set loRs = New Recordset
   
   sSQL = "SELECT ID, SiteName FROM mwcSites ORDER BY SiteName"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not ugAlertLog.ValueLists.Exists("Sites") Then
      ugAlertLog.ValueLists.Add ("Sites")
   Else
      ugAlertLog.ValueLists("Sites").ValueListItems.Clear
   End If
   
  Do While Not loRs.EOF
      nID = loRs!ID
      sValue = BlankNull(loRs!SiteName)
      ugAlertLog.ValueLists("Sites").ValueListItems.Add nID, sValue
      loRs.MoveNext
   Loop
      
   ugAlertLog.ValueLists("Sites").DisplayStyle = ssValueListDisplayStyleDisplayText
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).ValueList = "Sites"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).AutoEdit = True
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).ValueList = "Sites"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeyDest).AutoEdit = True
   
   loRs.Close
   
'Const UGAlertLog_mwcRoleTypeKeySource = 3 & 6 RTTarget

   sSQL = "SELECT ID, RoleTypeName FROM mwcRoleType ORDER BY RoleTypeName"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not ugAlertLog.ValueLists.Exists("RoleType") Then
      ugAlertLog.ValueLists.Add ("RoleType")
      ugAlertLog.ValueLists.Add ("TargetRT")
   Else
      ugAlertLog.ValueLists("RoleType").ValueListItems.Clear
      ugAlertLog.ValueLists("TargetRT").ValueListItems.Clear
   End If
   
  Do While Not loRs.EOF
      nID = loRs!ID
      'Added By N.Angelakis On 12th October 2009, Added blanknull
      sValue = BlankNull(loRs!RoleTypeName)
      ugAlertLog.ValueLists("RoleType").ValueListItems.Add nID, sValue
      ugAlertLog.ValueLists("TargetRT").ValueListItems.Add nID, sValue
      loRs.MoveNext
   Loop
      
   ugAlertLog.ValueLists("RoleType").DisplayStyle = ssValueListDisplayStyleDisplayText
   ugAlertLog.ValueLists("TargetRT").DisplayStyle = ssValueListDisplayStyleDisplayText
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).ValueList = "RoleType"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).AutoEdit = True
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).ValueList = "TargetRT"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).AutoEdit = True
   
   loRs.Close

'Const UGAlertLog_mwAlertLogStatusKey = 10

   sSQL = "SELECT ID, Description FROM mwAlertLogStatus ORDER BY ID"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not ugAlertLog.ValueLists.Exists("mwAlertLogStatus") Then
      ugAlertLog.ValueLists.Add ("mwAlertLogStatus")
   Else
      ugAlertLog.ValueLists("mwAlertLogStatus").ValueListItems.Clear
   End If
   
  Do While Not loRs.EOF
      nID = loRs!ID
      'Added By N.Angelakis On 12th October 2009, Added blanknull
      sValue = BlankNull(loRs!Description)
      ugAlertLog.ValueLists("mwAlertLogStatus").ValueListItems.Add nID, sValue
      loRs.MoveNext
   Loop
      
   ugAlertLog.ValueLists("mwAlertLogStatus").DisplayStyle = ssValueListDisplayStyleDisplayText
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).ValueList = "mwAlertLogStatus"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).AutoEdit = True
   
   loRs.Close

'Const UGAlertLog_mwAlertTypeKey = 11

   sSQL = "SELECT ID, Description FROM mwAlertType ORDER BY ID"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not ugAlertLog.ValueLists.Exists("mwAlertType") Then
      ugAlertLog.ValueLists.Add ("mwAlertType")
   Else
      ugAlertLog.ValueLists("mwAlertType").ValueListItems.Clear
   End If
   
  Do While Not loRs.EOF
      nID = loRs!ID
      'Added By N.Angelakis On 12th October 2009, Added blanknull
      sValue = BlankNull(loRs!Description)
      ugAlertLog.ValueLists("mwAlertType").ValueListItems.Add nID, sValue
      loRs.MoveNext
   Loop
      
   ugAlertLog.ValueLists("mwAlertType").DisplayStyle = ssValueListDisplayStyleDisplayText
   
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertTypeKey).ValueList = "mwAlertType"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertTypeKey).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertTypeKey).AutoEdit = True
   
   loRs.Close

'Const UGAlertLog_mwAlertEventsKey = 12
'Const UGAlertLog_mwEventTypeKey = 13
'Const UGAlertLog_mwEventDetailKey = 14

'Const UGAlertLog_mwcUsersKeySource = 18
'Const UGAlertLog_mwcUsersKeyTarget = 19
   
   sSQL = "SELECT ID, UserName FROM mwcUsers ORDER BY UserName"
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If Not ugAlertLog.ValueLists.Exists("UserSource") Then
      ugAlertLog.ValueLists.Add ("UserSource")
      ugAlertLog.ValueLists.Add ("UserTarget")
   Else
      ugAlertLog.ValueLists("UserSource").ValueListItems.Clear
      ugAlertLog.ValueLists("UserTarget").ValueListItems.Clear
   End If
   
  Do While Not loRs.EOF
      nID = loRs!ID
      sValue = BlankNull(loRs!username)
      ugAlertLog.ValueLists("UserSource").ValueListItems.Add nID, sValue
      ugAlertLog.ValueLists("UserTarget").ValueListItems.Add nID, sValue
      loRs.MoveNext
   Loop
      
   ugAlertLog.ValueLists("UserSource").DisplayStyle = ssValueListDisplayStyleDisplayText
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).ValueList = "UserSource"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).AutoEdit = True
   
   ugAlertLog.ValueLists("UserTarget").DisplayStyle = ssValueListDisplayStyleDisplayText
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).ValueList = "UserTarget"
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).Style = ssStyleDropDown
   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeyTarget).AutoEdit = True
   
   loRs.Close



'   ugAlertLog.Bands(0).Columns(UGAlertLog_mwAlertLogStatusKey).Header.VisiblePosition = 0
'   ugAlertLog.Bands(0).Columns(UGAlertLog_AlertDateTime).Header.VisiblePosition = 1
'   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcSitesKeySource).Header.VisiblePosition = 2
'   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeySource).Header.VisiblePosition = 3
'   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcUsersKeySource).Header.VisiblePosition = 4
'   ugAlertLog.Bands(0).Columns(UGAlertLog_Title).Header.VisiblePosition = 5
'   ugAlertLog.Bands(0).Columns(UGAlertLog_mwcRoleTypeKeyDest).Header.VisiblePosition = 12
   '
   ugAlertLog.Refresh ssFireInitializeRow
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.UgAlertLog_InitializeLayout ", Err.Number, Err.Description
End Sub

Private Sub cmdMarkRead_Click()
   Dim loRow As SSRow
   On Error GoTo SubError
   For Each loRow In ugAlertLog.Selected.Rows
      If loRow.Cells(UGAlertLog_mwAlertLogStatusKey).value = MW_ALERT_STATUS_SENT Then
         loRow.Cells(UGAlertLog_ReceivedDateTime).value = Format(Now, goSession.API.GetDisplayDateFormat)
         loRow.Cells(UGAlertLog_mwAlertLogStatusKey).value = MW_ALERT_STATUS_READ
         loRow.Update
      End If
      loRow.Selected = False
   Next loRow
   RefreshAlertLogView
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.UgAlertLog_InitializeLayout ", Err.Number, Err.Description
End Sub

Private Sub ugAlertLog_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
   On Error GoTo SubError
   Dim nStatus As Long
   Dim nType As Long
   On Error GoTo SubError
   ' Reqheader line graphic icon
   If Not IsNull(Row.Cells(UGAlertLog_mwAlertLogStatusKey).value) Then
      nStatus = Row.Cells(UGAlertLog_mwAlertLogStatusKey).value
      Select Case nStatus
         Case MW_ALERT_STATUS_SENT
'            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = _
'               LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & ALERT_SENT_ICON)
            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = "New"
'               LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & ALERT_SENT_ICON)
         
         Case MW_ALERT_STATUS_READ
            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = "Read"
'            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = _
'               LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & ALERT_READ_ICON)
         
         Case MW_ALERT_STATUS_CLOSED
            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = "Closed"
'            Row.Cells(UGAlertLog_AlertDateTime).Appearance.Picture = _
'               LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & ALERT_CLOSED_ICON)
      End Select
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.ugAlertLog_InitializeRow ", Err.Number, Err.Description
End Sub


Public Sub SetmwRoleTypeKey(mwRoleTypeKey As Long)
   m_mwRoleTypeKey = mwRoleTypeKey
End Sub
Public Sub SetShowNewAlerts(ShowNewAlerts As Boolean)
   m_ShowNewAlerts = ShowNewAlerts
'   RefreshAlertLogView
End Sub


Private Sub chkSent_Click()
   RefreshAlertLogView
End Sub

Private Sub chkClosed_Click()
   RefreshAlertLogView
End Sub

Private Sub chkCreated_Click()
   RefreshAlertLogView
End Sub

Private Sub chkRead_Click()
   RefreshAlertLogView
End Sub

Private Sub cmdReply_Click()
   MsgBox "Sorry you do not have authorization for this feature.", vbInformation, "Reply to Alert"
End Sub

Private Sub cmdClosed_Click()
   Dim loRow As SSRow
   On Error GoTo SubError
   For Each loRow In ugAlertLog.Selected.Rows
      If loRow.Cells(UGAlertLog_mwAlertLogStatusKey).value < MW_ALERT_STATUS_CLOSED Then
         loRow.Cells(UGAlertLog_ReceivedDateTime).value = Format(Now, goSession.API.GetDisplayDateFormat)
         loRow.Cells(UGAlertLog_mwAlertLogStatusKey).value = MW_ALERT_STATUS_CLOSED
         loRow.Update
      End If
      loRow.Selected = False
   Next loRow
   RefreshAlertLogView
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.cmdClosed_Click ", Err.Number, Err.Description

End Sub



'Private Sub moRsAlertLog_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'   Static IsBeginAdd As Boolean
'   Static IsBeginDelete As Boolean
'   On Error GoTo SubError
'   '
'   ' mwAlertLog must use this special WillChangeRecord handler because it is a
'   ' Site Specific table that does NOT have an mwcSitesKey.
'
'   ' DO NOT change this to the new-style mwReplicateWillChange class!
'
'   If adReason = adRsnAddNew Then
'      IsBeginAdd = True
'   ElseIf adReason = adRsnUpdate And IsBeginDelete Then
'      IsBeginDelete = False
'   ElseIf adReason = adRsnUpdate And IsBeginAdd Then
'      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRsAlertLog!ID, moRsAlertLog!mwcSitesKeyTarget, moRsAlertLog.Fields
'      Else
'         goSession.ReplicateWork.LogAddChange MWRT_mwAlertLog, moRsAlertLog!ID, goSession.Site.SiteKey, moRsAlertLog.Fields
'      End If
'      IsBeginAdd = False
'   ElseIf adReason = adRsnDelete Then
'      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRsAlertLog!ID, moRsAlertLog!mwcSitesKeyTarget
'      Else
'         goSession.ReplicateWork.LogDeleteChange MWRT_mwAlertLog, moRsAlertLog!ID, goSession.Site.SiteKey
'      End If
'      IsBeginDelete = True
'   ElseIf adReason <> adRsnFirstChange Then
'      If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRsAlertLog.Fields, moRsAlertLog!mwcSitesKeyTarget
'      Else
'         goSession.ReplicateWork.LogModifyChange MWRT_mwAlertLog, moRsAlertLog.Fields, goSession.Site.SiteKey
'      End If
'   End If
'
'   Exit Sub
'SubError:
'   goSession.RaisePublicError "General error in frmAlertCreate.moRs_WillChangeRecord. ", Err.Number, Err.Description
'End Sub

Private Sub chkMyAlerts_Click()
   On Error GoTo SubError
   
   ' toggle viewing MyAlerts with MySites
'   If goSession.Site.SiteType = SITE_TYPE_SHORE Then
'      If chkMyAlerts.value = 1 Then
'         chkMySites.Visible = True                 ' shore can view selected sites
'         cmdLookupMyUserSites.Visible = True
'         txtMyUserSiteNames.Visible = True
'      Else
'         chkMySites.Visible = False                ' ship does not need
'         cmdLookupMyUserSites.Visible = False
'         txtMyUserSiteNames.Visible = False
'      End If
'   End If
   
   RefreshAlertLogView
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.chkMyAlerts_Click ", Err.Number, Err.Description
End Sub

Private Sub cmdViewEventDetails_Click()
   Dim loCertShipWork As Object
   Dim loFleetSynchWork As Object
   Dim EventViewerType As EnumEventViewerType
   On Error GoTo SubError
   
   If IsRecordLoaded(moRsAlertLog) Then
   
      EventViewerType = ViewableEventType()
   
      Set loCertShipWork = CreateObject("mwSafety4.smCertShipWork")
      loCertShipWork.InitSession goSession
   
      Select Case EventViewerType
      
         Case EVT_CERT_SHIP
      
            loCertShipWork.ShowCertShipDetails ZeroNull(moRsAlertLog!mwEventDetailKey)
         
         Case EVT_OCCURRENCE
         
            loCertShipWork.DisplayEvent ZeroNull(moRsAlertLog!mwEventTypeKey), ZeroNull(moRsAlertLog!mwEventDetailKey), True
         
         Case EVT_REQ
         
            loCertShipWork.DisplayRequisition ZeroNull(moRsAlertLog!mwcSitesKeySource), ZeroNull(moRsAlertLog!mwEventDetailKey)
         
         Case EVT_WO
         
            loCertShipWork.DisplayWorkOrder ZeroNull(moRsAlertLog!mwEventDetailKey)
         
         Case EVT_PMS_CHANGE_REQUEST
         
            Set loFleetSynchWork = CreateObject("mwWorksCy.swFleetSynchWork")
            loFleetSynchWork.InitSession goSession
            
            loFleetSynchWork.ShowPMSChangeRequest ZeroNull(moRsAlertLog!mwEventDetailKey)
               
         Case Else
         
      End Select
      
      KillObject loCertShipWork
      KillObject loFleetSynchWork
      
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in frmAlertLog.cmdViewEventDetails_Click ", Err.Number, Err.Description
End Sub

Private Function ViewableEventType() As EnumEventViewerType
   Dim EventTableName As String
   Dim mwEventTypeKey As Long
   Dim mwEventDetailKey As Long
   On Error GoTo FuncError
   
   If IsRecordLoaded(moRsAlertLog) Then
   
      mwEventTypeKey = ZeroNull(moRsAlertLog!mwEventTypeKey)
      mwEventDetailKey = ZeroNull(moRsAlertLog!mwEventDetailKey)
      
      If mwEventTypeKey = MW_EVENT_Certificate_Ship And mwEventDetailKey > 0 Then
         ViewableEventType = EVT_CERT_SHIP
      ElseIf mwEventTypeKey = MW_EVENT_Eqpt_History And mwEventDetailKey > 0 Then
         ViewableEventType = EVT_WO
      ElseIf mwEventTypeKey = MW_EVENT_Requisiton_Header And mwEventDetailKey > 0 Then
         ViewableEventType = EVT_REQ
      ElseIf mwEventTypeKey = MW_EVENT_PMS_CHANGE_REQUEST And mwEventDetailKey > 0 Then
         ViewableEventType = EVT_PMS_CHANGE_REQUEST
      ElseIf mwEventTypeKey > 0 And mwEventDetailKey > 0 Then
      
         EventTableName = FetchEventTableName(mwEventTypeKey)
         
         If EventTableName = "SMOCCURRENCE" Then
            ViewableEventType = EVT_OCCURRENCE
         End If
      End If
      
   End If
   Exit Function
FuncError:
   goSession.RaisePublicError "General Error in frmAlertLog.ViewableEventType ", Err.Number, Err.Description
End Function

Private Function FetchEventTableName(mwEventTypeKey As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   Dim EventTableName As String
   On Error GoTo FunctionError
   
   sSQL = "SELECT TableName FROM mwEventType WHERE ID = " & mwEventTypeKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      FetchEventTableName = UCase(BlankNull(loRs!TableName))
   Else
      FetchEventTableName = ""
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   CloseRecordset loRs
   goSession.RaisePublicError "General Error in frmAlertLog.FetchEventTableName. ", Err.Number, Err.Description
   FetchEventTableName = ""
End Function


Private Sub ugAlertLog_AfterColPosChanged(ByVal Action As UltraGrid.Constants_PosChanged, ByVal Columns As UltraGrid.SSSelectedCols)
   Dim loCol As SSColumn
   On Error GoTo SubError

   ' Only save the settings if the InRefresh flag is False.
   If InRefresh = False Then
      For Each loCol In ugAlertLog.Bands(0).Columns
         moReg.ugSetWidth "ugAlertLog", loCol.Index, loCol.Width
         moReg.ugSetPosition "ugAlertLog", loCol.Index, loCol.Header.VisiblePosition
      Next loCol
   End If
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmAlertLog.ugAlertLog_AfterColPosChanged ", Err.Number, Err.Description
End Sub


'---------------


Private Sub chkMySites_Click()
   RefreshAlertLogView
End Sub


Private Sub cmdLookupMyUserSites_Click()
   Dim loWork As Object
   Dim sSQL As String
   Dim loRs As Recordset
   Dim sTargetSite As String
   Dim nShoreSiteKey As Long
   Dim IsShipType As Boolean
   Dim IsFirstItem As Boolean
   On Error GoTo SubError
   
   If goSession.Site.SiteType = SITE_TYPE_SHIP Then
      IsShipType = True
   End If
   nShoreSiteKey = goSession.Site.GetSiteKey(goSession.Site.TargetReplicateSiteID)
   Set loWork = CreateObject("mwUtility.mwFleetWork")
   loWork.InitSession goSession
   loWork.SetAlertVariables IsShipType, goSession.Site.SiteKey, nShoreSiteKey

   Set loRs = loWork.GetSites_KeysMultiSites(0, 0, "")
      
   If loWork.IsCancelled Then             ' cancelled?
      KillObject loWork
      Exit Sub
   ElseIf loWork.IsDeleted Then           ' deleted?
      sTargetSite = ""
      txtMyUserSiteNames.Text = ""
      txtMyUserSiteNames.ToolTipText = ""
      mMyUserSiteKeys = ""
      mMyUserSiteNames = ""
      
      KillObject loWork
      Exit Sub
   End If
   
   ' when IsUserSiteSpecific=1
   mMyUserSiteKeys = "( "                            ' Set IN (list)
   IsFirstItem = True
   
   ' initalize variables
   CloseRecordset moRsTargetSiteKeys
   sTargetSite = ""
   mMyUserSiteNames = ""
   txtMyUserSiteNames.Text = ""
   txtMyUserSiteNames.ToolTipText = ""
   Set moRsTargetSiteKeys = loRs
   
   moRsTargetSiteKeys.MoveFirst
   If moRsTargetSiteKeys.RecordCount > 0 Then
      Do While Not moRsTargetSiteKeys.EOF
         sTargetSite = goSession.Site.GetSiteName(moRsTargetSiteKeys!ID)
         If mMyUserSiteNames = "" Then
            mMyUserSiteNames = sTargetSite
         Else
            mMyUserSiteNames = mMyUserSiteNames & vbCrLf & sTargetSite
         End If
         txtMyUserSiteNames.Text = txtMyUserSiteNames.Text & sTargetSite & vbCrLf
         txtMyUserSiteNames.ToolTipText = Replace(mMyUserSiteNames, vbCrLf, ",")
         
         ' keys list for UserSiteSpecific RT list
         If IsFirstItem = True Then
            mMyUserSiteKeys = mMyUserSiteKeys & moRsTargetSiteKeys!ID
            IsFirstItem = False
         Else
            mMyUserSiteKeys = mMyUserSiteKeys & ", " & moRsTargetSiteKeys!ID
         End If
         
         moRsTargetSiteKeys.MoveNext
      Loop
      mMyUserSiteKeys = mMyUserSiteKeys & ")"             ' closeup IN (list)
   End If
   Set loRs = Nothing
   
   ' changed site selection also need to refresh log
   RefreshAlertLogView
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwAlertWork.cmdLookupSites_Click. ", Err.Number, Err.Description
   CloseRecordset loRs
End Sub

