VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmEventChooser 
   Caption         =   "Select Event Form Template"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   7935
      DesignHeight    =   6285
   End
   Begin VB.CommandButton cmdFormHelp 
      Caption         =   "Help"
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
      Left            =   600
      Picture         =   "frmEventChooser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Display the Online Reference Manual"
      Top             =   5400
      Width           =   1215
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
      Height          =   855
      Left            =   3480
      Picture         =   "frmEventChooser.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel Select "
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Picture         =   "frmEventChooser.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Selected highlighted record"
      Top             =   5400
      Width           =   1095
   End
   Begin UltraGrid.SSUltraGrid ug1 
      Height          =   4692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7692
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68158484
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColScrollRegions=   "frmEventChooser.frx":12DE
      Override        =   "frmEventChooser.frx":131C
      Caption         =   "Select Form Template"
   End
   Begin VB.Label LblMultiple 
      Caption         =   "Hold Down <Ctrl> Key to Select Multiple Rows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   1440
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   5052
   End
End
Attribute VB_Name = "frmEventChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmEventChooser - Choose forms and activities...
' 12/2002 ms
'

Option Explicit


Dim moRS As Recordset
Dim moRSLIC As Recordset
Dim mIsCancelled As Boolean
Dim mChooseType As EType
Dim mIsHistoryChooser As Boolean

Private Enum EType
   FormChooser = 1
   FactChooser = 2
   HistoryChooser = 3
   LinkChooser = 4
End Enum

   


' frmOffice UG Recordset mapping...
Const UG_FORM_mwEventFormTypeKey = 0
Const UG_FORM_TemplateID = 1
Const UG_FORM_DisplayIcon = 2
Const UG_FORM_Description = 3         '14
Const UG_FORM_FormHelpID_Override = 46
Const UG_FORM_FormHelpContextID = 47
Const HELP_MANUAL = "mwUser200OfficeForms.chm"

' mwEventFactType recordset
Const UG_ACT_ID = 0
Const UG_ACT_ActivityTitle = 6
Const UG_ACT_DisplayIcon = 4

' mwEventFactTypeSN recordset
Const UG_SN_FT_ID = 0
Const UG_SN_FT_mwEventFactCatSNKey = 2
Const UG_SN_FT_DisplayIcon = 5
Const UG_SN_FT_FactTitle = 8

' Fabricated mwEventFactType...
Const UG_FAB_ID = 0
Const UG_FAB_ActivityTitle = 1
Const UG_FAB_DisplayIcon = 2
Const UG_FAB_Category = 3
Const UG_FAB_LinkedFactTypeSNKey = 4


' mwEventHistoryType...
Const UG_HIST_ID = 0
Const UG_HIST_mwEventTypeKey = 1
Const UG_HIST_HistoryTitle = 2
Const UG_HIST_DisplaySequence = 3
Const UG_HIST_DisplayIcon = 4
Const UG_HIST_IsApproval = 5

' mwEventLinkType...
Const UG_LINK_ID = 0
Const UG_LINK_mwEventTypeKey = 1
Const UG_LINK_LinkTitle = 2
Const UG_LINK_DisplaySequence = 3
Const UG_LINK_DefaultDescription = 4
Const UG_LINK_DisplayIcon = 5
'etc
Private mReducedFactCategoryFilter As String
Private mFactTypeFilter As Integer
Public Property Let ReducedFactCategoryFilter(sFilter As String)
   '
   mReducedFactCategoryFilter = sFilter
End Property

Public Property Let FactTypeFilter(nFilter As Integer)
   mFactTypeFilter = nFilter
End Property

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function FormInitChooser(EventType As Long, Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loVfWork As mwEventFormWork
   On Error GoTo FunctionError
   ' Set form usage flag
   mChooseType = FormChooser
   Me.Caption = "Select Event Form Template"
   ug1.Caption = "Select Form Template"
   LblMultiple.Visible = True
   Set loVfWork = New mwEventFormWork
   Set moRS = loVfWork.FetchFormTemplatesRS(EventType, SiteKey)
   If moRS.RecordCount < 1 Then
      MsgBox "No form templates are available for EventType: " & goSession.EventTypes.Item(EventType).Description, vbInformation
      mIsCancelled = True
      CloseRecordset moRS
      FormInitChooser = False
      Exit Function
   End If
   Set ug1.DataSource = moRS
   HideUltragridColumns ug1, 0
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug1.Bands(0).Columns(UG_FORM_TemplateID).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_TemplateID).Width = 1500
   ug1.Bands(0).Columns(UG_FORM_TemplateID).Header.Caption = "Template ID"
   ' Description
   ug1.Bands(0).Columns(UG_FORM_Description).Hidden = False
   ug1.Bands(0).Columns(UG_FORM_Description).Width = 5000
   ug1.Bands(0).Columns(UG_FORM_Description).Header.Caption = "Description"
   FormInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FormInitChooser.", Err.Number, Err.Description
   FormInitChooser = False
End Function

Public Function EventFactInitChooser(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loCol As Collection
   Dim loVaWork As mwEventFactsWork
   On Error GoTo FunctionError
   ' Set form usage flag
   mChooseType = FactChooser
   LblMultiple.Visible = True
   ' Save these for validation after select...?
   Me.Caption = "Select New  Fact"
   ug1.Caption = "Select New Fact"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "SELECT * from mwEventFactType " & _
   '  " WHERE mwEventTypeKey=" & EventType & " and (IsMandatory=0 or IsMandatory is Null) and IsActive<>0 order by ActivityTitle"
   sSQL = "SELECT * from mwEventFactType " & _
     " WHERE mwEventTypeKey=" & EventType & " and (IsMandatory=0 or IsMandatory is Null) and IsActive<>0 order by DisplaySequence"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      MsgBox "No activities available for Type of Event: " & goSession.EventTypes.Item(EventType).Description, vbInformation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset loRs
      EventFactInitChooser = False
      Exit Function
   End If
   '
   ' Fetch Event Detail records...
   '
   Set loVaWork = New mwEventFactsWork
   Set loCol = loVaWork.FetchExpandedFieldList(EventType, EventDetailKey)
   '
   ' Fabricate recordset...
   '
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   If goSession.IsOracle Then
      moRS.Fields.Append loRs.Fields(UG_ACT_ID).Name, adInteger, 4
   Else
      moRS.Fields.Append loRs.Fields(UG_ACT_ID).Name, loRs.Fields(UG_ACT_ID).Type, loRs.Fields(UG_ACT_ID).DefinedSize
   End If
   moRS.Fields.Append loRs.Fields(UG_ACT_ActivityTitle).Name, loRs.Fields(UG_ACT_ActivityTitle).Type, loRs.Fields(UG_ACT_ActivityTitle).DefinedSize
   moRS.Fields.Append loRs.Fields(UG_ACT_DisplayIcon).Name, loRs.Fields(UG_ACT_DisplayIcon).Type, loRs.Fields(UG_ACT_DisplayIcon).DefinedSize
   moRS.Open
   '
   ' Parse activities...
   '
   Do While Not loRs.EOF
      If Not (IsInCollection(loCol, loRs!StartFieldTag) Or _
        IsInCollection(loCol, loRs!EndFieldTag) Or _
        IsInCollection(loCol, loRs!RemarksFieldTag)) Then
         '
         ' Append to fabricated recordset...
         '
         moRS.AddNew
         moRS!ID = loRs!ID
         moRS!ActivityTitle = loRs!ActivityTitle
         If Not IsNull(loRs!DisplayIcon) Then
            moRS!DisplayIcon = loRs!DisplayIcon
         Else
            moRS!DisplayIcon = "N/A"
         End If
         moRS.Update
      End If
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   Set loCol = Nothing
   '
   If moRS.RecordCount < 1 Then
      MsgBox "All available Facts have been added to this Event.", vbExclamation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset loRs
      EventFactInitChooser = False
      Exit Function
   End If
   moRS.MoveFirst
   'moRS.Sort = "ActivityTitle"
   'moRS.Sort = "DisplaySequence"
   Set ug1.DataSource = moRS
   HideUltragridColumns ug1, 0
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Hidden = False
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Width = 5000
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Header.Caption = "Activity Title"
   EventFactInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.EventFactInitChooser.", Err.Number, Err.Description
   EventFactInitChooser = False
End Function


'Public Function FetchTemplateID() As String
'   On Error GoTo FunctionError
'   If Not mIsFormChooser Then
'      FetchTemplateID = ""
'      Exit Function
'   End If
'   If ug1.ActiveRow Is Nothing Then
'      FetchTemplateID = ""
'   Else
'      FetchTemplateID = ug1.ActiveRow.Cells(UG_FORM_TemplateID).value
'   End If
'   Exit Function
'FunctionError:
'   goSession.RaiseError "General Error in frmEventChooser.FetchTemplateID.", err.Number, err.Description
'   FetchTemplateID = ""
'End Function

Public Function FetchFormTypeKey() As Long
   On Error GoTo FunctionError
   If Not mChooseType = FormChooser Then
      FetchFormTypeKey = -1
      Exit Function
   End If
   If ug1.ActiveRow Is Nothing Then
      FetchFormTypeKey = -1
   Else
      FetchFormTypeKey = ug1.ActiveRow.Cells(UG_FORM_mwEventFormTypeKey).value
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchFormTypeKey.", Err.Number, Err.Description
   FetchFormTypeKey = ""
End Function




Public Function FetchMwEventFactTypeKey() As Long
   On Error GoTo FunctionError
   If mChooseType = FormChooser Then
      FetchMwEventFactTypeKey = -1
      Exit Function
   End If
   If ug1.ActiveRow Is Nothing Then
      FetchMwEventFactTypeKey = -1
   Else
      If mChooseType = FactChooser Then
         FetchMwEventFactTypeKey = ug1.ActiveRow.Cells(UG_FAB_ID).value
      ElseIf mChooseType = HistoryChooser Then
         FetchMwEventFactTypeKey = ug1.ActiveRow.Cells(UG_HIST_ID).value
      ElseIf mChooseType = LinkChooser Then
         FetchMwEventFactTypeKey = ug1.ActiveRow.Cells(UG_LINK_ID).value
      Else
         FetchMwEventFactTypeKey = -1
      End If
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchMwEventFactTypeKey.", Err.Number, Err.Description
   FetchMwEventFactTypeKey = -1
End Function

'Public Function FetchActivityTitle() As String
'   On Error GoTo FunctionError
'   If mChooseType = FormChooser Then
'      FetchActivityTitle = ""
'      Exit Function
 '  End If
'   If ug1.ActiveRow Is Nothing Then
'      FetchActivityTitle = ""
'   Else
'      If mChooseType = FactChooser Then
'         FetchActivityTitle = ug1.ActiveRow.Cells(UG_FAB_ActivityTitle).value
'      ElseIf mChooseType = HistoryChooser Then
'         FetchActivityTitle = ug1.ActiveRow.Cells(UG_HIST_HistoryTitle).value
'      Else
'         FetchActivityTitle = ""
'      End If
'   End If
'   Exit Function
'FunctionError:
'   goSession.RaiseError "General Error in frmEventChooser.FetchActivityTitle.", err.Number, err.Description
'   FetchActivityTitle = ""
'End Function





Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdFormHelp_Click()
   goSession.API.ShowVbFormHelp Me.Name
End Sub

Private Sub cmdOK_Click()
   If ug1.Selected.Rows.Count < 1 Then
      If Not ug1.ActiveRow Is Nothing Then
         ug1.ActiveRow.Selected = True
      End If
   End If
   Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   CloseRecordset moRS
End Sub

Private Sub ug1_Click()
   Exit Sub
   'If ug1.ActiveRow Is Nothing Then
   '   Exit Sub
   'End If
   'If ug1.ActiveRow.Selected = True Then
   '   ug1.ActiveRow.Selected = False
   'Else
   '   ug1.ActiveRow.Selected = False
   'End If

End Sub

Private Sub ug1_DblClick()
   Me.Hide
End Sub


Private Sub ug1_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
   Dim nID As Long
   Dim sTitle As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo SubError
             
   If mChooseType = FactChooser Then
      ug1.ValueLists.Clear
      sSQL = "SELECT ID, CatDescription FROM mwEventFactCatSN ORDER BY CatDescription "
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      ug1.ValueLists.Add ("FactSNCat")
       
      ' need to use temporary variables because ug cannot use pointers for values
      Do While Not loRs.EOF
         nID = loRs!ID
         If IsNull(loRs!CatDescription) Then
            sTitle = ""
         Else
            sTitle = loRs!CatDescription
         End If
         ug1.ValueLists("FactSNCat").ValueListItems.Add nID, sTitle
         loRs.MoveNext
      Loop
      ug1.ValueLists("FactSNCat").DisplayStyle = ssValueListDisplayStyleDisplayText
      CloseRecordset loRs
      '3  mwEventFactCatSNKey
      ug1.Bands(0).Columns(UG_FAB_Category).ValueList = "FactSNCat"
      ug1.Bands(0).Columns(UG_FAB_Category).Style = ssStyleEdit
      ug1.Bands(0).Columns(UG_FAB_Category).AutoEdit = True
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.frmEventChooser.ug1_InitializeLayout. ", Err.Number, Err.Description
End Sub

Private Sub ug1_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
   On Error GoTo SubError
   Select Case mChooseType
      Case Is = FormChooser
         '
         If Not IsNull(Row.Cells(UG_FORM_DisplayIcon).value) Then
            Row.Cells(UG_FORM_TemplateID).Appearance.Picture = _
              LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & Row.Cells(UG_FORM_DisplayIcon).value)
         End If
      Case Is = FactChooser
         '
         If Not UCase(Row.Cells(UG_FAB_DisplayIcon).value) = "N/A" Then
            Row.Cells(UG_FAB_ActivityTitle).Appearance.Picture = _
              LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & Row.Cells(UG_FAB_DisplayIcon).value)
         End If
      Case Is = HistoryChooser
         '
         If Not IsNull(Row.Cells(UG_HIST_DisplayIcon).value) Then
            Row.Cells(UG_FAB_ActivityTitle).Appearance.Picture = _
              LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & Row.Cells(UG_HIST_DisplayIcon).value)
         End If
      Case Is = LinkChooser
         '
         If Not IsNull(Row.Cells(UG_LINK_DisplayIcon).value) Then
            Row.Cells(UG_LINK_LinkTitle).Appearance.Picture = _
              LoadPicture(goSession.GetAppPath() & "\icons\32x32\" & Row.Cells(UG_LINK_DisplayIcon).value)
         End If
         
   End Select
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmEventChooser.ug1_InitializeRow.", Err.Number, Err.Description
End Sub


Public Function EventHistoryInitChooser(EventType As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loCol As Collection
   Dim loVaWork As mwEventFactsWork
   On Error GoTo FunctionError
   ' Set form usage flag
   mChooseType = HistoryChooser
   ' Save these for validation after select...?
   Me.Caption = "Select New  History Type"
   ug1.Caption = "Select New History Type"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ' DEV-2091
   ' History Types 900-999 are system resrved...
   sSQL = "SELECT * from mwEventHistoryType " & _
     " WHERE mwEventTypeKey=" & EventType & " ORDER BY DisplaySequence "
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      MsgBox "No history types available for Type of Event: " & goSession.EventTypes.Item(EventType).Description, vbInformation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset loRs
      EventHistoryInitChooser = False
      Exit Function
   End If
   Set ug1.DataSource = loRs
   HideUltragridColumns ug1, 0
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug1.Bands(0).Columns(UG_HIST_HistoryTitle).Hidden = False
   ug1.Bands(0).Columns(UG_HIST_HistoryTitle).Width = 3500
   ug1.Bands(0).Columns(UG_HIST_HistoryTitle).Header.Caption = "History Title"
   EventHistoryInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.EventHistoryInitChooser.", Err.Number, Err.Description
   EventHistoryInitChooser = False
End Function

Public Function EventLinkInitChooser(EventType As Long) As Boolean
   Dim sSQL As String
   Dim loCol As Collection
   Dim loVaWork As mwEventFactsWork
   On Error GoTo FunctionError
   ' Set form usage flag
   mChooseType = LinkChooser
   ' Save these for validation after select...?
   Me.Caption = "Select New  Link Type"
   ug1.Caption = "Select New Link Type"
   Set moRSLIC = New Recordset
   moRSLIC.CursorLocation = adUseClient
   sSQL = "SELECT * from mwEventLinkType " & _
     " WHERE mwEventTypeKey=" & EventType
   moRSLIC.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If moRSLIC.RecordCount < 1 Then
      MsgBox "No Link types available for Type of Event: " & goSession.EventTypes.Item(EventType).Description, vbInformation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset moRSLIC
      EventLinkInitChooser = False
      Exit Function
   ElseIf moRSLIC.RecordCount = 1 Then
      '
      ' no display, choose the single record
      '
      
   End If
   Set ug1.DataSource = moRSLIC
   HideUltragridColumns ug1, 0
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug1.Bands(0).Columns(UG_LINK_LinkTitle).Hidden = False
   ug1.Bands(0).Columns(UG_LINK_LinkTitle).Width = 3500
   ug1.Bands(0).Columns(UG_LINK_LinkTitle).Header.Caption = "Link Title"
   
   ug1.Bands(0).Columns(UG_LINK_DefaultDescription).Hidden = False
   ug1.Bands(0).Columns(UG_LINK_DefaultDescription).Width = 3500
   ug1.Bands(0).Columns(UG_LINK_DefaultDescription).Header.Caption = "Description"
   
   EventLinkInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.EventLinkInitChooser.", Err.Number, Err.Description
   EventLinkInitChooser = False
End Function



Public Function FetchHistoryTitle() As String
   On Error GoTo FunctionError
   If mChooseType <> HistoryChooser Then
      FetchHistoryTitle = ""
      Exit Function
   End If
   If ug1.ActiveRow Is Nothing Then
      FetchHistoryTitle = ""
   Else
      FetchHistoryTitle = ug1.ActiveRow.Cells(UG_HIST_HistoryTitle).value
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchHistoryTitle.", Err.Number, Err.Description
   FetchHistoryTitle = ""
End Function

Public Function FetchLinkTitle() As String
   On Error GoTo FunctionError
   If mChooseType <> LinkChooser Then
      FetchLinkTitle = ""
      Exit Function
   End If
   If ug1.ActiveRow Is Nothing Then
      FetchLinkTitle = ""
   Else
      FetchLinkTitle = ug1.ActiveRow.Cells(UG_LINK_LinkTitle).value
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchLinkTitle.", Err.Number, Err.Description
   FetchLinkTitle = ""
End Function

Public Function FetchLinkKey() As Long
   On Error GoTo FunctionError
   If mChooseType <> LinkChooser Then
      Exit Function
   ElseIf moRSLIC.RecordCount = 1 Then
      moRSLIC.MoveFirst
      FetchLinkKey = moRSLIC!ID
   ElseIf Not ug1.ActiveRow Is Nothing Then
      FetchLinkKey = ug1.ActiveRow.Cells(UG_LINK_ID).value
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchLinkKey.", Err.Number, Err.Description
   FetchLinkKey = ""
End Function


Public Function IsSelectedRows() As Boolean
   If ug1.Selected.Rows.Count > 0 Then
      IsSelectedRows = True
   Else
      IsSelectedRows = False
   End If
End Function

Public Function FetchNextSelected(ByRef Title As String) As Long
   Dim loRow As SSRow
   On Error GoTo FunctionError
   If ug1.Selected.Rows.Count < 1 Then
      FetchNextSelected = -1
      Exit Function
   End If
   Select Case mChooseType
      Case Is = FactChooser
         Set loRow = ug1.Selected.Rows(0)
         FetchNextSelected = loRow.Cells(UG_FAB_ID).value
         Title = loRow.Cells(UG_FAB_ActivityTitle).value
         loRow.Selected = False
         Set loRow = Nothing
      Case Is = FormChooser
         Set loRow = ug1.Selected.Rows(0)
         FetchNextSelected = loRow.Cells(UG_FORM_mwEventFormTypeKey).value
         
         'VEL-206 Forms Compatibility Issue
         'BY N.Angelakis 10 April 2011
         'description may be null/empty added blanknull in order not to cause error
         'Title = loRow.Cells(UG_FORM_Description).value
         Title = BlankNull(loRow.Cells(UG_FORM_Description).value)
         
         loRow.Selected = False
         Set loRow = Nothing
      Case Else
         FetchNextSelected = -1
         Exit Function
   End Select
      
      
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchNextSelected.", Err.Number, Err.Description
   FetchNextSelected = -1
End Function

Public Function FetchNextSelectedFactSN(ByRef Title As String, ByRef LinkedFactTypeSNKey As Long) As Long
   Dim loRow As SSRow
   On Error GoTo FunctionError
   If ug1.Selected.Rows.Count < 1 Then
      FetchNextSelectedFactSN = -1
      Exit Function
   End If
   Select Case mChooseType
      Case Is = FactChooser
         Set loRow = ug1.Selected.Rows(0)
         FetchNextSelectedFactSN = loRow.Cells(UG_FAB_ID).value
         Title = loRow.Cells(UG_FAB_ActivityTitle).value
         LinkedFactTypeSNKey = ZeroNull(loRow.Cells(UG_FAB_LinkedFactTypeSNKey).value)
         loRow.Selected = False
         Set loRow = Nothing
      Case Is = FormChooser
         Set loRow = ug1.Selected.Rows(0)
         FetchNextSelectedFactSN = loRow.Cells(UG_FORM_mwEventFormTypeKey).value
         Title = loRow.Cells(UG_FORM_Description).value
         LinkedFactTypeSNKey = 0
         loRow.Selected = False
         Set loRow = Nothing
      Case Else
         FetchNextSelectedFactSN = -1
         Exit Function
   End Select
      
      
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.FetchNextSelectedFactSN.", Err.Number, Err.Description
   FetchNextSelectedFactSN = -1
End Function


Public Function EventFactSNInitChooser(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loCol As Collection
   Dim loWork As mwEventFactWorkSN
   On Error GoTo FunctionError
   ' Set form usage flag
   mChooseType = FactChooser
   LblMultiple.Visible = True
   ' Save these for validation after select...?
   Me.Caption = "Select New  Fact"
   ug1.Caption = "Select New Fact"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "SELECT * from mwEventFactTypeSN " & _
   '  " WHERE mwEventTypeKey=" & EventType & " and (IsMandatory=0 or IsMandatory is Null) and IsActive<>0 order by ActivityTitle"
   'sSQL = "SELECT * from mwEventFactTypeSN " & _
   '  " WHERE mwEventTypeKey=" & EventType & " and (IsMandatory=0 or IsMandatory is Null) and IsActive<>0 order by DisplaySequence"
   sSQL = "SELECT mwEventFactTypeSN.* " & _
    " FROM mwcFleetSites, mwcFleets, mwEventFactTypeSN, mwEventFactTypeSnFleet where " & _
    " mwEventTypeKey=" & EventType & " and  mwEventFactTypeSN.ID = mwEventFactTypeSnFleet.mwEventFactTypeSNKey and " & _
    " mwcFleets.ID = mwEventFactTypeSnFleet.mwcFleetsKey and " & _
    " mwcFleetSites.mwcFleetsKey = mwcFleets.ID and mwcFleetSites.mwcSitesKey=" & gAddEventFactSiteKey & _
    " and (IsMandatory=0 or IsMandatory is Null) and IsActive<>0 "
   
   ' reduced facts category filters (mwEventFactCatSNKey
   If mReducedFactCategoryFilter <> "" Then
      sSQL = sSQL & " And " & mReducedFactCategoryFilter
   End If
   
   If mFactTypeFilter = VRS_FACT_TYPE_CARGO_BOTH Then
      sSQL = sSQL & " AND (IsLoadPortFact <> 0 OR IsDiscPortFact <> 0)"
   ElseIf mFactTypeFilter = VRS_FACT_TYPE_CARGO_LOAD Then
      sSQL = sSQL & " AND IsLoadPortFact <> 0 "
   ElseIf mFactTypeFilter = VRS_FACT_TYPE_CARGO_DISCH Then
      sSQL = sSQL & " AND IsDiscPortFact <> 0"
   ElseIf (mFactTypeFilter = VRS_FACT_TYPE_CANAL Or mFactTypeFilter = VRS_FACT_TYPE_BUNKERING) Then
      sSQL = sSQL & " AND IsCanalBunkerFact <> 0"
   End If
   
   sSQL = sSQL & " ORDER BY DisplaySequence"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   
   If loRs.RecordCount < 1 Then
      MsgBox "No Facts available for this Event Type: " & goSession.EventTypes.Item(EventType).Description, vbInformation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset loRs
      EventFactSNInitChooser = False
      Exit Function
   End If
   '
   ' Fetch Event Detail records...
   '
   Set loWork = New mwEventFactWorkSN
   Set loCol = loWork.FetchExpandedFieldList(EventType, EventDetailKey)
   '
   ' Fabricate recordset...
   '
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Fields.Append "ID", adInteger, 4
   moRS.Fields.Append "FactTitle", adVarChar, 50
   moRS.Fields.Append "DisplayIcon", adVarChar, 50
   moRS.Fields.Append "mwEventFactCatSNKey", adInteger, 4
   moRS.Fields.Append "LinkedFactTypeSNKey", adInteger, 4, adFldIsNullable
   moRS.Open
   '
   ' Parse activities...
   '
   Do While Not loRs.EOF
'      If Not (IsInCollection(loCol, loRs!StartColumnName) Or _
'        IsInCollection(loCol, loRs!EndColumnName) Or _
'        IsInCollection(loCol, loRs!FactValueColumnName) Or _
'        IsInCollection(loCol, loRs!RemarksColumnName)) Then
'         '
'         ' Append to fabricated recordset...
'         '
'         moRS.AddNew
'         moRS!ID = loRs!ID
'         'moRS!ActivityTitle = loRS!FactTitle
'         moRS!FactTitle = loRs!FactTitle
'         If Not IsNull(loRs!DisplayIcon) Then
'            moRS!DisplayIcon = loRs!DisplayIcon
'         Else
'            moRS!DisplayIcon = "N/A"
'         End If
'         moRS.Update
'      End If
      moRS.AddNew
      moRS!ID = loRs!ID
      moRS!FactTitle = loRs!FactTitle
      If Not IsNull(loRs!DisplayIcon) Then
         moRS!DisplayIcon = loRs!DisplayIcon
      Else
         moRS!DisplayIcon = "N/A"
      End If
      moRS!mwEventFactCatSNkey = loRs!mwEventFactCatSNkey
      moRS!LinkedFactTypeSNKey = loRs!LinkedFactTypeSNKey
      moRS.Update
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   Set loCol = Nothing
   '
   
   If moRS.RecordCount < 1 Then
      MsgBox "All available Facts have been added to this Event.", vbExclamation, "Add New Fact"
      mIsCancelled = True
      CloseRecordset loRs
      EventFactSNInitChooser = False
      Exit Function
   End If
   moRS.MoveFirst
   Set ug1.DataSource = moRS
   HideUltragridColumns ug1, 0
   ug1.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   ' Sequence Text
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Hidden = False
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Width = 5000
   ug1.Bands(0).Columns(UG_FAB_ActivityTitle).Header.Caption = "Fact Title"
   
   ug1.Bands(0).Columns(UG_FAB_Category).Hidden = False
   ug1.Bands(0).Columns(UG_FAB_Category).Width = 2000
   ug1.Bands(0).Columns(UG_FAB_Category).Header.Caption = "Category"
   EventFactSNInitChooser = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in frmEventChooser.EventFactSNInitChooser.", Err.Number, Err.Description
   EventFactSNInitChooser = False
End Function

Private Sub Form_Load()
   On Error GoTo SubError
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   goSession.RaiseError "General Error in frmEventChooser.Form_Load.", Err.Number, Err.Description
End Sub