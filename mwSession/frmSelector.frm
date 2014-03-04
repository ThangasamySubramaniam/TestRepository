VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Begin VB.Form frmSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Commercial Operator"
   ClientHeight    =   3192
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   855
      Left            =   3480
      Picture         =   "frmSelector.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Selected highlighted record"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   1920
      Picture         =   "frmSelector.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel Select "
      Top             =   1800
      Width           =   975
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
      Height          =   855
      Left            =   120
      Picture         =   "frmSelector.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Display the Online Reference Manual"
      Top             =   1800
      Width           =   1215
   End
   Begin PVCOMBOLibCtl.PVComboBox pvcboSelector 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      _Version        =   524288
      _cx             =   6800
      _cy             =   661
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Select Commercial Operator"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moRS As Recordset
Dim mIsCancelled As Boolean
Dim mChooseType As EType
Dim mIsHistoryChooser As Boolean

Private Enum EType
   CommercialOperator = 1
End Enum

Public Function IsCancelled() As Boolean
   IsCancelled = mIsCancelled
End Function

Public Function FetchValue() As String
   'On Error GoTo FunctionError
   'Select Case mChooseType
   '   Case Is = CommercialOperator
   '      FetchValue = pvCbo.SubItem(pvCbo.ListIndex, 1)
   '   Case Else
   '      FetchValue = ""
   'End Select
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.frmSelector.FetchValue.", err.Number, err.Description
   FetchValue = ""
End Function


Public Function InitForm(WhichType As String) As Boolean
   On Error GoTo FunctionError
   Select Case UCase(WhichType)
      Case Is = "COMMERCIAL_OPERATOR"
         InitForm = InitCommercialOperator
   End Select
   
   goSession.SetDotNetTheme Me
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.frmSelector.FetchValue.", err.Number, err.Description
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   '
   ' Set new value ?
   '
   If pvcboSelector.Text <> "" And pvcboSelector.ListIndex > -1 Then
      goSession.Site.SetExtendedProperty "mwcCommercialOperatorKey", pvcboSelector.SubItem(pvcboSelector.ListIndex, 0)
   End If
   
   
   Me.Hide
End Sub

Private Function InitCommercialOperator() As Boolean
   Dim strSQL As String
   Dim strTemp As String
   Dim strOpKey As String
   Dim i As Long
   On Error GoTo FunctionError
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   strSQL = "select * from mwcCommercialOperator"
   moRS.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If moRS.RecordCount < 1 Then
      goSession.RaiseError "Error in mwSession.frmSelector, no Commercial Operator records found."
      InitCommercialOperator = False
   Else
      Set pvcboSelector.ListSource = moRS
      pvcboSelector.ColumnHidden(0) = True
      pvcboSelector.ColumnHidden(2) = True
      pvcboSelector.PrimaryColumn = 1
      '
      ' Fetch current value
      '
      strOpKey = goSession.Site.GetExtendedProperty("mwcCommercialOperatorKey")
      If strOpKey <> "" Then
         For i = 0 To pvcboSelector.ListCount - 1
            If pvcboSelector.SubItem(i, 0) = CLng(strOpKey) Then
               pvcboSelector.ListIndex = i
            End If
         Next i
      End If
      InitCommercialOperator = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.frmSelector.FetchValue.", err.Number, err.Description
End Function

