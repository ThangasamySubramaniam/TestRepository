VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmSetCurrentSite 
   Caption         =   "Set Current Site"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmSetCurrentSite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Click here to cancel this form without making changes"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Click here to implement your change. Careful, this is a No Twiddle Zone !"
      Top             =   2640
      Width           =   1455
   End
   Begin PVCOMBOLibCtl.PVComboBox pvcboSites 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Highlight a site and click Update to change the This Site Assignment"
      Top             =   1800
      Width           =   4215
      _Version        =   524288
      _cx             =   7435
      _cy             =   873
      Appearance      =   1
      Enabled         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Locked          =   0   'False
      Style           =   0
      Sorted          =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   7230
      DesignHeight    =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Set Current Site Configuration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Change To:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblCurrentSite 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "This is the current This Site Assignment"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Current Site: "
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
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmSetCurrentSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moCon As ADODB.Connection
Dim oRsThisSite As Recordset
Dim oRsSites As Recordset
Dim moParent As mwSession.Session

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdUpdate_Click()
   ' Nothing changes, get out...
   'If lblCurrentSite.Caption = pvcboSites.Text Then
   '   Me.Hide
   '   Exit Sub
   'End If
   If pvcboSites.ListIndex >= 0 Then
      oRsThisSite!ThisSite = pvcboSites.SubItem(pvcboSites.ListIndex, 0)
      oRsThisSite!IsThisSiteValidated = True
      oRsThisSite.Update
      MsgBox "System has been updated. Please re-enter ShipNet Fleet for changes to take effect.", vbInformation, "Set Current Site"
   Else
      MsgBox "System has not been updated.", vbExclamation, "Set Current Site"
   End If
   Me.Hide
      
End Sub

Private Sub Form_Load()
   On Error GoTo SubError
   goSession.LogIt mwl_Workstation, mwl_Information, "Entering Set Current Site Configuration, SiteID: " & goSession.Site.SiteID
   Set oRsThisSite = New Recordset
   oRsThisSite.CursorLocation = adUseClient
   oRsThisSite.Open "mwcThisSite", moCon, adOpenDynamic, adLockOptimistic, adCmdTable
   lblCurrentSite.Caption = goSession.Site.GetShipProperty(oRsThisSite!ThisSite, "SiteName")
   'pvcboSites.Text = oRsThisSite!ThisSite
   pvcboSites.Text = pvcboSites.Text = oRsThisSite!ThisSite
   
   Set oRsSites = New Recordset
   oRsSites.CursorLocation = adUseClient
   oRsSites.Open "SELECT SiteID,SiteName FROM mwcSites order by sitename", moCon, adOpenStatic, adLockReadOnly
   Set pvcboSites.ListSource = oRsSites
   pvcboSites.BoundColumn = "SiteID"
   pvcboSites.PrimaryColumn = 1              'category description column displayed
   pvcboSites.ColumnHidden(0) = True
   'pvcboSites.DataField = "SiteID"
   'pvcboSites.ColumnHidden(0) = True      ' make index column hidden
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
SubError:
   moParent.RaiseError "General Error in frmSetCurrentSite. ", Err.Number, Err.Description
   Me.Hide
End Sub


Public Function SetParentSession(ByRef ses As Session)
   Set moParent = ses
   Set moCon = moParent.DBConnection
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   moParent.CloseRecordset oRsThisSite
   moParent.CloseRecordset oRsSites
   goSession.LogIt mwl_Workstation, mwl_Information, "Exiting Set Current Site Configuration, SiteID: " & goSession.Site.SiteID

End Sub

