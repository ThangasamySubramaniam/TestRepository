VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{C2000000-FFFF-1100-8000-000000000001}#8.0#0"; "PVMask.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmailConfiguration 
   Caption         =   "ShipNet Fleet Email  Configuration"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   8790
   StartUpPosition =   1  'CenterOwner
   Begin PVMaskEditLib.PVMaskEdit pvmMailPassword 
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   5570
      Width           =   1095
      _Version        =   524288
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   253
      Text            =   "M"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Text            =   "M"
      Mask            =   "************"
      EditMode        =   0
      PasswordOnly    =   -1  'True
   End
   Begin VB.TextBox txtPop3EmailPortNumber 
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
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3630
      Width           =   1335
   End
   Begin VB.TextBox txtSMTPEmailPortNumber 
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
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2660
      Width           =   1335
   End
   Begin VB.TextBox txtFromEmailAddress 
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
      Left            =   3360
      TabIndex        =   9
      Top             =   4600
      Width           =   5175
   End
   Begin VB.TextBox txtPop3Server 
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
      Left            =   3360
      TabIndex        =   4
      Top             =   2175
      Width           =   5175
   End
   Begin VB.TextBox txtMailServer 
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1690
      Width           =   5175
   End
   Begin VB.TextBox txtMailUserID 
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
      Left            =   3360
      TabIndex        =   10
      Top             =   5085
      Width           =   5175
   End
   Begin VB.TextBox txtSendByFolder 
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
      Left            =   3360
      TabIndex        =   12
      Top             =   6060
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton cmdImportCrewList 
      Caption         =   "Import Crewlist"
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
      Left            =   3600
      Picture         =   "frmEmailConfiguration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Width           =   2055
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1800
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   8790
      DesignHeight    =   7545
   End
   Begin VB.CommandButton cmdFormHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      Picture         =   "frmEmailConfiguration.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   975
      Left            =   7560
      Picture         =   "frmEmailConfiguration.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   975
      Left            =   240
      Picture         =   "frmEmailConfiguration.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2400
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PVCOMBOLibCtl.PVComboBox pvcboEmailCarrier 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1205
      Width           =   2895
      _Version        =   524288
      _cx             =   5106
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
      DropEffect      =   1
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
   Begin PVCOMBOLibCtl.PVComboBox pvcboDefaultTransport 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   2895
      _Version        =   524288
      _cx             =   5106
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
      DropEffect      =   1
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
   Begin PVCOMBOLibCtl.PVComboBox pvcboSMTPEmailSecurityProtocol 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3150
      Width           =   2895
      _Version        =   524288
      _cx             =   5106
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
      DropEffect      =   1
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
   Begin PVCOMBOLibCtl.PVComboBox pvcboPop3EmailSecurityProtocol 
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   4110
      Width           =   2895
      _Version        =   524288
      _cx             =   5106
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
      DropEffect      =   1
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
   Begin VB.Label lblPop3SecurityProtocol 
      Caption         =   "Pop3 Security Protocol:"
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
      Left            =   360
      TabIndex        =   28
      Top             =   4175
      Width           =   2655
   End
   Begin VB.Label lblPop3EmailPortNumber 
      Caption         =   "Pop3 Email Port Number:"
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
      Left            =   360
      TabIndex        =   27
      Top             =   3690
      Width           =   2535
   End
   Begin VB.Label lblSMTPEmailPortNumber 
      Caption         =   "SMTP Email Port Number:"
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
      Left            =   360
      TabIndex        =   26
      Top             =   2720
      Width           =   2775
   End
   Begin VB.Label lblSMTPSecurityProtocol 
      Caption         =   "SMTP Security Protocol:"
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
      Left            =   360
      TabIndex        =   25
      Top             =   3205
      Width           =   2775
   End
   Begin VB.Label lblFromEmailAddress 
      Caption         =   "From Email Address:"
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
      Left            =   360
      TabIndex        =   24
      Top             =   4660
      Width           =   2295
   End
   Begin VB.Label lblMailPop3Server 
      Caption         =   "Pop3 Mail Server:"
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
      Left            =   360
      TabIndex        =   23
      Top             =   2235
      Width           =   2295
   End
   Begin VB.Label lblEmailCarrier 
      Caption         =   "Email Carrier"
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
      Left            =   360
      TabIndex        =   22
      Top             =   1265
      Width           =   1935
   End
   Begin VB.Label lblMailPassword 
      Caption         =   "Mail Password:"
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
      Left            =   360
      TabIndex        =   21
      Top             =   5630
      Width           =   2055
   End
   Begin VB.Label lblMailUserID 
      Caption         =   "Mail User ID:"
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
      Left            =   360
      TabIndex        =   20
      Top             =   5145
      Width           =   2175
   End
   Begin VB.Label lblMailServer 
      Caption         =   "SMTP Mail Server:"
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
      Left            =   360
      TabIndex        =   19
      Top             =   1750
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Default Transport:"
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
      Left            =   360
      TabIndex        =   18
      Top             =   780
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Send By Folder"
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
      Left            =   360
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Email Configuration"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2318
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmEmailConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmEmailConfiguration - ShipNet Fleet Configuration
' 1/7/2001 ms
'

Option Explicit


Const HELP_MANUAL = "mwUser800Configuration.chm"
Const Carrier_Pigeon = 1
Const Amos_Mail = 2
Const Lotus_Notes = 3
Const Outlook_2002 = 4
Const Microsoft_Outlook = 5
Const SMTP_POP3 = 6
Const Email_Carrier = 7
Const Demo_Loopback = 8
Const Send_By_Media = 9
Const Electronic_Mail = 10
Const Transport_Container = 11
Const Default_Transport = 12



Private Sub cmdCancel_Click()
   Me.Hide

End Sub


Private Sub cmdFormHelp_Click()
   On Error GoTo SubError
   goSession.API.ShowMwHelp HELP_MANUAL
   Exit Sub
SubError:
   goSession.RaiseError "General error in frmEmailConfiguration.cmdFormHelp_Click.", Err.Number, Err.Description
End Sub


Private Sub Form_Initialize()
   On Error GoTo SubError
SubError:
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim nEmailSecurityProtocol As Integer
   ' Init screen from current registry values
   '
   '
   ' Preferences
   '
   On Error GoTo Form_load_error
   
   
   If goSession.ThisSite.CompanyID <> "NOVOSHIP" Then
      cmdImportCrewList.Visible = False
   End If
   
   '
   ' Transport
   '
   '
   pvcboEmailCarrier.AppendItem "SMTP/POP3"
   pvcboEmailCarrier.AppendItem "Microsoft Outlook"
   pvcboEmailCarrier.AppendItem "Microsoft Exchange"
   pvcboEmailCarrier.AppendItem "Lotus Notes"
   pvcboEmailCarrier.AppendItem "AMOS Mail"
   pvcboEmailCarrier.AppendItem "AMOS Link"
   pvcboEmailCarrier.AppendItem "Outlook 2002/2007"
   pvcboEmailCarrier.AppendItem "Groupwise"
   pvcboEmailCarrier.AppendItem "MAPI"
   'DEV-1797 Sending Email with Outlook client 97
   'Added By N.Angelakis On 02 Feb 2010
   pvcboEmailCarrier.AppendItem "Outlook 97"
   
   'DEV-1809 Send By Media Error Notification
   'Added By N.Angelakis On 04 Feb 2010
   pvcboEmailCarrier.AppendItem "Send By Media Link"
   
   Select Case goSession.User.DefaultEmailCarrier
      Case Is = mw_SMTP
         pvcboEmailCarrier.ListIndex = 0
      Case Is = mw_OUTLOOK
         pvcboEmailCarrier.ListIndex = 1
      Case Is = mw_EXCHANGE
         pvcboEmailCarrier.ListIndex = 2
      Case Is = mw_NOTES
         pvcboEmailCarrier.ListIndex = 3
      Case Is = mw_AMOS_MAIL
         pvcboEmailCarrier.ListIndex = 4
      Case Is = mw_AMOS_LINK
         pvcboEmailCarrier.ListIndex = 5
      Case Is = mw_OUTLOOK_2002
         pvcboEmailCarrier.ListIndex = 6
      Case Is = mw_GROUPWISE
         pvcboEmailCarrier.ListIndex = 7
      Case Is = mw_MAPI
         pvcboEmailCarrier.ListIndex = 8
      'DEV-1797 Sending Email with Outlook client 97
      'Added By N.Angelakis On 02 Feb 2010
      Case Is = mw_OUTLOOK_97
         pvcboEmailCarrier.ListIndex = 9
         
      'DEV-1809 Send By Media Error Notification
      'Added By N.Angelakis On 04 Feb 2010
      Case Is = mw_SENDBYMEDIA_LINK
         pvcboEmailCarrier.ListIndex = 10
         
         
   End Select
'   If goSession.User.Security.AllowEmailCarrierOverride = False Then
'      pvcboEmailCarrier.Enabled = False
'      pvcboDefaultTransport.Enabled = False
'   End If
   '
   pvcboDefaultTransport.AppendItem "Transport Container"
   pvcboDefaultTransport.AppendItem "Electronic Mail"
   pvcboDefaultTransport.AppendItem "Send By Media"
   pvcboDefaultTransport.AppendItem "Demo Loopback"
   If goSession.User.DefaultTransport >= 0 And _
     goSession.User.DefaultTransport <= 3 Then
      pvcboDefaultTransport.ListIndex = goSession.User.DefaultTransport
   End If
   
'   If goSession.User.Security.AllowTransportOverride = False Then
'      pvcboDefaultTransport.Enabled = False
'   End If
   
   '
   txtMailServer.Text = goSession.User.MailServerName
   txtMailUserID.Text = goSession.User.MailUserID
   
   txtPop3Server.Text = goSession.User.GetExtendedProperty("Pop3MailServer")
   txtFromEmailAddress.Text = goSession.User.GetExtendedProperty("FromEmailAddress")
   txtSMTPEmailPortNumber.Text = goSession.User.GetExtendedProperty("EmailPortNumber")
   txtPop3EmailPortNumber.Text = goSession.User.GetExtendedProperty("Pop3EmailPortNumber")
   

   pvcboSMTPEmailSecurityProtocol.AppendItem "Implicit Auto"
   pvcboSMTPEmailSecurityProtocol.AppendItem "Implicit SSL 3.0"
   pvcboSMTPEmailSecurityProtocol.AppendItem "Implicit SSL 2.0"
   pvcboSMTPEmailSecurityProtocol.AppendItem "Implicit PCT 1.0"
   pvcboSMTPEmailSecurityProtocol.AppendItem "Implicit TLS 1.0"
   pvcboSMTPEmailSecurityProtocol.AppendItem "None"
   pvcboSMTPEmailSecurityProtocol.AppendItem "Explicit TLS 1.0"
   
   
   If goSession.User.GetExtendedProperty("EmailSecurityProtocol") <> "" Then
      nEmailSecurityProtocol = Val(goSession.User.GetExtendedProperty("EmailSecurityProtocol"))
      If nEmailSecurityProtocol >= 0 Then
         pvcboSMTPEmailSecurityProtocol.ListIndex = nEmailSecurityProtocol
      End If
   End If
   
   pvcboPop3EmailSecurityProtocol.AppendItem "Implicit Auto"
   pvcboPop3EmailSecurityProtocol.AppendItem "Implicit SSL 3.0"
   pvcboPop3EmailSecurityProtocol.AppendItem "Implicit SSL 2.0"
   pvcboPop3EmailSecurityProtocol.AppendItem "Implicit PCT 1.0"
   pvcboPop3EmailSecurityProtocol.AppendItem "Implicit TLS 1.0"
   pvcboPop3EmailSecurityProtocol.AppendItem "None"
   pvcboPop3EmailSecurityProtocol.AppendItem "Explicit TLS 1.0"
   
   If goSession.User.GetExtendedProperty("Pop3EmailSecurityProtocol") <> "" Then
      nEmailSecurityProtocol = Val(goSession.User.GetExtendedProperty("Pop3EmailSecurityProtocol"))
      If nEmailSecurityProtocol >= 0 Then
         pvcboPop3EmailSecurityProtocol.ListIndex = nEmailSecurityProtocol
      End If
   End If
   
   
   If Trim(goSession.User.MailPassword) = "-1" Then
      'Registry error if "blank" password
      pvmMailPassword.Text = ""
   Else
      pvmMailPassword.Text = goSession.User.MailPassword
   End If
   
   '
   'Select Case goSession.User.Security.UserConfigTransportAccess
   '   Case Is = MW_CONFIG_NO_ACCESS
   '      SSTab1.TabEnabled(2) = False
   '   Case Is = MW_CONFIG_READ_ONLY
   '      pvcboDefaultTransport.Enabled = False
   '      pvcboEmailCarrier.Enabled = False
   '      txtMailServer.Enabled = False
   '      txtMailUserID.Enabled = False
   '      pvmMailPassword.Enabled = False
   '      chkUserCanTransport.Enabled = False
   '      chkZipOnDuringSubmit.Enabled = False
   '   Case Is = MW_CONFIG_READ_WRITE
   '      'Nothing to do...
   'End Select
   
   goSession.SetDotNetTheme Me
   
   Exit Sub
Form_load_error:
   goSession.RaiseError "General error in frmEmailConfiguration.form_load.", Err.Number, Err.Description
   Me.Hide
End Sub


Private Sub pvcboDefaultTransport_Change()
   
   
   Select Case pvcboDefaultTransport.ListIndex
      Case Is = 0
         
         ' Transport Container
         lblEmailCarrier.Visible = False
         pvcboEmailCarrier.Visible = False
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
         
         lblSMTPEmailPortNumber.Visible = False
         lblPop3EmailPortNumber.Visible = False
         txtSMTPEmailPortNumber.Visible = False
         txtPop3EmailPortNumber.Visible = False
         lblSMTPSecurityProtocol.Visible = False
         lblPop3SecurityProtocol.Visible = False
         pvcboSMTPEmailSecurityProtocol.Visible = False
         pvcboPop3EmailSecurityProtocol.Visible = False
                   
      Case Is = 1
         ' Electronic Mail
         lblEmailCarrier.Visible = True
         pvcboEmailCarrier.Visible = True
         lblMailServer.Visible = True
         txtMailServer.Visible = True
         lblMailUserID.Visible = True
         txtMailUserID.Visible = True
         lblMailPassword.Visible = True
         pvmMailPassword.Visible = True
         
         
         ' Reset Email Controls after user changes Current transport...
         SetEmailCarrierControls
         
         
      Case Is = 2, 3                            ' 'Send by Media' or 'Demo Loopback' to hide carrier options
         lblEmailCarrier.Visible = False
         pvcboEmailCarrier.Visible = False
      
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
                
         lblSMTPEmailPortNumber.Visible = False
         lblPop3EmailPortNumber.Visible = False
         txtSMTPEmailPortNumber.Visible = False
         txtPop3EmailPortNumber.Visible = False
         lblSMTPSecurityProtocol.Visible = False
         lblPop3SecurityProtocol.Visible = False
         pvcboSMTPEmailSecurityProtocol.Visible = False
         pvcboPop3EmailSecurityProtocol.Visible = False
                
   End Select
    
End Sub


Private Sub pvcboEmailCarrier_Change()
   SetEmailCarrierControls
End Sub

Private Function SetEmailCarrierControls()

   ' clear old values
         txtMailServer.Text = ""
         txtMailUserID.Text = ""
         pvmMailPassword.Text = ""
         txtPop3Server.Text = ""
         txtFromEmailAddress.Text = ""
   
   lblSMTPEmailPortNumber.Visible = False
   lblPop3EmailPortNumber.Visible = False
   txtSMTPEmailPortNumber.Visible = False
   txtPop3EmailPortNumber.Visible = False
   lblSMTPSecurityProtocol.Visible = False
   lblPop3SecurityProtocol.Visible = False
   pvcboSMTPEmailSecurityProtocol.Visible = False
   pvcboPop3EmailSecurityProtocol.Visible = False
   
   txtSMTPEmailPortNumber.Text = ""
   pvcboSMTPEmailSecurityProtocol.ListIndex = -1
   txtPop3EmailPortNumber.Text = ""
   pvcboPop3EmailSecurityProtocol.ListIndex = -1
   
   Select Case pvcboEmailCarrier.ListIndex
      Case Is = 0
         ' SMTP Mail
         lblMailServer.Visible = True
         txtMailServer.Visible = True
         lblMailUserID.Visible = True
         txtMailUserID.Visible = True
         lblMailPassword.Visible = True
         pvmMailPassword.Visible = True
         lblMailServer.Caption = "SMTP Server:"
         lblMailUserID.Caption = "Mailbox Address: "
         lblMailPassword.Caption = "Mailbox Password:"
         lblMailPop3Server.Visible = True
         lblMailPop3Server.Caption = "Pop3 Server:"
         txtPop3Server.Visible = True
         lblFromEmailAddress.Visible = True
         lblFromEmailAddress.Caption = "From Email Address"
         
         txtFromEmailAddress.Visible = True
         lblSMTPEmailPortNumber.Visible = True
         txtSMTPEmailPortNumber.Visible = True
         lblSMTPSecurityProtocol.Visible = True
         pvcboSMTPEmailSecurityProtocol.Visible = True
         
         lblPop3EmailPortNumber.Visible = True
         txtPop3EmailPortNumber.Visible = True
         lblPop3SecurityProtocol.Visible = True
         pvcboPop3EmailSecurityProtocol.Visible = True
      Case Is = 1, 6, 9 'Added By N.Angelakis On 02 Feb 2010, 'DEV-1797 Sending Email with Outlook client 97
         ' Outlook/Outlook 2002
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         'lblMailUserID.Caption = "Profile Name: "
         'lblMailPassword.Caption = "Password: "
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
      
      Case Is = 2, 8
         ' Exchange
         lblMailServer.Visible = True
         txtMailServer.Visible = True
         lblMailUserID.Visible = True
         txtMailUserID.Visible = True
         lblMailPassword.Visible = True
         pvmMailPassword.Visible = True
         lblMailServer.Caption = "Exchange/MAPI Server:"
         lblMailUserID.Caption = "Mailbox Name: "
         lblMailPassword.Caption = "Mailbox Password"
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
      
      Case Is = 3
         ' Notes
         lblMailServer.Visible = True
         txtMailServer.Visible = True
         lblMailUserID.Visible = True
         txtMailUserID.Visible = True
         lblMailPassword.Visible = True
         pvmMailPassword.Visible = True
         lblMailServer.Caption = "Notes Server Name:"
         lblMailUserID.Caption = "Notes File Name: "
         lblMailPassword.Caption = "Notes Password:"
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
      
      Case Is = 4
         ' AMOS_MAIL
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         'lblMailUserID.Caption = "Send From Address: "
         'lblMailPassword.Caption = "Password: "
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
         
      Case Is = 5
         ' AMOS_LINK
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         'lblMailUserID.Caption = "Send From Address: "
         'lblMailPassword.Caption = "Password: "
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
         
         
      Case Is = 7
         ' Groupwise
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
         
         
      'DEV-1809 Send By Media Error Notification
      'Added By N.Angelakis On 04 Feb 2010
      Case Is = 10
         lblMailServer.Visible = False
         txtMailServer.Visible = False
         lblMailUserID.Visible = False
         txtMailUserID.Visible = False
         lblMailPassword.Visible = False
         pvmMailPassword.Visible = False
         'lblMailUserID.Caption = "Send From Address: "
         'lblMailPassword.Caption = "Password: "
         lblMailPop3Server.Visible = False
         txtPop3Server.Visible = False
         lblFromEmailAddress.Visible = False
         txtFromEmailAddress.Visible = False
   
   End Select
   
End Function



Private Sub cmdSave_Click()
   Dim i As Integer
   Dim oTempSession As Session
   
   '
   ' Transport
   '
   goSession.User.DefaultTransport = pvcboDefaultTransport.ListIndex
   '
   goSession.User.DefaultEmailCarrier = pvcboEmailCarrier.ListIndex
   '
   If txtMailServer.DataChanged Then
      goSession.User.MailServerName = txtMailServer.Text
   End If
   If txtMailUserID.DataChanged Then
      goSession.User.MailUserID = txtMailUserID.Text
   End If
   If Trim(pvmMailPassword.Text) = "" Then
         'Registry error if "blank" password
         goSession.User.MailPassword = "-1"
      Else
         goSession.User.MailPassword = pvmMailPassword.Text
   End If
   
   If txtPop3Server.DataChanged Then
      goSession.User.SetExtendedProperty "Pop3MailServer", txtPop3Server.Text, goSession.User.UserID
   End If
   If txtFromEmailAddress.DataChanged Then
      goSession.User.SetExtendedProperty "FromEmailAddress", txtFromEmailAddress.Text, goSession.User.UserID
   End If
   If txtSMTPEmailPortNumber.DataChanged Then
      goSession.User.SetExtendedProperty "EmailPortNumber", txtSMTPEmailPortNumber.Text, goSession.User.UserID
   End If
   If txtPop3EmailPortNumber.DataChanged Then
      goSession.User.SetExtendedProperty "Pop3EmailPortNumber", txtPop3EmailPortNumber.Text, goSession.User.UserID
   End If
   
   If pvcboSMTPEmailSecurityProtocol.ListIndex > -1 Then
      goSession.User.SetExtendedProperty "EmailSecurityProtocol", pvcboSMTPEmailSecurityProtocol.ListIndex, goSession.User.UserID
   Else
      goSession.User.SetExtendedProperty "EmailSecurityProtocol", "", goSession.User.UserID
   End If
   
   If pvcboPop3EmailSecurityProtocol.ListIndex > -1 Then
      goSession.User.SetExtendedProperty "Pop3EmailSecurityProtocol", pvcboPop3EmailSecurityProtocol.ListIndex, goSession.User.UserID
   Else
      goSession.User.SetExtendedProperty "Pop3EmailSecurityProtocol", "", goSession.User.UserID
   End If
   
   
   ' All done, let's get out of here...
   MsgBox "Configuration Settings have been updated.", vbInformation, "Email Configuration Change Complete"
   
   Me.Hide
End Sub

'
' Sample Novoship...
'
'TABNO*SURNAME*NAME*PATRONYMIC*DOB*POB*LASTRANK*SIGNONDAT*SIGNONPORT*LRW*ELT*ELTD*SEAMBNO*SEAMBID*SEAMBED*TOURPNO*TOURPID*TOURPED*CIVPASSNO*CIVPASSID*RUSLICNO*RUSLICTYPE*RUSLICID*ENDORNO*ENDORTYPE*ENDORID*ENDORED*MEDPASSNO*MEDPASSID*MEDPASSED*DRALDECID*DRALDECED*PERRLM105ID*PERRLM105ED*YELFEVID*YELFEVED*ARPANO*ARPAID*ARPAED*RADOBNO*RADOBID*RADOBED*RADIOTNO*RADIOTID*RADIOTED*BTTNO*BTTID*BTTED*GMDSSNO*GMDSSID*GMDSSED*COWNO*COWID*COWED*IGSNO*IGSID*IGSED*ABSTCNO*ABSTCID*ABSTCED*ATPOOTONO*ATPOOTOID*ATPOOTOED*ATPOCTONO*ATPOCTOID*ATPOCTOED*ACATFNO*ACATFTYPE*ACATFID*ACATFED*TCCTVRS*TCCTVRSID*TCCTVRSED*ATIFFNO*ATIFFID*ATIFFED*ATILNO*ATILID*ATILED*ATIFRBNO*ATIFRBID*ATIFRBED*ATIFFLBNO*ATIFFLBID*ATIFFLBED*MEDTRNO*MEDTRID*MEDTRED*ISMNO*ISMID*LIBSID*LIBSIDID*LIBSIDED*LIBLICNO*LIBLICRANK*LIBLICID*LIBLICED*LIBGMDSSNO*LIBGMDSSID*LIBGMDSSED*LIBCTGCNO*LIBCTGCID*LIBCTGCED*LIBCOWNO*LIBCOWID*LIBCOWED*LIBCFCTNO*LIBCFCTID*LIBCFCTED*NOKNAME*NOKSURNAME*NOKPATRONIMIC*NOKKINSHIP*NOKADDR*NOKTLF*NOKEMAIL
'481092*KISEL*YURIY*NIKOLAEVICH*27/09/1963*KRASNODAR  REG*CAPTAIN*18/07/2003*TAFT, LOUISIANA*1*80*15/04/1998*MF0046564*27/12/1995*22/05/2005*46N6140499*01/03/1999*01/03/2004***2010300245*DSC*31/05/2001*2010220480*OILCHEM*18/07/2002*25/04/2007*2388*01/06/2000*28/06/2003*01/09/2002*01/09/2003*28/06/2002*28/06/2003*25/05/2001*25/05/2011*2*05/02/1999*05/02/2004*2*29/01/1999*29/01/2004*242/1994*23/11/1994*23/11/1999*000474*15/07/2002*15/07/2007*16051*30/05/2001*30/05/2006*005932*21/05/2002*21/05/2007*005932*21/05/2002*21/05/2007*60889*10/04/2001*19/02/2004*005932*21/05/2002*21/05/2007*003759*24/04/2002*25/04/2007*3118*CAPTAIN*30/01/1998*30/01/2003*001319*21/05/2002*21/05/2007*27956*10/04/2001*26/02/2004*11387*18/05/2001*18/05/2006*11387*18/05/2001*18/05/2006****4728*11/04/2001*12/02/2004*000554*01/04/2002*504124*25/05/1999*25/05/2004*703719*CAPTAIN*26/07/2001*26/07/2006*703719*26/07/2001*26/07/2006*395238*25/05/1999*25/05/2004*395238*25/05/1999*25/05/2004*579684*26/07/2001*26/07/2006*MARINA*KISEL*NIKOL**** _
'492171*MATYAZH*NIKOLAY*VIKTOROVITCH*10/05/1949*...
'
' Novoship Only
'
Private Sub cmdImportCrewList_Click()
   Dim loRs As Recordset
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim i As Long
   Dim iCount As Long
   Dim s As String
   Dim sa() As String
   On Error GoTo FunctionError
   If goSession.ThisSite.CompanyID <> "NOVOSHIP" Then
      cmdImportCrewList.Visible = False
      Exit Sub
   End If
   cd1.CancelError = True
   cd1.ShowOpen
   ' Change manual, delete invalid link...
   If Trim(cd1.FileName) = "" Then
      Exit Sub
   End If
   ' Try and open file...
   Set fso = New FileSystemObject
   If Not fso.FileExists(cd1.FileName) Then
      goSession.RaisePublicError "Novoship Crew data file not found: " & cd1.FileName
      Set fso = Nothing
      Exit Sub
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "scPersonnel", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   goCon.Execute "delete from scPersonnel"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "scPersonnel", goCon, adOpenDynamic, adLockOptimistic, adCmdTable
   Set ts = fso.OpenTextFile(cd1.FileName, ForReading, False)
   ' skip first line
   s = ts.ReadLine
   ' meat and potatoes
   i = goSession.MakePK("scPersonnel")
   iCount = 0
   Do While Not ts.AtEndOfStream
      s = ts.ReadLine
      sa = Split(s, "*")
      If UBound(sa) > 3 Then
         With loRs
            .AddNew
            .Fields("ID") = i
            .Fields("EmployeeID") = sa(0)
            .Fields("FullName") = sa(1) & ", " & sa(2) & " " & sa(3)
            .Update
            iCount = iCount + 1
            i = i + 1
         End With
      End If
   Loop
   If i > 0 Then
      goSession.UpdatePrimaryKeySequence "scPersonnel", i
   End If
   CloseRecordset loRs
   ts.Close
   Set ts = Nothing
   Set fso = Nothing
   MsgBox str(iCount) & " crew records have been imported", vbInformation, "Import Crewlist"
   Exit Sub
FunctionError:
   If Err.Number <> 32755 Then
      goSession.RaisePublicError "General Error in frmKnowledgeGuides.ugList.BeforeCellActivate: ", Err.Number, Err.Description
   End If
   KillObject fso
End Sub


Private Sub txtSMTPEmailPortNumber_KeyPress(KeyAscii As Integer)
   If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtPop3EmailPortNumber_KeyPress(KeyAscii As Integer)
   If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub


