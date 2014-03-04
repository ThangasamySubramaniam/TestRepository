VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Begin VB.Form frmEmailBunkers 
   Caption         =   "Select Email N Grade Bunkers"
   ClientHeight    =   6048
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7788
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
   ScaleHeight     =   6048
   ScaleWidth      =   7788
   StartUpPosition =   1  'CenterOwner
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   7788
      DesignHeight    =   6048
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   975
      Left            =   1320
      Picture         =   "frmEmailBunkers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   975
      Left            =   5040
      Picture         =   "frmEmailBunkers.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin UltraGrid.SSUltraGrid ug 
      Height          =   4812
      Left            =   48
      TabIndex        =   0
      Top             =   120
      Width           =   7692
      _ExtentX        =   13568
      _ExtentY        =   8488
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   1
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "frmEmailBunkers.frx":0FD4
      Caption         =   "Email N Grade Bunkers"
   End
End
Attribute VB_Name = "frmEmailBunkers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moRS As Recordset
Public mIsCancelled As Boolean

' constants from mwSession.mwDataForm

 Const UG_BP_ID = 0
 Const UG_DetailType = 1
 Const UG_DetailKey = 2
 Const UG_vrsBPKey = 3
 Const UG_vrsBG_Description = 4     ' x
 Const UG_vrsBP_Description = 5     ' x
 Const UG_vrsBGKey = 6
 Const UG_UnitType = 7              ' x
 Const UG_Quantity = 8              ' x
 Const UG_IsShowInEmailTextBody = 9 ' x
 Const UG_IsShownOnPOSASS = 10
 Const UG_IsShownOnArrival = 11
 Const UG_IsShownOnDeparture = 12


Public Function InitForm(Rs As Recordset) As Boolean
   On Error GoTo FunctionError
   
   If IsNull(Rs) Then
      Exit Function
   End If
   
   CloseRecordset moRS
   Set moRS = Rs
   
   ' refreshUgColumns
   Set ug.DataSource = moRS
   HideUltragridColumns ug, 0
   ug.Override.HeaderClickAction = ssHeaderClickActionSortSingle
   
   
   ' bunker desc
   
   ug.Bands(0).Columns(UG_vrsBG_Description).Hidden = False
   ug.Bands(0).Columns(UG_vrsBG_Description).Width = 1100
   ug.Bands(0).Columns(UG_vrsBG_Description).Header.Caption = "Grade"
   ug.Bands(0).Columns(UG_vrsBG_Description).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   
   ug.Bands(0).Columns(UG_vrsBP_Description).Hidden = False
   ug.Bands(0).Columns(UG_vrsBP_Description).Width = 1600
   ug.Bands(0).Columns(UG_vrsBP_Description).Header.Caption = "Bunker"
   ug.Bands(0).Columns(UG_vrsBP_Description).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   
   ug.Bands(0).Columns(UG_UnitType).Hidden = False
   ug.Bands(0).Columns(UG_UnitType).Width = 900
   ug.Bands(0).Columns(UG_UnitType).Header.Caption = "Unit"
   ug.Bands(0).Columns(UG_UnitType).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   
   ug.Bands(0).Columns(UG_Quantity).Hidden = False
   ug.Bands(0).Columns(UG_Quantity).Width = 1000
   ug.Bands(0).Columns(UG_Quantity).Header.Caption = "Quantity"
   ug.Bands(0).Columns(UG_Quantity).CellAppearance.BackColor = goSession.GUI.UG_NoEdit_BackColor
   '
   ug.Bands(0).Columns(UG_IsShowInEmailTextBody).Hidden = False
   ug.Bands(0).Columns(UG_IsShowInEmailTextBody).Width = 1600
   ug.Bands(0).Columns(UG_IsShowInEmailTextBody).Header.Caption = "Show in Email"
   ug.Bands(0).Columns(UG_IsShowInEmailTextBody).Style = ssStyleCheckBox
   ug.Bands(0).Columns(UG_IsShowInEmailTextBody).CellAppearance.BackColor = goSession.GUI.UG_EditBackColor
   InitForm = True

   goSession.SetDotNetTheme Me
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.InitForm: ", err.Number, err.Description
End Function

Private Sub cmdCancel_Click()
   mIsCancelled = True
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   Me.Hide
   mIsCancelled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo SubError
   'CloseRecordset moRS
   Set ug.DataSource = Nothing
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.frmEmailTemplate: ", err.Number, err.Description
   CloseRecordset moRS
End Sub

Private Sub ug_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
   On Error GoTo SubError
   
   If Cell Is Nothing Then Exit Sub
   If Cell.Column.Index <> UG_IsShowInEmailTextBody Then Cancel = True
   
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwSession.frmEmailTemplate.ug_BeforeCellActivate ", err.Number, err.Description
End Sub

