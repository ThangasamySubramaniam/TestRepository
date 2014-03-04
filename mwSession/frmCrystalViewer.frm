VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form frmCrystalViewer 
   Caption         =   "MWS Crystal  Viewer"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "frmCrystalViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer cr1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _cx             =   19923
      _cy             =   15266
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "frmCrystalViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowReport(WindowTitle As String, oReport As CRAXDRT.Report, _
 Optional oRs As Recordset) As Boolean
   On Error GoTo FunctionError
   Me.Caption = WindowTitle
   goSession.SetDotNetTheme Me
   cr1.ReportSource = oReport
   cr1.DisplayToolbar = True
   cr1.EnableNavigationControls = True
   cr1.EnableRefreshButton = True
   cr1.EnableSelectExpertButton = True
   cr1.EnableSearchControl = True
   cr1.EnableStopButton = True
   cr1.EnableToolbar = True
   cr1.EnableZoomControl = True
   cr1.Refresh
   cr1.ViewReport
   '
   ' test code
   '
   If Not oRs Is Nothing Then
      oReport.Database.SetDataSource oRs
   End If
   ShowReport = True
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.frmCrystalViewer. ", err.Number, err.Description
   ShowReport = False
End Function

Private Sub Form_Resize()
   On Error Resume Next
   cr1.Height = Me.Height - 600
   cr1.Width = Me.Width - 100
End Sub

Public Function PrintReport(moRep As CRAXDRT.Report) As Boolean
   On Error GoTo FunctionError
   cr1.ReportSource = moRep
   cr1.PrintReport
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwSession.frmCrystalViewer. ", err.Number, err.Description
   
End Function
