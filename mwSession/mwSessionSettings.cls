VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwSessionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Registry base for session settings
Private Const BASE_REGISTRY As String = "Software\Maritime Systems Inc\SessionSettings"


Dim moReg As New Registry


' Form Window Settings
Private mWindowState As String
Private mWidth As String
Private mTop As String
Private mScaleHeight As String
Private mScaleWidth As String
Private mHeight As String
Private mLeft As String
Private mIsStartingValues As Boolean
Private mShowOutlookBar As String


'Menu Settings
Private mActiveMenu As String


Public Property Let WindowState(ByVal Data As String)
   mWindowState = Data
   moReg.SetReg "WindowState", Data
End Property

Public Property Get WindowState() As String
    WindowState = mWindowState
End Property

Public Property Let Width(ByVal Data As String)
   mWidth = Data
   moReg.SetReg "Width", Data
End Property

Public Property Get Width() As String
    Width = mWidth
End Property

Public Property Get Top() As String
    Top = mTop
End Property

Public Property Let Top(ByVal Data As String)
   mTop = Data
   moReg.SetReg "Top", Data
End Property

Public Property Get ScaleHeight() As String
    ScaleHeight = mScaleHeight
End Property

Public Property Let ScaleHeight(ByVal Data As String)
   mScaleHeight = Data
   moReg.SetReg "ScaleHeight", Data
End Property

Public Property Get ScaleWidth() As String
    ScaleWidth = mScaleWidth
End Property

Public Property Let ScaleWidth(ByVal Data As String)
   mScaleWidth = Data
   moReg.SetReg "ScaleWidth", Data
End Property

Public Property Get Height() As String
    Height = mHeight
End Property

Public Property Let Height(ByVal Data As String)
   mHeight = Data
   moReg.SetReg "Height", Data
End Property

Public Property Get ActiveMenu() As String
    ActiveMenu = mActiveMenu
End Property

Public Property Let ActiveMenu(ByVal Data As String)
   mActiveMenu = Data
   moReg.SetReg "ActiveMenu", Data
End Property

Public Property Get IsStartingValues() As Boolean
    IsStartingValues = mIsStartingValues
End Property

Public Property Get Left() As String
    Left = mLeft
End Property

Public Property Let Left(ByVal Data As String)
   mLeft = Data
   moReg.SetReg "Left", Data
End Property

Public Property Get ShowOutlookBar() As String
    ShowOutlookBar = mShowOutlookBar
End Property

Public Property Let ShowOutlookBar(ByVal Data As String)
   mShowOutlookBar = Data
   moReg.SetReg "ShowOutlookBar", Data
End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'   mA = Data
'   moReg.SetReg "A", Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'   mA = Data
'   moReg.SetReg "A", Data
'End Property


Private Sub Class_Initialize()
   On Error GoTo SubError
   Set moReg = New Registry
   moReg.BaseRegistry = BASE_REGISTRY
   mWindowState = moReg.GetReg("WindowState")
   mWidth = moReg.GetReg("Width")
   mTop = moReg.GetReg("Top")
   mScaleHeight = moReg.GetReg("ScaleHeight")
   mScaleWidth = moReg.GetReg("ScaleWidth")
   mHeight = moReg.GetReg("Height")
   mLeft = moReg.GetReg("Left")
   mActiveMenu = moReg.GetReg("ActiveMenu")
   mShowOutlookBar = moReg.GetReg("ShowOutlookBar")
   If moReg.Errors.Count = 0 Then
      mIsStartingValues = True
   Else
      mIsStartingValues = False
   End If
   Do While moReg.Errors.Count <> 0
      moReg.Errors.Remove 1
   Loop
   Exit Sub
SubError:
   Exit Sub
End Sub

Private Sub Class_Terminate()
   Set moReg = Nothing
End Sub

