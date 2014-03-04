VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A menu to be displayed to the desktop user"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MenuItems"
' Menu Class representing a menu (Group in OutlookBar terms) in WFDesktop
' 11/23/2000 ms  Copyright 2000 Maritime Systems Inc - All Rights Reserved
'


Option Explicit

Private mDisplayTitle As String 'local copy
Private mMenuID As String 'local copy
Private mImageList As String
Private mvarMenuItems As mwMenuItems


Public Property Get MenuItems() As mwMenuItems
    If mvarMenuItems Is Nothing Then
        Set mvarMenuItems = New mwMenuItems
    End If
    Set MenuItems = mvarMenuItems
End Property


Public Property Set MenuItems(vData As mwMenuItems)
    Set mvarMenuItems = vData
End Property

Private Sub Class_Terminate()
    Set mvarMenuItems = Nothing
End Sub



Public Property Let MenuID(ByVal vData As String)
Attribute MenuID.VB_Description = "General Purpose identifier, links to other configuration objects"
    mMenuID = vData
End Property


Public Property Get MenuID() As String
    MenuID = mMenuID
End Property

Public Property Let DisplayTitle(ByVal vData As String)
    mDisplayTitle = vData
End Property


Public Property Get DisplayTitle() As String
    DisplayTitle = mDisplayTitle
End Property

Public Property Let ImageList(ByVal vData As String)
    mImageList = vData
End Property

Public Property Get ImageList() As String
   ImageList = mImageList
End Property



