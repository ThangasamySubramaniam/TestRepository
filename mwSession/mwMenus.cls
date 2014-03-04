VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"WFD_Menu"
Attribute VB_Ext_KEY = "Member0" ,"WFD_Menu"
' mwMenus Collection class, used by mwUSER to store a collection of mwMENUS
' 11/24/2000 ms  Copyright 2000 Mariitme Systems Inc - All Rights Reserved
'
Option Explicit

'local variable to hold collection
Private mCol As Collection
Private mbIsAdded As Boolean
Private mw_rErrorMsg As String

Public Function Add(DisplayTitle As String, MenuID As String, ImageList As String, _
  Optional sKey As String) As mwMenu
   Dim MenuItems As Collection
   Dim objNewMember As mwMenu
   On Error GoTo add_error
   mw_rErrorMsg = ""
   Set objNewMember = New mwMenu
   objNewMember.DisplayTitle = Trim(DisplayTitle)
   objNewMember.ImageList = ImageList
   objNewMember.MenuID = MenuID
   If Len(sKey) = 0 Then
      mCol.Add objNewMember
   Else
      mCol.Add objNewMember, sKey
   End If
   Set Add = objNewMember
   Set objNewMember = Nothing
   mbIsAdded = True
   Exit Function
add_error:
   If err.Number = 457 Then
      goSession.RaiseError "Error in mwMenus.Add, " & DisplayTitle & " is a duplicate menu entry."
   Else
      goSession.RaiseError "General Error in mwMenus.Add, duplicate entry in mwcUserMenus.", err.Number, err.Description
   End If
   If Not objNewMember Is Nothing Then Set objNewMember = Nothing
   mbIsAdded = False
End Function

Public Property Get Item(vntIndexKey As Variant) As mwMenu
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get isAdded() As Boolean
   isAdded = mbIsAdded
End Property

Public Property Get ErrorMsg() As String
   ErrorMsg = mw_rErrorMsg
End Property
