VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"WFD_MenuItem"
Attribute VB_Ext_KEY = "Member0" ,"WFD_MenuItem"
' mwMenuItems Collection class, used by mwMENU to store a collection of mwMENUITEMS
' 11/24/2000 ms  Copyright 2000 Mariitme Systems Inc - All Rights Reserved
'


Option Explicit

Private mCol As Collection
Private mbIsAdded As Boolean
Private mw_rErrorMsg As String
Private mbIsItem As Boolean


Public Function Add(FormID As String, Optional sKey As String) As mwMenuItem
'Public Function Add(FormID As String, DisplayTitle As String, ImageReference As String, _
'   MenuType As String, Optional sKey As String) As mwMenuItem
    'create a new object
   Dim objNewMember As mwMenuItem
   On Error GoTo add_error

    Set objNewMember = New mwMenuItem


    'set the properties passed into the method
    objNewMember.FormID = FormID
    'objNewMember.DisplayTitle = DisplayTitle
    'objNewMember.ImageReference = ImageReference
    'objNewMember.FormType = MenuType
    'If IsObject(FormHandle) Then
    '    Set objNewMember.FormHandle = FormHandle
    'Else
    '    objNewMember.FormHandle = FormHandle
    'End If
    'objNewMember.IsExists = IsExists
    'objNewMember.FormID = FormID
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    mbIsAdded = True
   Exit Function
add_error:
   mw_rErrorMsg = err.Number & " - " & err.Description
   If Not objNewMember Is Nothing Then Set objNewMember = Nothing
   mbIsAdded = False


End Function

Public Property Get Item(vntIndexKey As Variant) As mwMenuItem
Attribute Item.VB_UserMemId = 0
   On Error GoTo item_error:
   Set Item = mCol(vntIndexKey)
   mbIsItem = True
   Exit Property
item_error:
   mbIsItem = False
   mw_rErrorMsg = "Error on Get Menu Item: " & vntIndexKey
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


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
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Property Get isAdded() As Boolean
   isAdded = mbIsAdded
End Property

Public Property Get ErrorMsg() As String
   ErrorMsg = mw_rErrorMsg
End Property

Public Property Get isItem() As Boolean
   isItem = mbIsItem
End Property


