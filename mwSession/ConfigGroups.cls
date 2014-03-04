VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ConfigGroup"
Attribute VB_Ext_KEY = "Member0" ,"ConfigGroup"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(ConfigGroup As String, Optional sKey As String) As ConfigGroup
    'create a new object
    On Error GoTo FunctionError
    Dim objNewMember As ConfigGroup
    Set objNewMember = New ConfigGroup


    'set the properties passed into the method
    objNewMember.ConfigGroup = ConfigGroup
    'Set objNewMember.ConfigKeys = ConfigKeys
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
FunctionError:
   If err.Number = 457 Then
      MsgBox "Unlogged Error in mwSession.ConfigGroups.Add, INI file contains duplicate section: " & ConfigGroup
   End If
   Debug.Print "General Error in ConfigGroups.Add. " & err.Number & " - " & err.Description
End Function

Public Property Get Item(vntIndexKey As Variant) As ConfigGroup
Attribute Item.VB_UserMemId = 0
   On Error GoTo PropertyError
  Set Item = mCol(vntIndexKey)
  Exit Property
PropertyError:
   'MsgBox "can't find configgroup: " & vntIndexKey
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
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


