VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ConfigKey"
Attribute VB_Ext_KEY = "Member0" ,"ConfigKey"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(KeyName As String, KeyValue As String, Optional sKey As String) As ConfigKey
    'create a new object
    On Error GoTo FunctionError
    Dim objNewMember As ConfigKey
    Set objNewMember = New ConfigKey


    'set the properties passed into the method
    objNewMember.KeyValue = KeyValue
    objNewMember.KeyName = KeyName
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
   'MsgBox "error."


End Function

Public Property Get Item(vntIndexKey As Variant) As ConfigKey
Attribute Item.VB_UserMemId = 0
   On Error GoTo PropertyError
  Set Item = mCol(vntIndexKey)
  Exit Property
PropertyError:
   
   'MsgBox "here"
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
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function GetKeyValue(Key As String) As String
   On Error GoTo FunctionError
   GetKeyValue = mCol.Item(UCase(Key)).KeyValue
   Exit Function
FunctionError:
   GetKeyValue = ""
End Function

Public Function GetBoolKeyValue(Key As String) As Boolean
   Dim strKeyValue As String
   
   On Error GoTo FunctionError
   
   strKeyValue = mCol.Item(UCase(Key)).KeyValue
   
   If strKeyValue = "" Then
      GetBoolKeyValue = False
   ElseIf strKeyValue = "0" Then
      GetBoolKeyValue = False
   ElseIf strKeyValue = "1" Then
      GetBoolKeyValue = True
   ElseIf strKeyValue = "-1" Then
      GetBoolKeyValue = True
   ElseIf strKeyValue = "True" Then
      GetBoolKeyValue = True
   Else
      GetBoolKeyValue = False
   End If
   
   Exit Function
FunctionError:
   GetBoolKeyValue = False
End Function

