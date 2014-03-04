VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwLogItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"LogItem"
Attribute VB_Ext_KEY = "Member0" ,"LogItem"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(MessageType As mwl_LogMessageType, Criticality As mwl_LogMessageCriticality, _
  Msg As String, Optional oError As mwError) As mwLogItem
    'create a new object
    On Error Resume Next
    Dim objNewMember As mwLogItem
    Set objNewMember = New mwLogItem


    'set the properties passed into the method
    'objNewMember.DateTime = Now()
    objNewMember.Message = Msg
    objNewMember.Criticality = Criticality
    objNewMember.MessageType = MessageType
    If Not IsMissing(oError) Then
      Set objNewMember.Error = oError
    End If
    mCol.Add objNewMember


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As mwLogItem
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

