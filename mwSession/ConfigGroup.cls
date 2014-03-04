VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ConfigKeys"
Attribute VB_Ext_KEY = "Member1" ,"ConfigGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mConfigGroup As String
Private mvarConfigKeys As ConfigKeys


Public Property Get ConfigKeys() As ConfigKeys
   On Error GoTo PropertyError
    If mvarConfigKeys Is Nothing Then
        Set mvarConfigKeys = New ConfigKeys
    End If


    Set ConfigKeys = mvarConfigKeys
    Exit Property
PropertyError:
   MsgBox "error in Config.ConfigKeys."
    
End Property


Public Property Set ConfigKeys(vData As ConfigKeys)
    Set mvarConfigKeys = vData
End Property

Private Sub Class_Initialize()
    Set mvarConfigKeys = New ConfigKeys

End Sub

Private Sub Class_Terminate()
    Set mvarConfigKeys = Nothing
End Sub

Public Property Let ConfigGroup(ByVal Data As String)
    mConfigGroup = UCase(Data)
End Property

Public Property Get ConfigGroup() As String
    ConfigGroup = UCase(mConfigGroup)
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


