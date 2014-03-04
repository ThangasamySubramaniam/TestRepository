VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' 1/10/02 ms mwEventForm - Event Form triggered by Event Queue Table
'
Option Explicit

Private moParent As mwForm
Private moSes As mwSession.Session
Dim moRS As Recordset
Private moRecords As ConfigGroups
Private mIsWritten As Boolean
Private moCon As Connection

Public Property Set Parent(oValue As mwForm)
   'If moParent Is Nothing Then
      Set moParent = oValue
   'End If
   
End Property

Public Property Get Parent() As mwForm
   Set Parent = moParent
End Property

Public Property Set Session(oSession As mwSession.Session)
   'If moSes Is Nothing Then
      Set moSes = oSession
   'End If
End Property

Public Property Get Session() As mwSession.Session
   Set Session = moSes
End Property


Public Property Set Records(ByVal Data As ConfigGroups)
    Set moRecords = Data
End Property

Public Property Get Records() As ConfigGroups
    Set Records = moRecords
End Property

Public Property Get IsWritten() As Boolean
    IsWritten = mIsWritten
End Property


