VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwDotNetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A menu to be displayed to the desktop user"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mFormHandle As Object
Private mVB6FormKey As Long


Public Property Let FormHandle(ByVal vData As Object)
Attribute FormHandle.VB_Description = "General Purpose identifier, links to other configuration objects"
    Set mFormHandle = vData
End Property


Public Property Get FormHandle() As Object
    Set FormHandle = mFormHandle
End Property

Public Property Let VB6FormKey(ByVal vData As Long)
    mVB6FormKey = vData
End Property


Public Property Get VB6FormKey() As Long
    VB6FormKey = mVB6FormKey
End Property







