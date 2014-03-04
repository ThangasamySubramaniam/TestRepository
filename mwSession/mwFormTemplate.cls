VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwFormTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwFormTemplate - Class containing form template information for
'                    XFR and WFR document container forms...
' 11/24/2000 ms
'
Option Explicit

Private mFormID As String
Private mFilePattern As String
Private mActiveConfigGroup As String
Private mPendingFolder As String
Private mSubmittedFolder As String
Private mArchivedFolder As String
Private mTemplateFolder As String
Private mSendToAddress As String
Private mDisplayTitle As String



Public Property Let FormID(ByVal Data As String)
    mFormID = Data
End Property


Public Property Get FormID() As String
    FormID = mFormID
End Property

Public Property Let FilePattern(ByVal Data As String)
    mFilePattern = Data
End Property


Public Property Get FilePattern() As String
    FilePattern = mFilePattern
End Property

Public Property Get PendingFolder() As String
    PendingFolder = mPendingFolder
End Property

Public Property Let PendingFolder(ByVal Data As String)
    mPendingFolder = Data
End Property


Public Property Get SubmittedFolder() As String
    SubmittedFolder = mSubmittedFolder
End Property

Public Property Let SubmittedFolder(ByVal Data As String)
    mSubmittedFolder = Data
End Property


Public Property Get ArchivedFolder() As String
    ArchivedFolder = mArchivedFolder
End Property

Public Property Let ArchivedFolder(ByVal Data As String)
    mArchivedFolder = Data
End Property


Public Property Get TemplateFolder() As String
    TemplateFolder = mTemplateFolder
End Property

Public Property Let TemplateFolder(ByVal Data As String)
    mTemplateFolder = Data
End Property

Public Property Get DisplayTitle() As String
    DisplayTitle = mDisplayTitle
End Property

Public Property Let DisplayTitle(ByVal Data As String)
    mDisplayTitle = Data
End Property

Public Property Let SendToAddress(ByVal Data As String)
    mSendToAddress = Data
End Property

Public Property Get SendToAddress() As String
    SendToAddress = mSendToAddress
End Property

Public Property Let ActiveConfigGroup(ByVal Data As String)
    mActiveConfigGroup = Data
End Property

Public Property Get ActiveConfigGroup() As String
    ActiveConfigGroup = mActiveConfigGroup
End Property


'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property



