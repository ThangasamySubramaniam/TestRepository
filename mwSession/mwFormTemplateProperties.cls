VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwFormTemplateProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwFormTemplateProperties
' 8/2001 ms
Option Explicit

Private mTemplateID As String
Private mDescription As String
Private mFilePrefix As String
Private mDefaultEmailAddress As String
Private mFormType As mw_FormType
Private moTemplatePlace As Place
Private mSourceFilename As String
Private mDeleteAfterTransmit As Boolean
Private mMakeInsertStatement As Boolean
Private mExcludeExternalForms As Boolean
Private mIsOverride As Boolean
Private mConvertToDatagram As Boolean
Private mDatagramTemplateID As String
Private mIsShipboardOnly As Boolean
Private mIsProtectWorkbook As Boolean


Public Property Let TemplateID(ByVal Data As String)
    mTemplateID = Data
End Property

Public Property Get TemplateID() As String
    TemplateID = mTemplateID
End Property

Public Property Let Description(ByVal Data As String)
    mDescription = Data
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let FilePrefix(ByVal Data As String)
    mFilePrefix = Data
End Property

Public Property Get FilePrefix() As String
    FilePrefix = mFilePrefix
End Property


Public Property Let DefaultEmailAddress(ByVal Data As String)
    mDefaultEmailAddress = Data
End Property

Public Property Get DefaultEmailAddress() As String
    DefaultEmailAddress = mDefaultEmailAddress
End Property

Public Property Let FormType(ByVal Data As mw_FormType)
    mFormType = Data
End Property

Public Property Get FormType() As mw_FormType
    FormType = mFormType
End Property

Public Property Set TemplatePlace(ByVal Data As Place)
    Set moTemplatePlace = Data
End Property

Public Property Get TemplatePlace() As Place
    Set TemplatePlace = moTemplatePlace
End Property

Public Property Let SourceFilename(ByVal Data As String)
    mSourceFilename = Data
End Property

Public Property Get SourceFilename() As String
    SourceFilename = mSourceFilename
End Property

Public Property Let DeleteAfterTransmit(ByVal Data As Boolean)
    mDeleteAfterTransmit = Data
End Property

Public Property Get DeleteAfterTransmit() As Boolean
    DeleteAfterTransmit = mDeleteAfterTransmit
End Property

Public Property Let MakeInsertStatement(ByVal Data As Boolean)
    mMakeInsertStatement = Data
End Property

Public Property Get MakeInsertStatement() As Boolean
    MakeInsertStatement = mMakeInsertStatement
End Property

Public Property Let ExcludeExternalForms(ByVal Data As Boolean)
    mExcludeExternalForms = Data
End Property

Public Property Get ExcludeExternalForms() As Boolean
    ExcludeExternalForms = mExcludeExternalForms
End Property

Public Property Let IsOverride(ByVal Data As Boolean)
    mIsOverride = Data
End Property

Public Property Get IsOverride() As Boolean
    IsOverride = mIsOverride
End Property

Public Property Let ConvertToDatagram(ByVal Data As String)
    mConvertToDatagram = Data
End Property

Public Property Get ConvertToDatagram() As String
    ConvertToDatagram = mConvertToDatagram
End Property

Public Property Let DatagramTemplateID(ByVal Data As String)
    mDatagramTemplateID = Data
End Property

Public Property Get DatagramTemplateID() As String
    DatagramTemplateID = mDatagramTemplateID
End Property

Public Property Let IsShipboardOnly(ByVal Data As Boolean)
    mIsShipboardOnly = Data
End Property

Public Property Get IsShipboardOnly() As Boolean
    IsShipboardOnly = mIsShipboardOnly
End Property

Public Property Let IsProtectWorkbook(ByVal Data As String)
    mIsProtectWorkbook = Data
End Property

Public Property Get IsProtectWorkbook() As String
    IsProtectWorkbook = mIsProtectWorkbook
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property



Private Sub Class_Initialize()
   Set moTemplatePlace = New Place
End Sub

Private Sub Class_Terminate()
   Set moTemplatePlace = Nothing

End Sub
