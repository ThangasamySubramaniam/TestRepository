VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwFormProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwFormProperties
'

Option Explicit

Private mFormID As String
Private mSiteID As String
Private mSiteKey As Long
Private mRoleType As String
Private mSubject As String
Private mFileName As String
Private mFullFilename As String
Private mCreatedDate As String
Private mModified As String
Private mEventType As Long
Private mEventRecordKey As Long
Private mTargetSiteID As String
Private mTargetSiteKey As Long
Private mIsFleetTransmission As Boolean

Private moParent As mwForm

Public Property Set Parent(oParent As mwForm)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As mwForm
   Set Parent = moParent
End Property

Public Property Let FormID(ByVal Data As String)
    mFormID = Data
End Property

Public Property Get FormID() As String
    FormID = mFormID
End Property

Public Property Let SiteID(ByVal Data As String)
    mSiteID = Data
End Property

Public Property Get SiteID() As String
    SiteID = mSiteID
End Property

Public Property Let RoleType(ByVal Data As String)
    mRoleType = Data
End Property

Public Property Get RoleType() As String
    RoleType = mRoleType
End Property

Public Property Let subject(ByVal Data As String)
    mSubject = Data
End Property

Public Property Get subject() As String
   subject = mSubject
End Property

Public Property Let FileName(ByVal Data As String)
    mFileName = Data
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let CreatedDate(ByVal Data As String)
    mCreatedDate = Data
End Property

Public Property Get CreatedDate() As String
    CreatedDate = mCreatedDate
End Property

Public Property Let Modified(ByVal Data As String)
    mModified = Data
End Property

Public Property Get Modified() As String
    Modified = mModified
End Property

Public Property Let FullFilename(ByVal Data As String)
    mFullFilename = Data
End Property

Public Property Get FullFilename() As String
    FullFilename = mFullFilename
End Property


Public Property Let EventType(ByVal Data As Long)
    mEventType = Data
End Property

Public Property Get EventType() As Long
    EventType = mEventType
End Property

Public Property Let EventRecordKey(ByVal Data As Long)
    mEventRecordKey = Data
End Property

Public Property Get EventRecordKey() As Long
    EventRecordKey = mEventRecordKey
End Property

Public Property Let SiteKey(ByVal Data As Long)
    mSiteKey = Data
End Property

Public Property Get SiteKey() As Long
    SiteKey = mSiteKey
End Property

Public Property Let TargetSiteID(ByVal Data As String)
    mTargetSiteID = Data
End Property

Public Property Get TargetSiteID() As String
    TargetSiteID = mTargetSiteID
End Property

Public Property Let TargetSiteKey(ByVal Data As Long)
    mTargetSiteKey = Data
End Property

Public Property Get TargetSiteKey() As Long
    TargetSiteKey = mTargetSiteKey
End Property

Public Property Let IsFleetTransmission(ByVal Data As Boolean)
    mIsFleetTransmission = Data
End Property

Public Property Get IsFleetTransmission() As Boolean
    IsFleetTransmission = mIsFleetTransmission
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


