VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mEventType As Long
Private mEventDetailKey As Long
Private mEventKey As Long
Private mIsSubmit As Boolean
Private mIsPrintForms As Boolean
Private mIsAddNew As Boolean
Private mIsAdded As Boolean
Private mIsDelete As Boolean
'Private moRS As Object
'Private moRS As Recordset
Private mIsSubmitted As Boolean
Private mIsDeleted As Boolean
Private mIsReopen As Boolean
Private mIsError As Boolean
Private mIsExplorerInterface As Boolean
Private mSiteKey As Long
Private mFleetKey As Long
Private mAbstractDescription As String
Private mVoyageID As String

Private mFormsPlaceID As String

Public Property Get AbstractDescription() As String
    AbstractDescription = mAbstractDescription
End Property

Public Property Let AbstractDescription(ByVal vData As String)
    mAbstractDescription = vData
End Property

Public Property Get SiteKey() As Long
    SiteKey = mSiteKey
End Property

Public Property Let SiteKey(ByVal vData As Long)
    mSiteKey = vData
End Property

Public Property Get FleetKey() As Long
    FleetKey = mFleetKey
End Property
Public Property Let FleetKey(ByVal vData As Long)
    mFleetKey = vData
End Property

Public Property Get EventType() As Long
    EventType = mEventType
End Property

Public Property Let EventType(ByVal vData As Long)
    mEventType = vData
End Property

Public Property Get EventDetailKey() As Long
    EventDetailKey = mEventDetailKey
End Property

Public Property Let EventDetailKey(ByVal vData As Long)
    mEventDetailKey = vData
End Property

Public Property Get IsSubmit() As Boolean
    IsSubmit = mIsSubmit
End Property

Public Property Let IsSubmit(ByVal vData As Boolean)
    mIsSubmit = vData
End Property

Public Property Get IsPrintForms() As Boolean
    IsPrintForms = mIsPrintForms
End Property

Public Property Let IsPrintForms(ByVal vData As Boolean)
    mIsPrintForms = vData
End Property

Public Property Get IsAddNew() As Boolean
    IsAddNew = mIsAddNew
End Property

Public Property Let IsAddNew(ByVal vData As Boolean)
    mIsAddNew = vData
End Property

Public Property Get IsDelete() As Boolean
    IsDelete = mIsDelete
End Property

Public Property Let IsDelete(ByVal vData As Boolean)
    mIsDelete = vData
End Property

Public Property Get IsExplorerInterface() As Boolean
    IsExplorerInterface = mIsExplorerInterface
End Property

Public Property Let IsExplorerInterface(ByVal vData As Boolean)
    mIsExplorerInterface = vData
End Property

Public Property Get IsSubmitted() As Boolean
    IsSubmitted = mIsSubmitted
End Property

Public Property Let IsSubmitted(ByVal vData As Boolean)
    mIsSubmitted = vData
End Property

Public Property Get IsDeleted() As Boolean
    IsDeleted = mIsDeleted
End Property

Public Property Let IsDeleted(ByVal vData As Boolean)
    mIsDeleted = vData
End Property

Public Property Get isAdded() As Boolean
    isAdded = mIsAdded
End Property

Public Property Let isAdded(ByVal vData As Boolean)
    mIsAdded = vData
End Property

Public Property Let IsError(ByVal vData As Boolean)
    mIsError = vData
End Property

Public Property Get IsError() As Boolean
    IsError = mIsError
End Property

Public Property Let IsReopen(ByVal vData As Boolean)
    mIsReopen = vData
End Property

Public Property Get IsReopen() As Boolean
    IsReopen = mIsReopen
End Property

Public Property Get FormsPlaceID() As String
    FormsPlaceID = mFormsPlaceID
End Property

Public Property Let FormsPlaceID(ByVal vData As String)
    mFormsPlaceID = vData
End Property

Public Property Get EventKey() As Long
    EventKey = mEventKey
End Property

Public Property Let EventKey(ByVal vData As Long)
    mEventKey = vData
End Property

Public Property Let VoyageID(ByVal Data As String)
    mVoyageID = Data
End Property

Public Property Get VoyageID() As String
    VoyageID = mVoyageID
End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property




