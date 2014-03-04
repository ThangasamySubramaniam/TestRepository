VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwEventType
' 10/2002 ms
'

Option Explicit

Private mID As Long
Private mDescription As String
Private mTableName As String
Private mImageName As String
Private mVbClassObject As String
Private mVbDefaultForm As String
Private mIsEnabled As Boolean
Private mIsReopenAllowed As Boolean
Private mIsDeleteAllowed As Boolean
Private mAbstractTemplate As String
Private mFormsPlaceID As String
Private mActvityRemarksShortWidth As Long
Private mActvityRemarksLongWidth As Long
Private mHelpFileID As String
Private mHelpContextID As String
Private mIsAutoCreateEvent As Boolean
Private mFactSetKey As Long
Private mVbClassObjectCustom As String

Dim mIsCreateEventAllowed As Boolean

Public Property Get FactSetKey() As Long
    FactSetKey = mFactSetKey
End Property

Public Property Let FactSetKey(ByVal vData As Long)
    mFactSetKey = vData
End Property



Public Property Get ID() As Long
    ID = mID
End Property

Public Property Let ID(ByVal vData As Long)
    mID = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal vData As String)
    mTableName = vData
End Property

Public Property Get ImageName() As String
    ImageName = mImageName
End Property

Public Property Let ImageName(ByVal vData As String)
    mImageName = vData
End Property

Public Property Get AbstractTemplate() As String
    AbstractTemplate = mAbstractTemplate
End Property

Public Property Let AbstractTemplate(ByVal vData As String)
    mAbstractTemplate = vData
End Property

Public Property Get VbClassObject() As String
    VbClassObject = mVbClassObject
End Property

Public Property Let VbClassObject(ByVal vData As String)
    mVbClassObject = vData
End Property

Public Property Get VbDefaultForm() As String
    VbDefaultForm = mVbDefaultForm
End Property

Public Property Let VbDefaultForm(ByVal vData As String)
    mVbDefaultForm = vData
End Property

Public Property Get FormsPlaceID() As String
    FormsPlaceID = mFormsPlaceID
End Property

Public Property Let FormsPlaceID(ByVal vData As String)
    mFormsPlaceID = vData
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = mIsEnabled
End Property

Public Property Let IsEnabled(ByVal vData As Boolean)
    mIsEnabled = vData
End Property

Public Property Get IsReopenAllowed() As Boolean
    IsReopenAllowed = mIsReopenAllowed
End Property

Public Property Let IsReopenAllowed(ByVal vData As Boolean)
    mIsReopenAllowed = vData
End Property

Public Property Get IsDeleteAllowed() As Boolean
    IsDeleteAllowed = mIsDeleteAllowed
End Property

Public Property Let IsDeleteAllowed(ByVal vData As Boolean)
    mIsDeleteAllowed = vData
End Property

Public Property Get IsCreateEventAllowed() As Boolean
    IsCreateEventAllowed = mIsCreateEventAllowed
End Property

Public Property Let IsCreateEventAllowed(ByVal vData As Boolean)
    mIsCreateEventAllowed = vData
End Property

Public Property Get ActvityRemarksShortWidth() As Long
    ActvityRemarksShortWidth = mActvityRemarksShortWidth
End Property

Public Property Let ActvityRemarksShortWidth(ByVal vData As Long)
    mActvityRemarksShortWidth = vData
End Property

Public Property Get ActvityRemarksLongWidth() As Long
    ActvityRemarksLongWidth = mActvityRemarksLongWidth
End Property

Public Property Let ActvityRemarksLongWidth(ByVal vData As Long)
    mActvityRemarksLongWidth = vData
End Property

Public Property Get HelpFileID() As String
    HelpFileID = mHelpFileID
End Property

Public Property Let HelpFileID(ByVal vData As String)
    mHelpFileID = vData
End Property

Public Property Get HelpContextID() As String
    HelpContextID = mHelpContextID
End Property

Public Property Let HelpContextID(ByVal vData As String)
    mHelpContextID = vData
End Property

Public Property Get IsAutoCreateEvent() As String
    IsAutoCreateEvent = mIsAutoCreateEvent
End Property

Public Property Let IsAutoCreateEvent(ByVal vData As String)
    mIsAutoCreateEvent = vData
End Property

Public Property Get VbClassObjectCustom() As String
    VbClassObjectCustom = mVbClassObjectCustom
End Property

Public Property Let VbClassObjectCustom(ByVal vData As String)
    mVbClassObjectCustom = vData
End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property



