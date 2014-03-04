VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwFormWorkflowProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'mwFormWorkflowProperties
Option Explicit


Private mWorkflowAction As mw_WorkflowActions
Private mMapScriptFile As String
Private mVbFormEventDLL As String
Private mWFStage As String
Private mDBConnectString As String
Private mConnectionID As String
Private mNotifyEmailOK As String
Private mIsUpdateAsRecordset As Boolean
Private mNotifyEmailError As String
Private mTolerateMissingColumns As Boolean
Private mTolerateDuplicateRecords As Boolean
Private mTolerateOtherErrors As Boolean

Private mIsRouteToFolderExclusive As Boolean
Private mRouteToFolder As String
Private mIsRouteToMailboxExclusive As Boolean
Private mRouteToEmailAddress As String
Private mIsPocketPcFormCompatible As Boolean
'Private mA As String
'Private mA As String
'Private mA As String
'Private mA As String




Public Property Let WorkflowAction(ByVal Data As mw_WorkflowActions)
    mWorkflowAction = Data
End Property

Public Property Get WorkflowAction() As mw_WorkflowActions
    WorkflowAction = mWorkflowAction
End Property

Public Property Let MapScriptFile(ByVal Data As String)
    mMapScriptFile = Data
End Property

Public Property Get MapScriptFile() As String
    MapScriptFile = mMapScriptFile
End Property

Public Property Let VbFormEventDLL(ByVal Data As String)
    mVbFormEventDLL = Data
End Property

Public Property Get VbFormEventDLL() As String
    VbFormEventDLL = mVbFormEventDLL
End Property

Public Property Let WFStage(ByVal Data As String)
    mWFStage = Data
End Property

Public Property Get WFStage() As String
    WFStage = mWFStage
End Property

Public Property Let DBConnectString(ByVal Data As String)
    mDBConnectString = Data
End Property

Public Property Get DBConnectString() As String
    DBConnectString = mDBConnectString
End Property

Public Property Let ConnectionID(ByVal Data As String)
    mConnectionID = Data
End Property

Public Property Get ConnectionID() As String
    ConnectionID = mConnectionID
End Property

Public Property Get NotifyEmailOK() As String
    NotifyEmailOK = mNotifyEmailOK
End Property

Public Property Let NotifyEmailOK(ByVal vData As String)
    mNotifyEmailOK = vData
End Property

Public Property Get NotifyEmailError() As String
    NotifyEmailError = mNotifyEmailError
End Property

Public Property Let NotifyEmailError(ByVal vData As String)
    mNotifyEmailError = vData
End Property

Public Property Get TolerateMissingColumns() As Boolean
    TolerateMissingColumns = mTolerateMissingColumns
End Property

Public Property Let IsUpdateAsRecordset(ByVal vData As Boolean)
    mIsUpdateAsRecordset = vData
End Property

Public Property Get IsUpdateAsRecordset() As Boolean
    IsUpdateAsRecordset = mIsUpdateAsRecordset
End Property


Public Property Let TolerateMissingColumns(ByVal vData As Boolean)
    mTolerateMissingColumns = vData
End Property

Public Property Get TolerateDuplicateRecords() As Boolean
    TolerateDuplicateRecords = mTolerateDuplicateRecords
End Property

Public Property Let TolerateDuplicateRecords(ByVal vData As Boolean)
    mTolerateDuplicateRecords = vData
End Property


Public Property Get TolerateOtherErrors() As Boolean
    TolerateOtherErrors = mTolerateOtherErrors
End Property

Public Property Let TolerateOtherErrors(ByVal vData As Boolean)
    mTolerateOtherErrors = vData
End Property

Public Property Get RouteToFolder() As String
    RouteToFolder = mRouteToFolder
End Property

Public Property Let RouteToFolder(ByVal vData As String)
    mRouteToFolder = vData
End Property

Public Property Get RouteToEmailAddress() As String
    RouteToEmailAddress = mRouteToEmailAddress
End Property

Public Property Let RouteToEmailAddress(ByVal vData As String)
    mRouteToEmailAddress = vData
End Property

Public Property Get IsRouteToFolderExclusive() As Boolean
    IsRouteToFolderExclusive = mIsRouteToFolderExclusive
End Property

Public Property Let IsRouteToFolderExclusive(ByVal vData As Boolean)
    mIsRouteToFolderExclusive = vData
End Property

Public Property Get IsRouteToMailboxExclusive() As Boolean
    IsRouteToMailboxExclusive = mIsRouteToMailboxExclusive
End Property

Public Property Let IsRouteToMailboxExclusive(ByVal vData As Boolean)
    mIsRouteToMailboxExclusive = vData
End Property

Public Property Get IsPocketPcFormCompatible() As Boolean
    IsPocketPcFormCompatible = mIsPocketPcFormCompatible
End Property

Public Property Let IsPocketPcFormCompatible(ByVal vData As Boolean)
    mIsPocketPcFormCompatible = vData
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



