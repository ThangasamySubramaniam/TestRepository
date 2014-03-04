VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mtmlRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' xmlRequisition - Generate XML Message for new Requisitions
' 23 March, 2002
'
Option Explicit



'
' Req Header Info...
'
Private mReqNumber As String
Private mReqDate As String
Private mDeliverByDate As String
Private mPreparationDate As String
Private mPreparationTime As String
Private mCurrencyCode As String
Private mFaxContact As String
Private mPhoneContact As String
Private mEmailContact As String
Private mContactName As String
Private mInternalComments As String
Private mCommentsToSupplier As String
Private mEtaPort As String
Private mEtdPort As String
Private mRequestedDeliveryDate
Private mExpirationDate As String


'
' Ships Particulars
'
Private mShipyardCode As String
Private mShipyardHullNo As String

' Module variables and objects
Private mXmlMessage As String
Private mIsRequisition As Boolean

Private moLineItems As mtmlLineItems
Private moParent As Object

' Interchange Sender substitution strings
' Only applies to required items...
Const IS_VESSEL_IMO_NUMBER = "[VESSEL_IMO_NUMBER]"
Const IS_SENDER_CODE_QUALIFIER = "[SENDER_CODE_QUALIFIER]"
Const IS_BUYERS_IMPA_ID = "[BUYERS_IMPA_ID]"
Const IS_RECIPIENT_CODE_QUALIFIER = "[RECIPIENT_CODE_QUALIFIER]"
Const IS_PREP_DATE = "[PREP_DATE]"
Const IS_PREP_TIME = "[PREP_TIME]"
Const IS_CONTROL_REFERENCE = "[CONTROL_REFERENCE]"

' Requisition Contents Qualifier substitution strings
Const RCQ_REQ_CONTENTS_QUALIFIER = "[REQ_CONTENTS_QUALIFIER]"
Const RCQ_CURRENCY = "[CURRENCY]"
Const RCQ_MESSAGE_NUMBER = "[MESSAGE_NUMBER]"
Const RCQ_MESSAGE_REFERENCE_NUMBER = "[MESSAGE_REFERENCE_NUMBER]"
Const RCQ_FUNCTION_CODE = "[FUNCTION_CODE]"
Const RCQ_LINE_ITEM_COUNT = "[LINE_ITEM_COUNT]"
Const RCQ_PRIORITY = "[PRIORITY]"
Const RCQ_VERSION_NUMBER = "[VERSION_NUMBER]"
Const RCQ_RELEASE_NUMBER = "[RELEASE_NUMBER]"
Const RCQ_CONTROLLING_AGENCY = "[CONTROLLING_AGENCY]"
Const RCQ_ASSOC_ASSIGNED_CODE = "[ASSOC_ASSIGNED_CODE]"


'Public Property Set Parent(oParent As swSession)
'   If moParent Is Nothing Then
'      Set moParent = oParent
'   End If
'End Property

'Public Property Get Parent() As swSession
'   Set Parent = moParent
'End Property


Public Property Let XmlMessage(ByVal Data As String)
    mXmlMessage = Data
End Property

Public Property Get XmlMessage() As String
    XmlMessage = mXmlMessage
End Property

Public Property Let ReqNumber(ByVal Data As String)
    mReqNumber = Data
End Property

Public Property Get ReqNumber() As String
    ReqNumber = mReqNumber
End Property

Public Property Let ReqDate(ByVal Data As String)
    mReqDate = Data
End Property

Public Property Get ReqDate() As String
    ReqDate = mReqDate
End Property

Public Property Let DeliverByDate(ByVal Data As String)
    mDeliverByDate = Data
End Property

Public Property Get DeliverByDate() As String
    DeliverByDate = mDeliverByDate
End Property

Public Property Let InternalComments(ByVal Data As String)
    mInternalComments = Data
End Property

Public Property Get InternalComments() As String
    InternalComments = mInternalComments
End Property

Public Property Let CommentsToSupplier(ByVal Data As String)
    mCommentsToSupplier = Data
End Property

Public Property Get CommentsToSupplier() As String
    CommentsToSupplier = mCommentsToSupplier
End Property

Public Property Let ShipyardCode(ByVal Data As String)
    mShipyardCode = Data
End Property

Public Property Get ShipyardCode() As String
    ShipyardCode = mShipyardCode
End Property

Public Property Let ShipyardHullNo(ByVal Data As String)
    mShipyardHullNo = Data
End Property

Public Property Get ShipyardHullNo() As String
    ShipyardHullNo = mShipyardHullNo
End Property

' Collection of Line Items...
Public Property Get LineItems() As mtmlLineItems
   If moLineItems Is Nothing Then
      Set moLineItems = New mtmlLineItems
   End If
   Set LineItems = moLineItems
End Property

Public Property Let LineItems(ByVal vData As mtmlLineItems)
   If moLineItems Is Nothing Then
      Set moLineItems = New mtmlLineItems
   End If
   Set moLineItems = vData
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

'
'
'Public Function MakeNewRequisition() As Boolean
'   Dim loRS As Recordset
'   Dim strSQL As String
'   On Error GoTo FunctionError
'
'   Set loRS = New Recordset
'   loRS.CursorLocation = adUseClient
'   '
'   ' 1. Get the mtmlCompany info...
'   '
'   strSQL = "select * from mtmlCompany"
'   'loRS.Open strSQL, moParent.DbConnection, adOpenForwardOnly, adLockReadOnly
'   If loRS.RecordCount < 1 Then
'      'moParent.RaiseError "Error in mtmlRequisition.MakeNewRequisition, No configuration in mtmlCompany."
'      'moParent.MwSes.CloseRecordset loRS
'      MakeNewRequisition = False
'      Exit Function
'   End If
'   'mInterchangeID = loRS!InterchangeID
'   'mInterchangeSenderTemplate = loRS!InterchangeSenderTemplate
'   'mReqContentsQualifierTemplate = loRS!ReqContentsQualifierTemplate
'   'mInterchangeSenderRecipient = loRS!InterchangeSenderRecipient
'   'mInterchangeControlReference = loRS!InterchangeControlReference
'   'loRS.Close
'   '
'   ' ShipParticulars information...
'   '
'
'
'
'
'   mIsRequisition = True
'   MakeNewRequisition = True
'   Exit Function
'FunctionError:
'   'moParent.RaiseError "General Error in mtmlRequisition.MakeNewRequisition", err.Number, err.Description
'   mIsRequisition = False
'   MakeNewRequisition = False
'   'moParent.MwSes.CloseRecordset loRS
'
'End Function


Private Sub Class_Terminate()
   'moParent.KillObject moLineItems
End Sub
