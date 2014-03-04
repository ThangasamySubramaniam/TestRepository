VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Site Class - Ships, Laptops, Shore Sites...

Option Explicit

'Intrinsic Values...
Private mSiteID As String
Private mSiteType As Long
Private mSiteKey As Long
Private mSiteName As String
Private mSiteSeedNo As Long
Private mShortSiteID As String
Private mYetAnotherSiteID As String
Private mCurrentVoyage As String
Private mNextPort As String
Private mLastPort As String
Private mNextPort_KeyPos As Long
Private mLastPort_KeyPos As Long
Private mNextPortETA As String
Private mLastPortDeparture As String
Private mLastPortArrival As String
Private mIsThisSiteValidated As Boolean
Private mIsShipWorksFleetModel As Boolean

Private mCurrentMaster As String
Private mCurrentPort As String
Private mTransportPlaceID As String
Private mWorkflowSendToAddress As String
Private mWorkflowSendFromAddress As String
Private mShipWorkflowSendToAddress As String
Private mWorkflowErrorNotifyAddress As String
Private mTargetReplicateSiteID As String

'MOL-143 Adding Ship IMO Number
Private mMtmlShipImoNumber As String

Private mDept As String
Private mBatchNbr As String
Private mscVesselGroupKey As Long
Private mscContractKey As Long
Private mGL_CO_TS As Long
Private mGL_CO_PLT As Long
Private mHR_CO_EMP As Long
Private mHR_CO_VIS As Long
Private mPROCESS_LVL As String
Private mNEXT_VISITOR_NUMBER As String

Private mIsZipInZip As Boolean

'By N.Angelakis On 29th June 2009
'MOLTOKYO-31 Add alternate VRS Shore Email Address
Private mVRSShoreEmailAddress As String

Private moParent As Session
Private moSiteKeys As ConfigKeys
Private mIsSiteKeysLoaded As Boolean


Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Set SiteKeys(oValue As ConfigKeys)
   Set moSiteKeys = oValue
End Property

Public Property Get SiteKeys() As ConfigKeys
   Set SiteKeys = moSiteKeys
End Property


'=========== Properties ======================

Public Property Get IsZipInZip() As Boolean
    IsZipInZip = mIsZipInZip
End Property

Public Property Let IsZipInZip(ByVal vData As Boolean)
    mIsZipInZip = vData
End Property

Public Property Get IsShipWorksFleetModel() As Boolean
    IsShipWorksFleetModel = mIsShipWorksFleetModel
End Property

Public Property Let IsShipWorksFleetModel(ByVal vData As Boolean)
    mIsShipWorksFleetModel = vData
End Property

Public Property Get SiteKey() As Long
    SiteKey = mSiteKey
End Property

Public Property Let SiteKey(ByVal vData As Long)
    mSiteKey = vData
End Property

Public Property Get SiteID() As String
    SiteID = mSiteID
End Property

Public Property Let SiteID(ByVal vData As String)
    mSiteID = vData
End Property

Public Property Get SiteName() As String
    SiteName = mSiteName
End Property

Public Property Let SiteName(ByVal vData As String)
    mSiteName = vData
End Property

Public Property Get ShortSiteID() As String
    ShortSiteID = mShortSiteID
End Property

Public Property Let ShortSiteID(ByVal vData As String)
    mShortSiteID = vData
End Property

Public Property Get CurrentVoyage() As String
   CurrentVoyage = mCurrentVoyage
End Property

Public Property Let CurrentVoyage(ByVal vData As String)
    mCurrentVoyage = vData
    SetSiteValue "CurrentVoyage", mCurrentVoyage
End Property

Public Property Get NextPort() As String
   NextPort = mNextPort
End Property

Public Property Let NextPort(ByVal vData As String)
    mNextPort = vData
    SetSiteValue "NextPort", mNextPort
End Property

Public Property Get LastPort() As String
   LastPort = mLastPort
End Property

Public Property Let LastPort(ByVal vData As String)
    mLastPort = vData
    SetSiteValue "LastPort", mLastPort
End Property

Public Property Get NextPort_KeyPos() As Long
   NextPort_KeyPos = mNextPort_KeyPos
End Property

Public Property Let NextPort_KeyPos(ByVal vData As Long)
    mNextPort_KeyPos = vData
    SetSiteValue "NextPort_KeyPos", mNextPort_KeyPos
End Property

Public Property Get LastPort_KeyPos() As Long
   LastPort_KeyPos = mLastPort_KeyPos
End Property

Public Property Let LastPort_KeyPos(ByVal vData As Long)
    mLastPort_KeyPos = vData
    SetSiteValue "LastPort_KeyPos", mLastPort_KeyPos
End Property

Public Property Get NextPortETA() As String
    NextPortETA = mNextPortETA
End Property

Public Property Let NextPortETA(ByVal vData As String)
   If IsDate(vData) Then
      mNextPortETA = vData
      SetSiteValue "NextPortETA", mNextPortETA
   End If
End Property

Public Property Get CurrentMaster() As String
   CurrentMaster = mCurrentMaster
   SetSiteValue "CurrentMaster", mCurrentMaster
End Property

Public Property Let CurrentMaster(ByVal vData As String)
    mCurrentMaster = vData
    SetSiteValue "CurrentMaster", mCurrentMaster
End Property

Public Property Get TransportPlaceID() As String
    TransportPlaceID = mTransportPlaceID
End Property

Public Property Let TransportPlaceID(ByVal vData As String)
    mTransportPlaceID = vData
End Property

Public Property Get WorkflowSendToAddress() As String
    WorkflowSendToAddress = mWorkflowSendToAddress
End Property

Public Property Let WorkflowSendToAddress(ByVal vData As String)
    mWorkflowSendToAddress = vData
End Property

Public Property Get WorkflowSendFromAddress() As String
    WorkflowSendFromAddress = mWorkflowSendFromAddress
End Property

Public Property Let WorkflowSendFromAddress(ByVal vData As String)
    mWorkflowSendFromAddress = vData
End Property

Public Property Get LastPortDeparture() As String
    LastPortDeparture = mLastPortDeparture
End Property

Public Property Let LastPortDeparture(ByVal vData As String)
   If IsDate(vData) Then
      mLastPortDeparture = vData
      SetSiteValue "LastPortDeparture", mLastPortDeparture
   End If
End Property

Public Property Get LastPortArrival() As String
    LastPortArrival = mLastPortArrival
End Property

Public Property Let LastPortArrival(ByVal vData As String)
   If IsDate(vData) Then
      mLastPortArrival = vData
      SetSiteValue "LastPortArrival", mLastPortArrival
   End If
End Property

Public Property Get IsThisSiteValidated() As Boolean
    IsThisSiteValidated = mIsThisSiteValidated
End Property


Public Property Get ShipWorkflowSendToAddress() As String
    ShipWorkflowSendToAddress = mShipWorkflowSendToAddress
End Property

Public Property Let ShipWorkflowSendToAddress(ByVal vData As String)
    mShipWorkflowSendToAddress = vData
End Property

Public Property Let WorkflowErrorNotifyAddress(ByVal vData As String)
    mWorkflowErrorNotifyAddress = vData
End Property

Public Property Get WorkflowErrorNotifyAddress() As String
    WorkflowErrorNotifyAddress = mWorkflowErrorNotifyAddress
End Property

Public Property Get YetAnotherSiteID() As String
    YetAnotherSiteID = mYetAnotherSiteID
End Property

Public Property Let YetAnotherSiteID(ByVal vData As String)
    mYetAnotherSiteID = vData
End Property

Public Property Get SiteSeedNo() As String
    SiteSeedNo = mSiteSeedNo
End Property

Public Property Let SiteSeedNo(ByVal vData As String)
    mSiteSeedNo = vData
End Property

Public Property Get SiteType() As Long
    SiteType = mSiteType
End Property

Public Property Let SiteType(ByVal vData As Long)
    mSiteType = vData
End Property

Public Property Get CurrentPort() As String
    CurrentPort = mCurrentPort
End Property

Public Property Let CurrentPort(ByVal vData As String)
    mCurrentPort = vData
   SetSiteValue "CurrentPort", mCurrentPort
End Property

Public Property Get TargetReplicateSiteID() As String
    TargetReplicateSiteID = mTargetReplicateSiteID
End Property

Public Property Let TargetReplicateSiteID(ByVal vData As String)
    mTargetReplicateSiteID = vData
End Property

Public Property Get Dept() As String
    Dept = mDept
End Property

Public Property Let Dept(ByVal vData As String)
    mDept = vData
   SetSiteValue "Dept", mDept
End Property

Public Property Get BatchNbr() As String
    BatchNbr = mBatchNbr
End Property

Public Property Let BatchNbr(ByVal vData As String)
    mBatchNbr = vData
   SetSiteValue "BatchNbr", mBatchNbr
End Property

Public Property Get scVesselGroupKey() As Long
    scVesselGroupKey = mscVesselGroupKey
End Property

Public Property Let scVesselGroupKey(ByVal vData As Long)
    mscVesselGroupKey = vData
   SetSiteValue "scVesselGroupKey", mscVesselGroupKey
End Property

Public Property Get scContractKey() As Long
    scContractKey = mscContractKey
End Property

Public Property Let scContractKey(ByVal vData As Long)
    mscContractKey = vData
   SetSiteValue "scContractKey", mscContractKey
End Property

Public Property Get GL_CO_TS() As Long
    GL_CO_TS = mGL_CO_TS
End Property

Public Property Let GL_CO_TS(ByVal vData As Long)
    mGL_CO_TS = vData
   SetSiteValue "GL_CO_TS", mGL_CO_TS
End Property

Public Property Get GL_CO_PLT() As Long
    GL_CO_PLT = mGL_CO_PLT
End Property

Public Property Let GL_CO_PLT(ByVal vData As Long)
    mGL_CO_PLT = vData
   SetSiteValue "GL_CO_PLT", mGL_CO_PLT
End Property

Public Property Get HR_CO_EMP() As Long
    HR_CO_EMP = mHR_CO_EMP
End Property

Public Property Let HR_CO_EMP(ByVal vData As Long)
    mHR_CO_EMP = vData
   SetSiteValue "HR_CO_EMP", mHR_CO_EMP
End Property

Public Property Get HR_CO_VIS() As Long
    HR_CO_VIS = mHR_CO_VIS
End Property

Public Property Let HR_CO_VIS(ByVal vData As Long)
    mHR_CO_VIS = vData
   SetSiteValue "HR_CO_VIS", mHR_CO_VIS
End Property

Public Property Get PROCESS_LVL() As String
    PROCESS_LVL = mPROCESS_LVL
End Property

Public Property Let PROCESS_LVL(ByVal vData As String)
    mPROCESS_LVL = vData
   SetSiteValue "PROCESS_LVL", mPROCESS_LVL
End Property

Public Property Get NEXT_VISITOR_NUMBER() As String
    NEXT_VISITOR_NUMBER = mNEXT_VISITOR_NUMBER
End Property

Public Property Let NEXT_VISITOR_NUMBER(ByVal vData As String)
    mNEXT_VISITOR_NUMBER = vData
   SetSiteValue "NEXT_VISITOR_NUMBER", mNEXT_VISITOR_NUMBER
End Property

'By N.Angelakis On 29th June 2009
'MOLTOKYO-31 Add alternate VRS Shore Email Address
Public Property Get VRSShoreEmailAddress() As String
    VRSShoreEmailAddress = mVRSShoreEmailAddress
End Property
Public Property Let VRSShoreEmailAddress(ByVal vData As String)
    mVRSShoreEmailAddress = vData
End Property

'MOL-143 Adding Ship IMO Number
Public Property Get ShipImoNumber() As String
    ShipImoNumber = mMtmlShipImoNumber
End Property

'MOL-143 Adding Ship IMO Number
Public Property Let ShipImoNumber(ByVal vData As String)
    mMtmlShipImoNumber = vData
   SetSiteValue "MtmlShipImoNumber", mMtmlShipImoNumber
End Property


'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

Friend Function LoadSiteKeys(Optional SiteID As Variant, Optional SiteRoot As String, _
 Optional UseGlobalConnections As Boolean) As Boolean
   Dim oRs As Recordset
   Dim strSiteID As String
   Dim sSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   If mIsSiteKeysLoaded Then
      LoadSiteKeys = True
      Exit Function
   End If
   '
   '
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   If IsMissing(SiteID) Then
      ' retired 27 March 2002 - added mwThisSite class...
      ' Get it from mwcThisSite
      'oRs.Open "Select * from mwcThisSite", moParent.DbConnection, adOpenForwardOnly, adLockReadOnly
      'If oRs.RecordCount < 1 Then
      '   moParent.RaiseError "Error in mwSession.mwSite.LoadSiteKeys, ThisSite Record Missing !", err.Number, err.Description
      '   moParent.CloseRecordset oRs
      '   LoadSiteKeys = False
      '   Exit Function
      'End If
      '
      ' ThisSite Values
      '
      'strSiteID = oRs!ThisSite
      'mIsThisSiteValidated = oRs!IsThisSiteValidated
      'If Not IsNull(oRs!SiteRoot) Then moParent.SiteRoot = oRs!SiteRoot
      'If Not IsNull(oRs!UseGlobalConnections) Then moParent.UseGlobalConnections = oRs!UseGlobalConnections
      'oRs.Close
      strSiteID = moParent.ThisSite.ThisSite
      mIsThisSiteValidated = moParent.ThisSite.IsThisSiteValidated
   
   
   Else
      strSiteID = SiteID
      If Not IsMissing(SiteRoot) Then moParent.SiteRoot = SiteRoot
      If Not IsMissing(UseGlobalConnections) Then moParent.UseGlobalConnections = UseGlobalConnections
   End If
   '
   ' Fetch it...
   '
   sSQL = "select * from mwcSites where SiteID='" & strSiteID & "'"
   oRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If oRs.RecordCount < 1 Then
      moParent.RaiseError "Error in mwSession.mwSite.LoadSiteKeys, Matching ThisSite Record Missing in Sites !", Err.Number, Err.Description
      moParent.CloseRecordset oRs
      LoadSiteKeys = False
      Exit Function
   End If
   '
   ' OK, populate the collection...
   '
   For Each loField In oRs.Fields
      Select Case UCase(loField.Name)
         Case Is = "SITEID"
            If Not IsNull(loField.value) Then mSiteID = loField.value
         Case Is = "SITETYPE"
            If Not IsNull(loField.value) Then mSiteType = loField.value
         Case Is = "SITENAME"
            If Not IsNull(loField.value) Then mSiteName = loField.value
         Case Is = "SHORTSITEID"
            If Not IsNull(loField.value) Then mShortSiteID = loField.value
         Case Is = "CURRENTVOYAGE"
            If Not IsNull(loField.value) Then mCurrentVoyage = loField.value
         Case Is = "NEXTPORT"
            If Not IsNull(loField.value) Then mNextPort = loField.value
         Case Is = "LASTPORT"
            If Not IsNull(loField.value) Then mLastPort = loField.value 'By N.Angelakis On 13th July 2009
         Case Is = "NEXTPORT_KEYPOS"
            If Not IsNull(loField.value) Then mNextPort_KeyPos = loField.value 'By N.Angelakis On 13th July 2009
         Case Is = "LASTPORT_KEYPOS"
            If Not IsNull(loField.value) Then mLastPort_KeyPos = loField.value 'By N.Angelakis On 13th July 2009
         Case Is = "NEXTPORTETA"
            If Not IsNull(loField.value) Then mNextPortETA = loField.value
         Case Is = "LASTPORTDEPARTURE"
            If Not IsNull(loField.value) Then mLastPortDeparture = loField.value
         Case Is = "LASTPORTARRIVAL"
            If Not IsNull(loField.value) Then mLastPortArrival = loField.value
         Case Is = "CURRENTMASTER"
            If Not IsNull(loField.value) Then mCurrentMaster = loField.value
         Case Is = "TRANSPORTPLACEID"
            If Not IsNull(loField.value) Then mTransportPlaceID = loField.value
         Case Is = "WORKFLOWSENDTOADDRESS"
            If Not IsNull(loField.value) Then mWorkflowSendToAddress = loField.value
         Case Is = "WORKFLOWSENDFROMADDRESS"
            If Not IsNull(loField.value) Then mWorkflowSendFromAddress = loField.value
         Case Is = "WORKFLOWERRORNOTIFYADDRESS"
            If Not IsNull(loField.value) Then mWorkflowErrorNotifyAddress = loField.value
         Case Is = "SHIPWORKFLOWSENDTOADDRESS"
            If Not IsNull(loField.value) Then mShipWorkflowSendToAddress = loField.value
         Case Is = "YETANOTHERSITEID"
            If Not IsNull(loField.value) Then mYetAnotherSiteID = loField.value
         Case Is = "SITESEEDNO"
            If Not IsNull(loField.value) Then mSiteSeedNo = loField.value
         Case Is = "CURRENTPORT"
            If Not IsNull(loField.value) Then mCurrentPort = loField.value
         Case Is = "ID"
            mSiteKey = loField.value
         Case Is = "TARGETREPLICATESITEID"
            If Not IsNull(loField.value) Then mTargetReplicateSiteID = loField.value
         Case Is = "ISSHIPWORKSFLEETMODEL"
            If Not IsNull(loField.value) Then mIsShipWorksFleetModel = loField.value
         
         Case Is = "DEPT"
            If Not IsNull(loField.value) Then mDept = loField.value
         Case Is = "BATCHNBR"
            If Not IsNull(loField.value) Then mBatchNbr = loField.value
         Case Is = "SCVESSELGROUPKEY"
            If Not IsNull(loField.value) Then mscVesselGroupKey = loField.value
         Case Is = "SCCONTRACTKEY"
            If Not IsNull(loField.value) Then mscContractKey = loField.value
         Case Is = "GL_CO_TS"
            If Not IsNull(loField.value) Then mGL_CO_TS = loField.value
         Case Is = "GL_CO_PLT"
            If Not IsNull(loField.value) Then mGL_CO_PLT = loField.value
         Case Is = "HR_CO_EMP"
            If Not IsNull(loField.value) Then mHR_CO_EMP = loField.value
         Case Is = "HR_CO_VIS"
            If Not IsNull(loField.value) Then mHR_CO_VIS = loField.value
         Case Is = "PROCESS_LVL"
            If Not IsNull(loField.value) Then mPROCESS_LVL = loField.value
         Case Is = "NEXT_VISITOR_NUMBER"
            If Not IsNull(loField.value) Then mNEXT_VISITOR_NUMBER = loField.value
            
         Case Is = "ISZIPINZIP"
            If IsNull(loField.value) Then
               mIsZipInZip = False
            Else
               mIsZipInZip = loField.value
            End If
            
         'By N.Angelakis On 29th June 2009
         'MOLTOKYO-31 Add alternate VRS Shore Email Address
         Case Is = "VRSSHOREEMAILADDRESS"
            If Not IsNull(loField.value) Then mVRSShoreEmailAddress = loField.value
            
         'MOL-143 Adding Ship IMO Number
         Case Is = "MTMLSHIPIMONUMBER"
            If Not IsNull(loField.value) Then mMtmlShipImoNumber = loField.value
            
         Case Else
            '
            ' Otherwise, Add to the Collection of configkeys... (extended properties)
            '
            If IsNull(loField.value) Then
               moSiteKeys.Add UCase(loField.Name), "", UCase(loField.Name)
            Else
               moSiteKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
            End If
      End Select
   Next
   moParent.CloseRecordset oRs
   LoadSiteKeys = True
   mIsSiteKeysLoaded = True
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwSite.LoadSiteKeys.", Err.Number, Err.Description
   LoadSiteKeys = False
End Function



Public Function GetExtendedProperty(Key As String) As String
   On Error GoTo FunctionError
   GetExtendedProperty = moSiteKeys(UCase(Key)).KeyValue
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(Key As String, value As Variant) As Boolean
   On Error GoTo FunctionError
   SetSiteValue Key, value
   moSiteKeys(Key).KeyValue = value
   SetExtendedProperty = True
   Exit Function
FunctionError:
   'moParent.RaiseError "Error in mwSite.SetExtendedProperty. Invalid Key ?: " & key, err.Number, err.Description
   SetExtendedProperty = False
End Function


Private Sub Class_Initialize()
   Set moSiteKeys = New ConfigKeys
End Sub

Private Sub Class_Terminate()
   Set moSiteKeys = Nothing
End Sub

Private Function SetSiteValue(FieldName As String, FieldValue As Variant)
   Dim oRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   sSQL = "select * from mwcSites where SiteID='" & mSiteID & "'"
   oRs.Open sSQL, moParent.DBConnection, adOpenStatic, adLockOptimistic
   If oRs.RecordCount < 1 Then
      moParent.RaiseError "Error in mwSession.mwSite.SetSiteValue, Site record not found: " & mSiteID, Err.Number, Err.Description
   Else
      oRs.Fields(FieldName).value = FieldValue
      oRs.Update
   End If
   moParent.CloseRecordset oRs
   Exit Function
FunctionError:
   If Err.Number = 3265 Then
      moParent.RaiseError "Error in mwSite.SetSiteValue. Field missing in table: " & FieldName
   Else
      moParent.RaiseError "General Error in mwSite.SetSiteValue. Field is: " & FieldName, Err.Number, Err.Description
   End If
   moParent.CloseRecordset oRs


End Function


Public Function GetShipWorkflowSendToAddress(Optional SiteID As String, Optional SiteKey As Long)
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If Trim(SiteID) <> "" Then
      sSQL = "select ShipWorkflowSendToAddress from mwcSites where SiteID='" & SiteID & "'"
   ElseIf SiteKey > 0 Then
      sSQL = "select ShipWorkflowSendToAddress from mwcSites where ID=" & SiteKey
   Else
      moParent.RaiseError "Error in mwSession.mwSite.GetShipWorkflowSendToAddress, no valid Site ID/Key provided."
      Exit Function
   End If
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetShipWorkflowSendToAddress = ""
   Else
      If IsNull(loRs!ShipWorkflowSendToAddress) Then
         GetShipWorkflowSendToAddress = ""
      Else
         GetShipWorkflowSendToAddress = loRs!ShipWorkflowSendToAddress
      End If
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetShipWorkflowSendToAddress.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetShipWorkflowSendToAddress = ""

End Function

Public Function GetShipProperty(SiteID As String, FieldName As String, Optional SiteKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If Trim(SiteID) <> "" Then
      sSQL = "select * from mwcSites where SiteID='" & SiteID & "'"
   ElseIf SiteKey > 0 Then
      sSQL = "select * from mwcSites where ID=" & SiteKey
   Else
      'moParent.RaiseError "Error in mwSession.mwSite.GetShipProperty, no valid Site ID/Key provided."
      Exit Function
   End If
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetShipProperty = ""
   Else
      If Not IsNull(loRs.Fields(FieldName).value) Then
         GetShipProperty = loRs.Fields(FieldName).value
      Else
         GetShipProperty = ""
      End If
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetShipProperty.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetShipProperty = ""

End Function

Public Function GetSiteID(SiteKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select SiteID from mwcSites where ID=" & SiteKey
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetSiteID = ""
   Else
      GetSiteID = loRs!SiteID
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetSiteID.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetSiteID = ""

End Function

Public Function GetSiteName(SiteKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select SiteName from mwcSites where ID=" & SiteKey
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetSiteName = ""
   Else
      GetSiteName = BlankNull(loRs!SiteName)
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetSiteName.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetSiteName = ""

End Function

Public Function GetFleetKey(FleetType As Long) As Long
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwcFleets.ID" & _
'     " FROM mwcFleetType INNER JOIN (mwcFleets INNER JOIN mwcFleetSites ON mwcFleets.ID = mwcFleetSites.mwcFleetsKey) ON mwcFleetType.ID = mwcFleets.mwcFleetTypeKey " & _
'     " WHERE mwcFleetType.ID=" & FleetType & " AND mwcFleetSites.mwcSitesKey=" & mSiteKey
   
   sSQL = "SELECT mwcFleets.ID" & _
      " FROM mwcFleetType , mwcFleets , mwcFleetSites " & _
      " WHERE mwcFleets.ID = mwcFleetSites.mwcFleetsKey AND " & _
      " mwcFleetType.ID = mwcFleets.mwcFleetTypeKey AND " & _
      " mwcFleetType.ID = " & FleetType & " AND " & _
      " mwcFleetSites.mwcSitesKey = " & mSiteKey
   
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetFleetKey = -1
   Else
      GetFleetKey = loRs.Fields(0).value
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetFleetKey.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetFleetKey = -1

End Function

Public Function GetFleetName(mwcFleetsKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   
   'DEV-1803 LTIF Reports User Interface and crystal reports
   'Added By N.Angelakis On 24 Feb 2010
   
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT FleetName FROM mwcFleets  WHERE ID=" & mwcFleetsKey
   
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetFleetName = ""
   Else
      GetFleetName = loRs.Fields(0).value
   End If
   moParent.CloseRecordset loRs
   
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetFleetName.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetFleetName = ""

End Function


Public Function GetSiteKey(SiteID As String) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select ID from mwcSites where SiteID='" & SiteID & "'"
   loRs.Open sSQL, moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      GetSiteKey = -1
   Else
      GetSiteKey = loRs!ID
   End If
   moParent.CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwSite.GetSitekey.", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   GetSiteKey = -1

End Function


Public Function GetSiteListRS(Optional IncludeExternalShips As Boolean) As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Dim loRs As Recordset
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If IncludeExternalShips Then
      sSQL = "select ID, SiteName, SiteID from mwcSites WHERE NOT SiteName Is Null"
   Else
      sSQL = "select ID, SiteName, SiteID from mwcSites WHERE SiteName Is NOT Null " & _
       " And (IsExternalShip=0 or IsExternalShip is Null)"
   End If
   If goSession.IsOracle Then
      sSQL = sSQL & " ORDER BY UPPER(SiteName) "
   Else
      sSQL = sSQL & " ORDER BY SiteName "
   End If
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set GetSiteListRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwSite.GetSiteListRS. ", Err.Number, Err.Description
End Function


