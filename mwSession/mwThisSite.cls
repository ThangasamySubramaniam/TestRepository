VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwThisSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwThisSite - Current Site and company properties...
' 27 March 2002 ms
'
Option Explicit

Private mThisSite As String
Private mSiteRoot As String
Private mUseGlobalConnections As Boolean
Private mWorkflowCfg As String
Private mIsValidatePassword As Boolean
Private mIsValidateEmployeeID As Boolean
Private mCompanyID As String
Private mIsThisSiteValidated As Boolean
Private mIsSatelliteDB As Boolean
Private mSiteSeedOffset As String
Private mAppSequenceOffset As String
Private mSatelliteDBName As String
Private mIsReplicateSchemaTransactions As Boolean
Private mSatelliteSeedNo As Long
Private mSiteSeedMaxOffset As Long

Private mValidationDate As String

'By N.Angelakis On 22 April 2009
'DEV-1174 Advance Password Settings
Private mPasswordExpireNoDays As Long
Private mIsPasswordStrong As Boolean
Private mLoginFailNotifyRoleTypeID As Long
Private mPasswordFailedAttempts As Integer

Private mDisableBackupOnExit As Boolean

'DEV-1846 Show/Hide Site Specific users column
'By N.Angelakis 28 APril 2010
Private mIsUserSiteSpecific  As Boolean

Private mIsLoaded As Boolean

Private mIsLogIncomingTransactions As Boolean

' Working Objects...
Private moExtendedKeys As ConfigKeys
Private moParent As Session

Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property


Public Property Get ThisSite() As String
    ThisSite = mThisSite
End Property

Public Property Let ThisSite(ByVal vData As String)
    mThisSite = vData
End Property

Public Property Get SiteRoot() As String
    SiteRoot = mSiteRoot
End Property

Public Property Let SiteRoot(ByVal vData As String)
    mSiteRoot = vData
End Property

Public Property Get UseGlobalConnections() As Boolean
    UseGlobalConnections = mUseGlobalConnections
End Property

Public Property Let UseGlobalConnections(ByVal vData As Boolean)
    mUseGlobalConnections = vData
End Property

Public Property Get WorkflowCfg() As String
    WorkflowCfg = mWorkflowCfg
End Property

Public Property Let WorkflowCfg(ByVal vData As String)
    mWorkflowCfg = vData
End Property

Public Property Get isValidatePassword() As Boolean
    isValidatePassword = mIsValidatePassword
End Property

Public Property Let isValidatePassword(ByVal vData As Boolean)
    mIsValidatePassword = vData
End Property

Public Property Get isValidateEmployeeID() As Boolean
    isValidateEmployeeID = mIsValidateEmployeeID
End Property

Public Property Let isValidateEmployeeID(ByVal vData As Boolean)
    mIsValidateEmployeeID = vData
End Property

Public Property Get CompanyID() As String
    CompanyID = mCompanyID
End Property

Public Property Let CompanyID(ByVal vData As String)
    mCompanyID = vData
End Property

Public Property Get IsThisSiteValidated() As Boolean
    IsThisSiteValidated = mIsThisSiteValidated
End Property

Public Property Let IsThisSiteValidated(ByVal vData As Boolean)
    mIsThisSiteValidated = vData
End Property

Public Property Get IsSatelliteDB() As Boolean
    IsSatelliteDB = mIsSatelliteDB
End Property

Public Property Let IsSatelliteDB(ByVal vData As Boolean)
    mIsSatelliteDB = vData
End Property

Public Property Get SiteSeedOffset() As String
    SiteSeedOffset = mSiteSeedOffset
End Property

Public Property Let SiteSeedOffset(ByVal vData As String)
    mSiteSeedOffset = vData
End Property

Public Property Get AppSequenceOffset() As String
    AppSequenceOffset = mAppSequenceOffset
End Property

Public Property Let AppSequenceOffset(ByVal vData As String)
    mAppSequenceOffset = vData
End Property

Public Property Get SatelliteDBName() As String
    SatelliteDBName = mSatelliteDBName
End Property

Public Property Let SatelliteDBName(ByVal vData As String)
    mSatelliteDBName = vData
End Property

Public Property Get IsReplicateSchemaTransactions() As Boolean
    IsReplicateSchemaTransactions = mIsReplicateSchemaTransactions
End Property

Public Property Let IsReplicateSchemaTransactions(ByVal vData As Boolean)
    mIsReplicateSchemaTransactions = vData
End Property

Public Property Get SatelliteSeedNo() As Long
    SatelliteSeedNo = mSatelliteSeedNo
End Property

Public Property Let SatelliteSeedNo(ByVal vData As Long)
    mSatelliteSeedNo = vData
End Property

Public Property Get SiteSeedMaxOffset() As Long
    SiteSeedMaxOffset = mSiteSeedMaxOffset
End Property

Public Property Let SiteSeedMaxOffset(ByVal vData As Long)
    mSiteSeedMaxOffset = vData
End Property

'By N.Angelakis On 22 April 2009
'DEV-1174 Advance Password Settings
Public Property Get PasswordExpireNoDays() As Long
    PasswordExpireNoDays = mPasswordExpireNoDays
End Property
Public Property Let PasswordExpireNoDays(ByVal vData As Long)
    mPasswordExpireNoDays = vData
End Property
Public Property Get IsPasswordStrong() As Boolean
    IsPasswordStrong = mIsPasswordStrong
End Property
Public Property Let IsPasswordStrong(ByVal vData As Boolean)
    mIsPasswordStrong = vData
End Property
Public Property Get LoginFailNotifyRoleTypeID() As Long
    LoginFailNotifyRoleTypeID = mLoginFailNotifyRoleTypeID
End Property
Public Property Let LoginFailNotifyRoleTypeID(ByVal vData As Long)
    mLoginFailNotifyRoleTypeID = vData
End Property
Public Property Get PasswordFailedAttempts() As Long
    PasswordFailedAttempts = mPasswordFailedAttempts
End Property
Public Property Let PasswordFailedAttempts(ByVal vData As Long)
    mPasswordFailedAttempts = vData
End Property

Public Property Get DisableBackupOnExit() As Boolean
    DisableBackupOnExit = mDisableBackupOnExit
End Property
Public Property Let DisableBackupOnExit(ByVal vData As Boolean)
    mDisableBackupOnExit = vData
End Property

'DEV-1846 Show/Hide Site Specific users column
'By N.Angelakis 28 APril 2010
Public Property Get IsUserSiteSpecific() As Boolean
    IsUserSiteSpecific = mIsUserSiteSpecific
End Property
Public Property Let IsUserSiteSpecific(ByVal vData As Boolean)
    mIsUserSiteSpecific = vData
End Property

Public Property Get IsLogIncomingTransactions() As Boolean
    IsLogIncomingTransactions = mIsLogIncomingTransactions
End Property
Public Property Let IsLogIncomingTransactions(ByVal vData As Boolean)
    mIsLogIncomingTransactions = vData
End Property

Public Property Get ValidationDate() As String
    ValidationDate = mValidationDate
End Property

Public Property Let ValidationDate(ByVal vData As String)
    mValidationDate = vData
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


Public Function LoadConfiguration() As Boolean
   Dim loRs As Recordset
   Dim strSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   ' Get it from mwcThisSite
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "Select * from mwcThisSite", moParent.DBConnection, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      moParent.RaiseError "Error in mwSession.mwSite.LoadSiteKeys, ThisSite Record Missing !", Err.Number, Err.Description
      moParent.CloseRecordset loRs
      LoadConfiguration = False
      Exit Function
   End If
       
   If Not moExtendedKeys Is Nothing Then
      Set moExtendedKeys = Nothing
   End If
   Set moExtendedKeys = New ConfigKeys
   
   mIsSatelliteDB = False
   mSiteSeedOffset = ""
   mAppSequenceOffset = ""
   mSiteSeedMaxOffset = 0
   
   For Each loField In loRs.Fields
      Select Case UCase(loField.Name)
   
         Case Is = UCase("ThisSite")
            mThisSite = loRs!ThisSite
         Case Is = UCase("IsThisSiteValidated")
            mIsThisSiteValidated = loRs!IsThisSiteValidated
         Case Is = UCase("SiteRoot")
            If Not IsNull(loRs!SiteRoot) Then
               mSiteRoot = loRs!SiteRoot
               moParent.SiteRoot = loRs!SiteRoot
            End If
         Case Is = UCase("UseGlobalConnections")
            If Not IsNull(loRs!UseGlobalConnections) Then
               moParent.UseGlobalConnections = loRs!UseGlobalConnections
            End If
         Case Is = UCase("isValidatePassword")
            mIsValidatePassword = loRs!isValidatePassword
         Case Is = UCase("isValidateEmployeeID")
            mIsValidateEmployeeID = loRs!isValidateEmployeeID
         Case Is = UCase("WorkflowCfg")
            If Not IsNull(loRs!WorkflowCfg) Then
               mWorkflowCfg = loRs!WorkflowCfg
            End If
         Case Is = UCase("CompanyID")
            If Not IsNull(loRs!CompanyID) Then
               mCompanyID = loRs!CompanyID
            End If
         Case Is = UCase("IsSatelliteDB")
            If Not IsNull(loRs!IsSatelliteDB) Then
               mIsSatelliteDB = loRs!IsSatelliteDB
            End If
         Case Is = UCase("SiteSeedOffset")
            If Not IsNull(loRs!SiteSeedOffset) Then
               mSiteSeedOffset = loRs!SiteSeedOffset
            End If
         Case Is = UCase("AppSequenceOffset")
            If Not IsNull(loRs!AppSequenceOffset) Then
               mAppSequenceOffset = loRs!AppSequenceOffset
            End If
         Case Is = UCase("SatelliteDBName")
            If Not IsNull(loRs!SatelliteDBName) Then
               mSatelliteDBName = loRs!SatelliteDBName
            End If
         Case Is = UCase("IsReplicateSchemaTransactions")
            If Not IsNull(loRs!IsReplicateSchemaTransactions) Then
               mIsReplicateSchemaTransactions = loRs!IsReplicateSchemaTransactions
            End If
         Case Is = UCase("SatelliteSeedNo")
            If Not IsNull(loRs!SatelliteSeedNo) Then
               mSatelliteSeedNo = loRs!SatelliteSeedNo
            End If
         Case Is = UCase("SiteSeedMaxOffset")
            If Not IsNull(loRs!SiteSeedMaxOffset) Then
               mSiteSeedMaxOffset = loRs!SiteSeedMaxOffset
            End If
         
         'By N.Angelakis On 22 April 2009
         'DEV-1174 Advance Password Settings
         Case Is = UCase("PasswordExpireNoDays")
            mPasswordExpireNoDays = ZeroNull(loRs!PasswordExpireNoDays)
         Case Is = UCase("IsPasswordStrong")
            mIsPasswordStrong = BoolNull(loRs!IsPasswordStrong) 'Added BoolNull By N.ANgelakis 4th October 2010
         Case Is = UCase("LoginFailNotifyRoleTypeID")
            mLoginFailNotifyRoleTypeID = ZeroNull(loRs!LoginFailNotifyRoleTypeID)
         Case Is = UCase("PasswordFailedAttempts")
            mPasswordFailedAttempts = ZeroNull(loRs!PasswordFailedAttempts)
            
         Case Is = UCase("DisableBackupOnExit")
            If Not IsNull(loRs!DisableBackupOnExit) Then
               mDisableBackupOnExit = loRs!DisableBackupOnExit
            End If
         
         'DEV-1846 Show/Hide Site Specific users column
         'By N.Angelakis 28 APril 2010
         Case Is = UCase("IsUserSiteSpecific")
            If Not IsNull(loRs!IsUserSiteSpecific) Then
               mIsUserSiteSpecific = BoolNull(loRs!IsUserSiteSpecific) 'Added BoolNull By N.ANgelakis 4th October 2010
            End If
            
         Case Is = UCase("ValidationDate")
            If Not IsNull(loRs!ValidationDate) Then
               ValidationDate = loRs!ValidationDate
            End If
         Case Is = UCase("IsLogIncomingTransactions")
            If Not IsNull(loRs!IsLogIncomingTransactions) Then
               mIsLogIncomingTransactions = loRs!IsLogIncomingTransactions
            End If

         Case Else
            '
            ' Otherwise, Add to the Collection of configkeys... (extended properties)
            '
            If IsNull(loField.value) Then
               moExtendedKeys.Add UCase(loField.Name), "", UCase(loField.Name)
            Else
               If goSession.IsOracle And loField.Type = adNumeric And loField.Precision = 1 Then
                  If loField.value = 0 Then
                     moExtendedKeys.Add UCase(loField.Name), "False", UCase(loField.Name)
                  Else
                     moExtendedKeys.Add UCase(loField.Name), "True", UCase(loField.Name)
                  End If
               Else
                  moExtendedKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
               End If
            End If
            
      End Select
   Next
            
            
            
   moParent.CloseRecordset loRs
   LoadConfiguration = True
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwThisSite.LoadConfiguration. ", Err.Number, Err.Description
   moParent.CloseRecordset loRs
   LoadConfiguration = False
End Function


Public Function GetExtendedProperty(KeyName As String) As String
   On Error GoTo FunctionError
   GetExtendedProperty = moExtendedKeys(UCase(KeyName)).KeyValue
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(KeyName As String, KeyValue As String) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   strSQL = "select * from mwcThisSite"
   If moParent.UpdateExtendedKeyData(KeyName, KeyValue, strSQL) Then
      moExtendedKeys(UCase(KeyName)).KeyValue = KeyValue
   End If
   SetExtendedProperty = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwThisSite.SetExtendedProperty. Invalid Key ?: " & KeyName, Err.Number, Err.Description
   SetExtendedProperty = False
End Function

