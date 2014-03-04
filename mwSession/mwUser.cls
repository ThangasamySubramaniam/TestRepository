VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mRoleType As String
Private mRoleTypeKey As Long
Private mEmployeeID As String
Private mEmployeeKey As Long
Private mIsLoadUser As Boolean

Private mUserKey As Long
Private mUserID As String
Private mLogFile As String
Private mResizeForms As Boolean
Private mdebugmode As Boolean
Private mNextFormSequenceNumber As String
Private mDateFormat As String
Private mw_StartOfficeMaximized As Boolean
'
Private mDefaultTransport As mwt_TransportMethod
Private mDefaultEmailCarrier As mw_EmailCarrier

Private mSubmitZipDefaultOn As Boolean
Private mMailUserID As String
Private mMailPassword As String
Private mMailServerName As String

Private momwSecurity As mwUserSecurity
Private moExtendedKeys As ConfigKeys
Private moParent As Session

Private mMailIsSharedOutlook As Boolean

'By N.Angelakis On 22 April 2009
'DEV-1174 Advance Password Settings
Private mPasswordLastChangedDate As Date
Private mIsShoreUser As Boolean

Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"

Public Function IsPasswordValid(strPassword As String) As Boolean
   Dim loEncrypt As New mwEncrypt
   Dim MwPassword As String
   MwPassword = GetExtendedProperty("MwPassword")
   If GetExtendedProperty("PasswordEncrypted") = "1" Or UCase(GetExtendedProperty("PasswordEncrypted")) = "TRUE" Then
      If loEncrypt.EnableEncryption(ENCRYPT_PSWD) Then
         MwPassword = loEncrypt.DecryptString(MwPassword)
      End If
   End If
   
   If strPassword = MwPassword Then
      IsPasswordValid = True
   Else
      IsPasswordValid = False
   End If
End Function
Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Set Security(vData As mwUserSecurity)
   Set momwSecurity = vData
End Property

Public Property Get Security() As mwUserSecurity
   Set Security = momwSecurity
End Property



' Following exposes error collection for public use...
'Public Property Get Errors() As mwErrors
'   Set Errors = moReg.Errors
'End Property

Public Property Let MailServerName(ByVal vData As String)
    mMailServerName = vData
   UpdateUserData "MailServerName", vData
End Property

Public Property Get MailServerName() As String
    MailServerName = mMailServerName
End Property

Public Property Let MailUserID(ByVal vData As String)
    mMailUserID = vData
   UpdateUserData "MailUserID", vData
End Property

Public Property Get MailUserID() As String
    MailUserID = mMailUserID
End Property

Public Property Let MailPassword(ByVal vData As String)
    mMailPassword = vData
   UpdateUserData "MailPassword", vData
End Property

Public Property Get MailPassword() As String
    MailPassword = mMailPassword
End Property

Public Property Let MailIsSharedOutlook(ByVal Data As Boolean)
    mMailIsSharedOutlook = Data
End Property

Public Property Get MailIsSharedOutlook() As Boolean
    MailIsSharedOutlook = mMailIsSharedOutlook
End Property

'By N.Angelakis On 22 April 2009
'DEV-1174 Advance Password Settings
Public Property Get PasswordLastChangedDate() As Date
    PasswordLastChangedDate = mPasswordLastChangedDate
End Property
Public Property Let PasswordLastChangedDate(ByVal Data As Date)
    mPasswordLastChangedDate = Data
End Property
Public Property Get IsShoreUser() As Boolean 'ok
    IsShoreUser = mIsShoreUser
End Property
Public Property Let IsShoreUser(ByVal Data As Boolean)
    mIsShoreUser = Data
End Property




Public Property Let DebugMode(ByVal vData As Boolean)
    mdebugmode = vData
   UpdateUserData "DebugMode", vData
End Property

Public Property Get DebugMode() As Boolean
    DebugMode = mdebugmode
End Property

Public Property Let ResizeForms(ByVal vData As Boolean)
    mResizeForms = vData
   UpdateUserData "ResizeForms", vData
End Property

Public Property Get ResizeForms() As Boolean
    ResizeForms = True
End Property

Public Property Let LogFile(ByVal vData As String)
    mLogFile = vData
   UpdateUserData "LogFile", vData
End Property

Public Property Get LogFile() As String
    LogFile = mLogFile
End Property

Public Property Let RoleType(ByVal vData As String)
    mRoleType = vData
   UpdateUserData "RoleType", vData
End Property

Public Property Get RoleType() As String
    RoleType = mRoleType
End Property


Public Property Let DateFormat(ByVal vData As String)
    mDateFormat = vData
   UpdateUserData "DateFormat", vData
End Property

Public Property Get DateFormat() As String
    DateFormat = mDateFormat
End Property

Public Property Let UserID(ByVal vData As String)
    mUserID = vData
   UpdateUserData "UserID", vData
End Property

Public Property Get UserID() As String
    UserID = mUserID
End Property

Public Property Let StartOfficeMaximized(ByVal vData As Boolean)
    mw_StartOfficeMaximized = vData
   UpdateUserData "StartOfficeMaximized", vData
End Property

Public Property Get StartOfficeMaximized() As Boolean
    StartOfficeMaximized = mw_StartOfficeMaximized
End Property

Public Property Let NextFormSequenceNumber(ByVal vData As String)
   If IsNumeric(vData) Then
      mNextFormSequenceNumber = Format(vData, "0000")
      UpdateUserData "NextFormSequenceNumber", vData
   End If
End Property

Public Property Get NextFormSequenceNumber() As String
    NextFormSequenceNumber = Format(mNextFormSequenceNumber, "0000")
End Property

Public Property Let DefaultTransport(ByVal vData As mwt_TransportMethod)
   mDefaultTransport = vData
   UpdateUserData "DefaultTransport", vData
End Property

Public Property Get DefaultTransport() As mwt_TransportMethod
   DefaultTransport = mDefaultTransport
End Property

Public Property Let SubmitZipDefaultOn(ByVal vData As Boolean)
   If vData = "0" Or vData = "1" Then
      mSubmitZipDefaultOn = vData
      UpdateUserData "SubmitZipDefaultOn", vData
   End If
End Property

Public Property Get SubmitZipDefaultOn() As Boolean
    SubmitZipDefaultOn = mSubmitZipDefaultOn
End Property

Public Property Get UserKey() As Long
    UserKey = mUserKey
End Property

Public Property Let UserKey(ByVal vData As Long)
    mUserKey = vData
End Property
'

Public Property Get RoleTypeKey() As Long
    RoleTypeKey = mRoleTypeKey
End Property

Public Property Let RoleTypeKey(ByVal vData As Long)
    mRoleTypeKey = vData
End Property
'


Public Property Get EmployeeKey() As Long
    EmployeeKey = mEmployeeKey
End Property

Public Property Let EmployeeKey(ByVal vData As Long)
    mEmployeeKey = vData
End Property

Public Property Get DefaultEmailCarrier() As mw_EmailCarrier
    DefaultEmailCarrier = mDefaultEmailCarrier
End Property

Public Property Let DefaultEmailCarrier(ByVal vData As mw_EmailCarrier)
    mDefaultEmailCarrier = vData
   UpdateUserData "DefaultEmailCarrier", vData
   moParent.Message.Carrier = vData
End Property



'Public Property Get HasAmosEmail() As String
'    HasAmosEmail = mHasAmosEmail
'End Property

'Public Property Let HasAmosEmail(ByVal vData As String)
'    mHasAmosEmail = vData
'   UpdateUserData "HasAmosEmail", vData
'End Property
'

'

'

Public Property Get EmployeeID() As String
    EmployeeID = mEmployeeID
End Property

Public Property Let EmployeeID(ByVal vData As String)
    mEmployeeID = vData
   UpdateUserData "EmployeeID", vData
End Property


'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'   UpdateUserData "A", vData
'End Property
'
'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'   UpdateUserData "A", vData
'End Property
'
'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'   UpdateUserData "A", vData
'End Property
'

'
' 7/1/2001 ms - Convert to Database Configuration
' 7/16/2001 ms - Converted to Friend Function, called by Login...
'
Friend Function LoadUser(ByRef oRs As Recordset) As Boolean
   '
   Dim loField As field
   On Error GoTo FunctionError
   '
   ' Don't update User record set when loading it !!!
   '
   mIsLoadUser = True
   '
   ' 3/25/2002 ms - Extended Formtemplate Properties...
       
   If Not moExtendedKeys Is Nothing Then
      Set moExtendedKeys = Nothing
   End If
   Set moExtendedKeys = New ConfigKeys
   For Each loField In oRs.Fields
      'If Not IsNull(loField.value) Then
         Select Case UCase(loField.Name)
      
            Case Is = UCase("RoleType")
               'mRoleType = oRs.Fields("RoleType")
            Case Is = UCase("mwcRoleTypeKey")
               If Not IsNull(oRs.Fields("mwcRoleTypeKey").value) Then
                  mRoleTypeKey = oRs.Fields("mwcRoleTypeKey").value
               End If
             Case Is = UCase("LogFile")
               mLogFile = BlankNull(oRs.Fields("LogFile"))
            Case Is = UCase("DebugMode")
               mdebugmode = ZeroNull(oRs.Fields("DebugMode"))
            Case Is = UCase("NextFormSequenceNumber")
               mNextFormSequenceNumber = Format(ZeroNull(oRs.Fields("NextFormSequenceNumber")), "0000")
            Case Is = UCase("DateFormat")
               mDateFormat = BlankNull(oRs.Fields("DateFormat"))
            Case Is = UCase("UserID")
               mUserID = BlankNull(oRs.Fields("UserID"))
            Case Is = UCase("ID")
               mUserKey = oRs.Fields("ID")
            Case Is = UCase("StartOfficeMaximized")
               mw_StartOfficeMaximized = ZeroNull(oRs.Fields("StartOfficeMaximized"))
               ' Transport Switches
            Case Is = UCase("DefaultTransport")
               If Not IsNull(oRs.Fields("DefaultTransport")) Then
                  mDefaultTransport = oRs.Fields("DefaultTransport")
               End If
            Case Is = UCase("AllowTransportOverride")
               momwSecurity.AllowTransportOverride = ZeroNull(oRs.Fields("AllowTransportOverride"))
            Case Is = UCase("AllowTransmitContainer")
               momwSecurity.AllowTransmitContainer = ZeroNull(oRs.Fields("AllowTransmitContainer"))
            Case Is = UCase("SubmitZipDefaultOn")
               mSubmitZipDefaultOn = ZeroNull(oRs.Fields("SubmitZipDefaultOn"))
            Case Is = UCase("DefaultEmailCarrier")
               If Not IsNull(oRs.Fields("DefaultEmailCarrier")) Then
                  mDefaultEmailCarrier = oRs.Fields("DefaultEmailCarrier")
               End If
            Case Is = UCase("MailUserID")
               If Not IsNull(oRs.Fields("MailUserID")) Then
                  mMailUserID = oRs.Fields("MailUserID")
               End If
            Case Is = UCase("MailPassword")
               If Not IsNull(oRs.Fields("MailPassword")) Then
                  mMailPassword = oRs.Fields("MailPassword")
               End If
            Case Is = UCase("MailServerName")
               If Not IsNull(oRs.Fields("MailServerName")) Then
                  mMailServerName = oRs.Fields("MailServerName")
               End If
            Case Is = UCase("EmployeeID")
               If Not IsNull(oRs!EmployeeID) Then
                  mEmployeeID = oRs!EmployeeID
               End If
            Case Is = UCase("scPersonnelKey")
               If Not IsNull(oRs!scPersonnelKey) Then
                  mEmployeeKey = oRs!scPersonnelKey
               End If
               ' Security
            Case Is = UCase("UserConfigAccess")
               momwSecurity.UserConfigAccess = ZeroNull(oRs.Fields("UserConfigAccess"))
            Case Is = UCase("UserConfigPreferencesAccess")
               momwSecurity.UserConfigPreferencesAccess = ZeroNull(oRs.Fields("UserConfigPreferencesAccess"))
            Case Is = UCase("UserConfigWorkflowAccess")
               momwSecurity.UserConfigWorkflowAccess = ZeroNull(oRs.Fields("UserConfigWorkflowAccess"))
            Case Is = UCase("UserConfigTransportAccess")
               momwSecurity.UserConfigTransportAccess = ZeroNull(oRs.Fields("UserConfigTransportAccess"))
            Case Is = UCase("UserConfigAdvancedAccess")
               momwSecurity.UserConfigAdvancedAccess = ZeroNull(oRs.Fields("UserConfigAdvancedAccess"))
            Case Is = UCase("UserConfigLicensingAccess")
               momwSecurity.UserConfigLicensingAccess = ZeroNull(oRs.Fields("UserConfigLicensingAccess"))
            Case Is = UCase("UserConfigLicenseDistribution")
               momwSecurity.UserConfigLicenseDistribution = ZeroNull(oRs.Fields("UserConfigLicenseDistribution"))
            Case Is = UCase("AllowEmailCarrierOverride")
               momwSecurity.AllowEmailCarrierOverride = ZeroNull(oRs.Fields("AllowEmailCarrierOverride"))
            Case Is = UCase("AllowEmailAddressOverride")
               momwSecurity.AllowEmailAddressOverride = ZeroNull(oRs.Fields("AllowEmailAddressOverride"))
            Case Is = UCase("MailIsSharedOutlook")
               MailIsSharedOutlook = ZeroNull(oRs.Fields("MailIsSharedOutlook"))
       
            'By N.Angelakis On 22 April 2009
            'DEV-1174 Advance Password Settings
            Case Is = UCase("PasswordChangedDate")
               PasswordLastChangedDate = BlankNull(oRs.Fields("PasswordChangedDate"))
            Case Is = UCase("IsShoreUser")
               IsShoreUser = ZeroNull(oRs.Fields("IsShoreUser"))
               
               
            Case Else
               '
               ' Otherwise, Add to the Collection of configkeys... (extended properties)
               '
               'Debug.Print loField.Name
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
      'End If
   Next
   
   mIsLoadUser = False
   LoadUser = True
   Exit Function
FunctionError:
   'Resume Next
   moParent.RaiseError "General Error mwSession.User.LoadUser: ", Err.Number, Err.Description
   LoadUser = False
End Function




Public Function GetNextFormSequenceNumber() As String
   Dim strSQL As String
   Dim iTryCount As Integer
   Dim i As Integer
   Dim oRs As Recordset
   On Error GoTo FunctionError
   '
   ' Don't Update User when loading user...
   '
   If mIsLoadUser Then
      Exit Function
   End If
   '
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   strSQL = "select * from mwcUsers where UserID='" & mUserID & "'"
TryAgain:
   oRs.Open strSQL, moParent.DBConnection, adOpenDynamic, adLockPessimistic
   If oRs.RecordCount < 1 Then
      moParent.RaiseError "Error in mwUser.GetFormSequenceNumber, User record missing: " & mUserID
   Else
      If IsNull(oRs!NextFormSequenceNumber) Then
         GetNextFormSequenceNumber = "0000"
      Else
         GetNextFormSequenceNumber = oRs!NextFormSequenceNumber
      End If
      
      oRs!NextFormSequenceNumber = GetNextFormSequenceNumber + 1
      
      If oRs!NextFormSequenceNumber > 32767 Then oRs!NextFormSequenceNumber = 1
      
      On Error Resume Next
      oRs.Update
      If Err Then
         If Err.Number = -2147217864 And iTryCount < 2 Then
            moParent.Logger.LogIt mwl_Transport, mwl_Warning, " Contention error in mwUser.GetNextFormSequenceNumber."
            MsgBox "Record Lock Error - Another User may be logged in as: " & moParent.User.UserID & ". That is not cool." & vbCrLf & _
            " Press OK to try again."
            iTryCount = iTryCount + 1
            oRs.CancelUpdate
            oRs.Close
            On Error GoTo FunctionError
            GoTo TryAgain
         Else
            moParent.RaiseError "General Error in mwUser.GetFormSequenceNumber. ", Err.Number, Err.Description
            moParent.CloseRecordset oRs
            Exit Function
         End If
      End If
   End If
   moParent.CloseRecordset oRs
  Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwUser.GetFormSequenceNumber. ", Err.Number, Err.Description
   moParent.CloseRecordset oRs

End Function


Private Sub Class_Initialize()
   Set momwSecurity = New mwUserSecurity
   Set momwSecurity.Parent = Me
   'Set moReg = New Registry
   'moReg.BaseRegistry = BASE_REGISTRY
End Sub

Private Sub Class_Terminate()
   Set momwSecurity = Nothing
   'Set moReg = Nothing
   'Set RegErrors = Nothing
End Sub

Friend Function UpdateUserData(field As String, value As Variant)
   Dim strSQL As String
   Dim oRs As Recordset
   On Error GoTo FunctionError
   '
   ' Don't Update User when loading user...
   '
   If mIsLoadUser Then
      Exit Function
   End If
   '
   Set oRs = New Recordset
   oRs.CursorLocation = adUseClient
   strSQL = "select * from mwcUsers where UserID='" & mUserID & "'"
   oRs.Open strSQL, moParent.DBConnection, adOpenDynamic, adLockPessimistic
   If oRs.RecordCount < 1 Then
      moParent.RaiseError "Error in mwSession.mwuser.UpdateUserData, User record missing: " & mUserID, Err.Number, Err.Description
   Else
      oRs.Fields(field) = value
      oRs.Update
   End If
   moParent.CloseRecordset oRs
  Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwuser.UpdateUserData. ", Err.Number, Err.Description
   moParent.CloseRecordset oRs
End Function

Public Function GetExtendedProperty(KeyName As String) As String
   On Error GoTo FunctionError
   If mIsLoadUser Then
      GetExtendedProperty = ""
      Exit Function
   End If
   GetExtendedProperty = moExtendedKeys.GetKeyValue(UCase(KeyName))
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Public Function SetExtendedProperty(KeyName As String, KeyValue As String, Optional UserID As String) As Boolean
   Dim strSQL As String
   On Error GoTo FunctionError
   If mIsLoadUser Then
      SetExtendedProperty = False
      Exit Function
   End If
   If Trim(UserID) = "" Then
      strSQL = "select * from mwcUsers where UserID='" & mUserID & "'"
   Else
      strSQL = "select * from mwcUsers where UserID='" & UserID & "'"
   End If
   If moParent.UpdateExtendedKeyData(KeyName, KeyValue, strSQL) Then
      moExtendedKeys(UCase(KeyName)).KeyValue = KeyValue
   End If
   SetExtendedProperty = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwUser.SetExtendedProperty. Invalid Key ?: " & KeyName, Err.Number, Err.Description
   SetExtendedProperty = False
End Function


Public Function ConfigureEmail() As Boolean
   On Error GoTo FunctionError
   Dim loform As frmEmailConfiguration
   Set loform = New frmEmailConfiguration
   loform.Show vbModal
   Unload loform
   Set loform = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwUser.ConfigureEmail.", Err.Number, Err.Description
End Function

Public Function GetUserRoleTypeKey(UsersKey As Long) As Long
'PBT-1639
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   sSQL = " SELECT mwcRoletypeKey FROM mwcUsers WHERE ID = " & UsersKey
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      GetUserRoleTypeKey = loRs!mwcRoleTypekey
   Else
      GetUserRoleTypeKey = 0
   End If
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwManuals.frmRespView.GetUserRoleTypeKey. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

