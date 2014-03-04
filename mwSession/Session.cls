VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Session"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Logger"
Attribute VB_Ext_KEY = "Member1" ,"Registry"
Attribute VB_Ext_KEY = "Member2" ,"Zipper"
' Session - ShipNet Fleet Session  - SDK COM Object Entry Point
'
' 11/13/2000 ms    Copyright 2000/2001 Maritime Systems Inc, All Rights Reserved
' 12/9/2000 ms
' 7/15/2001 ms - Version 2.0 Consolidated DLL for SDK Release
' 9/21/2001 ms - Version 2.0 CrypKey Validation and License Form
'              - Login Form
'

Option Explicit


'******** WFD User Session Staging *******************************************

' Registry staging values...
Private Const HKEY_CURRENT_USER = &H80000001
Private Const BASE_REGISTRY As String = "Software\Maritime Systems Inc"
Private Const REG_SZ As Long = 1
Private Const MAX_LOGIN_COUNT = 5
Private mIsWriting As Boolean

' global shore use of SiteName, SiteKey
Private mSiteNameGlobal As String
Private mSiteKeyGlobal As Long
Private mApplySiteKeyGlobalOverride As Boolean
Private mSiteFleetNameGlobal As String
Private mSiteFleetKeyGlobal As String
 
Private Const MWRT_moRsErrorLog = 1000
 
Public Enum msUserRegistryAccess
   NO_ACCESS = 0
   READ_ONLY = 1
   FULL_ACCESS = 2
End Enum

Private WithEvents moRsErrorLog As Recordset
Attribute moRsErrorLog.VB_VarHelpID = -1
'
' Break Compatibility - Deprecate 12/4/2005 ms
'
'Public Enum mwFeatures
'   Central_Workflow_Agent = 1
'   Edit_Management_Lists = 2
'   Non_Conformities_Management = 3
'   Data_Logger = 4
'   Sdk_Developer_Kit = 5
'   ShipWorks_Equipment = 6
'   Distribute_Packages = 7
'   ShipWorks_Ordering = 8
'   ShipWorks_Maintenance = 9
'End Enum


'
'**** mw_ Constants **************************
'

Public Enum mwt_TransportMethod
   mwt_TransportContainer = 0
   mwt_ElectronicMail = 1
   mwt_Media = 2
   mwt_DemoLoopback = 3
End Enum

Public Enum mw_GlobePriority
   mw_gUrgent
   mw_gMediumHigh
   mw_gMedium
   mw_gMediumLow
   mw_gLow
End Enum

Public Enum mw_GetPlaceType
   mw_DriveLetter = 1
   mw_DrivePath = 2
   mw_DriveLetterPath = 3
   mw_RootPath = 4
   mw_SiteRootDrivePath = 5
End Enum

Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"

Private DotNetFormActive As Boolean
Private mDotNetBackColor As Long
Private mDotNetFontName As String
Private mDotNetFontSize As String
 
Private mdebugmode As Boolean
Private mIsLicensed As Boolean
Private mIsLicenseKeyDisabled As Boolean
Private mIsLicenseServiceOff As Boolean
Private mDisplayErrors As Boolean
Private mRequestBackupOnExit As Boolean
Private mIsWFAProcessNow As Boolean
'

'
' Objects supported by DLL
'
Private moFso As FileSystemObject
' Collections
Private moMenus As mwMenus
Private moDotNetForms As mwDotNetForms
Private moErrors As mwErrors
Private moAppKeys As ConfigGroups
Private moSites As mwSites
Private moPlaces As mwPlaces
Private moEventTypes As mwEventTypes
Private moReplicateWork As mwReplicateWork
Private moRoleType As mwRoleType
'Private moCompany As mwCompany


' Class Objects
Private moSessionSettings As mwSessionSettings
Private moUser As mwUser
Private moEncrypt As mwEncrypt
Private moSite As mwSite
Private moRegistry As Registry
Private moZipper As Zipper
Private moLogger As Logger
Private moMemoryLogger As Logger
Private moMailbox As mwMailbox
Private WithEvents moMessage As mwMessage
Attribute moMessage.VB_VarHelpID = -1
Private moContainer As mwTransportContainer
Private moDbConnection As ADODB.Connection
Private moDbShapeConnection As ADODB.Connection
Private moDbBlobConnection As ADODB.Connection
Private momwAPI As mwAPI
Private momwGUI As mwGUI
Private WithEvents moWorkflow As mwWorkflow
Attribute moWorkflow.VB_VarHelpID = -1
Private moThisSite As mwThisSite
'
' Properties
'
Private mDecryptedDBConnectString As String
Private mDecryptedDbShapeConnectString As String
Private mDBConnectString As String
Private mDbShapeConnectString As String
'Private mMwUserID As String
'Private mMwPassword As String
'Private mMwEmployeeID As String
Private mUserKey As Long

'Private mCrystalDLL As String
'Private mCrystalServer As String
'Private mCrystalDatabaseName As String
'Private mCrystalUser As String
'Private mCrystalPswd As String

Private mUseGlobalConnections As Boolean
Private mTransportMethod As mwt_TransportMethod
Private mIsLoggedIn As Boolean
Private mSiteRoot As String
Private mSuppressLoginWindow As Boolean
Private mIsValidatePassword As Boolean
Private mIsValidateEmployeeID As Boolean

Private mBuildNumber As Long

'
' mwSession Events
'
Event MessageStatus(Msg As String, progress As Integer, Err As Boolean)
Event WorkflowStatus(Msg As String, ErrNumber As Long, ErrDescription As String)
Event mwSessionError(Msg As String, ErrNumber As Long, ErrDescription As String)
Event RequestExitProgram()



'*********** Classes exposed via the Session Object ********************

Public Property Get User() As mwUser
   'If Not IsLoggedIn Then
   '   Exit Property
   'End If
   If moUser Is Nothing Then
      Set moUser = New mwUser
      Set moUser.Parent = Me
   End If
   Set User = moUser
End Property
Public Property Set Encrypt(vData As mwEncrypt)
   If moEncrypt Is Nothing Then
      Set moEncrypt = New mwEncrypt
   End If
   Set Encrypt = vData
End Property

Public Property Get Encrypt() As mwEncrypt
   If moEncrypt Is Nothing Then
      Set moEncrypt = New mwEncrypt
   End If
   Set Encrypt = moEncrypt
End Property

'Public Property Set FSO(vData As FileSystemObject)
'   If moFso Is Nothing Then
'      Set moFso = New FileSystemObject
'   End If
'   Set FSO = vData
'End Property

'Public Property Get FSO() As FileSystemObject
'   If moFso Is Nothing Then
'      Set moFso = New FileSystemObject
'   End If
'   Set FSO = moFso
'End Property

Public Property Set ThisSite(vData As mwThisSite)
   If moThisSite Is Nothing Then
      Set moThisSite = New mwThisSite
      Set moThisSite.Parent = Me
   End If
   Set ThisSite = vData
End Property

Public Property Get ThisSite() As mwThisSite
   If moThisSite Is Nothing Then
      Set moThisSite = New mwThisSite
      Set moThisSite.Parent = Me
   End If
   Set ThisSite = moThisSite
End Property

Public Property Set ReplicateWork(vData As mwReplicateWork)
   If moReplicateWork Is Nothing Then
      Set moReplicateWork = New mwReplicateWork
   End If
   Set ReplicateWork = vData
End Property

Public Property Get ReplicateWork() As mwReplicateWork
   If moReplicateWork Is Nothing Then
      Set moReplicateWork = New mwReplicateWork
   End If
   Set ReplicateWork = moReplicateWork
End Property


Public Property Set RoleType(vData As mwRoleType)
   If moRoleType Is Nothing Then
      Set moRoleType = New mwRoleType
   End If
   Set RoleType = vData
End Property

Public Property Get RoleType() As mwRoleType
   If moRoleType Is Nothing Then
      Set moRoleType = New mwRoleType
   End If
   Set RoleType = moRoleType
End Property



'Public Property Set Company(vData As mwCompany)
'   If moCompany Is Nothing Then
'      Set moCompany = New mwCompany
'      Set moCompany.Parent = Me
'   End If
'   Set Company = vData
'End Property

'Public Property Get Company() As mwCompany
'   If moCompany Is Nothing Then
'      Set moCompany = New mwCompany
'      Set moCompany.Parent = Me
'   End If
'   Set Company = moCompany
'End Property

Public Property Set User(vData As mwUser)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moUser Is Nothing Then
      Set moUser = New mwUser
      Set moUser.Parent = Me
   End If
   Set User = vData
End Property

Public Property Get EventTypes() As mwEventTypes
   If moEventTypes Is Nothing Then
      Set moEventTypes = New mwEventTypes
      LoadMwEventTypes
   End If
   Set EventTypes = moEventTypes
End Property



Public Property Set EventTypes(vData As mwEventTypes)
   If moEventTypes Is Nothing Then
      Set moEventTypes = New mwEventTypes
      LoadMwEventTypes
   End If
   Set moEventTypes = vData
End Property


Public Property Get Errors() As mwErrors
   If moErrors Is Nothing Then
      Set moErrors = New mwErrors
   End If
   Set Errors = moErrors
End Property


Public Property Set Errors(vData As mwErrors)
   If moErrors Is Nothing Then
      Set moErrors = New mwErrors
   End If
   Set moErrors = vData
End Property


Public Property Get Menus() As mwMenus
   If moMenus Is Nothing Then
      Set moMenus = New mwMenus
   End If
   Set Menus = moMenus
End Property


Public Property Set Menus(vData As mwMenus)
   If moMenus Is Nothing Then
      Set moMenus = New mwMenus
   End If
   Set moMenus = vData
End Property

Public Property Get DotNetForms() As mwDotNetForms
   If moDotNetForms Is Nothing Then
      Set moDotNetForms = New mwDotNetForms
   End If
   Set DotNetForms = moDotNetForms
End Property


Public Property Set DotNetForms(vData As mwDotNetForms)
   If moDotNetForms Is Nothing Then
      Set moDotNetForms = New mwDotNetForms
   End If
   Set moDotNetForms = vData
End Property

Public Property Get Places() As mwPlaces
   If moPlaces Is Nothing Then
      Set moPlaces = New mwPlaces
   End If
   Set Places = moPlaces
End Property


Public Property Set Places(vData As mwPlaces)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moPlaces Is Nothing Then
      Set moPlaces = New mwPlaces
   End If
    Set Places = vData
End Property

Public Property Set SessionSettings(ByVal vData As mwSessionSettings)
   If moSessionSettings Is Nothing Then
      Set moSessionSettings = New mwSessionSettings
   End If
   Set moSessionSettings = vData
End Property


Public Property Get SessionSettings() As mwSessionSettings
   If moSessionSettings Is Nothing Then
      Set moSessionSettings = New mwSessionSettings
   End If
   Set SessionSettings = moSessionSettings
End Property


Public Property Get Registry() As Registry
   If moRegistry Is Nothing Then
      Set moRegistry = New Registry
   End If
   Set Registry = moRegistry
End Property

Public Property Set Registry(ByVal vData As Registry)
   If moRegistry Is Nothing Then
      Set moRegistry = New Registry
   End If
   Set moRegistry = vData
End Property

Public Property Get Zipper() As Zipper
   If moZipper Is Nothing Then
      Set moZipper = New Zipper
      'Set moZipper.Parent = Me
   End If
   Set Zipper = moZipper
End Property

Public Property Set Zipper(ByVal vData As Zipper)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moZipper Is Nothing Then
      Set moZipper = New Zipper
      'moZipper.Parent = Me
   End If
   Set moZipper = vData
End Property

Public Property Get Logger() As Logger
   If moLogger Is Nothing Then
      Set moLogger = New Logger
      Set moLogger.Parent = Me
   End If
   Set Logger = moLogger
End Property

Public Property Set Logger(ByVal vData As Logger)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moLogger Is Nothing Then
      Set moLogger = New Logger
      Set moLogger.Parent = Me
   End If
   Set moLogger = vData
End Property


Public Property Get Mailbox() As mwMailbox
   If moMailbox Is Nothing Then
      Set moMailbox = New mwMailbox
      Set moMailbox.Parent = Me
   End If
   Set Mailbox = moMailbox
End Property

Public Property Let Mailbox(ByVal vData As mwMailbox)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moMailbox Is Nothing Then
      Set moMailbox = New mwMailbox
      Set moMailbox.Parent = Me
   End If
   Set moMailbox = vData
End Property

' Object Template
Public Property Get Message() As mwMessage
   If moMessage Is Nothing Then
      Set moMessage = New mwMessage
      Set moMessage.Parent = Me
   End If
   Set Message = moMessage
End Property

Public Property Set Message(ByVal vData As mwMessage)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moMessage Is Nothing Then
      Set moMessage = New mwMessage
      Set moMessage.Parent = Me
   End If
   Set moMessage = vData
End Property

' Object Template
Public Property Get Container() As mwTransportContainer
   If moContainer Is Nothing Then
      Set moContainer = New mwTransportContainer
      Set moContainer.Parent = Me
   End If
   Set Container = moContainer
End Property

Public Property Set Container(ByVal vData As mwTransportContainer)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moContainer Is Nothing Then
      Set moContainer = New mwTransportContainer
      Set moContainer.Parent = Me
   End If
   Set moContainer = vData
End Property

'
' Nice try, but calling EXE can't use the connection...Works
' Except in DLL Debugger Mode...
'

' ADO Data Connection object
Public Property Get DBConnection() As ADODB.Connection
   If moDbConnection Is Nothing Then
      Set moDbConnection = goCon
   End If
   Set DBConnection = goCon
End Property

Public Property Set DBConnection(ByRef vData As ADODB.Connection)
   MsgBox "Set DBConnection no longer supported!"
'   If moDbConnection Is Nothing Then
'      Set moDbConnection = New ADODB.Connection
'   End If
'   Set moDbConnection = vData
'   Set goCon = vData
End Property

' ADO Shape Data Connection object
Public Property Get DbShapeConnection() As ADODB.Connection
   If moDbShapeConnection Is Nothing Then
      CreateDbShapeConnection
   End If
   Set DbShapeConnection = moDbShapeConnection
End Property

Public Property Let DbShapeConnection(ByVal vData As ADODB.Connection)
   If moDbShapeConnection Is Nothing Then
      CreateDbShapeConnection
   End If
   Set moDbShapeConnection = vData
End Property

' Site Class Object
Public Property Get Site() As mwSite
   If moSite Is Nothing Then
      Set moSite = New mwSite
      Set moSite.Parent = Me
   End If
   Set Site = moSite
End Property

Public Property Let Site(ByVal vData As mwSite)
   If Not IsLoggedIn Then
      Exit Property
   End If
   If moSite Is Nothing Then
      Set moSite = New mwSite
      Set moSite.Parent = Me
   End If
   Set moSite = vData
End Property

' API Class
Public Property Get API() As mwAPI
   If momwAPI Is Nothing Then
      Set momwAPI = New mwAPI
   Set momwAPI.Parent = Me
   End If
   Set API = momwAPI
End Property

Public Property Get GUI() As mwGUI
   If momwGUI Is Nothing Then
      Set momwGUI = New mwGUI
   End If
   Set GUI = momwGUI
End Property


' workflow Object
Public Property Get Workflow() As mwWorkflow
   If moWorkflow Is Nothing Then
      Set moWorkflow = New mwWorkflow
      Set moWorkflow.Parent = Me
      Set moWorkflow.Package.Session = Me
      moWorkflow.LoadConfiguration
   End If
   Set Workflow = moWorkflow
End Property

Public Property Let Workflow(ByVal vData As mwWorkflow)
   If moWorkflow Is Nothing Then
      Set moWorkflow = New mwWorkflow
      Set moWorkflow.Parent = Me
      moWorkflow.LoadConfiguration
   End If
   Set moWorkflow = vData
End Property

' Object Template
Public Property Get AppKeys() As ConfigGroups
   If moAppKeys Is Nothing Then
      Set moAppKeys = New ConfigGroups
   End If
   Set AppKeys = moAppKeys
End Property

Public Property Let AppKeys(ByVal vData As ConfigGroups)
   If moAppKeys Is Nothing Then
      Set moAppKeys = New ConfigGroups
   End If
   Set moAppKeys = vData
End Property

' Object Template
'Public Property Get A() As Object
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set A = moA
'End Property

'Public Property Set A(ByVal vData As Object)
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set moA = vData
'End Property

Public Function GetDecryptedDBConnectString(EncryptPassword As String) As String
   If EncryptPassword = ENCRYPT_PSWD Then
      GetDecryptedDBConnectString = mDecryptedDBConnectString
   End If
End Function
Public Property Get DBConnectString() As String
    DBConnectString = mDBConnectString
End Property

Public Property Let DBConnectString(ByVal vData As String)
   On Error GoTo PropertyError
   Dim SysMDW As String
   
   Dim iEnd As Integer
   mDecryptedDBConnectString = vData
   '
   iEnd = InStr(1, vData, "User")
   If iEnd > 0 Then
      mDBConnectString = Left(vData, iEnd - 1)
   Else
      mDBConnectString = vData
   End If
   
   iEnd = InStr(1, mDBConnectString, "Database Password")
   If iEnd > 0 Then
      mDBConnectString = Left(mDBConnectString, iEnd - 1)
   End If
   
   If InStr(1, vData, "Microsoft.Jet.OLEDB") > 0 Then
      If Right(vData, 1) <> ";" Then
         SysMDW = ";Jet OLEDB:System Database=" & goSession.GetAppPath & "\mwDatabaseSystem.MDW;"
      Else
         SysMDW = "Jet OLEDB:System Database=" & goSession.GetAppPath & "\mwDatabaseSystem.MDW;"
      End If
      mDecryptedDBConnectString = mDecryptedDBConnectString & SysMDW
      mDBConnectString = mDBConnectString & SysMDW
   End If
   
   Exit Property
   
PropertyError:
   RaiseError "Error in mwSession.Let_DBConnectString", Err.Number, Err.Description
End Property
Public Function GetDecryptedDbShapeConnectString(EncryptPassword As String) As String
   If EncryptPassword = ENCRYPT_PSWD Then
      GetDecryptedDbShapeConnectString = mDecryptedDbShapeConnectString
   End If
End Function

Public Property Get DbShapeConnectString() As String
    DbShapeConnectString = mDbShapeConnectString
End Property

Public Property Let DbShapeConnectString(ByVal vData As String)
   On Error GoTo PropertyError
   Dim SysMDW As String
   Dim iEnd As Integer
   
   mDecryptedDbShapeConnectString = vData
   
   iEnd = InStr(1, vData, "User")
   If iEnd > 0 Then
      mDbShapeConnectString = Left(vData, iEnd - 1)
   Else
      mDbShapeConnectString = vData
   End If
   
   iEnd = InStr(1, mDbShapeConnectString, "Database Password")
   If iEnd > 0 Then
      mDbShapeConnectString = Left(mDbShapeConnectString, iEnd - 1)
   End If
   
   If InStr(1, vData, "Microsoft.Jet.OLEDB") > 0 Then
      If Right(vData, 1) <> ";" Then
         SysMDW = ";Jet OLEDB:System Database=" & goSession.GetAppPath & "\mwDatabaseSystem.MDW;"
      Else
         SysMDW = "Jet OLEDB:System Database=" & goSession.GetAppPath & "\mwDatabaseSystem.MDW;"
      End If
   
'      SysMDW = "Jet OLEDB:System Database=" & goSession.GetAppPath & "\mwDatabaseSystem.MDW;"
   
      mDecryptedDbShapeConnectString = mDecryptedDbShapeConnectString & SysMDW
      mDbShapeConnectString = mDbShapeConnectString & SysMDW
   End If
   
   Exit Property
PropertyError:
   RaiseError "Error in mwSession.Let_DbShapeConnectString", Err.Number, Err.Description
End Property

Public Property Get RequestBackupOnExit() As Boolean
    RequestBackupOnExit = mRequestBackupOnExit
End Property

Public Property Let RequestBackupOnExit(ByVal vData As Boolean)
    mRequestBackupOnExit = vData
End Property

Public Function RequestExitProgram()
   RaiseEvent RequestExitProgram
End Function

'Public Property Get CrystalDLL() As String
'    CrystalDLL = mCrystalDLL
'End Property
'
'Public Property Let CrystalDLL(ByVal vData As String)
'    mCrystalDLL = vData
'End Property
'
'Public Property Get CrystalServer() As String
'    CrystalServer = mCrystalServer
'End Property
'
'Public Property Let CrystalServer(ByVal vData As String)
'    mCrystalServer = vData
'End Property
'
'Public Property Get CrystalDatabaseName() As String
'    CrystalDatabaseName = mCrystalDatabaseName
'End Property
'
'Public Property Let CrystalDatabaseName(ByVal vData As String)
'    mCrystalDatabaseName = vData
'End Property

'Public Property Get CrystalUser() As String
'    CrystalUser = mCrystalUser
'End Property
'
'Public Property Let CrystalUser(ByVal vData As String)
'    mCrystalUser = vData
'End Property
'
'Public Property Get CrystalPswd() As String
'    CrystalPswd = mCrystalPswd
'End Property
'
'Public Property Let CrystalPswd(ByVal vData As String)
'    mCrystalPswd = vData
'End Property

Public Property Let TransportMethod(ByVal Data As mwt_TransportMethod)
    mTransportMethod = Data
End Property

Public Property Get TransportMethod() As mwt_TransportMethod
    TransportMethod = mTransportMethod
End Property


Public Property Get UseGlobalConnections() As Boolean
    UseGlobalConnections = mUseGlobalConnections
End Property

Friend Property Let UseGlobalConnections(ByVal vData As Boolean)
    mUseGlobalConnections = vData
End Property

Public Property Get IsLicensed() As Boolean
    IsLicensed = mIsLicensed
End Property

Public Property Get IsLoggedIn() As Boolean
    IsLoggedIn = mIsLoggedIn
End Property

Friend Property Let SiteRoot(ByVal vData As String)
    mSiteRoot = vData
End Property

Public Property Get SiteRoot() As String
    SiteRoot = mSiteRoot
End Property

Public Property Get DebugMode() As Boolean
    DebugMode = mdebugmode
End Property

Public Property Let DebugMode(ByVal vData As Boolean)
    mdebugmode = vData
End Property


Public Property Get SuppressLoginWindow() As Boolean
    SuppressLoginWindow = mSuppressLoginWindow
End Property

Public Property Let SuppressLoginWindow(ByVal vData As Boolean)
    mSuppressLoginWindow = vData
End Property

Friend Property Let UserKey(ByVal vData As String)
    mUserKey = vData
End Property

Public Property Get DisplayErrors() As Boolean
    DisplayErrors = mDisplayErrors
End Property

Public Property Let DisplayErrors(ByVal vData As Boolean)
    mDisplayErrors = vData
End Property

Public Property Get IsSqlServer() As Boolean
    IsSqlServer = gIsSqlServer
End Property
Public Property Get IsOracle() As Boolean
    IsOracle = gIsOracle
End Property
Public Property Get IsAccess() As Boolean
    IsAccess = gIsAccess
End Property

Public Property Get GetAppPath() As String
   
   If gAppPath = "" Then
      
      gAppPath = Environ("MWS_DEV_DEBUG")
      If gAppPath = "" Then
         gAppPath = App.Path
      End If
   End If
   
   GetAppPath = gAppPath
End Property

Public Property Get BuildNumber() As Long
   BuildNumber = mBuildNumber
End Property

Public Property Let BuildNumber(ByVal vData As Long)
   mBuildNumber = vData
End Property

Public Property Get IsWFAProcessNow() As Boolean
    IsWFAProcessNow = mIsWFAProcessNow
End Property

Public Property Let IsWFAProcessNow(ByVal vData As Boolean)
    mIsWFAProcessNow = vData
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

' Object Template

'*************** Public Functions ******************************

Private Sub Class_Initialize()
   On Error GoTo SubError
   Set moErrors = New mwErrors
   Set moAppKeys = New ConfigGroups
   Set moMenus = New mwMenus
   Set moDotNetForms = New mwDotNetForms
   Set moSites = New mwSites
   Set moPlaces = New mwPlaces
   'Set moUser = New mwUser
   'Set moUser.Parent = Me
   Set moSessionSettings = New mwSessionSettings
   mUseGlobalConnections = True
   Set goSession = Me
   DotNetFormActive = False
   
   Exit Sub
SubError:
   RaiseError "Error in mwSession.Class_Initialize", Err.Number, Err.Description
End Sub

Private Sub Class_Terminate()
   'On Error GoTo SubError
   On Error Resume Next
   KillObject moErrors
   KillObject moAppKeys
   KillObject moMenus
   KillObject moDotNetForms
   KillObject moUser
   KillObject moSessionSettings
   KillObject moSites
   KillObject moPlaces
   KillObject moWorkflow
   KillObject moSite
   KillObject moRegistry
   
   If Not moZipper Is Nothing Then
      moZipper.TerminateClass
   End If
   
   KillObject moZipper
   KillObject moLogger
   KillObject moMailbox
   KillObject moMessage
   KillObject moContainer
   CloseConnection moDbConnection
   CloseConnection moDbShapeConnection
   KillObject moDbConnection
   KillObject moDbShapeConnection
   KillObject momwAPI
   KillObject goSession
   'KillObject moMwDocument
   Exit Sub
SubError:
   RaiseError "Error in mwSession.Class_Terminate", Err.Number, Err.Description
   
End Sub

Public Function GetAppKey(Group As String, Key As String) As String
   On Error GoTo GetAppKey_error
   GetAppKey = moAppKeys(UCase(Group)).ConfigKeys(UCase(Key)).KeyValue
   Exit Function
GetAppKey_error:
   GetAppKey = ""
End Function

Public Function SetAppKey(Group As String, Key As String, NewKey As String) As Boolean
   On Error GoTo SetAppKey_error
   moAppKeys(UCase(Group)).ConfigKeys(UCase(Key)).KeyValue = NewKey
   SetAppKey = True
   Exit Function
SetAppKey_error:
   RaiseError "Error in mwSession.SetAppKey, Key Missing. Group=" & Group & ", Key=" & Key, Err.Number, Err.Description
   SetAppKey = False
End Function



'
' General Error Handler
'
Friend Function RaiseError(Msg As String, Optional ByVal ErrNumber As Long, Optional ByVal ErrDescription As String)
   Dim iErr As Long
   Dim strDescription As String
   Dim nRet As Long
   On Error GoTo FunctionError
   iErr = ErrNumber
   strDescription = ErrDescription
   
   '
   ' Log it
   '
   Logger.LogIt mwl_Transport, mwl_Critical, Msg & " " & str(iErr) & strDescription
   
   If mIsWFAProcessNow = True Then
      If IsNetworkAlive(nRet) = 0 Or ErrNumber = -2147467259 Then
         RequestExitProgram
      End If
   End If
   
   If Not Workflow.IsSilentErrors Then
      RaiseEvent mwSessionError(Msg, iErr, strDescription)
      ' raise the error as workflow status anyway so it appears on error logs...
      RaiseEvent WorkflowStatus(Msg, ErrNumber, ErrDescription)
      
      If DotNetFormActive = True Then
         If iErr <> 0 Then
            goSession.GUI.ImprovedMsgBox "VB6Error Encountered: " & Msg & vbCrLf & _
            iErr & " - " & strDescription, vbCritical, "ShipNet Fleet Session Error"
         Else
            goSession.GUI.ImprovedMsgBox "VB6Error Encountered: " & Msg, vbCritical
         End If
      End If
      
   Else
      If mDisplayErrors Then
         goSession.GUI.ImprovedMsgBox Msg & " - " & iErr & " - " & strDescription
      Else
         RaiseEvent WorkflowStatus(Msg, ErrNumber, ErrDescription)
      End If
      moWorkflow.IsWorkflowError = True
   End If
   
   If Not Workflow.IsStopUpdateDBOnError Then
      If Not mIsWriting Then
         mIsWriting = True
         If WriteDbError(Msg, ErrNumber, ErrDescription) Then
            mIsWriting = False
         End If
      End If
   End If
   'Added By N.Angelakis On 12th August 2010
   'The mIsWriting flag appears to be designed to prevent recursive errors
   'from happening if the WriteDbError function gets an error.
   'however we need to catch all errors so after writting error we need to reset
   'in order to write next error
   mIsWriting = False
   
'   '
'   ' Log it
'   '
'   Logger.LogIt mwl_Transport, mwl_Critical, Msg & " " & str(iErr) & strDescription
   'Set oErr = Nothing
   Exit Function
FunctionError:
   MsgBox "Unlogged Error in mwSession.RaiseError. Error Encountered in the Error Handler, Bummer: " & Err.Description
End Function


Public Function LoadMWDesktopConfiguration() As Boolean
   Dim oRsUserMenus As Recordset
   Dim oRsRoleMenus As Recordset
   Dim loRs As Recordset
   Dim strSQL As String
   ' Possibly Null Guys
   Dim strFormID As String
   Dim strDisplayTitle As String
   Static IsDesktopLoaded As Boolean
   Dim cntRoleMenus As Integer
   Dim IsUsemwcFleetMenus As Boolean
   
   On Error GoTo FunctionError
   
   cntRoleMenus = 0
   
   ' Run Once Restriction
   If IsDesktopLoaded Then
      LoadMWDesktopConfiguration = True
      Exit Function
   End If
   
   If Not mIsLoggedIn = True Then
      If Not Login() Then
         LoadMWDesktopConfiguration = False
         Exit Function
      End If
   End If
   
   On Error Resume Next
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select count(*) from mwcFleetMenus"
   loRs.Open strSQL, moDbConnection, adOpenForwardOnly, adLockReadOnly
   
   If Err.Number <> 0 Then
      IsUsemwcFleetMenus = False
   ElseIf loRs.RecordCount > 0 Then
      If loRs.Fields(0).value > 0 Then
         IsUsemwcFleetMenus = True
      Else
         IsUsemwcFleetMenus = False
      End If
   Else
      IsUsemwcFleetMenus = False
   End If
   CloseRecordset loRs
   
   On Error GoTo FunctionError

   '
   ' Get User mwMenus...
   '
   Set oRsUserMenus = New Recordset
   oRsUserMenus.CursorLocation = adUseClient
   strSQL = "select * from mwcUserMenus where UserKey=" & mUserKey & " order by DisplaySequence"
   oRsUserMenus.Open strSQL, moDbConnection, adOpenForwardOnly, adLockReadOnly
   
   Set oRsRoleMenus = New Recordset
   oRsRoleMenus.CursorLocation = adUseClient
   strSQL = "select * from mwcRoleMenus where RoleKey=" & goSession.User.RoleTypeKey & " order by DisplaySequence"
   
   On Error Resume Next
   
   oRsRoleMenus.Open strSQL, moDbConnection, adOpenForwardOnly, adLockReadOnly
   
   cntRoleMenus = oRsRoleMenus.RecordCount
   
   On Error GoTo FunctionError
   
   If oRsUserMenus.RecordCount < 1 And cntRoleMenus < 1 Then
      '
      RaiseError "No Menus defined for user: " & moUser.UserID, Err.Number, Err.Description
      LoadMWDesktopConfiguration = False
      CloseRecordset oRsUserMenus
      Exit Function
   End If
   
   If cntRoleMenus > 0 Then
      LoadMenus oRsRoleMenus, IsUsemwcFleetMenus
   End If
   LoadMenus oRsUserMenus, IsUsemwcFleetMenus
   '
   ' Close up shop
   '
   CloseRecordset oRsUserMenus
   CloseRecordset oRsRoleMenus
   LoadMWDesktopConfiguration = True
   Exit Function
FunctionError:
   'Resume Next
   RaiseError "General Error in mwSession.LoadMWDesktopConfiguration: ", Err.Number, Err.Description
   CloseRecordset oRsUserMenus
   CloseRecordset oRsRoleMenus
   LoadMWDesktopConfiguration = False
End Function
Private Function LoadMenus(oRsUserMenus As Recordset, IsUsemwcFleetMenus As Boolean) As Boolean
   Dim oRsMenus As Recordset
   Dim oRsMenuItems As Recordset
   Dim oMenu As mwMenu
   Dim oItem As mwMenuItem
   Dim loField As field
   Dim strSQL As String
   Dim strFilter As String
   ' Possibly Null Guys
   Dim strFormID As String
   Dim strDisplayTitle As String
   Static IsDesktopLoaded As Boolean
   
   On Error GoTo FunctionError
   
   '
   ' Load all mwMenus for processing
   '
   strSQL = "select * from mwcMenus"
   
   If IsUsemwcFleetMenus Then
      strSQL = strSQL & " WHERE ID IN (SELECT MenuKey from mwcFleetMenus WHERE FleetKey IN " & _
                        " (SELECT mwcFleetsKey from mwcFleetSites WHERE mwcSitesKey = " & _
                        goSession.Site.SiteKey & _
                        " ))"
   End If
   
   Set oRsMenus = New Recordset
   oRsMenus.CursorLocation = adUseClient
   oRsMenus.Open strSQL, moDbConnection, adOpenStatic, adLockReadOnly
   If oRsMenus.RecordCount < 1 Then
      '
      RaiseError "Error in mwSession.LoadMenus, No mwMenus defined in mwDatabase. ", Err.Number, Err.Description
      If IsUsemwcFleetMenus Then
         strSQL = "select * from mwcMenus"
         oRsMenus.Close
         oRsMenus.Open strSQL, moDbConnection, adOpenStatic, adLockReadOnly
      End If
      If oRsMenus.RecordCount < 1 Then
         LoadMenus = False
         CloseRecordset oRsUserMenus
         CloseRecordset oRsMenus
         Exit Function
      End If
   End If
   '
   ' Loop through mwMenus for the user
   '
   Set oRsMenuItems = New Recordset
   oRsMenuItems.CursorLocation = adUseClient
   strSQL = "select * from mwcMenuItems order by DisplaySequence"
   oRsMenuItems.Open strSQL, moDbConnection, adOpenStatic, adLockReadOnly
   Do While Not oRsUserMenus.EOF
      '
      ' Find the menu...
      'oRsMenus.MoveFirst
      oRsMenus.Filter = "ID=" & str(oRsUserMenus!MenuKey)
      If Not oRsMenus.EOF Then
         '
         ' Create Menu object
         '
         
         If IsNull(oRsMenus.Fields("DisplayTitle").value) Then
              Set oMenu = moMenus.Add("No Title", _
               oRsMenus!ID, oRsMenus.Fields("ImageList"), oRsMenus!ID)
         Else
              Set oMenu = moMenus.Add(oRsMenus.Fields("DisplayTitle"), _
              oRsMenus!ID, oRsMenus.Fields("ImageList"), oRsMenus!ID)
         End If
         'Set oMenu = moMenus.Add(oRsMenus.Fields("DisplayTitle"), _
         '  oRsMenus!ID, oRsMenus.Fields("ImageList"), oRsMenus!ID)
         
         If Not oMenu Is Nothing Then
            oRsMenuItems.Filter = "MenuKey=" & oRsMenus!ID
            
            Do While Not oRsMenuItems.EOF
               '
               ' 3/2002 Extended Properties
               '
               If IsNull(oRsMenuItems.Fields("FormID")) Then
                  strFormID = " "
               Else
                  strFormID = oRsMenuItems.Fields("FormID")
               End If
               '
               ' Create Menu Item object
               '
               Set oItem = oMenu.MenuItems.Add(strFormID)
               Set oItem.Session = Me
               For Each loField In oRsMenuItems.Fields
                  Select Case UCase(loField.Name)
                     
                     Case Is = UCase("FormID")
                     Case Is = UCase("DateTimeCreated")
                     Case Is = UCase("DateTimeModified")
                     Case Is = UCase("DisplayTitle")
                        If IsNull(oRsMenuItems.Fields("DisplayTitle")) Then
                           oItem.DisplayTitle = ""
                        Else
                           oItem.DisplayTitle = oRsMenuItems.Fields("DisplayTitle")
                        End If
                     Case Is = UCase("ItemTitle")
                        If Not IsNull(oRsMenuItems.Fields("ItemTitle")) Then
                           oItem.ItemTitle = oRsMenuItems.Fields("ItemTitle")
                        End If
                     Case Is = UCase("ImageName")
                        If Not IsNull(oRsMenuItems.Fields("ImageName")) Then
                           oItem.ImageReference = oRsMenuItems.Fields("ImageName")
                        End If
                     Case Is = UCase("FormType")
                        If Not IsNull(oRsMenuItems.Fields("FormType")) Then
                           oItem.FormType = oRsMenuItems.Fields("FormType")
                        End If
                     Case Is = UCase("AutoLoad")
                        If IsNull(loField.value) Then
                           oItem.AutoLoad = False
                        Else
                           oItem.AutoLoad = oRsMenuItems.Fields("AutoLoad")
                        End If
                     Case Is = UCase("PendingPlaceID")
                        If IsNull(oRsMenuItems.Fields("PendingPlaceID")) Then
                           oItem.FormTemplate.PendingFolder = " "
                        Else
                           oItem.FormTemplate.PendingFolder = oRsMenuItems.Fields("PendingPlaceID")
                        End If
                     Case Is = UCase("SubmittedPlaceID")
                        If IsNull(oRsMenuItems.Fields("SubmittedPlaceID")) Then
                           oItem.FormTemplate.SubmittedFolder = " "
                        Else
                           oItem.FormTemplate.SubmittedFolder = oRsMenuItems.Fields("SubmittedPlaceID")
                        End If
                     Case Is = UCase("ArchivedPlaceID")
                        If IsNull(oRsMenuItems.Fields("ArchivedPlaceID")) Then
                           oItem.FormTemplate.ArchivedFolder = " "
                        Else
                           oItem.FormTemplate.ArchivedFolder = oRsMenuItems.Fields("ArchivedPlaceID")
                        End If
                     Case Is = UCase("TemplatePlaceID")
                        If IsNull(oRsMenuItems.Fields("TemplatePlaceID")) Then
                           oItem.FormTemplate.TemplateFolder = " "
                        Else
                           oItem.FormTemplate.TemplateFolder = oRsMenuItems.Fields("TemplatePlaceID")
                        End If
                     Case Is = UCase("FilePattern")
                        If IsNull(oRsMenuItems.Fields("FilePattern")) Then
                           oItem.FormTemplate.FilePattern = " "
                        Else
                           oItem.FormTemplate.FilePattern = oRsMenuItems.Fields("FilePattern")
                        End If
                     Case Is = UCase("SendToAddress")
                        If IsNull(oRsMenuItems.Fields("SendToAddress")) Then
                           oItem.FormTemplate.SendToAddress = " "
                        Else
                           oItem.FormTemplate.SendToAddress = oRsMenuItems.Fields("SendToAddress")
                        End If
                     Case Is = UCase("ActiveConfigGroup")
                        If IsNull(oRsMenuItems.Fields("ActiveConfigGroup")) Then
                           oItem.FormTemplate.ActiveConfigGroup = " "
                        Else
                           oItem.FormTemplate.ActiveConfigGroup = oRsMenuItems.Fields("ActiveConfigGroup")
                        End If
                     Case Is = UCase("ID")
                        oItem.MenuItemKey = oRsMenuItems!ID
                     Case Else
                        '
                        ' Otherwise, Add to the Collection of configkeys... (extended properties)
                        '
                        If IsNull(loField.value) Then
                           oItem.ExtendedKeys.Add UCase(loField.Name), "", UCase(loField.Name)
                        Else
                           If goSession.IsOracle And loField.Type = adNumeric And loField.Precision = 1 Then
                              If loField.value = 0 Then
                                 oItem.ExtendedKeys.Add UCase(loField.Name), "False", UCase(loField.Name)
                              Else
                                 oItem.ExtendedKeys.Add UCase(loField.Name), "True", UCase(loField.Name)
                              End If
                           Else
                              oItem.ExtendedKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
                           End If
                        End If
                        
                  End Select
               Next
               oRsMenuItems.MoveNext
            Loop
         End If
         oRsMenuItems.Filter = adFilterNone
      End If
      oRsUserMenus.MoveNext
   Loop
   '
   ' Close up shop
   '
   CloseRecordset oRsMenuItems
   CloseRecordset oRsMenus
   LoadMenus = True
   Exit Function
FunctionError:
   'Resume Next
   RaiseError "General Error in mwSession.LoadMenus: ", Err.Number, Err.Description
   CloseRecordset oRsMenuItems
   CloseRecordset oRsMenus
   LoadMenus = False
End Function

'
' Login to Database Connection. If LoginID is ommitted, use registry value for current user...
'
Public Function Login( _
  Optional MwUserID As String, _
  Optional MwPassword As String, _
  Optional DBConnectString As String, _
  Optional DbShapeConnectString As String, _
  Optional EmployeeID As String) As Boolean
   Dim loDW As mwDataWork
   Dim oLoginForm As frmLogin
   Dim sFullComputerName As String
   Dim sDomainName As String
   On Error GoTo FunctionError
   
   If Trim(Environ("MWS_DEV_DEBUG")) <> "" Then
      mdebugmode = True
   Else
      mdebugmode = False
   End If
   Set oLoginForm = New frmLogin
   'oLoginForm.SetLoginParameters MwUserID, MwPassword, DBConnectString, _
     DbShapeConnectString, EmployeeID
   oLoginForm.SetParent Me
   '
   
   If Not oLoginForm.SetLoginParameters(MwUserID, MwPassword, DBConnectString, _
     DbShapeConnectString, EmployeeID) Then
      Login = False
      Unload oLoginForm
      Exit Function
   End If
   
   ' Load login info from registry and mwSession.ini, as required.
   If Not oLoginForm.DetermineLoginParameters() Then
      Login = False
      Unload oLoginForm
      Exit Function
   End If
   
   ' Prevalidate User, or Show the form for logging in...
   If Not oLoginForm.ValidateUser() Then
      '
      ' Display User Login
      '
      
      If Not oLoginForm.PreActivateForm() Then
         Unload oLoginForm
         Login = False
         Exit Function
      End If
      
      oLoginForm.Show
      While oLoginForm.Visible = True
        DoEvents
        goSession.API.SleepApp 10
      Wend
      If Not oLoginForm.IsLoggedIn Then
         Unload oLoginForm
         Login = False
         Exit Function
      End If
      Unload oLoginForm
   End If
   
   If Not oLoginForm Is Nothing Then
      Unload oLoginForm
   End If

   ' OK, we be logged in...
   '
   FinishDBConnection
   
'   Set loDW = New mwDataWork
'   gIsSqlServer = loDW.IsSqlServer()
'   gIsOracle = loDW.IsOracle()
'   gIsAccess = loDW.IsAccess()
'
'   Set loDW = Nothing
   
   LoadAppKeys
   LoadPlaces
'   If moSite Is Nothing Then
'      Set moSite = New mwSite
'      Set moSite.Parent = Me
'   End If
'   moSite.LoadSiteKeys
   '
   ' CrypKey validation..
   '
   #If NO_CRYPKEY Then
      mIsLicensed = True
      LogIt mwl_Workstation, mwl_Information, "Logging in with Licensing Service Disabled."
   #Else
      If ValidateCrypKey() Then
         mIsLicensed = True
      Else
         mIsLicensed = False
      End If
   #End If
   gIsAllowWin9xDelay = False
   gWin9xMilliseconds = 0
   
   sFullComputerName = API.GetMachName
   sDomainName = API.GetLongDomainName
   If sDomainName <> "" Then
      sFullComputerName = sFullComputerName & "." & sDomainName
   End If
   Login = True
   mIsLoggedIn = True
   Workflow.ReceiveRoleTypeNotification
   Logger.LogIt mwl_Workstation, mwl_Information, "mwSession Login, Version: " & _
     App.Major & "." & App.Minor & "." & Left(Format(App.Revision, "0000"), 1) & " Build " & CStr(Val(mID(Format(App.Revision, "0000"), 2, 2))) & " , Workstation OS: " & API.WindowsVersion & ", User: " _
   & goSession.User.UserID & ", Site: " & goSession.Site.SiteID & ", Machine: " & sFullComputerName & ", OS User: " & API.GetOSUserName
   
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.Login. ", Err.Number, Err.Description
   Login = False
End Function
Public Sub FinishDBConnection()
   Dim loDW As mwDataWork
   
   On Error GoTo SubError

   ' OK, we be logged in...
   '
   Set loDW = New mwDataWork
   gIsSqlServer = loDW.IsSqlServer()
   gIsOracle = loDW.IsOracle()
   gIsAccess = loDW.IsAccess()
   
   Set loDW = Nothing
   
   If moSite Is Nothing Then
      Set moSite = New mwSite
      Set moSite.Parent = Me
   End If
   moSite.LoadSiteKeys

   Exit Sub
SubError:
   RaiseError "General Error in mwSession.FinishDBConnection. ", Err.Number, Err.Description

End Sub

Private Function ValidateCrypKey() As Boolean
   Dim loCryp As mwCrypkey
   Dim iCrypInitError As Integer
   Dim iCrypAuthError As Integer
   Dim strBaseReg As String
   Dim loReg As Registry
   On Error GoTo FunctionError
   'MsgBox "debug: before New mwCrypkey"
   Set loCryp = New mwCrypkey
   'MsgBox "debug: before loCryp.InitCrypkeyGE"
   iCrypInitError = loCryp.InitCrypkeyGE
   ' Initialize Crypkey
   If iCrypInitError = -102 Then
      MsgBox "Licensing Service is not working, did you re-install without rebooting ?" & vbCrLf & _
        "MW will try to install the service during next reboot. Try rebooting now.", vbCritical, "ShipNet Fleet Licensing Error"
      Logger.LogIt mwl_Workstation, mwl_Information, "Licensing Service not working, will try RunOnce License Service instaalation during reboot."
      mIsLoggedIn = False
      mIsLicenseServiceOff = True
      '
      ' Set the RunOnce registry key
      '
      Set loReg = New Registry
      strBaseReg = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
      loReg.SetHKLMRegKey strBaseReg, "CrypKeyInstall", goSession.GetAppPath() & "\nt2000\setupex.exe"
      Set loReg = Nothing
      '
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   ElseIf iCrypInitError <> 0 Then
      RaiseError "License Initialization Failed. Error Code is: " & str(iCrypInitError)
      mIsLoggedIn = False
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   End If
   
   ' Authorization for Crypkey
   'MsgBox "debug: before loCryp.GetAuthorizationGE"
   iCrypAuthError = loCryp.GetAuthorizationGE
   If iCrypAuthError = -4 Then
      MsgBox "This Workstation is not licensed, you will need to request a license key." & vbCrLf & _
        "The Licensing Service Manager will be displayed. ", vbInformation
      Logger.LogIt mwl_Workstation, mwl_Information, "ShipNet Fleet First Time startup, the Licensing Service Manager will be displayed. "
      mIsLoggedIn = False
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   ElseIf iCrypAuthError = -9 Or iCrypAuthError = -8 Then
      MsgBox "The ShipNet Fleet license has been disabled, the system clock has been altered more than 90 minutes. " & vbCrLf & _
        "You will need to request a new license key." & vbCrLf & _
        "The Licensing Service Manager will be displayed. ", vbInformation
      Logger.LogIt mwl_Workstation, mwl_Information, "The ShipNet Fleet license is damaged, Error -15, the Licensing Service Manager will be displayed. "
      mIsLoggedIn = False
      ValidateCrypKey = False
      mIsLicenseKeyDisabled = True
      KillObject loCryp
      Exit Function
   
   ElseIf iCrypAuthError = -10 Then
      MsgBox "The ShipNet Fleet license has expired, you will need to request a new license key." & vbCrLf & _
        "The Licensing Service Manager will be displayed. ", vbInformation
      Logger.LogIt mwl_Workstation, mwl_Information, "The ShipNet Fleet license is damaged, Error -15, the Licensing Service Manager will be displayed. "
      mIsLoggedIn = False
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   
   ElseIf iCrypAuthError = -15 Then
      MsgBox "The ShipNet Fleet license is damaged. Either license files were deleted or the application folder was moved ?" & vbCrLf & _
      "The Licensing Service Manager will be displayed, you will need to request a new License Key. ", vbCritical
      Logger.LogIt mwl_Workstation, mwl_Information, "The ShipNet Fleet license is damaged, Error -15, the Licensing Service Manager will be displayed. "
      mIsLoggedIn = False
      ValidateCrypKey = False
      mIsLicenseKeyDisabled = True
      KillObject loCryp
      Exit Function
   
   ElseIf iCrypAuthError <> 0 Then
      RaiseError "License Authorization Failed. Error Code is: " & str(iCrypAuthError)
      mIsLoggedIn = False
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   End If
   If Not loCryp.ckChallengeMS() Then
      RaiseError "License Challenge Authorization Failed."
      mIsLoggedIn = False
      ValidateCrypKey = False
      KillObject loCryp
      Exit Function
   End If
   ValidateCrypKey = True
   KillObject loCryp
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.ValidateCrypKey. ", Err.Number, Err.Description
   ValidateCrypKey = False
End Function

'
' Logout Database Connection
'
Public Function Logout() As Boolean
   On Error GoTo FunctionError
   If Not mIsLoggedIn Then
      Logout = True
      Exit Function
   End If
   CloseConnection moDbConnection
   mIsLoggedIn = False
   Logout = True
   Exit Function
FunctionError:
   RaiseError "Error in mwSession.Logout. ", Err.Number, Err.Description
   Logout = False
End Function


Private Function LoadAppKeys() As Boolean
   Dim oRsAppGroup As Recordset
   Dim oRsAppKey As Recordset
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strSQL As String
   Dim strKeyName As String
   Dim strKeyValue As String
   Dim strPlace As String
   On Error GoTo FunctionError
   '
   ' Database Connection
   '
   If Trim(mDBConnectString) = "" Then
      '
      RaiseError "No Database Connection String Defined for mwDatabase. ", Err.Number, Err.Description
      LoadAppKeys = False
      Exit Function
   End If
   
   Set oRsAppGroup = New Recordset
   oRsAppGroup.CursorLocation = adUseClient
   Set oRsAppKey = New Recordset
   oRsAppKey.CursorLocation = adUseClient
   On Error Resume Next
   oRsAppGroup.Open "select * from mwcApplicationKeys order by ID", _
     moDbConnection, adOpenForwardOnly, adLockReadOnly
   If oRsAppGroup.State = adStateOpen Then
      On Error GoTo FunctionError
      oRsAppKey.Open "select * from mwcAppKey order by GroupKey", _
        moDbConnection, adOpenForwardOnly, adLockReadOnly
      '
      ' All Config groups...
      '
      Do While Not oRsAppGroup.EOF
         Set oGroup = moAppKeys.Add(oRsAppGroup.Fields("GroupID"), oRsAppGroup.Fields("GroupID"))
         strSQL = "GroupKey=" & oRsAppGroup.Fields("ID")
         oRsAppKey.Filter = strSQL
         Do While Not oRsAppKey.EOF
            '
            ' All keys for the Group
            '
            If IsNull(oRsAppKey.Fields("KeyName")) Then
               strKeyName = ""
            Else
               strKeyName = UCase(oRsAppKey.Fields("KeyName"))
            End If
            If IsNull(oRsAppKey.Fields("KeyValue")) Then
               strKeyValue = ""
            Else
               strKeyValue = oRsAppKey.Fields("KeyValue")
            End If
            
            Set oKey = oGroup.ConfigKeys.Add(strKeyName, strKeyValue, strKeyName)
            oRsAppKey.MoveNext
         Loop
         oRsAppKey.Filter = adFilterNone
         oRsAppGroup.MoveNext
      Loop
   End If
   On Error GoTo FunctionError
   '
   ' Close Up Shop
   '
   CloseRecordset oRsAppGroup
   CloseRecordset oRsAppKey
   '
   ' Dec/Jan 2002 - New tables without AutoNumber Key...mwcAppGroup/mwcAppSwitch
   '
   On Error GoTo skip_error
   Set oRsAppGroup = New Recordset
   oRsAppGroup.CursorLocation = adUseClient
   Set oRsAppKey = New Recordset
   oRsAppKey.CursorLocation = adUseClient
   oRsAppGroup.Open "select * from mwcAppGroup", _
     moDbConnection, adOpenForwardOnly, adLockReadOnly
   oRsAppKey.Open "select * from mwcAppGroupKey order by GroupID", _
     moDbConnection, adOpenForwardOnly, adLockReadOnly
   '
   ' All Config groups...
   '
   Do While Not oRsAppGroup.EOF
      Set oGroup = moAppKeys.Add(oRsAppGroup.Fields("GroupID"), oRsAppGroup.Fields("GroupID"))
      strSQL = "GroupID='" & oRsAppGroup.Fields("GroupID") & "'"
      oRsAppKey.Filter = strSQL
      Do While Not oRsAppKey.EOF
         '
         ' All keys for the Group
         '
         If IsNull(oRsAppKey.Fields("KeyName")) Then
            strKeyName = ""
         Else
            strKeyName = UCase(oRsAppKey.Fields("Keyname"))
         End If
         If IsNull(oRsAppKey.Fields("KeyValue")) Then
            strKeyValue = ""
         Else
            strKeyValue = oRsAppKey.Fields("KeyValue")
         End If
         
         Set oKey = oGroup.ConfigKeys.Add(strKeyName, strKeyValue, strKeyName)
         oRsAppKey.MoveNext
      Loop
      oRsAppKey.Filter = adFilterNone
      oRsAppGroup.MoveNext
   Loop
   '
   ' Close Up Shop
   '
   CloseRecordset oRsAppGroup
   CloseRecordset oRsAppKey
   
   LoadAppKeys = True
   Exit Function

FunctionError:
   RaiseError "General Error mwSession.LoadAppKeys.", Err.Number, Err.Description
   LoadAppKeys = False
   CloseRecordset oRsAppGroup
   CloseRecordset oRsAppKey
   Exit Function
skip_error:
   ' assume tables not there until retire old tables...
   CloseRecordset oRsAppGroup
   CloseRecordset oRsAppKey
   LoadAppKeys = True

End Function




Private Function LoadPlaces()
   Dim oRsPlaces As Recordset
   On Error GoTo FunctionError
   Set oRsPlaces = New Recordset
   oRsPlaces.CursorLocation = adUseClient
   oRsPlaces.Open "select * from mwcPlaces", moDbConnection, adOpenForwardOnly, adLockReadOnly
   Do While Not oRsPlaces.EOF
      moPlaces.Add UCase(oRsPlaces!PlaceID)
      With moPlaces.Item(oRsPlaces!PlaceID)
         If Not IsNull(oRsPlaces!DriveLetter) Then .DriveLetter = oRsPlaces!DriveLetter
         If Not IsNull(oRsPlaces!SubFolderPath) Then .SubFolderPath = oRsPlaces!SubFolderPath
         If Not IsNull(oRsPlaces!Notation) Then .Notation = oRsPlaces!Notation
         If Not IsNull(oRsPlaces!PlaceID) Then .PlaceID = oRsPlaces!PlaceID
         If Not IsNull(oRsPlaces!RootPath) Then .RootPath = oRsPlaces!RootPath
      End With
      oRsPlaces.MoveNext
   Loop
   CloseRecordset oRsPlaces
   Exit Function
FunctionError:
   RaiseError "Error in mwSession.LoadPlaces.", Err.Number, Err.Description
   CloseRecordset oRsPlaces
End Function



Public Function GetPlace(PlaceID As String, GetType As mw_GetPlaceType) As String
   Dim strTemp As String
   ' Test if in collection...
   On Error Resume Next
   strTemp = Places(PlaceID).PlaceID
   If Err Then
      GetPlace = ""
      Exit Function
   End If
   On Error GoTo FunctionError
   Select Case GetType
      Case Is = mw_DriveLetter
         GetPlace = Places(PlaceID).DriveLetter
      Case Is = mw_DrivePath
         GetPlace = Places(PlaceID).GetSubFolderPath
      Case Is = mw_DriveLetterPath
         GetPlace = Places(PlaceID).GetSubFolderPath
      Case Is = mw_RootPath
         GetPlace = Places(PlaceID).RootPath
      Case Is = mw_SiteRootDrivePath
         '
         ' Special Puppy: SiteRoot + DrivePath
         '
         ' Test SiteRoot for trailing "\"...
         '
         If mID(mSiteRoot, Len(SiteRoot), 1) = "\" Then
            strTemp = mID(mSiteRoot, 1, (Len(SiteRoot) - 1))
         Else
            strTemp = mSiteRoot
         End If
         '  Assumes SiteRoot is Drive and path (c:\mw_demo), or UNC (\\machine\share)....
         'Where DrivePath=
         '= c:\foo
         If InStr(1, ":\", Places(PlaceID).GetSubFolderPath) Then
            GetPlace = strTemp & "\" & mID(Places(PlaceID).GetSubFolderPath, 3)
         '= c:foo
         ElseIf InStr(1, ":", Places(PlaceID).GetSubFolderPath) Then
            GetPlace = strTemp & "\" & mID(Places(PlaceID).GetSubFolderPath, 2)
         '= \foo
         ElseIf mID(Places(PlaceID).GetSubFolderPath, 1, 1) = "\" Then
            GetPlace = strTemp & Places(PlaceID).GetSubFolderPath
         Else
            '= foo\bla
            GetPlace = strTemp & "\" & Places(PlaceID).GetSubFolderPath
         End If
   End Select
   Exit Function
FunctionError:
   'RaiseError "Error in mwSession.GetPlace. ", err.Number, err.Description
   GetPlace = ""
End Function

Public Function GetPlace2(PlaceID As String, GetType As mw_GetPlaceType) As String
'PBT-1671
   Dim sTemp As String
   Dim sSubFolderPath As String
   Dim sSQL As String
   Dim loRsPlaces As Recordset
   
   On Error GoTo FunctionError
   
   'validate
   If Len(PlaceID) < 1 Then
      goSession.RaisePublicError "Error In GetPlace2, no PlaceID Value"
      Exit Function
   End If
   
   If goSession.IsOracle Then
      sSQL = "SELECT * FROM mwcPlaces WHERE UPPER(PlaceID)  = '" & UCase(PlaceID) & "'"
   Else
      sSQL = "SELECT * FROM mwcPlaces WHERE PlaceID  = '" & PlaceID & "'"
   End If
    
   Set loRsPlaces = New Recordset
   loRsPlaces.CursorLocation = adUseClient
   loRsPlaces.Open sSQL, moDbConnection, adOpenForwardOnly, adLockReadOnly
   
   If loRsPlaces.RecordCount > 0 Then
        
        sSubFolderPath = loRsPlaces!DriveLetter & ":" & "\" & loRsPlaces!RootPath & "\" & loRsPlaces!SubFolderPath
        
        Select Case GetType
         Case Is = mw_DriveLetter
            GetPlace2 = loRsPlaces!DriveLetter
         Case Is = mw_DrivePath
            GetPlace2 = sSubFolderPath
         Case Is = mw_DriveLetterPath
            GetPlace2 = sSubFolderPath
         Case Is = mw_RootPath
            GetPlace2 = loRsPlaces!RootPath
         Case Is = mw_SiteRootDrivePath
            If mID(mSiteRoot, Len(SiteRoot), 1) = "\" Then
               sTemp = mID(mSiteRoot, 1, (Len(SiteRoot) - 1))
            Else
               sTemp = mSiteRoot
            End If
            
            If InStr(1, ":\", sSubFolderPath) Then
               GetPlace2 = sTemp & "\" & mID(sSubFolderPath, 3)
            '= c:foo
            ElseIf InStr(1, ":", sSubFolderPath) Then
               GetPlace2 = sTemp & "\" & mID(sSubFolderPath, 2)
            '= \foo
            ElseIf mID(sSubFolderPath, 1, 1) = "\" Then
               GetPlace2 = sTemp & sSubFolderPath
            Else
               '= foo\bla
               GetPlace2 = sTemp & "\" & sSubFolderPath
            End If
        End Select
   Else
       GetPlace2 = ""
   End If
   
   CloseRecordset loRsPlaces
   
   Exit Function
FunctionError:
   'RaiseError "Error in mwSession.GetPlace2. ", err.Number, err.Description
   GetPlace2 = ""
   CloseRecordset loRsPlaces
End Function
Public Function GetPlaceObject(PlaceID As String) As Place
      
   On Error GoTo FunctionError
   Set GetPlaceObject = Places(PlaceID)
   Exit Function
FunctionError:
   Set GetPlaceObject = Nothing
End Function



'
' 7/10/2001 ms   Carry code over from msDTransport...
'
'

Public Function SendNotification(subject As String, txtBody As String, _
  Optional sendTo As Variant, Optional Attachment As Variant) As Boolean
   Dim loform As mwForm
   Dim strSubject As String
   Dim strTxtBody As String
   On Error GoTo FunctionError
   SendNotification = True
   ' Set Default email if none provided...
   If IsMissing(sendTo) Then
      sendTo = moSite.WorkflowErrorNotifyAddress
   End If
   '
   '
'   If goSession.Site.SiteType <> SITE_TYPE_SHORE And _
'    (moUser.DefaultTransport = mwt_Media Or _
'     moUser.DefaultTransport = mwt_TransportContainer Or _
'     moUser.DefaultTransport = mwt_DemoLoopback) And Not _
'     moWorkflow.GetExtendedProperty("IsAlwaysSendNotificationAsEml") = "True" Then
   If moUser.DefaultTransport <> mwt_ElectronicMail And _
     moWorkflow.GetExtendedProperty("IsAlwaysSendNotificationAsEml") <> "True" Then
      '
      ' Create Datagram
      '
      Set loform = OpenNewForm
      If Not loform.MakeNewForm("ERROR_NOTIFY") Then
         'don't use RaiseError - Could be infinite loop....
         MsgBox "ERROR_NOTIFY Form Template missing. Cannot send Error Notification Datagram."
      End If
      With loform.DataForm.AddRecord
         .Add "Subject", subject
         .Add "TxtBody", Replace(txtBody, vbCrLf, "|")
      End With
      If Not loform.DataForm.WriteDataForm(moWorkflow.SendByMediaFolder) Then
         MsgBox "Error writing Notification Datagram. Process cancelled"
         SendNotification = False
         KillObject loform
         Exit Function
      Else
         SendNotification = True
      End If
      TearDownForm loform
      RaiseWorkflowStatus "Sent Error Notification Datagram."
   ElseIf goSession.Site.SiteType = SITE_TYPE_SHORE Or moUser.DefaultTransport = mwt_ElectronicMail Or moWorkflow.GetExtendedProperty("IsAlwaysSendNotificationAsEml") = "True" Then
   
      With Message
         .Body = txtBody
         .docID = "NOTIFICATION DATAGRAM"
         .Server = moUser.MailServerName
         .SendFromAddress = moUser.GetExtendedProperty("FromEmailAddress")
         .SendToAddress = sendTo
         .subject = subject
         If .Send() Then
            SendNotification = True
         Else
            SendNotification = False
         End If
      End With
      RaiseWorkflowStatus "Sent Error Notification Email..."
   Else
      '
      ' Do nothing... Demo loopback will create infinite loop...
      '
      'MsgBox "Demo Loopback in effect - no notification sent..."
   End If
   'Set moMessage = Nothing
   Exit Function
FunctionError:
   SendNotification = False
End Function


Public Function LogIt(msgType As mwl_LogMessageType, Criticality As mwl_LogMessageCriticality, _
  Msg As String, Optional oErr As mwError)
   On Error GoTo FunctionError
   If moLogger Is Nothing Then
      Set moLogger = New Logger
      Set moLogger.Parent = Me
   End If
   If oErr Is Nothing Then
      moLogger.LogIt msgType, Criticality, Msg
   Else
      moLogger.LogIt msgType, Criticality, Msg, oErr
   End If
   Exit Function
FunctionError:
   MsgBox "General Error in mwSession.Logit: " & Err.Number & " - " & Err.Description, vbCritical
End Function

'
' General Message event to calling program
'
Private Sub moMessage_Status(Msg As String, percent_complete As Integer, Err As Boolean)
   RaiseEvent MessageStatus(Msg, percent_complete, Err)
End Sub


Public Function CloseRecordset(ByRef oRs As Recordset, Optional IsWorkflow As Boolean)
   On Error GoTo FunctionError
   If Not oRs Is Nothing Then
      If oRs.State = adStateOpen Then
         If oRs.LockType <> adLockReadOnly Then
            If Not (oRs.EOF Or oRs.BOF) Then
               'oRs.Update
               If oRs.RecordCount > 0 Then
                  oRs.Move (0)
               End If
            End If
         End If
         oRs.Close
      End If
      Set oRs = Nothing
   End If
   Exit Function
FunctionError:
   Debug.Print "Session.CloseRecordset: Check error number -2147467259 in Oracle."
   If Err.Number = -2147467259 Then
      oRs.Cancel
      Set oRs = Nothing
   Else
      RaiseError "General Error in Session.CloseRecordset ", Err.Number, Err.Description
   End If
End Function

Public Function CloseConnection(ByRef oCon As Connection)
   On Error GoTo FunctionError
   'Exit Function
   If Not oCon Is Nothing Then
      If oCon.State = adStateOpen Then
         oCon.Close
      End If
      Set oCon = Nothing
   End If
   Exit Function
FunctionError:
   RaiseError "General Error in Session.CloseConnection ", Err.Number, Err.Description
End Function

Public Function KillObject(obj As Object)
   If Not obj Is Nothing Then
      Set obj = Nothing
   End If
End Function



Private Function CreateDbShapeConnection()
   On Error GoTo FunctionError
   Set goConShape = Nothing
   CloseConnection moDbShapeConnection
   Set moDbShapeConnection = New ADODB.Connection
   moDbShapeConnection.Provider = "MSDataShape"
   moDbShapeConnection.CursorLocation = adUseServer
   moDbShapeConnection.Open gDbShapeConnectString
   Set goConShape = moDbShapeConnection
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.CreateDbShapeConnection", Err.Number, Err.Description
End Function


Public Function OpenNewForm() As mwForm
   Set OpenNewForm = New mwForm
   Set OpenNewForm.Parent = Me
End Function

Public Function TearDownForm(ByRef TheForm As mwForm, Optional LeaveAppOpen As Boolean)
   If Not TheForm Is Nothing Then
      TheForm.TearDown LeaveAppOpen
      Set TheForm = Nothing
   End If
End Function

Public Function OpenNewZipper() As Zipper
   Set OpenNewZipper = New Zipper
   Set OpenNewZipper.Parent = Me
End Function



Public Function LoadConfigGroupsFile(ByVal strFile As String) As ConfigGroups
   Dim i As Integer
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim oGroups As ConfigGroups
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strBuffer As String
   Dim strElements() As String
   Dim strGroup As String
   Dim loEncrypt As New mwEncrypt
   Dim FirstRecord As Boolean
   Dim FileEncrypted As Boolean
   Dim strTemp As String
   
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   If oGroups Is Nothing Then
      Set oGroups = New ConfigGroups
   End If
   If Not fso.FileExists(strFile) Then
      RaiseError "Error mwSession.LoadConfigGroupsFile, file not found: " & strFile
      Set LoadConfigGroupsFile = Nothing
      KillObject fso
      Exit Function
   End If
   FirstRecord = True
   FileEncrypted = False
   Set ts = fso.OpenTextFile(strFile)
   Do While Not ts.AtEndOfStream
      strBuffer = ts.ReadLine
      
      If FirstRecord = True Then
         FirstRecord = False
         If loEncrypt.EnableEncryption(ENCRYPT_PSWD) Then
            strTemp = loEncrypt.DecryptString(strBuffer)
            If strTemp = "ENCRYPTED" Then
               strBuffer = " "
               FileEncrypted = True
            End If
         End If
      ElseIf FileEncrypted = True Then
         strBuffer = loEncrypt.DecryptString(strBuffer)
      End If
      If Len(strBuffer) = 0 Then strBuffer = " "
      strElements = Split(strBuffer, "=", , vbBinaryCompare)
      If UBound(strElements) > 1 Then
         'more than one "=" symbol found, concat them back together again...
         For i = 2 To UBound(strElements)
            strElements(1) = strElements(1) & "=" & strElements(i)
         Next i
      End If
      If Left(strElements(0), 1) = "[" Then
         ' add new config group
         strGroup = mID(strElements(0), 2, Len(strElements(0)) - 2)
         'strGroup = MassageKey(strGroup)
         Set oGroup = oGroups.Add(strGroup, strGroup)
      ElseIf Left(Trim(strElements(0)), 1) = ";" Or Left(Trim(strElements(0)), 1) = "'" Then
         ' This is a comment, ignore it.
      ElseIf UBound(strElements) > 0 Then
         If IsObject(oGroup) Then
            strElements(0) = UCase(MassageKey(strElements(0)))
            strElements(1) = MassageKey(strElements(1))
            oGroup.ConfigKeys.Add strElements(0), strElements(1), strElements(0)
         End If
      End If
   Loop
   ts.Close
   KillObject ts
   KillObject fso
   Set LoadConfigGroupsFile = oGroups
   Exit Function
FunctionError:
   If Err.Number = 91 Then
      Resume Next
   Else
      RaiseError "Error in msSession.LoadConfigGroupsFile, check mwSession.ini file.", Err.Number, Err.Description
      Set LoadConfigGroupsFile = Nothing
      KillObject ts
      KillObject fso
   End If
End Function

Public Function WriteConfigGroupsFile(ByRef oConfigGroups As ConfigGroups, ByVal strFile As String) As Boolean
   Dim i As Integer
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strBuffer As String
   Dim strElements() As String
   Dim strGroup As String
   Dim strBakFile As String
   Dim strTemp As String
   
   On Error GoTo WriteConfigGroupsFile_error
   If oConfigGroups Is Nothing Then
      WriteConfigGroupsFile = False
      Exit Function
   End If
   If fso.FileExists(strFile) Then
      'Save last instance with ".bak"
      strBakFile = strFile & ".bak"
      If fso.FileExists(strBakFile) Then fso.DeleteFile strBakFile
      fso.MoveFile strFile, strBakFile
   End If
   Set ts = fso.OpenTextFile(strFile, ForWriting, True)
   ' Outer Loop - Groups
   For Each oGroup In oConfigGroups
      strBuffer = "[" & oGroup.ConfigGroup & "]"
      ts.WriteLine strBuffer
      For Each oKey In oGroup.ConfigKeys
         
         strTemp = FixString(oKey.KeyValue)
         
         strBuffer = oKey.KeyName & " = " & strTemp
         ts.WriteLine strBuffer
      Next oKey
      ts.WriteLine ""
      DoEvents
   Next oGroup
   ts.Close
   Set ts = Nothing
   Set fso = Nothing
   WriteConfigGroupsFile = True
   Exit Function
   
WriteConfigGroupsFile_error:
   If Err.Number = 5 Then
      On Error Resume Next
      goSession.RaiseError "Error in mwsession.WriteConfigGroupsFile, Bad Unicode Data Encountered - skipping Table/Key/Column: " & _
        oGroup.ConfigKeys("table").KeyValue & "/" & oGroup.ConfigKeys("ID").KeyValue & "/" & oKey.KeyName
      Resume Next
   Else
      RaiseError "General Error in mwSession.WriteConfigGroupsFile.", Err.Number, Err.Description
      WriteConfigGroupsFile = False
      KillObject fso
      KillObject ts
   End If
End Function
Private Function FixString(strVal As String) As String
   Dim strTemp As String
   Dim ByteArray() As Byte
   Dim xx As Long
   
   ' Look for and fix UniCode characters.
   
   On Error GoTo FunctionError
   
   strTemp = Replace(strVal, vbCrLf, "|")
   strTemp = Replace(strTemp, vbLf, "|")
   
   ByteArray = strTemp
   
   For xx = 1 To UBound(ByteArray) Step 2
      If ByteArray(xx) <> 0 Then
         ByteArray(xx - 1) = 32     ' Replace uniCode character with a blank
         ByteArray(xx) = 0
      End If
      If moWorkflow.IsStripOffExtendedAsciiChars Then
         If ByteArray(xx - 1) <> 0 Then
            If ByteArray(xx - 1) >= &H80 And ByteArray(xx - 1) <= &HFF Then
               ByteArray(xx - 1) = 32
            End If
         End If
      End If
   Next
   
   FixString = ByteArray
   
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.FixString.", Err.Number, Err.Description
   FixString = strVal
End Function
Public Function WriteConfigGroupsFileXML(ByRef oConfigGroups As ConfigGroups, ByVal strFile As String) As Boolean
   Dim i As Integer
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strBuffer As String
   Dim strElements() As String
   Dim strGroup As String
   Dim strBakFile As String
   
   On Error GoTo WriteConfigGroupsFileXML_error
   If oConfigGroups Is Nothing Then
      WriteConfigGroupsFileXML = False
      Exit Function
   End If
   If fso.FileExists(strFile) Then
      'Save last instance with ".bak"
      strBakFile = strFile & ".bak"
      If fso.FileExists(strBakFile) Then fso.DeleteFile strBakFile
      fso.MoveFile strFile, strBakFile
   End If
   Set ts = fso.OpenTextFile(strFile, ForWriting, True)
   '
   ' XML Header
   '
   
   ' Outer Loop - Groups
   For Each oGroup In oConfigGroups
      strBuffer = "[" & oGroup.ConfigGroup & "]"
      ts.WriteLine strBuffer
      For Each oKey In oGroup.ConfigKeys
         strBuffer = oKey.KeyName & " = " & oKey.KeyValue
         ts.WriteLine strBuffer
      Next oKey
      ts.WriteLine ""
   Next oGroup
   ts.Close
   Set ts = Nothing
   Set fso = Nothing
   WriteConfigGroupsFileXML = True
   Exit Function
   
WriteConfigGroupsFileXML_error:
   RaiseError "General Error in mwSession.WriteConfigGroupsFileXML.", Err.Number, Err.Description
   WriteConfigGroupsFileXML = False
   KillObject fso
   KillObject ts
End Function



Private Function MassageKey(ByVal str As String) As String
   str = LTrim(str)
   str = RTrim(str)
   MassageKey = str
End Function

Private Sub moWorkflow_Status(Msg As String)
   On Error GoTo SubError
   LogIt mwl_User_Defined, mwl_Information, Msg
   RaiseEvent WorkflowStatus(Msg, 0, " ")
   Exit Sub
SubError:
   MsgBox "Unlogged Error in mwSession.moWorkflow_Status. " & Err.Description
End Sub


Public Function RaiseWorkflowStatus(Msg As String, Optional ErrNumber As Long, Optional ErrDescription As String)
   On Error GoTo FunctionError
   If ErrNumber <> 0 Then
      LogIt mwl_User_Defined, mwl_Critical, Msg & ": " & ErrNumber & " - " & ErrDescription
      RaiseEvent WorkflowStatus(Msg, ErrNumber, ErrDescription)
   Else
      LogIt mwl_User_Defined, mwl_Information, Msg
      RaiseEvent WorkflowStatus(Msg, 0, " ")
   End If
   Exit Function
FunctionError:
   MsgBox "Unlogged Error in mwSession.RaiseWorkflowStatus. " & Err.Description
End Function


Public Function ManageLicenses()
   Dim loform As frmLicensing
   On Error GoTo FunctionError
   '
   '
   If mIsLicenseServiceOff Then
      LogIt mwl_Workstation, mwl_Critical, "Licensing Service is not turned on. Cannot display Manage Licenses form."
      'MsgBox "Licensing Service is not turned on. Try rebooting to re-activate.", vbCritical
      Exit Function
   End If
   Set loform = New frmLicensing
   loform.SetParentSession Me
   If mIsLicenseKeyDisabled Then
      loform.SetIsLicenseKeyDisabled
   End If
   
   loform.Show vbModal
   Set loform = Nothing
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.ManageLicenses", Err.Number, Err.Description
End Function

Public Function ManageCurrentSite()
   Dim oForm As frmSetCurrentSite
   On Error GoTo FunctionError
   If Not mIsLoggedIn = True Then
      RaiseError "Error in mwSession.ManageCurrentSite, must be logged in."
      Exit Function
   End If
   Set oForm = New frmSetCurrentSite
   oForm.SetParentSession Me
   oForm.Show vbModal
   Set oForm = Nothing
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.ManageCurrentSite", Err.Number, Err.Description
End Function

Public Function ManageConfiguration()
   Dim oForm As frmCfgBrowser
   On Error GoTo FunctionError
   If Not mIsLoggedIn = True Then
      RaiseError "Error in mwSession.ManageConfiguration, must be logged in."
      Exit Function
   End If
   Set oForm = New frmCfgBrowser
   oForm.SetParentSession Me
   oForm.Show vbModal
   Unload oForm
   Set oForm = Nothing
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.ManageConfiguration", Err.Number, Err.Description
End Function

Public Function IsFeatureLicensed(Feature As Integer) As Boolean
   Dim oCrypt As mwCrypkey
   Dim iFeature As Integer
   On Error GoTo FunctionError
   
   If Feature = LIC_10_MarineAssurance Then
      IsFeatureLicensed = IsMALicensed()
      Exit Function
   End If
   
' LIC_10_MarineAssurance

   #If NO_CRYPKEY Then
      IsFeatureLicensed = True
      Exit Function
   #End If
   Set oCrypt = New mwCrypkey
   oCrypt.InitCrypkeyGE
   iFeature = Feature
   If oCrypt.IsOptionValid(iFeature) Then
      IsFeatureLicensed = True
   Else
      IsFeatureLicensed = False
   End If
   KillObject oCrypt
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.IsFeatureLicensed. ", Err.Number, Err.Description
   KillObject oCrypt
   IsFeatureLicensed = False
End Function
Private Function IsMALicensed() As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo FunctionError

   IsMALicensed = False

   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "SELECT VettingAuth FROM mwcThisSite", goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount > 0 Then
      If loRs!VettingAuth = "8PEAssXRuvY" Then
         IsMALicensed = True
      End If
   End If
   
   CloseRecordset loRs
   
   Exit Function
   
FunctionError:
   IsMALicensed = False

End Function

Public Function RemoteSchemaUpdater(Optional IsReplicatedSchema As Boolean) As Boolean

   Dim loWork As mwSchemaUpdateWork
   On Error GoTo FunctionError
   Set loWork = New mwSchemaUpdateWork
   RemoteSchemaUpdater = loWork.RemoteSchemaUpdater(IsReplicatedSchema)
   Set loWork = Nothing
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.RemoteSchemaUpdater. ", Err.Number, Err.Description
   Set loWork = Nothing
End Function



Friend Function UpdateExtendedKeyData(FieldName As String, FieldValue As String, SelectString As String) As Boolean
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open SelectString, DBConnection, adOpenDynamic, adLockPessimistic
   If loRs.RecordCount < 1 Then
      RaiseError "Error in mwSession.UpdateExtendedKeyData, No record found for Select: " & SelectString
      CloseRecordset loRs
      UpdateExtendedKeyData = False
   End If
   Do While Not loRs.EOF
      loRs.Fields(FieldName) = FieldValue
      loRs.Update
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   UpdateExtendedKeyData = True
  Exit Function
FunctionError:
   RaiseError "General Error in mwSession.UpdateExtendedKeyData. ", Err.Number, Err.Description
   CloseRecordset loRs
   UpdateExtendedKeyData = False
End Function


Public Function RaisePublicError(Msg As String, Optional ByVal ErrNumber As Long, Optional ByVal ErrDescription As String)
   Dim iErr As Long
   Dim strDescription As String
   Dim nRet As Long
   On Error GoTo FunctionError
   iErr = ErrNumber
   strDescription = ErrDescription
   'Msg = "Public Error Reported: " & Msg
   
   '
   ' Log it
   '
   Logger.LogIt mwl_Transport, mwl_Critical, Msg & " - " & str(iErr) & " - " & strDescription
   
   If mIsWFAProcessNow = True Then
      If IsNetworkAlive(nRet) = 0 Or ErrNumber = -2147467259 Then
         RequestExitProgram
      End If
   End If
   
   If Not Workflow.IsSilentErrors Then
      If mDisplayErrors Then
         goSession.GUI.ImprovedMsgBox Msg & " - " & iErr & " - " & strDescription, vbCritical
      Else
         RaiseEvent mwSessionError(Msg, iErr, strDescription)
         
      If DotNetFormActive = True Then
         If iErr <> 0 Then
            goSession.GUI.ImprovedMsgBox "VB6Error Encountered: " & Msg & vbCrLf & _
            iErr & " - " & strDescription, vbCritical, "ShipNet Fleet Session Error"
         Else
            goSession.GUI.ImprovedMsgBox "VB6Error Encountered: " & Msg, vbCritical
         End If
      End If
      
      End If
   Else
      RaiseEvent WorkflowStatus(Msg, ErrNumber, ErrDescription)
      moWorkflow.IsWorkflowError = True
   End If
   If Not Workflow.IsStopUpdateDBOnError Then
      If Not mIsWriting Then
         mIsWriting = True
         If WriteDbError(Msg, ErrNumber, ErrDescription) Then
            mIsWriting = False
         End If
      End If
   End If
   'Added By N.Angelakis On 12th August 2010
   'The mIsWriting flag appears to be designed to prevent recursive errors
   'from happening if the WriteDbError function gets an error.
   'however we need to catch all errors so after writting error we need to reset
   'in order to write next error
   mIsWriting = False
   
   
'   '
'   ' Log it
'   '
'   Logger.LogIt mwl_Transport, mwl_Critical, Msg & " - " & str(iErr) & " - " & strDescription
   'Set oErr = Nothing
   Exit Function
FunctionError:
   MsgBox "Error in mwSession.RaiseError. Error Encountered in the Error Handler, Bummer: " & Err.Description
End Function


Public Function GetDBConnectStrings(ConnectionID As String, ByRef DBConnectString As String, ByRef DbShapeConnectString As String) As Boolean
   
   ' Removed by T. McSally, 08/04/2004
   goSession.RaiseError "mwSession.GetDBConnectStrings is a deprecated function."
   GetDBConnectStrings = False
   
'   Dim strSQL As String
'   Dim loRsDB As Recordset
'   Dim strTemp As String
'   On Error GoTo FunctionError
'   Set loRsDB = New Recordset
'   loRsDB.CursorLocation = adUseClient
'   strSQL = "select * from mwcDbConnection where ConnectionID='" & UCase(ConnectionID) & "'"
'   loRsDB.Open strSQL, moDbConnection, adOpenForwardOnly, adLockReadOnly
'   If loRsDB.RecordCount = 1 Then
'      If IsNull(loRsDB!DBConnectString) Then
'         '
'         ' Use MW Session default...
'         '
'         DBConnectString = mDBConnectString
'         DbShapeConnectString = mDbShapeConnectString
'      Else
'         '
'         ' Use different connect strings...
'         '
'         DBConnectString = loRsDB!DBConnectString
'         strTemp = loRsDB!ShapeConnectString
'         DbShapeConnectString = strTemp
'      End If
'      GetDBConnectStrings = True
'   Else
'      GetDBConnectStrings = False
'   End If
'   CloseRecordset loRsDB
'   Exit Function
'FunctionError:
'   RaiseError "General error in mwSession.GetDbConnectStrings. ", err.Number, err.Description
'   CloseRecordset loRsDB
'   GetDBConnectStrings = False
End Function


Public Function OpenNewDataWork() As mwDataWork
   Set OpenNewDataWork = New mwDataWork
   Set OpenNewDataWork.Parent = Me
End Function

Public Function MakePK(TableName As String, Optional oCon As Connection) As Long
   Dim strSQL As String
   Dim iStart As Long
   Dim iMax As Long
   Dim loRs As Recordset
   
   On Error GoTo FunctionError

   MakePK = UsePrimaryKeySequence(TableName, oCon)
   
   If MakePK <> 0 Then
      Exit Function
   End If
   
   If goSession.ThisSite.IsSatelliteDB Then
      If goSession.ThisSite.SatelliteSeedNo > 0 Then
         iStart = goSession.ThisSite.SatelliteSeedNo * 1000000
         If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
            iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
         Else
            iMax = (goSession.ThisSite.SatelliteSeedNo + 1) * 1000000
         End If
      Else
         iStart = (moSite.SiteSeedNo * 1000000) + goSession.ThisSite.SiteSeedOffset
         If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
            iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
         Else
            iMax = iStart + 1 + 10000
         End If
      End If
   Else
      iStart = moSite.SiteSeedNo * 1000000
      If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
         iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
      Else
         iMax = (moSite.SiteSeedNo + 1) * 1000000
      End If
   End If
   
   'strSQL = "select max(ID) from " & Tablename & " where ID>=" & Str(iStart) & " And ID < " & Str(iMax)
   strSQL = "select max(ID) from " & TableName & " where ID < " & str(iMax)
   If oCon Is Nothing Then
      Set loRs = moDbConnection.Execute(strSQL)
   Else
      Set loRs = oCon.Execute(strSQL)
   End If
   If loRs.Fields(0) Is Nothing Then
      MakePK = iStart
   Else
      If loRs.Fields(0).value >= iStart Then
         MakePK = loRs.Fields(0).value + 1
      Else
         MakePK = iStart
      End If
   End If
   loRs.Close
   Set loRs = Nothing
   Exit Function
FunctionError:
   If Not mIsWriting Then
      RaisePublicError "General Error in mwSession.MakePK. ", Err.Number, Err.Description
   End If
   CloseRecordset loRs
   MakePK = -1
End Function
Public Function AddRecord(TableName As String, mwrChangeTableKey As Long, Optional mwcSitesKey As Long) As Recordset
   Dim loRs As Recordset
   Dim sSQL As String
   Dim NewID As Long
   Dim NumTrys As Integer
   
   On Error Resume Next
   
   NumTrys = 10
   
   goSession.GUI.SetCursorWait "Inserting new record"
   
   Do While NumTrys > 0
   
      NewID = MakePK(TableName)
      If NewID < 0 Then
         goSession.GUI.SetCursorNormal
         Set AddRecord = Nothing
         Exit Function
      End If
      sSQL = " SELECT * FROM " & TableName & " WHERE ID = " & NewID
      
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      
      If loRs.RecordCount = 0 Then
         loRs.AddNew
         loRs!ID = NewID
         If mwcSitesKey > 0 Then
            loRs!mwcSitesKey = mwcSitesKey
         End If
         loRs.Update
         
         If Err.Number = 0 Then
            If goSession.ReplicateWork.LogAddChange(mwrChangeTableKey, loRs!ID, mwcSitesKey, loRs.Fields) Then
               loRs.Close
               
               If goSession.IsAccess Then
                  goSession.API.SleepApp 1000
                  DoEvents
               End If
               loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
               
               Set AddRecord = loRs
               goSession.GUI.SetCursorNormal
               Exit Function
            Else
               CloseRecordset loRs
               Set AddRecord = Nothing
               goSession.GUI.SetCursorNormal
               Exit Function
            End If
         End If
         
'         If goSession.IsAccess Then
            DoEvents
            goSession.API.SleepApp 1000
            DoEvents
'         End If
         
      End If
      
      CloseRecordset loRs
      NumTrys = NumTrys - 1
   Loop
   
   goSession.GUI.SetCursorNormal
   CloseRecordset loRs
   Set AddRecord = Nothing
   RaisePublicError "Error in mwSession.AddRecord, unable to add new " & TableName & " record. "
End Function


Private Function LoadMwEventTypes() As Boolean
   Dim loRs As Recordset
   Dim loEventType As mwEventType
   On Error GoTo FunctionError
   If moEventTypes Is Nothing Then
      Set moEventTypes = New mwEventTypes
   End If
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open "mwEventType", moDbConnection, adOpenForwardOnly, adLockReadOnly, adCmdTable
   loRs.Sort = "ID"
   Do While Not loRs.EOF
      Set loEventType = moEventTypes.Add(loRs!ID, "A" & LTrim(str(loRs!ID)))
      If Not IsNull(loRs!Description) Then
         loEventType.Description = loRs!Description
      End If
      If Not IsNull(loRs!TableName) Then
         loEventType.TableName = loRs!TableName
      End If
      If Not IsNull(loRs!ImageName) Then
         loEventType.ImageName = loRs!ImageName
      End If
      If Not IsNull(loRs!AbstractTemplate) Then
         loEventType.AbstractTemplate = loRs!AbstractTemplate
      End If
      If Not IsNull(loRs!VbClassObject) Then
         loEventType.VbClassObject = loRs!VbClassObject
      End If
      If Not IsNull(loRs!VbClassObjectCustom) Then
         loEventType.VbClassObjectCustom = loRs!VbClassObjectCustom
      End If
      If Not IsNull(loRs!VbDefaultForm) Then
         loEventType.VbDefaultForm = loRs!VbDefaultForm
      End If
      If Not IsNull(loRs!FormsPlaceID) Then
         loEventType.FormsPlaceID = loRs!FormsPlaceID
      End If
      If Not IsNull(loRs!IsEnabled) Then
         loEventType.IsEnabled = loRs!IsEnabled
      End If
      If Not IsNull(loRs!IsReopenAllowed) Then
         loEventType.IsReopenAllowed = loRs!IsReopenAllowed
      End If
      If Not IsNull(loRs!IsDeleteAllowed) Then
         loEventType.IsDeleteAllowed = loRs!IsDeleteAllowed
      End If
      If Not IsNull(loRs!ActvityRemarksShortWidth) Then
         loEventType.ActvityRemarksShortWidth = loRs!ActvityRemarksShortWidth
      End If
      If Not IsNull(loRs!ActvityRemarksLongWidth) Then
         loEventType.ActvityRemarksLongWidth = loRs!ActvityRemarksLongWidth
      End If
      If Not IsNull(loRs!HelpFileID) Then
         loEventType.HelpFileID = loRs!HelpFileID
      End If
      If Not IsNull(loRs!HelpContextID) Then
         loEventType.HelpContextID = loRs!HelpContextID
      End If
      If Not IsNull(loRs!IsCreateEventAllowed) Then
         loEventType.IsCreateEventAllowed = loRs!IsCreateEventAllowed
      End If
      If Not IsNull(loRs!IsAutoCreateEvent) Then
         loEventType.IsAutoCreateEvent = loRs!IsAutoCreateEvent
      End If
      If Not IsNull(loRs!mwFactSetKey) Then
         loEventType.FactSetKey = loRs!mwFactSetKey
      End If
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   LoadMwEventTypes = True
   Exit Function
FunctionError:
   RaisePublicError "General Error in mwSession.LoadMwEventTypes. ", Err.Number, Err.Description
   CloseRecordset loRs
   LoadMwEventTypes = False
End Function


Public Function TearDownSession() As Boolean
   On Error GoTo FunctionError
   
   KillObject moErrors
   KillObject moAppKeys
   KillObject moMenus
   
   If Not moDotNetForms Is Nothing Then
      UnloadAllDotNetForms
   End If
   
   KillObject moDotNetForms
   KillObject moUser
   KillObject moSessionSettings
   KillObject moSites
   KillObject moPlaces
   KillObject moWorkflow
   KillObject moSite
   KillObject moRegistry
   
   If Not moZipper Is Nothing Then
      moZipper.TerminateClass
   End If
   
   KillObject moZipper
   KillObject moLogger
   KillObject moMailbox
   KillObject moMessage
   KillObject moContainer
   
   CloseConnection moDbConnection
   CloseConnection moDbShapeConnection
   KillObject moDbConnection
   KillObject moDbShapeConnection
   KillObject momwAPI
   KillObject goSession
   
   'KillObject moMwDocument
   Exit Function
FunctionError:
   MsgBox "Unlogged error in mwSession.TearDownSession: " & Err.Number & Err.Description
End Function

Public Function GetEventSecurityKeys(EventType As Long) As ConfigKeys
   Dim loRs As Recordset
   Dim loKeys As ConfigKeys
   Dim sSQL As String

   'Added by N.Angelakis on 16 June 2010
   'DEV-1875 Company Wide security switches
   'added filter to query to exclude All Users security profile " (mwEventProfile.IsAllUsers = 0  OR  mwEventProfile.IsAllUsers is null) "
   Static IsAllUsersCompatible As Boolean
   Dim sIsAllUsersCompatible As String
   
   On Error GoTo FunctionError
   
   'Added by N.Angelakis on 16 June 2010
   'DEV-1875 Company Wide security switches
   'added filter to query to exclude All Users security profile " (mwEventProfile.IsAllUsers = 0  OR  mwEventProfile.IsAllUsers is null) "
   If IsAllUsersCompatible = False Then IsAllUsersCompatible = GetIsAllUsersCompatible
   If IsAllUsersCompatible = True Then
      sIsAllUsersCompatible = " AND (mwEventProfile.IsAllUsers = 0  OR  mwEventProfile.IsAllUsers is null) "
   End If
   

   Set loKeys = New ConfigKeys
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventTypeSwitchType.KeyName, mwEventProfileETSwitch.KeyValue" & _
'    " FROM mwEventTypeSwitchType INNER JOIN (((mwEventProfile INNER JOIN mwEventProfileET " & _
'    " ON mwEventProfile.ID = mwEventProfileET.mwEventProfileKey) INNER JOIN mwEventProfileRole " & _
'    " ON mwEventProfile.ID = mwEventProfileRole.mwEventProfileKey) INNER JOIN mwEventProfileETSwitch " & _
'    " ON mwEventProfileET.ID = mwEventProfileETSwitch.mwEventProfileETKey) ON mwEventTypeSwitchType.ID " & _
'    " = mwEventProfileETSwitch.mwEventTypeSwitchTypeKey" & _
'    " Where mwEventProfileRole.mwcRoleTypeKey=" & moUser.RoleTypeKey & " And mwEventProfileET.mwEventTypeKey=" & _
'    EventType & " And mwEventProfileET.IsActive <> 0 And mwEventTypeSwitchType.IsActive <> 0 And mwEventProfileRole.IsActive <> 0" & _
'    " ORDER BY mwEventProfileRole.LoadOrder "

    sSQL = "SELECT mwEventTypeSwitchType.KeyName, mwEventProfileETSwitch.KeyValue " & _
      " FROM mwEventTypeSwitchType , mwEventProfile , mwEventProfileET, " & _
      "      mwEventProfileRole, mwEventProfileETSwitch " & _
      " WHERE mwEventProfile.ID = mwEventProfileET.mwEventProfileKey AND " & _
      " mwEventProfile.ID = mwEventProfileRole.mwEventProfileKey AND " & _
      " mwEventProfileET.ID = mwEventProfileETSwitch.mwEventProfileETKey AND " & _
      " mwEventTypeSwitchType.ID = mwEventProfileETSwitch.mwEventTypeSwitchTypeKey AND " & _
      " mwEventProfileRole.mwcRoleTypeKey = " & moUser.RoleTypeKey & " AND " & _
      " mwEventProfileET.mwEventTypeKey = " & EventType & " AND " & _
      " mwEventProfileET.IsActive <> 0 AND " & _
      " mwEventProfileRole.IsActive <> 0 " & _
      sIsAllUsersCompatible & _
      " ORDER BY mwEventProfileRole.LoadOrder "

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      '
      'Debug.Print loRs!KeyName
      If Not IsNull(loRs!KeyName) And Not IsNull(loRs!KeyValue) Then
         On Error Resume Next
         loKeys.Add loRs!KeyName, loRs!KeyValue, loRs!KeyName
         On Error GoTo FunctionError
      End If
      loRs.MoveNext
   Loop
   
   'Added by N.Angelakis on 16 June 2010
   'DEV-1875 Company Wide security switches
   If IsAllUsersCompatible = True Then
      GetAllUsersSecurityKeys EventType, loKeys
   End If
   
   Set GetEventSecurityKeys = loKeys
   CloseRecordset loRs
   Exit Function
FunctionError:
   If IsAllUsersCompatible = True Then
      IsAllUsersCompatible = False
      Resume Next
   Else
      RaiseError "General Error in mwSession.GetEventSecurityKeys. ", Err.Number, Err.Description
   End If
End Function
Private Function GetAllUsersSecurityKeys(EventType As Long, loKeys As ConfigKeys) As ConfigKeys
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   'DEV-1875 Company Wide security switches
   'By N.Angelakis 16 June 2010
   'implemented for company wide switches, a default company wide 'All users' profile exists and
   'individual profiles can override. If Entry is not found from individual role then we look at default 'All Users'
   
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT mwEventTypeSwitchType.KeyName, mwEventProfileETSwitch.KeyValue " & _
      " FROM      mwEventProfile, mwEventProfileET, mwEventTypeSwitchType, mwEventProfileETSwitch " & _
      " WHERE     mwEventProfile.ID = mwEventProfileET.mwEventProfileKey " & _
      " AND       mwEventProfileET.ID = mwEventProfileETSwitch.mwEventProfileETKey " & _
      " AND       mwEventTypeSwitchType.ID = mwEventProfileETSwitch.mwEventTypeSwitchTypeKey " & _
      " AND       mwEventProfileET.mwEventTypeKey = " & EventType & _
      " AND       mwEventProfileET.IsActive <> 0 " & _
      " AND       mwEventProfile.IsAllUsers <> 0 "

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      If Not IsNull(loRs!KeyName) And Not IsNull(loRs!KeyValue) Then
         On Error Resume Next
         loKeys.Add loRs!KeyName, loRs!KeyValue, loRs!KeyName
         On Error GoTo FunctionError
      End If
      loRs.MoveNext
   Loop
   
   Set GetAllUsersSecurityKeys = loKeys
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.GetAllUsersSecurityKeys. ", Err.Number, Err.Description
End Function

Public Function MemoryLogIt(sMessage As String)
   Static LastStat As Long
   Dim ThisStat As Long
   Dim TheDiff As Long
   Dim sDiff As String
   Dim StatString As String
   
   On Error GoTo FunctionError
   If moMemoryLogger Is Nothing Then
      Exit Function
   End If
   
   ThisStat = API.MemoryStats
   TheDiff = ThisStat - LastStat
   sDiff = Format(TheDiff, "###,###,###,###")
   sDiff = Left("               ", (13 - Len(sDiff))) & sDiff
   StatString = Format(ThisStat, "###,###,###,###,###") & " " & sDiff
   LastStat = ThisStat
   
   moMemoryLogger.LogIt mwl_Workstation, mwl_Information, StatString & " " & sMessage

   Exit Function
FunctionError:
   MsgBox "General Error in mwSession.Logit: " & Err.Number & " - " & Err.Description, vbCritical
End Function

Public Sub MemoryLogFile(LogFile As String)
   On Error GoTo SubError
   If moMemoryLogger Is Nothing Then
      Set moMemoryLogger = New Logger
      Set moMemoryLogger.Parent = Me
   End If
   
   moMemoryLogger.LogFile = LogFile
   
   Exit Sub
SubError:
   MsgBox "General Error in mwSession.MemoryLogFile: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Function WriteDbError(Msg As String, ErrNumber As Long, ErrMsg As String) As Boolean
   On Error GoTo FunctionError
   If Site.SiteKey > 0 And User.UserKey > 0 Then
      Set moRsErrorLog = AddRecord("mwErrorLog", MWRT_moRsErrorLog, Site.SiteKey)
      If moRsErrorLog Is Nothing Then
         Exit Function
      End If
      moRsErrorLog!mwcSitesKey = Site.SiteKey
      moRsErrorLog!mwcUsersKey = User.UserKey
      moRsErrorLog!Msg = Left(Msg, 200)
      moRsErrorLog!ErrNumber = ErrNumber
      moRsErrorLog!ErrDescription = Left(ErrMsg, 200)
      moRsErrorLog!LogDateTime = Now()
      moRsErrorLog!WorkstationName = Left(API.GetMachName, 100)
      moRsErrorLog!BuildNumber = Val(mID(Format(App.Revision, "0000"), 2, 2))
      moRsErrorLog!AppVersion = App.Major & "." & App.Minor & "." & Format(App.Revision, "0000")
      moRsErrorLog.Update
      CloseRecordset moRsErrorLog
   End If
   Exit Function
FunctionError:
   Logger.LogIt mwl_Transport, mwl_Critical, "General Error in mwSession.WriteDbError: " & Err.Number & Err.Description
End Function

Private Sub moRsErrorLog_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As mwReplicateWillChange
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = New mwReplicateWillChange
      If Not loWork.Initialize("mwErrorLog") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset
   Exit Sub
SubError:
   Logger.LogIt mwl_Transport, mwl_Critical, "General Error in mwSession.moRsErrorLog_WillChangeRecord: " & Err.Number & Err.Description
End Sub

'
' for shore consistency among modules this semi-fixed user supplied constant
'
Public Property Get SiteNameGlobal() As String
   SiteNameGlobal = mSiteNameGlobal
End Property
Public Property Let SiteNameGlobal(ByVal sData As String)
   mSiteNameGlobal = sData
End Property

Public Property Get SiteKeyGlobal() As String
   SiteKeyGlobal = mSiteKeyGlobal
End Property
Public Property Let SiteKeyGlobal(ByVal sData As String)
   mSiteKeyGlobal = ZeroNull(sData)
End Property

Public Property Get ApplySiteKeyGlobalOverride() As String
   ApplySiteKeyGlobalOverride = mApplySiteKeyGlobalOverride
End Property
Public Property Let ApplySiteKeyGlobalOverride(ByVal sData As String)
   mApplySiteKeyGlobalOverride = BoolNull(sData)
End Property

Public Property Get SiteFleetNameGlobal() As String
   SiteFleetNameGlobal = mSiteFleetNameGlobal
End Property
Public Property Let SiteFleetNameGlobal(ByVal sData As String)
   mSiteFleetNameGlobal = sData
End Property

Public Property Get SiteFleetKeyGlobal() As String
   SiteFleetKeyGlobal = mSiteFleetKeyGlobal
End Property
Public Property Let SiteFleetKeyGlobal(ByVal sData As String)
   mSiteFleetKeyGlobal = ZeroNull(sData)
End Property

Private Function GetIsAllUsersCompatible() As Boolean
   'Added by N.Angelakis on 16 June 2010
   'DEV-1875 Company Wide security switches
   
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT count(IsAllUsers) FROM mwEventProfile"

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   GetIsAllUsersCompatible = True

   CloseRecordset loRs
   
   Exit Function
FunctionError:
   GetIsAllUsersCompatible = False
End Function

'-----------------------

Public Function GetToolTipText(sKeyName As String) As String
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   ' control or grid level column tooltiphelp (KeyName=Project.Form.ControlName)
   
   sSQL = "SELECT ToolTipText FROM mwcHelpToolTIp WHERE KeyName LIKE '" & sKeyName & "'"
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If IsRecordLoaded(loRs) Then
      If Not IsNull(loRs!ToolTipText) Then
         GetToolTipText = loRs!ToolTipText
      End If
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.Session.GetToolTipText. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Property Get GetLogPath() As String
   Dim sLogFile As String
   Dim objFso As FileSystemObject
   
   If Not moLogger Is Nothing Then
      Set objFso = New FileSystemObject
      GetLogPath = objFso.GetParentFolderName(moLogger.LogFile())
      Set objFso = Nothing
   Else
      GetLogPath = App.Path
   End If
End Property

Private Function UsePrimaryKeySequence(TableName As String, Optional oCon As Connection) As Long
   Dim sSQL As String
   Dim loRs As Recordset

   On Error GoTo FunctionError

   UsePrimaryKeySequence = 0
   
   If goSession.IsOracle Then
      sSQL = "SELECT UsePkIncrementTable FROM mwrChangeTable WHERE UPPER(TableName) = '" & UCase(TableName) & "'"
   Else
      sSQL = "SELECT UsePkIncrementTable FROM mwrChangeTable WHERE TableName = '" & TableName & "'"
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   If oCon Is Nothing Then
      loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Else
      loRs.Open sSQL, oCon, adOpenDynamic, adLockOptimistic
   End If
   
   If loRs.RecordCount > 0 Then
      If BoolNull(loRs!UsePkIncrementTable) = True Then
         UsePrimaryKeySequence = MakePrimaryKeySequence(TableName, oCon)
      End If
   End If
   
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Session.UsePrimaryKeySequence. ", Err.Number, Err.Description

End Function


Private Function MakePrimaryKeySequence(TableName As String, Optional oCon As Connection) As Long
   Dim sSQL As String
   Dim loRsPK As Recordset
   Dim loRs As Recordset
   Dim iStart As Long
   Dim iMax As Long
   Dim nSiteSeedNo As Long
   Dim nThisSeedNo As Long
   On Error GoTo FunctionError

   MakePrimaryKeySequence = 0
   
'   Private Const SeedFactor = 1000000
'   Private Const Offset = 10000

   If goSession.ThisSite.IsSatelliteDB Then
      If goSession.ThisSite.SatelliteSeedNo > 0 Then
         nSiteSeedNo = goSession.ThisSite.SatelliteSeedNo
      Else
         nSiteSeedNo = goSession.Site.SiteSeedNo
      End If
   Else
      nSiteSeedNo = goSession.Site.SiteSeedNo
   End If

   If goSession.IsOracle Then
      sSQL = "SELECT * FROM mwrPrimaryKeySequence WHERE UPPER(TableName) = '" & UCase(TableName) & "'"
   Else
      sSQL = "SELECT * FROM mwrPrimaryKeySequence WHERE TableName = '" & TableName & "'"
   End If
   
   Set loRsPK = New Recordset
   loRsPK.CursorLocation = adUseClient
   
   If oCon Is Nothing Then
      loRsPK.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Else
      loRsPK.Open sSQL, oCon, adOpenDynamic, adLockOptimistic
   End If
   
   If loRsPK.RecordCount > 0 Then
   
      nThisSeedNo = ZeroNull(loRsPK!NextValue)
      
      If nThisSeedNo <> 0 Then
         nThisSeedNo = nThisSeedNo / 1000000
      End If
      
      If nThisSeedNo <> nSiteSeedNo Then
         MakePrimaryKeySequence = GetMaxKey(TableName, oCon) + 1
      Else
         MakePrimaryKeySequence = ZeroNull(loRsPK!NextValue)
      End If
      
      loRsPK!NextValue = MakePrimaryKeySequence + 1
      loRsPK.Update
      CloseRecordset loRsPK
      Exit Function
   Else
      MakePrimaryKeySequence = GetMaxKey(TableName, oCon) + 1
      
      loRsPK.AddNew
      loRsPK!TableName = TableName
      loRsPK!NextValue = MakePrimaryKeySequence + 1
      loRsPK.Update
   
   End If
   
   CloseRecordset loRsPK
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Session.MakePrimaryKeySequence. ", Err.Number, Err.Description

End Function

Public Function UpdatePrimaryKeySequence(TableName As String, NextValue As Long) As Boolean
   Dim sSQL As String
   Dim loRsPK As Recordset
   On Error GoTo FunctionError

   If goSession.IsOracle Then
      sSQL = "SELECT * FROM mwrPrimaryKeySequence WHERE UPPER(TableName) = '" & UCase(TableName) & "'"
   Else
      sSQL = "SELECT * FROM mwrPrimaryKeySequence WHERE TableName = '" & TableName & "'"
   End If
   
   Set loRsPK = New Recordset
   loRsPK.CursorLocation = adUseClient
   loRsPK.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If loRsPK.RecordCount > 0 Then
      loRsPK!NextValue = NextValue
      loRsPK.Update
   End If
   
   CloseRecordset loRsPK
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Session.UpdatePrimaryKeySequence. ", Err.Number, Err.Description

End Function

Private Function GetMaxKey(TableName As String, Optional oCon As Connection) As Long
   Dim sSQL As String
   Dim loRs As Recordset
   Dim iStart As Long
   Dim iMax As Long
   
   On Error GoTo FunctionError

   GetMaxKey = 0
   
   If goSession.ThisSite.IsSatelliteDB Then
      If goSession.ThisSite.SatelliteSeedNo > 0 Then
         iStart = goSession.ThisSite.SatelliteSeedNo * 1000000
         If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
            iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
         Else
            iMax = (goSession.ThisSite.SatelliteSeedNo + 1) * 1000000
         End If
      Else
         iStart = (goSession.Site.SiteSeedNo * 1000000) + goSession.ThisSite.SiteSeedOffset
         If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
            iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
         Else
            iMax = iStart + 1 + 10000
         End If
      End If
   Else
      iStart = goSession.Site.SiteSeedNo * 1000000
      If goSession.ThisSite.SiteSeedMaxOffset > 0 Then
         iMax = iStart + goSession.ThisSite.SiteSeedMaxOffset
      Else
         iMax = (goSession.Site.SiteSeedNo + 1) * 1000000
      End If
   End If
   
   'ssql = "select max(ID) from " & Tablename & " where ID>=" & Str(iStart) & " And ID < " & Str(iMax)
   sSQL = "select max(ID) from " & TableName & " where ID < " & str(iMax)
      
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   If oCon Is Nothing Then
      loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Else
      loRs.Open sSQL, oCon, adOpenDynamic, adLockOptimistic
   End If
   
   If loRs.RecordCount > 0 Then
      
      If loRs.Fields(0) Is Nothing Then
         GetMaxKey = iStart
      Else
         If loRs.Fields(0).value >= iStart Then
            GetMaxKey = loRs.Fields(0).value
         Else
            GetMaxKey = iStart
         End If
      End If
   Else
      GetMaxKey = iStart
   End If
   
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Session.GetMaxKey. ", Err.Number, Err.Description

End Function


Public Function ShowFleetForm(strModuleName As String, strFormName As String) As Long
   Dim loFormObject As Object
   Dim loFormSession As Object
   
   Dim sVbDLL As String
   On Error GoTo FunctionError
   
   '
   'Visual Basic Exrternal DLL
   '
   sVbDLL = strModuleName & ".FormSession"
   
   Set loFormSession = CreateObject(sVbDLL)
   
   Set loFormSession.ActiveMwSession = goSession
   
   
   On Error GoTo FunctionError
   '
   Set loFormObject = loFormSession.LoadCustomForm(strFormName)
    
   ChangeFormBorder loFormObject, vbBSNone
          
   ShowFleetForm = loFormObject.hwnd

   Exit Function

FunctionError:
   MsgBox "Error " & Err.Number & " " & Err.Description
   ShowFleetForm = 0
End Function

Public Function ShowVB6FormFromMenuItem(mwcMenuItemsKey As Long) As Long
   
   Dim loMenuItem As mwMenuItem
   Dim nKey As Long
   
   On Error GoTo FunctionError
   
   DotNetFormActive = True
   
   Set loMenuItem = CreateMenuItem(mwcMenuItemsKey)
'   If loMenuItem.IsExists Then
'      'Show the form...
'      loMenuItem.FormHandle.Show
'      frmWFD.MousePointer = fmMousePointerDefault
'      Exit Sub
'   End If
   '
   ' Aug 2011 - WIN32 Launch Intercept
   If loMenuItem.FormType = "WIN32" Then
      
      goSession.API.RunShellExecute loMenuItem.GetExtendedProperty("Win32ApiLaunch")
'      frmWFD.MousePointer = fmMousePointerDefault
      ShowVB6FormFromMenuItem = 0
      Exit Function
   End If
   If Not CreatePVForm(loMenuItem) Then
'      frmWFD.MousePointer = fmMousePointerDefault
      ShowVB6FormFromMenuItem = 0
      Exit Function
   End If
'   loMenuItem.FormHandle.Show

   ChangeFormBorder loMenuItem.FormHandle, vbBSNone
          
   ShowVB6FormFromMenuItem = loMenuItem.FormHandle.hwnd

'   MsgBox "ShowVB6FormFromMenuItem 1 "

   moDotNetForms.Add loMenuItem.FormHandle, loMenuItem.FormHandle.hwnd
   
'   MsgBox "ShowVB6FormFromMenuItem 2 "
   Exit Function

FunctionError:
   MsgBox "Error " & Err.Number & " " & Err.Description
   ShowVB6FormFromMenuItem = 0
End Function

'Public Function UnloadVB6Form(VB6Form As Object) As Boolean
'   On Error GoTo FunctionError
'
'   Unload VB6Form
'
'   Exit Function
'
'FunctionError:
'   MsgBox "Error in UnloadVB6Form " & Err.Number & " " & Err.Description
'
'End Function
Public Function CreateMenuItem(mwcMenuItemsKey As Long) As mwMenuItem
   Dim oRsMenuItems As Recordset
   Dim oItem As mwMenuItem
   Dim loField As field
   Dim strSQL As String
   ' Possibly Null Guys
   Dim strFormID As String
   
   On Error GoTo FunctionError
   
   '
   ' Loop through mwMenus for the user
   '
   Set oRsMenuItems = New Recordset
   oRsMenuItems.CursorLocation = adUseClient
   strSQL = "select * from mwcMenuItems WHERE ID = " & mwcMenuItemsKey
   oRsMenuItems.Open strSQL, moDbConnection, adOpenStatic, adLockReadOnly
            
   If oRsMenuItems.EOF = False Then
      '
      ' 3/2002 Extended Properties
      '
      If IsNull(oRsMenuItems.Fields("FormID")) Then
         strFormID = " "
      Else
         strFormID = oRsMenuItems.Fields("FormID")
      End If
      '
      ' Create Menu Item object
      '
      Set oItem = New mwMenuItem
      
'      Set oItem = oMenu.MenuItems.Add(strFormID)
      oItem.FormID = strFormID
      Set oItem.Session = Me
      
      For Each loField In oRsMenuItems.Fields
         Select Case UCase(loField.Name)
            
            Case Is = UCase("FormID")
            Case Is = UCase("DateTimeCreated")
            Case Is = UCase("DateTimeModified")
            Case Is = UCase("DisplayTitle")
               If IsNull(oRsMenuItems.Fields("DisplayTitle")) Then
                  oItem.DisplayTitle = ""
               Else
                  oItem.DisplayTitle = oRsMenuItems.Fields("DisplayTitle")
               End If
            Case Is = UCase("ItemTitle")
               If Not IsNull(oRsMenuItems.Fields("ItemTitle")) Then
                  oItem.ItemTitle = oRsMenuItems.Fields("ItemTitle")
               End If
            Case Is = UCase("ImageName")
               If Not IsNull(oRsMenuItems.Fields("ImageName")) Then
                  oItem.ImageReference = oRsMenuItems.Fields("ImageName")
               End If
            Case Is = UCase("FormType")
               If Not IsNull(oRsMenuItems.Fields("FormType")) Then
                  oItem.FormType = oRsMenuItems.Fields("FormType")
               End If
            Case Is = UCase("AutoLoad")
               If IsNull(loField.value) Then
                  oItem.AutoLoad = False
               Else
                  oItem.AutoLoad = oRsMenuItems.Fields("AutoLoad")
               End If
            Case Is = UCase("PendingPlaceID")
               If IsNull(oRsMenuItems.Fields("PendingPlaceID")) Then
                  oItem.FormTemplate.PendingFolder = " "
               Else
                  oItem.FormTemplate.PendingFolder = oRsMenuItems.Fields("PendingPlaceID")
               End If
            Case Is = UCase("SubmittedPlaceID")
               If IsNull(oRsMenuItems.Fields("SubmittedPlaceID")) Then
                  oItem.FormTemplate.SubmittedFolder = " "
               Else
                  oItem.FormTemplate.SubmittedFolder = oRsMenuItems.Fields("SubmittedPlaceID")
               End If
            Case Is = UCase("ArchivedPlaceID")
               If IsNull(oRsMenuItems.Fields("ArchivedPlaceID")) Then
                  oItem.FormTemplate.ArchivedFolder = " "
               Else
                  oItem.FormTemplate.ArchivedFolder = oRsMenuItems.Fields("ArchivedPlaceID")
               End If
            Case Is = UCase("TemplatePlaceID")
               If IsNull(oRsMenuItems.Fields("TemplatePlaceID")) Then
                  oItem.FormTemplate.TemplateFolder = " "
               Else
                  oItem.FormTemplate.TemplateFolder = oRsMenuItems.Fields("TemplatePlaceID")
               End If
            Case Is = UCase("FilePattern")
               If IsNull(oRsMenuItems.Fields("FilePattern")) Then
                  oItem.FormTemplate.FilePattern = " "
               Else
                  oItem.FormTemplate.FilePattern = oRsMenuItems.Fields("FilePattern")
               End If
            Case Is = UCase("SendToAddress")
               If IsNull(oRsMenuItems.Fields("SendToAddress")) Then
                  oItem.FormTemplate.SendToAddress = " "
               Else
                  oItem.FormTemplate.SendToAddress = oRsMenuItems.Fields("SendToAddress")
               End If
            Case Is = UCase("ActiveConfigGroup")
               If IsNull(oRsMenuItems.Fields("ActiveConfigGroup")) Then
                  oItem.FormTemplate.ActiveConfigGroup = " "
               Else
                  oItem.FormTemplate.ActiveConfigGroup = oRsMenuItems.Fields("ActiveConfigGroup")
               End If
            Case Is = UCase("ID")
               oItem.MenuItemKey = oRsMenuItems!ID
            Case Else
               '
               ' Otherwise, Add to the Collection of configkeys... (extended properties)
               '
               If IsNull(loField.value) Then
                  oItem.ExtendedKeys.Add UCase(loField.Name), "", UCase(loField.Name)
               Else
                  If goSession.IsOracle And loField.Type = adNumeric And loField.Precision = 1 Then
                     If loField.value = 0 Then
                        oItem.ExtendedKeys.Add UCase(loField.Name), "False", UCase(loField.Name)
                     Else
                        oItem.ExtendedKeys.Add UCase(loField.Name), "True", UCase(loField.Name)
                     End If
                  Else
                     oItem.ExtendedKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
                  End If
               End If
               
         End Select
      Next
   End If
   '
   ' Close up shop
   '
   CloseRecordset oRsMenuItems
   
   Set CreateMenuItem = oItem
   Exit Function
FunctionError:
   'Resume Next
   RaiseError "General Error in mwSession.CreateMenuItem: ", Err.Number, Err.Description
   CloseRecordset oRsMenuItems
End Function

Private Function CreatePVForm(ByRef oMenuItem As mwMenuItem) As Boolean
   Dim loFormObject As Object
   Dim loFormSession As Object
   
   Dim sVbDLL As String
   On Error GoTo FunctionError
   'Create the form
   If oMenuItem.FormType = MW_CONTAINER_FORM Then
'      '
'      'Intrinsic msWorkstation Form: Type "F" (only WF Agent left...)
'      '
'      oMenuItem.FormHandle = CreateCustomForm(oMenuItem)
'      oMenuItem.IsExists = True
'      Item.WindowObject1 = oMenuItem.FormHandle.hwnd
'      If Not oMenuItem.AutoLoad Then
'         oMenuItem.FormHandle.Show
'      End If
'      frmWFD.MousePointer = fmMousePointerDefault
   Else
      '
      ' Instrinsic Reusable forms...
      '
      Select Case oMenuItem.FormType
'         Case Is = MW_CONTAINER_OFFICE
'            'Office Component
'            Set loFormObject = New frmOfficeMain
'            SetFolders loFormObject, oMenuItem
'            ' MenuItemKey links forms to menu items...
'            loFormObject.SetMenuItem oMenuItem
'         Case Is = MW_CONTAINER_MW_EVENTS
'            'Voyage Events Component
'            Set loFormObject = New frmMwEvents
'            ' MenuItemKey links forms to menu items...
'            loFormObject.SetMenuItem oMenuItem
'            'ActiveConfigGroup contains the "vrsEventGroup" record for this module
'            loFormObject.InitForm
'         Case Is = MW_CONTAINER_WF_EVENTS
'            'Voyage Events Component
'            Set loFormObject = New frmMwEventsWF
'            ' MenuItemKey links forms to menu items...
'            loFormObject.SetMenuItem oMenuItem
'            'ActiveConfigGroup contains the "vrsEventGroup" record for this module
'            loFormObject.InitForm
         Case Is = MW_CONTAINER_VBDLL
            '
            'Visual Basic Exrternal DLL
            '
            sVbDLL = oMenuItem.GetExtendedProperty("VbClassObject")
            If Trim(sVbDLL) = "" Then
               MsgBox "Error in mwSession.Session.CreatePVForm, Configuration Error. " & vbCrLf & _
                 "A menu item has been defined to use an extrnal form, but no 'VbClassObject' has been defined."
               CreatePVForm = False
               Exit Function
            End If
            sVbDLL = sVbDLL & ".FormSession"
            Set loFormSession = CreateObject(sVbDLL)
            Set loFormSession.ActiveMwSession = goSession
            '
            ' VBDLL may need to MenuItem object for parameters...
            '
            On Error Resume Next
            loFormSession.SetMenuItem oMenuItem
            On Error GoTo FunctionError
            '
            If Trim(oMenuItem.GetExtendedProperty("VbDefaultForm")) <> "" Then
               Set loFormObject = loFormSession.LoadCustomForm(oMenuItem.GetExtendedProperty("VbDefaultForm"))
            Else
               Set loFormObject = loFormSession.LoadCustomForm
            End If
            '
            ' Cast form from generic object...
            '
'            Set loFormObject = moObject

            
'         Case Is = MW_CONTAINER_VISIO
'            'Visio Form
'            Set loFormObject = New frmVisioMain
'            SetFolders loFormObject, oMenuItem
'         Case Is = MW_CONTAINER_CRYSTAL
'            'Crystal Report form type
'            Set loFormObject = New frmCrystalMain
'            SetFolders loFormObject, oMenuItem
'         Case Is = MW_CONTAINER_MANUAL
'            'Crystal Report form type
'            Set loFormObject = New frmManualViewer
'         Case Is = MW_CONTAINER_FLUKE
'            'Crystal Report form type
'            Set loFormObject = New frmFluke
         Case Else
            MsgBox "Error in mwSession.Session.CreatePVForm, invalid  FormType: " & oMenuItem.FormType
            CreatePVForm = False
            Exit Function
       End Select
      oMenuItem.FormHandle = loFormObject
      ' Used to save document lists...
      loFormObject.TAG = oMenuItem.FormID
      'loFormObject.lblTitle = oMenuItem.FormTemplate.DisplayTitle
      If oMenuItem.FormType <> MW_CONTAINER_VBDLL And oMenuItem.FormType <> MW_CONTAINER_MW_EVENTS _
        And oMenuItem.FormType <> MW_CONTAINER_WF_EVENTS And oMenuItem.FormType <> MW_CONTAINER_WF_EXPLORE Then
         ' MW Proforma...
         loFormObject.lblTitle = oMenuItem.DisplayTitle
         '
         loFormObject.SetTemplate goSession.GetPlace(oMenuItem.FormTemplate.TemplateFolder, _
           mw_DriveLetterPath), _
         oMenuItem.FormTemplate.FilePattern, oMenuItem.FormTemplate.ActiveConfigGroup
         loFormObject.SetSendtoAddress oMenuItem.FormTemplate.SendToAddress
      End If
      '
      ' Function that prepares form for use by user
      '
'      Item.WindowObject1 = oMenuItem.FormHandle.hwnd
      'oMenuItem.FormHandle.Show
   End If
   '
   oMenuItem.IsExists = True
'   frmWFD.MousePointer = fmMousePointerDefault
   CreatePVForm = True
   Exit Function

FunctionError:

   If loFormObject Is Nothing Then
      MsgBox "Could not load specified Menu Item."
   Else
      MsgBox "Error in mwSession.Session.CreatePVForm. ", Err.Number, Err.Description
   End If
   If Not loFormObject Is Nothing Then
      Unload loFormObject
      Set loFormObject = Nothing
   End If
   CreatePVForm = False
End Function

Private Function SetFolders(oForm As Form, oItem As mwMenuItem)
   Dim strPlace As String
   Dim fso As FileSystemObject
   Dim loTools As Object

   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   
   Set loTools = CreateObject("mwUtility.mwTools")
   loTools.InitSession goSession
   
   
   ' 8/8/2001 ms - Place object always used...
   '
   ' Submitted
   '
   If Trim(oItem.FormTemplate.PendingFolder) = "" Then
      'Disable tab
      oForm.SSTab1.TabVisible(0) = False
   Else
      loTools.MakeFolderExist goSession.GetPlace(oItem.FormTemplate.PendingFolder, mw_DriveLetterPath)
   
      If fso.FolderExists(goSession.GetPlace(oItem.FormTemplate.PendingFolder, mw_DriveLetterPath)) Then
         oForm.flbPending.Path = goSession.GetPlace(oItem.FormTemplate.PendingFolder, mw_DriveLetterPath)
      Else
         goSession.RaisePublicError "Error in mwSession.Session.SetFolders, Pending Folder does not exist: " & oItem.FormTemplate.PendingFolder
         oForm.SSTab1.TabVisible(0) = False
      End If
   End If
   
   ' 8/8/2001 ms - Place object always used...
   '
   ' Submitted
   '
   If Trim(oItem.FormTemplate.SubmittedFolder) = "" Then
      'Disable tab
      oForm.SSTab1.TabVisible(1) = False
   Else
      loTools.MakeFolderExist goSession.GetPlace(oItem.FormTemplate.SubmittedFolder, mw_DriveLetterPath)
      
      If fso.FolderExists(goSession.GetPlace(oItem.FormTemplate.SubmittedFolder, mw_DriveLetterPath)) Then
         oForm.flbSubmitted.Path = goSession.GetPlace(oItem.FormTemplate.SubmittedFolder, mw_DriveLetterPath)
      Else
         goSession.RaisePublicError "Error in mwSession.Session.SetFolders, Submitted Folder does not exist: " & oItem.FormTemplate.SubmittedFolder
         oForm.SSTab1.TabVisible(1) = False
      End If
   End If
   

   ' 8/8/2001 ms - Place object always used...
   '
   ' Archived
   '
   If Trim(oItem.FormTemplate.ArchivedFolder) = "" Then
      'Disable tab
      oForm.SSTab1.TabVisible(2) = False
   Else
      loTools.MakeFolderExist goSession.GetPlace(oItem.FormTemplate.ArchivedFolder, mw_DriveLetterPath)
      
      If fso.FolderExists(goSession.GetPlace(oItem.FormTemplate.ArchivedFolder, mw_DriveLetterPath)) Then
         oForm.flbArchived.Path = goSession.GetPlace(oItem.FormTemplate.ArchivedFolder, mw_DriveLetterPath)
      Else
         goSession.RaisePublicError "Error in mwSession.Session.SetFolders, Archived Folder does not exist: " & oItem.FormTemplate.ArchivedFolder
         oForm.SSTab1.TabVisible(2) = False
      End If
   End If

   ' File Pattern...
   oForm.flbPending.Pattern = oItem.FormTemplate.FilePattern
   oForm.flbSubmitted.Pattern = oItem.FormTemplate.FilePattern
   oForm.flbArchived.Pattern = oItem.FormTemplate.FilePattern
   KillObject loTools
   KillObject fso
   Exit Function
FunctionError:
   MsgBox "Error in mwSession.Session.SetFolders.", Err.Number, Err.Description
   KillObject loTools
   KillObject fso

End Function

Public Sub UnloadAllDotNetForms()
   Dim oDotNetForm As mwDotNetForm
   
   '
   ' Unload forms that have been created...
   '
   If moDotNetForms.Count <> 0 Then
      For Each oDotNetForm In moDotNetForms
         On Error Resume Next   'in case forms were not created...
         Unload oDotNetForm.FormHandle
         moDotNetForms.Remove moDotNetForms.MakeFormHandleKey(oDotNetForm.VB6FormKey)
      Next
   End If

End Sub

Public Sub UnloadDotNetForm(VB6FormKey As Long)
   Dim oDotNetForm As mwDotNetForm
   Dim FormHandleKey As String
      Dim nKey As Long

   '
   ' Unload forms that have been created...
   '
'   MsgBox "UnloadDotNetForm 1"
   FormHandleKey = moDotNetForms.MakeFormHandleKey(VB6FormKey)
   
   If moDotNetForms.Count <> 0 Then
      Set oDotNetForm = moDotNetForms(FormHandleKey)
         
      On Error Resume Next   'in case forms were not created...
'   nKey = FormHandleKey
'   MsgBox "UnloadDotNetForm 2 Key = " & nKey
      Unload oDotNetForm.FormHandle
      moDotNetForms.Remove FormHandleKey
   End If
'   MsgBox "UnloadDotNetForm 3"

End Sub

'
' Login to Database Connection. If LoginID is ommitted, use registry value for current user...
'
Public Function VB6Login( _
  Optional MwUserID As String, _
  Optional MwPassword As String, _
  Optional DBConnectString As String, _
  Optional DbShapeConnectString As String, _
  Optional EmployeeID As String) As Boolean
   Dim loDW As mwDataWork
   Dim oLoginForm As frmLogin
   Dim sFullComputerName As String
   Dim sDomainName As String
   On Error GoTo FunctionError
   
   DotNetFormActive = True
   
   If Trim(Environ("MWS_DEV_DEBUG")) <> "" Then
      mdebugmode = True
   Else
      mdebugmode = False
   End If
   Set oLoginForm = New frmLogin
   'oLoginForm.SetLoginParameters MwUserID, MwPassword, DBConnectString, _
     DbShapeConnectString, EmployeeID
   oLoginForm.SetParent Me
   '
   
   If Not oLoginForm.SetLoginParameters(MwUserID, MwPassword, DBConnectString, _
     DbShapeConnectString, EmployeeID) Then
      VB6Login = False
      Unload oLoginForm
      Exit Function
   End If
   
   ' Load VB6Login info from registry and mwSession.ini, as required.
   If Not oLoginForm.DetermineLoginParameters() Then
      VB6Login = False
      Unload oLoginForm
      Exit Function
   End If
   
   ' Prevalidate User, or Show the form for logging in...
   If Not oLoginForm.ValidateUser() Then
      VB6Login = False
      Unload oLoginForm
      Exit Function
      
   End If
   
   If Not oLoginForm Is Nothing Then
      Unload oLoginForm
   End If

   ' OK, we be logged in...
   '
   FinishDBConnection
   
'   Set loDW = New mwDataWork
'   gIsSqlServer = loDW.IsSqlServer()
'   gIsOracle = loDW.IsOracle()
'   gIsAccess = loDW.IsAccess()
'
'   Set loDW = Nothing
   
   LoadAppKeys
   LoadPlaces
'   If moSite Is Nothing Then
'      Set moSite = New mwSite
'      Set moSite.Parent = Me
'   End If
'   moSite.LoadSiteKeys
   '
   ' CrypKey validation..
   '
   #If NO_CRYPKEY Then
      mIsLicensed = True
      LogIt mwl_Workstation, mwl_Information, "Logging in with Licensing Service Disabled."
   #Else
      If ValidateCrypKey() Then
         mIsLicensed = True
      Else
         mIsLicensed = False
      End If
   #End If
   gIsAllowWin9xDelay = False
   gWin9xMilliseconds = 0
   
   sFullComputerName = API.GetMachName
   sDomainName = API.GetLongDomainName
   If sDomainName <> "" Then
      sFullComputerName = sFullComputerName & "." & sDomainName
   End If
   VB6Login = True
   mIsLoggedIn = True
   Workflow.ReceiveRoleTypeNotification
   Logger.LogIt mwl_Workstation, mwl_Information, "mwSession Login, Version: " & _
     App.Major & "." & App.Minor & "." & Left(Format(App.Revision, "0000"), 1) & " Build " & CStr(Val(mID(Format(App.Revision, "0000"), 2, 2))) & " , Workstation OS: " & API.WindowsVersion & ", User: " _
   & goSession.User.UserID & ", Site: " & goSession.Site.SiteID & ", Machine: " & sFullComputerName & ", OS User: " & API.GetOSUserName
   
   DotNetFormActive = True

   Exit Function
FunctionError:
   RaiseError "General Error in mwSession.VB6Login. ", Err.Number, Err.Description
   VB6Login = False
End Function


Public Property Let DotNetBackColor(ByVal sDotNetBackColor As String)
   mDotNetBackColor = HtmlToVBHex(sDotNetBackColor)
End Property

Public Property Get DotNetBackColor() As String
   DotNetBackColor = mDotNetBackColor
End Property

Public Property Let DotNetFontName(ByVal sDotNetFontName As String)
   mDotNetFontName = sDotNetFontName
End Property

Public Property Let DotNetFontSize(ByVal sDotNetFontSize As String)
   mDotNetFontSize = sDotNetFontSize
End Property

Public Property Get IsDotNetFormActive() As Boolean
   IsDotNetFormActive = DotNetFormActive
End Property

Public Sub SetDotNetTheme(FormInstance As Object)
   On Error Resume Next
   Dim ctlInstance As Control

   If DotNetFormActive = False Then
      Exit Sub
   End If

   FormInstance.BackColor = mDotNetBackColor

   For Each ctlInstance In FormInstance.Controls
      If TypeOf ctlInstance Is TextBox Then
         ctlInstance.FontName = mDotNetFontName
'           ctlInstance.FontSize = mDotNetFontSize
      ElseIf TypeOf ctlInstance Is Frame Then
         ctlInstance.BackColor = mDotNetBackColor
      ElseIf TypeOf ctlInstance Is SSUltraGrid Then
         ctlInstance.Appearance.BackColor = mDotNetBackColor
         ctlInstance.Appearance.Font.Name = mDotNetFontName
         ctlInstance.Override.HeaderAppearance.BackColor = mDotNetBackColor
         ctlInstance.CaptionAppearance.BackColor = mDotNetBackColor
 '          ctlInstance.Appearance.Font.Size = mDotNetFontSize
      Else
         ctlInstance.BackColor = mDotNetBackColor
         ctlInstance.Font.Name = mDotNetFontName
  '         ctlInstance.Font.Size = mDotNetFontSize
       End If
   Next

End Sub


Public Function HtmlToVBHex(sHex As String) As Long
On Error Resume Next
   Dim sVBHex As String
   sHex = Trim("" & sHex)
   If Len(sHex) = 7 Then
      If Left(sHex, 1) = "#" Then sHex = Right(sHex, 6)
   End If
   If Len(sHex) = 6 Then
      sVBHex = "&H" & Right(sHex, 2) & mID(sHex, 3, 2) & Left(sHex, 2) & "&"
   Else
      sVBHex = "&H000000&"
   End If
   HtmlToVBHex = Val(sVBHex)
End Function

Public Sub SetDotNetThemeForUC(objUC As Object)
   On Error Resume Next
   Dim i As Integer
   Dim ctlInstance As Control

   If DotNetFormActive = False Then
      Exit Sub
   End If

   For i = 0 To objUC.Count
      Set ctlInstance = objUC(i)
      If TypeOf ctlInstance Is TextBox Then
         ctlInstance.FontName = mDotNetFontName
'           ctlInstance.FontSize = mDotNetFontSize
      ElseIf TypeOf ctlInstance Is Frame Then
         ctlInstance.BackColor = mDotNetBackColor
      ElseIf TypeOf ctlInstance Is SSUltraGrid Then
         ctlInstance.Appearance.BackColor = mDotNetBackColor
         ctlInstance.Appearance.Font.Name = mDotNetFontName
         ctlInstance.Override.HeaderAppearance.BackColor = mDotNetBackColor
         ctlInstance.CaptionAppearance.BackColor = mDotNetBackColor
 '          ctlInstance.Appearance.Font.Size = mDotNetFontSize
      Else
         ctlInstance.BackColor = mDotNetBackColor
         ctlInstance.Font.Name = mDotNetFontName
  '         ctlInstance.Font.Size = mDotNetFontSize
       End If
   Next

End Sub

