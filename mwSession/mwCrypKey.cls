VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwCrypkey"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The CrypKey Class provides a distributable component wrapper for all of the CrypKey SLAPI calls."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
' Cryptkey Class - Adapated from Samples from Kenonic
' 1/26/2001 ms Maritime Systems Inc



Option Explicit

' keys used in the InitCrypKey call, values gotten from Crypkey for Shipnet Fleet
Private mUserKey As String 'local copy
Private mMasterKey As String 'local copy
Private mCompanyNumber As String 'local copy
Private mPassword As String 'local copy
Private mCustomerServiceNumber As String 'local copy

Private mSiteKey As String 'local copy
Private mCrypKeyVersion As Long 'local copy
Private mSiteCode As String 'local copy
Private mFileName As String 'local copy
Private mNetworkMaxCheckTime As Integer 'local copy
Private mNumOptions As Integer

'CrypKey SLAPI

Private Declare Function ckChallenge% Lib "CRP32DLL.DLL" (ByVal random1%, ByVal random2%)
Private Declare Function crypkeyVersion& Lib "CRP32DLL.DLL" ()
Private Declare Function DirectTransfer& Lib "CRP32DLL.DLL" (ByVal directory$)
Private Declare Sub EndCrypKey Lib "CRP32DLL.DLL" ()
Private Declare Sub ExplainErr2 Lib "CRP32DLL.DLL" (ByVal FunctionCode&, ByVal errcode&, ByVal buffer$)
Private Declare Function Get1RestInfo& Lib "CRP32DLL.DLL" (ByVal which&)
Private Declare Function GetAuthorization& Lib "CRP32DLL.DLL" (oplevel&, ByVal decrement&)

Private Declare Function GetAuthorization2& Lib "CRP32DLL.DLL" (ByVal decrement&)
Private Declare Function GetLevel& Lib "CRP32DLL.DLL" (ByVal numDefineOpts&)
Private Declare Function GetNetHandle& Lib "CRP32DLL.DLL" ()
Private Declare Function GetNumCopies& Lib "CRP32DLL.DLL" ()
Private Declare Function GetNumMultiUsers& Lib "CRP32DLL.DLL" ()
Private Declare Function GetOption& Lib "CRP32DLL.DLL" (ByVal numDefineOpts&, ByVal optnum&)
Private Declare Function GetRestrictionInfo& Lib "CRP32DLL.DLL" (authopt&, start_date&, num_allowed&, num_used&)
Private Declare Function GetSiteCode& Lib "CRP32DLL.DLL" (ByVal SiteCode$)
Private Declare Function InitCrypkey& Lib "CRP32DLL.DLL" (ByVal filepath$, ByVal masterkey$, ByVal userkey$, ByVal floppy&, ByVal checktime&)
Private Declare Function KillLicense& Lib "CRP32DLL.DLL" (ByVal ConfirmCode$)
Private Declare Function ReadyToTry& Lib "CRP32DLL.DLL" (ByVal oplevel&, ByVal NumDays&)
Private Declare Function readyToTryDays& Lib "CRP32DLL.DLL" (ByVal oplevel&, ByVal NumDays&, ByVal Version&, ByVal copies&)
Private Declare Function readyToTryRuns& Lib "CRP32DLL.DLL" (ByVal oplevel&, ByVal NumRuns&, ByVal Version&, ByVal copies&)
Private Declare Function RegisterTransfer& Lib "CRP32DLL.DLL" (ByVal directory$)
Private Declare Function SaveSiteKey& Lib "CRP32DLL.DLL" (ByVal SiteKey$)
Private Declare Function TransferIn& Lib "CRP32DLL.DLL" (ByVal directory$)
Private Declare Function TransferOut& Lib "CRP32DLL.DLL" (ByVal directory$)
Private Declare Sub SetNetHandle Lib "CRP32DLL.DLL" (ByVal net_handle&)


'Private Constants
'** restriction types **
Private Const RESTR_NONE = 0
Private Const RESTR_TIME = 1
Private Const RESTR_RUNS = 2


'** general error return codes ***'
'--NOTE: These errors can come from any function listed below
Private Const FILE_INFO_FAIL = -100
Private Const NETWORK_DISCONNECTED = -101
Private Const NETWORK_NO_CKSERVE = -102
Private Const NETWORK_BAD_REPLY = -103
' currently only checked on Win32, DOS and Windows will add this saftey check later
Private Const INIT_NOT_SUCCEEDED = -104
Private Const THUNK_FAILURE = -105       '' only on Win32



'** return codes for InitCrypkey&() **'
Private Const INIT_OK = 0
Private Const INIT_FILE_NOT_FOUND = -1
Private Const INIT_MASTERKEY_CRC_FAILURE = -2
Private Const INIT_BAD_PRODUCT_NAME = -3
Private Const INIT_KEYFILE_CREATION_FAIL = -4
Private Const INIT_NETWORK_NOT_PURCHASED = -5
Private Const INIT_NT_NOT_PURCHASED = -6
Private Const INIT_NT32BIT_NOT_PURCHASED = -7
Private Const INIT_WIN95_NOT_PURCHASED = -8
Private Const INIT_WIN32S_NOT_PURCHASED = -9
' currently only checked on Win32, DOS and Windows will add this saftey check later
Private Const INIT_MULTIPLE_CALL_TO_INIT = -10
Private Const INIT_THUNK_LIB_NOT_FOUND = -11    '' only on Win32



'** return codes for GetAuthorization&() **'
Private Const AUTH_OK = 0
Private Const AUTH_INIT_FAIL = -1
Private Const AUTH_DISALLOW_FLOPPY = -2
Private Const AUTH_BAD_PATH = -3
Private Const AUTH_NOT_PRESENT = -4
Private Const AUTH_DIFFERENT = -5
Private Const AUTH_BAD_MASTERKEY = -6
Private Const AUTH_SITEKEY_CRC = -7
Private Const AUTH_TIME_TOO_EARLY = -8
Private Const AUTH_TIME_SETBACK = -9
Private Const AUTH_TIME_RUNOUT = -10
Private Const AUTH_RUNS_RESTR = -11
Private Const AUTH_NOT_ENOUGH_RUNS = -12
Private Const MISSING_RST_FILE = -13
Private Const AUTH_MISSING_RST_FILE = -13
Private Const RST_BAD_CRC = -14
Private Const AUTH_RST_BAD_CRC = -14
Private Const RST_BAD_LOCATION = -15
Private Const AUTH_RST_BAD_LOCATION = -15
Private Const ENTRY_CHECK_FAIL = -16
Private Const AUTH_ENTRY_CHECK_FAIL = -16
Private Const AUTH_NETTABLEFILE_FAIL = -17
Private Const AUTH_NETMAX_EXCEEDED = -18
Private Const AUTH_NETWORK_NOT_ALLOWED = -19

'** return codes for GetSiteCode&() **'
Private Const GSC_OK = 0
Private Const GSC_CRYPKEY_NOT_INITIALIZED = -1
Private Const GSC_ENTRY_FILEOPEN_FAIL = -2

'** return codes for SaveSiteKey&() **'
Private Const SITE_KEY_OK = 0
Private Const SITE_KEY_ENTRY_CHECK_FAIL = -1
Private Const SITE_KEY_ENTRY_CRC_FAIL = -2
Private Const SITE_KEY_FILEWRITE_FAILURE = -3

'** return codes for RegisterTransfer&() **'
Private Const REG_OK = 0
Private Const REG_THIS_ALREADY_AUTHORIZED = -1
Private Const REG_COULDNOT_OPEN_TARGET_REGFILE = -2
Private Const REG_TARGET_ALREADY_REGISTERED = -3
Private Const REG_SOURCE_ALREADY_REGISTERED = -4
Private Const REG_CANNOT_OPEN_REGFILE = -5
Private Const REG_CANNOT_WRITE_REGFILE = -6

'** return codes for TransferOut&() **'
Private Const TO_OK = 0
Private Const TO_THIS_NOT_AUTHORIZED = -1
Private Const TO_REGFILE_NOT_FOUND = -2
Private Const TO_REGFILE_CRC_FAILURE = -3
Private Const TO_DIFFERENT_APPLICATION = -4
Private Const TO_TARGET_ALREADY_HAS_LICENSE = -5
Private Const TO_SITEKEYFILE_NOT_FOUND = -6
Private Const TO_COULDNOT_WRITE_SITEKEYFILE = -7
Private Const TO_RSTKEYFILE_NOT_FOUND = -8
Private Const TO_COULDNOT_WRITE_RSTKEYFILE = -9
Private Const TO_COULDNOT_WRITE_REGKEYFILE = -10
Private Const TO_SOURCE_WRITE_PROTECTED = -11

'** return codes for TransferIn&() **'
Private Const TI_OK = 0
Private Const TI_THIS_ALREADY_AUTHORIZED = -1
Private Const TI_HARDDISK_REGFILE_NOT_FOUND = -2
Private Const TI_HARDDISK_REGFILE_CRC_FAILURE = -3
Private Const TI_REGFILE_NOT_FOUND = -4
Private Const TI_REGFILE_CRC_FAILURE = -5
Private Const TI_HARDDISK_REGFILE_MOVED = -6
Private Const TI_REG_FILES_DONT_MATCH = -7
Private Const TI_SOURCE_HAS_NO_LICENSE = -8
Private Const TI_SITEKEYFILE_NOT_FOUND = -9
Private Const TI_DIFFERENT_SITEKEY = -10
Private Const TI_COULDNOT_WRITE_SITEKEYFILE = -11
Private Const TI_RSTKEYFILE_NOT_FOUND = -12
Private Const TI_DIFFERENT_RSTFILE = -13
Private Const TI_COULDNOT_WRITE_RSTKEYFILE = -14

'** return codes for DirectTransfer&() **'
Private Const DT_OK = 0
Private Const DT_THIS_NOT_AUTHORIZED = -1
Private Const DT_TARGET_APP_NOT_FOUND = -2
Private Const DT_FLOPPY_NOT_ALLOWED = -3
Private Const DT_SITEKEYFILE_NOT_FOUND = -4
Private Const DT_RSTKEYFILE_NOT_FOUND = -5
Private Const DT_COULDNOT_WRITE_SITEKEYFILE = -6
Private Const DT_SOURCE_WRITE_PROTECTED = -7
Private Const DT_SOURCE_SAME_AS_TARGET = -8

'** return codes for ReadyToTry&() **'
'**Note - Most of these error codes will never be seen - they are only given
         'so that you can report them to us if they occur. The only code you
         'may be interested in is RTT_DONE_THIS, which will come up often,
         'and signifies the trial license has already been given.
         'If you get any other error, report it to us. **'
Private Const RTT_OK = 0
Private Const RTT_COULD_NOT_GET_SITE_CODE = -1
Private Const RTT_COULD_NOT_GET_SITE_KEY = -2
Private Const RTT_COULD_NOT_SAVE_SITE_KEY = -3
Private Const RTT_RESERVED_4 = -4
Private Const RTT_BAD_DOS = -5
Private Const RTT_BAD_TRUENAME = -6
Private Const RTT_NO_REDIRECT = -7
Private Const RTT_NO_DPB = -8
Private Const RTT_NO_MEM = -9
Private Const RTT_CANT_GET_CLUSTER = -10
Private Const RTT_BAD_STAT = -11
Private Const RTT_NO_ROOM = -12
Private Const RTT_BAD_SECTOR_READ = -13
Private Const RTT_BAD_SECTOR_WRITE = -14
Private Const RTT_FILE_SEARCH = -15
Private Const RTT_FILE_ACCESS = -16
Private Const RTT_FILE_NOT_FOUND = -17
Private Const RTT_FILE_OPEN = -18
Private Const RTT_DONE_THIS = -19
Private Const RTT_NO_SIG = -20
Private Const RTT_NO_LISTFILE = -21
Private Const RTT_CANT_FIND_DRIVE = -22
Private Const RTT_NO_GOOD_PUTS = -23
Private Const RTT_NO_REAL_DRIVE = -24
Private Const RTT_32BIT_FILE_ACCESS = -25
Private Const RTT_CLOSE_ALL_FILES = -26

'** return codes for KillLicense&() **'
Private Const KL_OK = 0
Private Const KL_CRYPKEY_NOT_INITIALIZED = -1
Private Const KL_CRYPKEY_NOT_AUTHORIZED = -2
Private Const KL_LICENSE_WRITE_PROTECTED = -3

'** function codes for Get1RestInfo&() **'
Private Const G1_OUT_OF_RANGE = -1

'** function codes for GetNumCopies&() **'
Private Const GNC_CRYPKEY_NOT_INITIALIZED = -1

'** function codes for GetOption&() **'
Private Const OPTION_ON = 1
Private Const OPTION_OFF = 0
Private Const OPTION_UNAUTHORIZED = -1

'** function codes for GetRestrictionInfo&() **'
Private Const GRI_OK = 0
Private Const GRI_INVALID = -1

'** function codes for ExplainErr&() **'
Private Const EXP_AUTH_ERR = 1            'GetAuthorization&() errcodes
Private Const EXP_GET_SITECODE_ERR = 2    'GetSiteCode&() errcodes
Private Const EXP_SAVE_SITEKEY_ERR = 3    'SaveSiteKey&() errcodes
Private Const EXP_REG_ERR = 4             'RegisterTransfer&() errorcodes
Private Const EXP_TO_ERR = 5              'TransferOut&() errorcodes
Private Const EXP_TI_ERR = 6              'TransferIn&() errorcodes
Private Const EXP_DT_ERR = 7              'DirectTransfer&() errorcodes
Private Const EXP_INIT_ERR = 8            'InitCrypkey&() errorcodes
Private Const EXP_RTT_ERR = 9             'ReadyToTry&() errorcodes
Private Const EXP_KL_ERR = 10             'KillLicense&() errorcodes


Public Property Get IsDaysRestricted() As Boolean
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long

    CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

    Select Case AuthOptions
        Case RESTR_NONE
            IsDaysRestricted = False
        Case RESTR_TIME
            IsDaysRestricted = True
        Case RESTR_RUNS
            IsDaysRestricted = False
    End Select
End Property

Public Property Get GetNumCurrentNetUsers() As Integer
   'Function Net Available
    GetNumCurrentNetUsers = 0
End Property



Public Function KillLicenseGE() As String
   Dim ConfirmCode As String * 80
   KillLicense ConfirmCode
   KillLicenseGE = Trim(ConfirmCode)
End Function


Public Function ReadyToTryDaysMS() As Boolean
   Dim CrypKeyResult As Integer
   Dim numOpts As Integer
   Dim NumDays As Integer
   Dim numVersion As Integer
   Dim NumCopies As Integer
    numOpts = 0
    NumDays = 3
    numVersion = 2
    NumCopies = 1
    CrypKeyResult = readyToTryDays(numOpts, NumDays, numVersion, NumCopies)
    If CrypKeyResult <> RTT_OK Then
      Call ShowCrypKeyError(FunctionCode:=EXP_RTT_ERR, ErrorCode:=CrypKeyResult)
      ReadyToTryDaysMS = False
   Else
      ReadyToTryDaysMS = True
   End If
End Function

Public Property Get IsUnlimited() As Boolean
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long
   
   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

   Select Case AuthOptions
       Case RESTR_NONE
           IsUnlimited = True
       Case RESTR_TIME
           IsUnlimited = False
       Case RESTR_RUNS
           IsUnlimited = False
   End Select
   ' -1 indicates invalid information...
   If CrypResult < 0 Then
      IsUnlimited = False
   End If
End Property

Public Function IsOptionValid(OptionCheck As Integer) As Boolean
    IsOptionValid = GetOption(mNumOptions, OptionCheck)
End Function

Public Property Get IsRunsRestricted() As Boolean
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long
   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

   Select Case AuthOptions
       Case RESTR_NONE
           IsRunsRestricted = False
       Case RESTR_TIME
           IsRunsRestricted = False
       Case RESTR_RUNS
           IsRunsRestricted = True
   End Select
End Property

' Doesn't work...
'Public Property Get StartDateofAuthorization() As Date
'   Dim AuthOptions As Long
'   Dim StartDate As Long
'   Dim NumAllowed As Long
'   Dim NumUsed As Long
'   Dim CrypResult As Long
'   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)
'
'   StartDateofAuthorization = Date 'Process(StartDate)
'
'End Property

Public Property Get NumberOfDaysRunsUsed() As Integer
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long
   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

   NumberOfDaysRunsUsed = NumUsed
End Property

Public Property Get NumberOfDaysRunsAllowed() As Long
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long
   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

   NumberOfDaysRunsAllowed = NumAllowed
End Property

Public Property Get NumberOfDaysRunsRemaining() As Integer
   Dim AuthOptions As Long
   Dim StartDate As Long
   Dim NumAllowed As Long
   Dim NumUsed As Long
   Dim CrypResult As Long
   CrypResult = GetRestrictionInfo(AuthOptions, StartDate, NumAllowed, NumUsed)

   NumberOfDaysRunsRemaining = NumAllowed - NumUsed
End Property



Public Property Let FileName(vData As String)
    mFileName = vData
End Property

Public Property Get GetNumMultiUsersGE() As Integer
    GetNumMultiUsersGE = GetNumMultiUsers()
End Property

Public Property Get GetNumCopiesGE() As Integer
    GetNumCopiesGE = GetNumCopies()
End Property

Public Property Get GetSiteCodeGE() As String
   Dim SiteCodeText As String * 30
   GetSiteCodeGE = GetSiteCode(SiteCodeText)
   GetSiteCodeGE = Trim(SiteCodeText)
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Get GetLevelGE() As Integer
    GetLevelGE = GetLevel(mNumOptions)
End Property

Public Property Let NumOptions(vData As Integer)
    mNumOptions = vData
End Property

Public Property Get GetOptionGE(OptionNumber As Integer) As Integer
    GetOptionGE = GetOption(mNumOptions, OptionNumber)
End Property

Public Property Let NetworkMaxCheckTime(vData As Integer)
    mNetworkMaxCheckTime = vData
End Property

Public Property Get NetworkMaxCheckTime() As Integer
    NetworkMaxCheckTime = mNetworkMaxCheckTime
End Property

Public Sub SaveSiteKeyGE(ByVal SiteKey As String)
Attribute SaveSiteKeyGE.VB_Description = "Used to save  to a file the Site Key that has been accquired by the user. The key is checked first before it is saved"
   Dim CrypReturn As Integer
   CrypReturn = SaveSiteKey(SiteKey)
   If CrypReturn <> 0 Then
      ShowCrypKeyError EXP_SAVE_SITEKEY_ERR, CrypReturn
   End If
End Sub

Public Property Let SiteCode(ByVal vData As String)
Attribute SiteCode.VB_Description = "Returns the Site Code for the program location. The Site Code must be reported by the user before the devloper can issue a Site Key."
    mSiteCode = vData
End Property


Public Property Get SiteCode() As String
    SiteCode = mSiteCode
End Property



Public Property Get GetAuthorizationGE() As Integer
    GetAuthorizationGE = GetAuthorization2(1)
End Property

Public Function InitCrypkeyGE() As Integer
Attribute InitCrypkeyGE.VB_Description = "Initializes CrypKey with runtime information and must be called before any other function is SLAPI."
   On Error GoTo FunctionError
   Dim CrypKeyError As Integer
   InitCrypkeyGE = InitCrypkey(mFileName, mMasterKey, mUserKey, 0, mNetworkMaxCheckTime)
   Exit Function
FunctionError:
   goSession.RaiseError "General error in ICK, ", err.Number, err.Description
   InitCrypkeyGE = -1
End Function

Public Sub ShowCrypKeyError(FunctionCode As Integer, ErrorCode As Integer)
   Dim ErrorText As String * 80
   Dim Msg As String
   Call ExplainErr2(FunctionCode, ErrorCode, ErrorText)
   Msg = "An error has occured in the licensing protection software. Please note the following error:" & vbCrLf & vbCrLf
   Msg = Msg & "   Error Code: " & ErrorCode & vbCrLf
   Msg = Msg & "   Error Text: " & Trim(ErrorText)
   MsgBox Msg, vbInformation, "Licensing Protection Error"
End Sub
    
Private Function GetCrypKeyError(FunctionCode As Integer, ErrorCode As Integer) As String
   Dim ErrorText As String * 80
   Dim Msg As String
   Call ExplainErr2(FunctionCode, ErrorCode, ErrorText)
   GetCrypKeyError = Trim(ErrorText)
End Function

Public Sub EndCrypKeyGE()
Attribute EndCrypKeyGE.VB_Description = "Notifies CrypKey that the program is terminating. Useful in Floating License models."
    EndCrypKey
End Sub



Public Property Get CrypKeyVersionGE() As String
Attribute CrypKeyVersionGE.VB_Description = "Returns the version number of the CrypKey library currently in use."
   Dim CrypKeyResult As Integer
   CrypKeyResult = crypkeyVersion
   CrypKeyVersionGE = Left$(CStr(CrypKeyResult), 1) & "." & Right$(CStr(CrypKeyResult), 1)
End Property



Public Property Let SiteKeyGE(ByVal vData As String)
Attribute SiteKeyGE.VB_Description = "You will type this Site Key into your Site Key generator in order to activate its licensing capabilites for you product."
    mSiteKey = vData
End Property


Public Property Get SiteKeyGE() As String
    SiteKeyGE = mSiteKey
End Property



Public Property Let MasterKeyGE(ByVal vData As String)
Attribute MasterKeyGE.VB_Description = "An encrypted form of information specific to your product being protected and is provided to you by Kenonic."
    mMasterKey = vData
End Property


Public Property Get MasterKeyGE() As String
    MasterKeyGE = mMasterKey
End Property


Public Property Let UserKeyGE(ByVal vData As String)
Attribute UserKeyGE.VB_Description = "An encrypted form of the developer/product specific password which is provided to you by Kenonic."
    mUserKey = vData
End Property


Public Property Get UserKeyGE() As String
    UserKeyGE = mUserKey
End Property



Private Sub Class_Initialize()
   mFileName = goSession.GetAppPath() & "\mswf.lic"
   ' version 5.7
   mMasterKey = "DAAF 55C4 A789 A3CD 3E68 8718 D2F3 3B16 703C 3B16 6450"
   mUserKey = "DB36 D744 D7AD 4F5C CE24 9250 BA"
   
   ' version 7.1
'   mMasterKey = "7ac2fbaf0c2a24c69a14b3266f873cfd250ec395edc9d5db3611eab2e3c569122d1898750c50cfe82b6c653b7c07d26e8667c922d0183f11cb904c377d1daf1691e9671a366788824888062c775ac1d3aa66c351b64655a78541c816e7571db39eb65e82c4e38b39db97a3366e843e77374ec5f1959a95913134473fee7758ce"
'   mUserKey = "DB36D744D7AD4F5CCE249250BA"
   
   ' Not used in software, for information purposes only
   mCompanyNumber = "7956707"
   mPassword = "53296430SAM"
   mCustomerServiceNumber = "C03185"
   
   mNetworkMaxCheckTime = 1800
   mNumOptions = 19

End Sub

Private Sub Class_Terminate()
   'Let License Manager know we are leaving
   '    cCrypKey.EndCrypKeyGE
End Sub

Public Function ckChallengeMS() As Boolean
   Dim iR1 As Integer
   Dim iR2 As Integer
   Dim i As Integer
   Dim iResult As Long
   Dim iReturn As Long
   iResult = 0
   iR1 = Rnd(1) * 456
   iR2 = Rnd(1) * 699
   For i = 0 To 10
      iResult = (iResult * iR1 + 7956707) Mod 16381 + _
        (iResult * iR2 + 1298084087) Mod 16369
   Next i
   iReturn = ckChallenge(ByVal iR1, ByVal iR2)
   If iReturn <> iResult Then
      ckChallengeMS = False
   Else
      ckChallengeMS = True
   End If

End Function



'
' Step 1: Computer Requests a license
'
Public Function TransferRequestMS(directory As String) As Boolean
   Dim iRet As Integer
   Dim strDirectory As String
   Dim fso As FileSystemObject
   On Error GoTo TransferRequestMS_error
   Set fso = New FileSystemObject
   If Not fso.FolderExists(directory) Then
      iRet = MsgBox("Press any key when Transfer Media is Ready", vbOKCancel, "ShipNet Fleet Licensing")
      If iRet = vbCancel Then
         TransferRequestMS = False
         Set fso = Nothing
         Exit Function
      End If
   End If
   Set fso = Nothing
   strDirectory = directory & Chr$(0)
   iRet = RegisterTransfer(strDirectory)
   Select Case iRet
      Case Is = REG_OK
         MsgBox "License Request has been placed on media.", vbInformation, "ShipNet Fleet Licensing"
         TransferRequestMS = True
      Case Is = REG_THIS_ALREADY_AUTHORIZED
         MsgBox "This site is already authorized.", vbInformation, "ShipNet Fleet Licensing"
         TransferRequestMS = False
      Case Is = REG_COULDNOT_OPEN_TARGET_REGFILE
         MsgBox "This site is already authorized.", vbInformation, "ShipNet Fleet Licensing"
         TransferRequestMS = False
      Case Else
         MsgBox "Error On Transfer Request: " & iRet, vbCritical, "ShipNet Fleet Licensing"
         TransferRequestMS = False
   End Select
   Exit Function
TransferRequestMS_error:
   MsgBox "VB error in Transfer License: " & err.Number & err.Description
   TransferRequestMS = False
End Function

'
' Step 2: Computer with a license transfers out
'
Public Function TransferOutMS(directory As String) As Boolean
   Dim iRet As Integer
   Dim strDirectory As String
   Dim fso As FileSystemObject
   On Error GoTo TransferOutMS_error
   Set fso = New FileSystemObject
   strDirectory = directory & Chr$(0)
   iRet = TransferOut(strDirectory)
   Select Case iRet
      Case Is = REG_OK
         MsgBox "License has been transferred out.", vbInformation, "ShipNet Fleet Licensing"
         TransferOutMS = True
      Case Is = REG_THIS_ALREADY_AUTHORIZED
         MsgBox "This site is already authorized.", vbCritical, "ShipNet Fleet Licensing"
         TransferOutMS = False
      Case Is = REG_COULDNOT_OPEN_TARGET_REGFILE
         MsgBox "Could not open registry on media.", vbCritical, "ShipNet Fleet Licensing"
         TransferOutMS = False
      Case Else
         MsgBox "Error On Transfer Request: " & iRet, vbCritical, "ShipNet Fleet Licensing"
         TransferOutMS = False
   End Select
   Exit Function
TransferOutMS_error:
   MsgBox "VB error in Transfer License: " & err.Number & err.Description
   TransferOutMS = False
End Function

'
' Step 3: Computer needing a license Transfers In...
'

Public Function TransferInMS(directory As String) As Boolean
   Dim iRet As Integer
   Dim strDirectory As String
   Dim fso As FileSystemObject
   On Error GoTo TransferInMS_error
   Set fso = New FileSystemObject
   If Not fso.FolderExists(directory) Then
   Select Case iRet
      Case Is = REG_OK
         MsgBox "License has been imported.", vbInformation, "ShipNet Fleet Licensing"
         TransferInMS = True
      Case Is = REG_THIS_ALREADY_AUTHORIZED
         MsgBox "This site is already authorized.", vbInformation, "ShipNet Fleet Licensing"
         TransferInMS = False
      Case Is = REG_COULDNOT_OPEN_TARGET_REGFILE
         MsgBox "Could not open registry on media.", vbInformation, "ShipNet Fleet Licensing"
         TransferInMS = False
      Case Else
         MsgBox "Error On Transfer Request: " & iRet, vbCritical, "ShipNet Fleet Licensing"
         TransferInMS = False
   End Select
   End If
   Set fso = Nothing
   strDirectory = directory & Chr$(0)
   iRet = TransferIn(strDirectory)
   If iRet <> 0 Then
      
      TransferInMS = False
      Exit Function
   End If
   'iRet = TransferIn(strDirectory)
   'If iRet <> 0 Then
   '   Call ShowCrypKeyError(FunctionCode:=EXP_INIT_ERR, ErrorCode:=iRet)
   '   TransferInMS = False
   '   Exit Function
   'End If
   TransferInMS = True
   Exit Function
TransferInMS_error:
   MsgBox "VB error in Transfer License: " & err.Number & err.Description
   TransferInMS = False
End Function

