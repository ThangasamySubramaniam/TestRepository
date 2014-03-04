VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' mwAPI - ShipNet Fleet API Declarations and constants



'
' App Shell Constants - from Sybex VB 6 Developer's Handbook - Page 750
'
' 2/2001 - First used to shell to Internet Explorer to view HTML documents in frmManualViewer
'

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3

Private Const GW_CHILD = 5
Private Const GW_MAX = 5

'Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const GW_OWNER = 4
Private Const GWL_STYLE = -16
Private Const WS_DISABLED = &H8000000
Private Const WS_CANCELMODE = &H1F
Private Const WM_CLOSE = &H10

Private Const MAX_PATH = 260
Private Const SW_SHOWNORMAL = 1

Private Const USER_SESSION_BASE_REGISTRY As String = "Software\Maritime Systems Inc\SessionSettings"
Private Const FLEET_APP_BASE_REGISTRY As String = "Software\Maritime Systems Inc"

' Locale Information
  Private Const vbZLString As String = ""
  Private Const API_FALSE As Long = &H0&
  Private Const API_TRUE As Long = &H1&

  Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
  Private Const LOCALE_USER_DEFAULT As Long = &H400
  Private Const LOCALE_SDECIMAL As Long = &HE         '  decimal separator
  Private Const LOCALE_STHOUSAND As Long = &HF         '  thousand separator
  Private Const LOCALE_SGROUPING As Long = &H10        '  digit grouping
  Private Const LOCALE_IDIGITS As Long = &H11        '  number of fractional digits
  Private Const LOCALE_ILZERO As Long = &H12        '  leading zeros for decimal
  Private Const LOCALE_SDATE As Long = &H1D        '  date separator
  Private Const LOCALE_STIME As Long = &H1E        '  time separator
  Private Const LOCALE_SSHORTDATE As Long = &H1F        '  short date format string
  Private Const LOCALE_SLONGDATE As Long = &H20        '  long date format string
  Private Const LOCALE_STIMEFORMAT As Long = &H1003      '  time format string
  Private Const LOCALE_IDATE As Long = &H21        '  short date format ordering
  Private Const LOCALE_ILDATE As Long = &H22        '  long date format ordering
  Private Const LOCALE_ITIME As Long = &H23        '  time format specifier
  Private Const LOCALE_ICENTURY As Long = &H24        '  century format specifier
  Private Const LOCALE_ITLZERO As Long = &H25        '  leading zeros in time field
  Private Const LOCALE_IDAYLZERO As Long = &H26        '  leading zeros in day field
  Private Const LOCALE_IMONLZERO As Long = &H27        '  leading zeros in month field
  Private Const LOCALE_S1159 As Long = &H28        '  AM designator
  Private Const LOCALE_S2359 As Long = &H29        '  PM designator
  
  'VEL-207 Printing adjusted Inventories doesnt Print all the fields in A4 size paper
  'BY N.Angelakis 1 April 2011
  Private Const LOCALE_SMEASUREMENT As Long = &HD    'REGIONAL SETTINGS MEASUREMENT 0-metric, 1-U.S. inches

Public Enum mwLocaleStrings
   LOCALE_TIME_SEPARATOR = 1
   LOCALE_DATE_SEPARATOR = 2
   LOCALE_SHORT_DATE = 3
   LOCALE_LONG_DATE = 4
   LOCALE_TIME = 5
   LOCALE_DECIMAL_SEPARATOR = 6
   LOCALE_THOUSANDS_SEPARATOR = 7
   
  'VEL-207 Printing adjusted Inventories doesnt Print all the fields in A4 size paper
  'BY N.Angelakis 1 April 2011
   LOCALE_MEASUREMENT = 8
End Enum

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

'
' FromWin32 API programming - Page 121-123
'
Private Type OSVERSIONINFO
   OSVersionInfoSize As Long
   MajorVersion As Long
   MinorVersion As Long
   BuildNumber As Long
   PlatformID As Long
   szCSDVersion As String * 128
End Type

Public Type MEMORYSTATUS
   dwLength  As Long          'sizeof(MEMORYSTATUS)
   dwMemoryLoad  As Long      'percent of memory in use
   dwTotalPhys  As Long       'bytes of physical memory
   dwAvailPhys  As Long       'free physical memory bytes
   dwTotalPageFile  As Long   'bytes of paging file
   dwAvailPageFile  As Long   'free bytes of paging file
   dwTotalVirtual  As Long    'user bytes of address space
   dwAvailVirtual  As Long    'free user bytes
End Type



'
' App Shell Win API Declarations - from Sybex VB 6 Developer's Handbook - Page 750
' See modMWStaging for constants...
'
' 2/2001 - First used to shell to Internet Explorer to view HTML documents in frmManualViewer
'

Private Declare Function GetComputerName Lib "kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle _
   As Long, ByVal dwProcessId As Long) As Long
   
Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long
   
Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) _
    As Boolean
    
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
'Const INFINITE = &HFFFF
Const STARTF_USESHOWWINDOW = &H1
Private Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum

Private Type SECURITY_ATTRIBUTES
    nlength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
'Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
'   (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
'   lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, _
'   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, _
'   ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, _
'   lpProcessInformation As PROCESS_INFORMATION) As Long
   
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) _
    As Long
    
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal wCmd As Long) As Long
    
Private Declare Function PostMessage Lib "user32" Alias _
    "PostMessageA" (ByVal hwnd As Long, ByVal wMsg _
    As Long, ByVal wParam _
    As Long, ByVal lParam As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function GetParent Lib "user32" (ByVal hwnd _
  As Long) As Long
  
Private Declare Function GetWindowTextLength Lib "user32" Alias _
    "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias _
    "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString _
    As String, ByVal cch As Long) As Long


' File Associations
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
  (ByVal lpFile As String, ByVal lpDirectory As String, _
  ByVal lpResult As String) As Long
  
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
  "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
  lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Windows Version Information
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale&, ByVal LCType&, ByVal lpLCData$, ByVal cchData&) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)


Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer


' HTML Help Viewer API
'//Constant Declarations
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_CONTEXT = &HF
Private Const HH_CLOSE_ALL = &H12
Private Const HH_DISPLAY_TEXT_POPUP = &HE
Private Const HH_DISPLAY_TOC = &H1

'Private Const MST_OLD_PREFIX = "mk:@MSITStore:"

'//API Declaration
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias _
    "HtmlHelpA" (ByVal hwndCaller As Long, _
    ByVal pszFile As String, _
    ByVal uCommand As Long, _
    ByVal dwData As Long) As Long
    
'
' Win32 System Time Functions... Support Functions to convert times between Local and GMT
'

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, pvTime As Date) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vTime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Type PROCESS_HEAP_ENTRY
  lpData As Long
  cbData As Long
  cbOverhead As Byte
  iRegionIndex As Byte
  wFlags As Integer
  dwCommittedSize As Long
  dwUnCommittedSize As Long
  lpFirstBlock As Long
  lpLastBlock As Long
End Type

Private Const PROCESS_HEAP_REGION As Integer = &H1
Private Const PROCESS_HEAP_UNCOMMITTED_RANGE As Integer = &H2
Private Const PROCESS_HEAP_ENTRY_BUSY As Integer = &H4
Private Const PROCESS_HEAP_ENTRY_MOVEABLE As Integer = &H10
Private Const PROCESS_HEAP_ENTRY_DDESHARE As Integer = &H20

Private Declare Function GetProcessHeaps Lib "kernel32" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long

Private Declare Function HeapWalk Lib "kernel32" (ByVal hHeap As Long, lpEntry As PROCESS_HEAP_ENTRY) As Long



'
' Create random file snippet...
'
'Private Declare Function GetTempFileName Lib "kernel32" _
'     Alias "GetTempFileNameA" (ByVal lpszPath As String, _
'     ByVal lpPrefixString As String, ByVal wUnique As Long, _
'     ByVal lpTempFileName As String) As Long

'Private Sub Command1_Click()
'   Dim sPath As String
'   Dim sPrefix As String
'   Dim lUnique As Long
'   Dim sTempFileName As String
'
'   sPath = "c:\amoslink"
'   sPrefix = "fVB"
'   lUnique = 0
'
'   sTempFileName = Space$(100)
'   GetTempFileName sPath, sPrefix, lUnique, sTempFileName
'   sTempFileName = mID$(sTempFileName, 1, InStr(sTempFileName, Chr$(0)) - 1)
'   GenTempName = sTempFileName
'End Sub




'-------------------------------------------- end time zone stuff

' Numega memory trapping instrumentation (requires SmartCheck.exe)

'Release Notes for NuMega SmartCheck (TM) and NuMega BoundsChecker (TM)
'Event Reporting / Memory Analysis API Version 6.2
'Memory Analysis APIs
'---------------------
'
'BoundsChecker and SmartCheck now have additional memory analysis
'features. These functions are designed to allow a more detailed analysis
'of the code.
'
'     long NMMemMark ();
'     long NMMemSave (const char *pszFileName,
'                     const char *pszComment,
'                     long bSaveAll);
'     long NMMemPopup ();
'
'NMMemMark "marks" all currently allocated memory.  All future memory
'allocations will not be marked, so you can identify what has been
'allocated since the mark.
'
'NMMemSave writes information about the currently allocated memory blocks
'to a file.  The file is specified by the first parameter, pszFileName.
'pszComment specifies a comment that is written to the file as the first
'record.  Any carriage returns or line feeds in the comment will be
'converted into spaces.  If bSaveAll is non-zero, information about all
'the currently allocated memory blocks will be written to the file.  If
'bSaveAll is zero, information about only those memory blocks allocated
'since the last mark are written.
'
'NMMemPopup brings up a dialog containing all the statistics for the
'current memory started with the MemMark command, as well as the overall
'program.
'

Private Declare Function NMMemMark Lib "NMAPI.dll" () As Long
Private Declare Function NMMemSave Lib "NMAPI.dll" (ByVal strFilename As String, _
                                                    ByVal strComment As String, _
                                                    blnSaveAll As Boolean) As Long
Private Declare Function NMMemPopup Lib "NMAPI.dll" () As Long

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Private mWindowsVersion As String

Private moParent As Session

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
            (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Get WindowsVersion() As String
    WindowsVersion = mWindowsVersion
End Property

Public Property Let WindowsVersion(ByVal Data As String)
    mWindowsVersion = Data
End Property

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


Public Function GetGmtNow() As Date
   Dim st As SYSTEMTIME
   Dim utc As Date
   GetSystemTime st
   SystemTimeToVariantTime st, utc
   GetGmtNow = utc
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwAPI.GetGmtNow", Err.Number, Err.Description
   GetGmtNow = Now()
End Function

Public Function GmtToLocalTime(GmtDate As Date) As Date
   Dim stUTC As SYSTEMTIME
   Dim stLocal As SYSTEMTIME
   Dim ftUTC As FILETIME
   Dim ftLocal As FILETIME
   VariantTimeToSystemTime GmtDate, stUTC
   SystemTimeToFileTime stUTC, ftUTC
   FileTimeToLocalFileTime ftUTC, ftLocal
   FileTimeToSystemTime ftLocal, stLocal
   SystemTimeToVariantTime stLocal, GmtToLocalTime
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwAPI.GmtToLocalTime", Err.Number, Err.Description
   GmtToLocalTime = GmtDate
End Function

Public Function LocalToGmtTime(LocalDate As Date) As Date
   Dim stUTC As SYSTEMTIME
   Dim stLocal As SYSTEMTIME
   Dim ftUTC As FILETIME
   Dim ftLocal As FILETIME
   VariantTimeToSystemTime LocalDate, stLocal
   SystemTimeToFileTime stLocal, ftLocal
   LocalFileTimeToFileTime ftLocal, ftUTC
   FileTimeToSystemTime ftUTC, stUTC
   SystemTimeToVariantTime stUTC, LocalToGmtTime
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwAPI.LocalToGmtTime", Err.Number, Err.Description
   LocalToGmtTime = LocalDate
End Function


'Public Function RunShellExecute(sfile As String, ByRef formHandle As Long) As Long
Public Function RunShellExecute(sFile As String) As Long
   Dim lresp As Long
   Dim dummy As Long
   On Error GoTo FunctionError
   dummy = 1

'   The RunCreateProcess function is a replacement for ShellExecute but it doesn't work yet.
'   Uncomment it and comment the ShellExecute call to test new function

   lresp = ShellExecute(dummy, "open", sFile, vbNullString, vbNullString, SW_SHOWNORMAL)
'   lresp = RunCreateProcess(sFile)
   RunShellExecute = lresp
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwAPI.RunShellExecute", Err.Number, Err.Description
   RunShellExecute = lresp
End Function

'Public Function RunWinExe(sProg As String, sParms As String) As Long
'   Dim lresp As Long
'   Dim hRet As Long
'   Dim sCmdLine As String
'
'   On Error GoTo FunctionError
'
'   sCmdLine = sProg & " " & sParms
'
'   If sCmdLine = "" Then
'      hRet = 0
'   Else
'      hRet = WinExec(sCmdLine, SW_NORMAL)
'   End If
'   If hRet = 0 Then
'      moParent.RaiseError "Error in mwAPI.RunWinExe, Calling external program: " & sCmdLine
'   End If
'
'   RunWinExe = hRet
'
'   Exit Function
'FunctionError:
'   moParent.RaiseError "Error in mwSession.mwAPI.RunWinExe", err.Number, err.Description
'   RunWinExe = hRet
'End Function


Public Function GetWindowsVersion() As String

   Dim lret As Long
   Dim osverinfo As OSVERSIONINFO
   osverinfo.OSVersionInfoSize = Len(osverinfo)
   lret = GetVersionEx(osverinfo)
   If lret = 0 Then
      GetWindowsVersion = ""
      Exit Function
   End If
   Select Case osverinfo.PlatformID & "/" & osverinfo.MajorVersion & "/" & osverinfo.MinorVersion
      Case Is = "1/4/0"
         GetWindowsVersion = "WIN95"
      Case Is = "1/4/10"
         GetWindowsVersion = "WIN98"
      Case Is = "1/4/90"
         GetWindowsVersion = "WINME"
      Case Is = "1/4/0"
         GetWindowsVersion = "WIN95"
      Case Is = "2/4/0"
         GetWindowsVersion = "WINNT"
      Case Is = "2/5/0"
         GetWindowsVersion = "WIN2K"
      Case Is = "2/5/1"
         GetWindowsVersion = "WINXP"
      Case Else
         GetWindowsVersion = "Unknown OS: " & osverinfo.PlatformID & "/" & osverinfo.MajorVersion & "/" & osverinfo.MinorVersion
   End Select
   'GetWindowsVersion = str((osverinfo.MajorVersion * 100) + osverinfo.MinorVersion)

End Function



Public Function ShowMwHelp(HelpName As String, Optional HelpContextID As String, _
  Optional IsHelpFileID As Boolean, Optional IsTemplateID As Boolean)
   Dim strFile As String
   Dim IsChmFile As Boolean
   Dim IsDocFile As Boolean
   Dim IsPDFFile As Boolean
   Dim strExtension As String
   Dim strManualName As String
   Dim strTopic As String
   Dim strAppPath As String
   Dim loRs As Recordset
   Dim strSQL As String
   Dim sCmdLine As String
   Dim hRet As Long
   On Error GoTo FunctionError
   If IsHelpFileID Then
      '
      ' Fetch Help File Name...
      '
      strManualName = FetchHelpFileSourceName(HelpName)
      If strManualName = "" Then
         goSession.RaiseError "Error in mwAPI.ShowMwHelp, help record/source name  missing for: " & HelpName
         CloseRecordset loRs
      End If
      CloseRecordset loRs
   ElseIf IsTemplateID Then
      '
      ' Fetch felp information from form template record...
      '
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      strSQL = "select * from mwwfFormTemplate where TemplateID='" & HelpName & "'"
      loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If loRs.RecordCount > 0 Then
         If Not IsNull(loRs!FormHelpID_Override) Then
            strManualName = FetchHelpFileSourceName(loRs!FormHelpID_Override)
         Else
            'error
            Beep
            'goSession.RaiseError "Error in mwAPI.ShowMwHelp, TemplateID Help File Reference is missing for: " & HelpName
            CloseRecordset loRs
            Exit Function
         End If
         If Not IsNull(loRs!FormHelpContextID) Then
            HelpContextID = loRs!FormHelpContextID
         End If
      Else
         'error
         goSession.RaiseError "Error in mwAPI.ShowMwHelp, Template missing in Form Template table: " & HelpName
         CloseRecordset loRs
         Exit Function
      End If
      CloseRecordset loRs
   Else
      strManualName = HelpName
   End If
   
   Dim fso As New FileSystemObject
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   If InStr(1, strManualName, "\") < 1 Then
      strAppPath = goSession.GetAppPath & "\HelpManuals\"
   End If
   strExtension = fso.GetExtensionName(strManualName)
   If Not fso.FileExists(strManualName) Then
      strFile = strAppPath & strManualName
   Else
      strFile = strManualName
   End If
   
   If Trim(strExtension) = "" Then
      IsChmFile = True
   ElseIf UCase(strExtension) = "CHM" Then
      'strFile = strAppPath & strManualName
      IsChmFile = True
   ElseIf UCase(strExtension) = "DOC" Then
      IsDocFile = True
   ElseIf UCase(strExtension) = "DOCX" Then
      IsDocFile = True
   ElseIf UCase(strExtension) = "DOCM" Then
      IsDocFile = True
   ElseIf UCase(strExtension) = "PDF" Then
      IsPDFFile = True
   Else
      moParent.RaiseError "Error in mwAPI.ShowMwHelp, only CHM, DOC, or PDF help files supported: " & strManualName
   End If
   '
   ' Does file exist ?
   '
   If Not fso.FileExists(strFile) Then
      moParent.RaiseError "Error in mwAPI.ShowMwHelp, Help File is Missing: " & strManualName
      Set fso = Nothing
      Exit Function
   End If
   '
   ' OK...
   '
   If Trim(HelpContextID) = "" Or IsPDFFile Then
   

'      hRet = RunCreateProcess(strFile)
      If IsChmFile Then
         '
         ' Use WinExec
         '
         If fso.FileExists(strFile) Then
            sCmdLine = "hh.exe " & strFile
         Else
            sCmdLine = "hh.exe " & strAppPath & "\" & strFile
            If Not fso.FileExists(strAppPath & "\" & strFile) Then
               moParent.RaiseError "Error in mwAPI.ShowMwHelp, Help File is Missing: " & _
                strAppPath & "\" & strFile
               Set fso = Nothing
               Exit Function
            End If
         End If
         
         If sCmdLine = "" Then
            hRet = 0
         Else
            hRet = WinExec(sCmdLine, SW_NORMAL)
         End If
         If hRet = 0 Then
            moParent.RaiseError "Error in mwAPI.ShowMwHelp, Calling HelpContextID in File: " & strFile
         End If
         
      Else
         hRet = RunShellExecute(strFile)
         '
         ' Empirical - 2=badboy, 42=goodboy ?
         '
         If hRet = 2 Then
            moParent.RaiseError "Error in mwAPI.ShowMwHelp, opening Help File: " & strFile
         End If
      End If
   Else
      If IsChmFile Then
         '
         ' Chm File... do context help call in...
         '
         strTopic = strAppPath & "\" & strFile & "::/" & HelpContextID
         'hRet = HtmlHelp(0, strTopic, HH_DISPLAY_TOPIC, ByVal &O0)
         hRet = HtmlHelp(0, strFile, HH_HELP_CONTEXT, HelpContextID)
         'sCmdLine = "hh.exe " & strFile & "::/" & HelpContextID
         'sCmdLine = "HtmlHelp(0, strFile, HH_HELP_CONTEXT, HelpContextID)
         'If sCmdLine = "" Then
         '   hRet = 0
         'Else
         '   hRet = WinExec(sCmdLine, SW_NORMAL)
         'End If
         If hRet = 0 Then
            moParent.RaiseError "Error in mwAPI.ShowMwHelp, Calling HelpContextID in File: " & strFile
         End If
      Else
         '
         ' Word File -Treat Context as a bookmark...
         '
         
      End If
   End If
   
      
   'hRet = HtmlHelp(0, strFile, HH_DISPLAY_TOPIC, ByVal 0&)
      'MsgBox hRet
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.ShowMwHelp. ", Err.Number, Err.Description
End Function


Private Function FetchHelpFileSourceName(HelpFileID) As String
   Dim loRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      strSQL = "select * from mwcHelpFile where HelpFileID='" & HelpFileID & "'"
      loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If loRs.RecordCount > 0 Then
         If Not IsNull(loRs!SourceFile) Then
            FetchHelpFileSourceName = loRs!SourceFile
         Else
            FetchHelpFileSourceName = ""
         End If
      Else
         FetchHelpFileSourceName = ""
      End If
      CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.FetchHelpFileSourceName. ", Err.Number, Err.Description
End Function



Public Function LaunchExternalLink(FileName As String, Optional HelpContextID As String)
   Dim strExtension As String
   Dim strTopic As String
   Dim strAcroRd32 As String * 255
   Dim hRet As Long
   Dim fso As FileSystemObject
   Dim dummy As String
   On Error GoTo FunctionError
   
   dummy = ""
   strAcroRd32 = Space(255)
   
   Set fso = New FileSystemObject
   strExtension = UCase(fso.GetExtensionName(FileName))
   '
   ' Does file exist ?
   '
   If Not fso.FileExists(FileName) Then
      ' Not a file, so just shell execute
      
'      RunCreateProcess FileName
      RunShellExecute FileName
      'moParent.RaiseError "Error in mwAPI.LaunchExternalLink, Help File is Missing: " & FileName
      Set fso = Nothing
      Exit Function
   End If
   '
   ' OK...
   '
 If strExtension = "CHM" And Trim(HelpContextID) <> "" And HelpContextID <> "-1" Then
         '
         ' Chm File... do context help call in...
         '
         strTopic = FileName & "::/" & HelpContextID
         'hRet = HtmlHelp(0, strTopic, HH_DISPLAY_TOPIC, ByVal &O0)
         hRet = HtmlHelp(0, FileName, HH_HELP_CONTEXT, HelpContextID)
         If hRet = 0 Then
            moParent.RaiseError "Error in mwAPI.LaunchExternalLink, Calling HelpContextID in File: " & FileName
         End If
   ElseIf strExtension = "PDF" And Trim(HelpContextID) <> "" And HelpContextID <> "-1" Then

      hRet = FindExecutable(FileName, dummy, strAcroRd32)
      strAcroRd32 = Trim(strAcroRd32)

'      strAcroRd32 = "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
      strTopic = " /A ""page=" & HelpContextID & "=OpenActions"" """ & FileName & """"
      hRet = RunExe(strAcroRd32, strTopic)
      '
      ' Empirical - 2=badboy, 42=goodboy ?
      '
      If hRet = 2 Then
         moParent.RaiseError "Error in mwAPI.LaunchExternalLink, opening File: " & FileName
      ElseIf hRet = 31 Then
         moParent.RaiseError "Can not open File: " & FileName & ", there is no program associated with that file extension."
      End If
   Else
'      hRet = RunCreateProcess(FileName)
      hRet = RunShellExecute(FileName)
      '
      ' Empirical - 2=badboy, 42=goodboy ?
      '
      If hRet = 2 Then
         moParent.RaiseError "Error in mwAPI.LaunchExternalLink, opening File: " & FileName
      ElseIf hRet = 31 Then
         moParent.RaiseError "Can not open File: " & FileName & ", there is no program associated with that file extension."
      End If
   End If
   Set fso = Nothing
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.LaunchExternalLink. ", Err.Number, Err.Description
End Function
Private Function RunExe(sFile As String, sParms As String) As Long
   Dim lresp As Long
   Dim dummy As Long
   On Error GoTo FunctionError
   dummy = 1


'   The RunCreateProcess function is a replacement for ShellExecute but it doesn't work yet.
'   Uncomment it and comment the ShellExecute call to test new function

   lresp = ShellExecute(dummy, "open", sFile, sParms, vbNullString, SW_SHOWNORMAL)
'   lresp = RunCreateProcess(sFile)
   RunExe = lresp
   Exit Function
FunctionError:
   moParent.RaiseError "Error in mwSession.mwAPI.RunExe", Err.Number, Err.Description
   RunExe = lresp
End Function

Public Function SleepApp(Milliseconds As Long)
   On Error GoTo FunctionError
   Sleep Milliseconds
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.SleepApp. ", Err.Number, Err.Description
End Function

Private Sub Class_Initialize()
   mWindowsVersion = GetWindowsVersion()
End Sub

'
' source \downloads.020\GetLocaleInfo
' 12/2002
Public Function GetLocaleString(ByVal WhichInfo As mwLocaleStrings) As String
   ' this is the function where all of the work is done.
   Dim sFormatString$, nStrLen&
   Dim InfoConstant As Long
   On Error GoTo FunctionError
   Select Case WhichInfo
      Case Is = LOCALE_TIME_SEPARATOR
         InfoConstant = LOCALE_STIME
         'InfoConstant = LOCALE_SDECIMAL
      Case Is = LOCALE_DATE_SEPARATOR
         InfoConstant = LOCALE_SDATE
      Case Is = LOCALE_SHORT_DATE
         InfoConstant = LOCALE_SSHORTDATE
      Case Is = LOCALE_LONG_DATE
         InfoConstant = LOCALE_SLONGDATE
      Case Is = LOCALE_TIME
         InfoConstant = LOCALE_STIMEFORMAT
      Case Is = LOCALE_DECIMAL_SEPARATOR
         InfoConstant = LOCALE_SDECIMAL
      Case Is = LOCALE_THOUSANDS_SEPARATOR
         InfoConstant = LOCALE_STHOUSAND
         
      'BY N.Angelakis1 April 2011
      'VEL-207 Printing adjusted Inventories doesnt Print all the fields in A4 size paper
      Case Is = LOCALE_MEASUREMENT
         InfoConstant = LOCALE_SMEASUREMENT
   
   End Select
  ' get the length of the format string LOCALE_USER_DEFAULT
  'nStrLen = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, InfoConstant, sFormatString, API_FALSE)
  nStrLen = GetLocaleInfo(LOCALE_USER_DEFAULT, InfoConstant, sFormatString, API_FALSE)
  ' if the return wasn't zero...
  If nStrLen > 0 Then
    ' allocate string space and get the format string
    sFormatString = String$(nStrLen + 1, vbNullChar)
    'nStrLen = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, InfoConstant, sFormatString, Len(sFormatString))
    nStrLen = GetLocaleInfo(LOCALE_USER_DEFAULT, InfoConstant, sFormatString, Len(sFormatString))
    
    ' if the return wasn't zero...
    If nStrLen > 0 Then
      ' strip any nulls and convert if necessarry
      sFormatString = StripNulls(sFormatString)
  
      ' check to see if the string is unicode
      If LenB(sFormatString) > 1 Then
        Dim aBytes(1) As Byte
      
        CopyMemory ByVal VarPtr(aBytes(0)), ByVal StrPtr(sFormatString), 2&
        
        If (aBytes(0) <> 0) And (aBytes(1) <> 0) Then sFormatString = StrConv(sFormatString, vbUnicode)
      
      Else
        sFormatString = StrConv(sFormatString, vbUnicode)
      End If
      
      ' return the string
      GetLocaleString = sFormatString
    End If
  End If
  Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.GetLocaleInfo. ", Err.Number, Err.Description
End Function


Public Function StripNulls(ByVal sText As String) As String
  ' strips any nulls from the end of a string
  Dim nPosition&
  
  StripNulls = sText
  
  nPosition = InStr(sText, vbNullChar)
  If nPosition Then StripNulls = Left$(sText, nPosition - 1)
  If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = vbZLString
End Function

Public Function UserSessionSettingSet(KeyName As String, KeyValue As String)
   Dim loReg As Registry
   Set loReg = New Registry
   loReg.BaseRegistry = USER_SESSION_BASE_REGISTRY
   loReg.SetReg KeyName, KeyValue
   Set loReg = Nothing
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.UserSessionSettingSet. ", Err.Number, Err.Description

End Function

Public Function UserSessionSettingGet(KeyName As String) As String
   Dim loReg As Registry
   Set loReg = New Registry
   loReg.BaseRegistry = USER_SESSION_BASE_REGISTRY
   UserSessionSettingGet = loReg.GetReg(KeyName)
   Set loReg = Nothing
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.UserSessionSettingGet. ", Err.Number, Err.Description

End Function

Public Function SmartCheck_MemMark()
   Dim lResult As Long
   On Error GoTo FunctionError
   lResult = NMMemMark
   If lResult = -1 Then
       goSession.LogIt mwl_Workstation, mwl_Warning, "Smartcheck debugger code enabled, no Smartcheck loaded."
       Exit Function
   End If
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.SmartCheck_MemMark. ", Err.Number, Err.Description
End Function

Public Function SmartCheck_MemShow()
   Dim lResult As Long
   On Error GoTo FunctionError
   lResult = NMMemPopup
   If lResult = -1 Then
       goSession.LogIt mwl_Workstation, mwl_Warning, "Smartcheck debugger code enabled, no Smartcheck loaded."
       Exit Function
   End If
    
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.SmartCheck_MemShow. ", Err.Number, Err.Description
End Function

Public Function SmartCheck_MemSave()
   Dim lResult As Long
   Dim strFile As String
   On Error GoTo FunctionError
   strFile = goSession.GetAppPath & "\SmartCheck" & Format(Now(), "yyyymmddhhmmss")
   lResult = NMMemSave(strFile, "", False)
   If lResult = -1 Then
       goSession.LogIt mwl_Workstation, mwl_Warning, "Smartcheck debugger code enabled, no Smartcheck loaded."
       Exit Function
   End If
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.SmartCheck_MemSave. ", Err.Number, Err.Description
End Function


Public Function GetDisplayDateFormat(Optional ShortDate As Boolean) As String
   Dim i As Integer
   Dim s As String
   On Error GoTo FunctionError
   If ShortDate Then
      '
      s = GetLocaleString(LOCALE_SHORT_DATE)
      i = InStr(1, s, GetLocaleString(LOCALE_DATE_SEPARATOR))
      i = InStr(i + 1, s, GetLocaleString(LOCALE_DATE_SEPARATOR))
      If UCase(Left(GetLocaleString(LOCALE_SHORT_DATE), 2)) = "YY" Then
         ' Swedish or other year leading countries...
         'GetDisplayDateFormat = "dd-MM HH:mm"
         GetDisplayDateFormat = mID$(GetLocaleString(LOCALE_SHORT_DATE), 6) & " HH:mm"
      Else
         GetDisplayDateFormat = Left(GetLocaleString(LOCALE_SHORT_DATE), i - 1) & " HH:mm"
      End If
      
   Else
      GetDisplayDateFormat = GetLocaleString(LOCALE_SHORT_DATE) & " HH:mm"
   End If
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwAPI.GetDisplayDateFormat: ", Err.Number, Err.Description
End Function


Public Function GetDisplayCurrencyFormat() As String
   Dim s As String
   On Error GoTo FunctionError
   GetDisplayCurrencyFormat = "###" & GetLocaleString(LOCALE_THOUSANDS_SEPARATOR) _
    & "###" & GetLocaleString(LOCALE_THOUSANDS_SEPARATOR) _
    & "##0" & GetLocaleString(LOCALE_DECIMAL_SEPARATOR) & "00"
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwAPI.GetDisplayCurrencyFormat: ", Err.Number, Err.Description
End Function



Private Function Convert2MwsDateTime(DateString As String) As String

   Dim s As String
   Dim nDay As Integer
   Dim nNowDay As Integer
   Dim nMonth As Integer
   Dim nYear As Integer
   Dim nHour As Integer
   Dim nMinute As Integer
   Dim sDate As String
   Dim sTimeSep As String
   Dim sDateSep As String
   Dim nDateSep As Integer
   Dim sDisplayHelp As String
   On Error GoTo FunctionError
   sTimeSep = GetLocaleString(LOCALE_TIME_SEPARATOR)
   sDateSep = GetLocaleString(LOCALE_DATE_SEPARATOR)
   s = DateString
   ' Check mm/dd hh:mm... or dd/mm of course...
   If IsDate(s) And (InStr(1, s, sDateSep) = 3 Or InStr(1, s, sDateSep) = 2) Then
      If InStr(1, s, sTimeSep) = 0 Then
         MsgBox "Date entry must include time separator (example: 2300 should be 23" & sTimeSep & "00)", , "Date Entry Error"
         Convert2MwsDateTime = ""
         Exit Function
      End If
      nYear = DatePart("yyyy", s)
      nMonth = DatePart("m", s)
      nDay = DatePart("d", s)
      nHour = DatePart("h", s)
      nMinute = DatePart("n", s)
      Convert2MwsDateTime = DateSerial(nYear, nMonth, nDay) & " " & Format(LTrim(str(nHour)), "00") & sTimeSep & Format(LTrim(str(nMinute)), "00")
      Exit Function
   End If
   ' Check hhmm/dd
   nDateSep = InStr(1, s, "/")
   If nDateSep = 0 Then
      MsgBox sDisplayHelp, , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   ElseIf nDateSep <> 5 Then
      MsgBox "Time must be exactly four digits (example: 9 AM must be 0900).", , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   ElseIf Not IsNumeric(Left(s, 4)) Then
      MsgBox "First four characters must be the time of day (example: 2130/12 is 9:30 PM on the 12th of the month).", , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   ElseIf CInt(Left(s, 4)) > 2359 Then
      MsgBox "First four characters must be between 0000 and 2359 (which is the time of day).", , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   ElseIf Not IsNumeric(mID(s, 6)) Then
      MsgBox "Last characters must be the day of month (example: 2130/12 is 9:30 PM on the 12th of the month).", , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   ElseIf mID(s, 3, 2) > 59 Then
      MsgBox "Minutes cannot be greater than 59.", , "Date Entry Error"
      Convert2MwsDateTime = ""
      Exit Function
   End If
   nHour = Left(s, 2)
   nMinute = mID(s, 3, 2)
   
   ' Calculate the month...
   nDay = mID(s, 6)
   nNowDay = CInt(Format(Now(), "dd"))
   If nDay = nNowDay Then
      nMonth = CInt(Format(Now(), "mm"))
      nYear = CInt(Format(Now(), "yyyy"))
   ElseIf (nDay < nNowDay) And ((nNowDay - nDay) > 15) Then
      ' must be next month
      nMonth = Format(DateAdd("m", 1, Now()), "mm")
      ' next year ?
      If nMonth < CInt(Format(Now(), "mm")) Then
         nYear = Format(DateAdd("yyyy", 1, Now()), "yyyy")     ' year needs to be yyyy not y
      Else
         nYear = CInt(Format(Now(), "yyyy"))
      End If
   ElseIf (nDay > nNowDay) And ((nDay - nNowDay) > 15) Then
      ' must be last month
      nMonth = Format(DateAdd("m", -1, Now()), "mm")
      ' last year ?
      If nMonth = 12 Then
         nYear = Format(DateAdd("yyyy", -1, Now()), "yyyy")    ' year needs to be yyyy not y
      Else
         nYear = CInt(Format(Now(), "yyyy"))
      End If
   Else
      ' assume this month
      nMonth = CInt(Format(Now(), "mm"))
      nYear = CInt(Format(Now(), "yyyy"))
   End If
   '
   ' Put it all together...
   '
   Convert2MwsDateTime = DateSerial(nYear, nMonth, nDay) & " " & Format(LTrim(str(nHour)), "00") & sTimeSep & Format(LTrim(str(nMinute)), "00")
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwAPI.SetDateTimeHelpString: ", Err.Number, Err.Description
   Convert2MwsDateTime = ""
End Function

Public Function DateChecker(DateString As String) As String
   
   Dim s As String
   Dim nDay As Integer
   Dim nNowDay As Integer
   Dim nMonth As Integer
   Dim nYear As Integer
   Dim nHour As Integer
   Dim nMinute As Integer
   Dim sDate As String
   Dim sTimeSep As String
   Dim sDateSep As String
   Dim nDateSep As Integer
   On Error GoTo FunctionError
   If Trim(DateString) = "" Then
      DateChecker = ""
      Exit Function
   End If
   sTimeSep = GetLocaleString(LOCALE_TIME_SEPARATOR)
   sDateSep = GetLocaleString(LOCALE_DATE_SEPARATOR)
   s = DateString
   '
   ' If time only, add "/d"
   '
   If Len(s) = 4 Then
      s = s & "/" & LTrim(Format(Now, "d"))
   End If
   
   
   ' Check mm/dd hh:mm... or dd/mm of course...
   ' or Swedish full date...
   If IsDate(s) And ((InStr(1, s, sDateSep) = 3 Or InStr(1, s, sDateSep) = 2) Or _
    IsDate(s) And (InStr(1, s, sDateSep) = 5 Or InStr(1, s, sDateSep) = 8) And Len(s) > 7) Then
      If InStr(1, s, sTimeSep) = 0 Then
         MsgBox "Date entry must include time separator (example: 2300 should be 23" & sTimeSep & "00)", , "Date Entry Error"
         DateChecker = ""
         Exit Function
      End If
      nYear = DatePart("yyyy", s)
      nMonth = DatePart("m", s)
      nDay = DatePart("d", s)
      nHour = DatePart("h", s)
      nMinute = DatePart("n", s)
      DateChecker = DateSerial(nYear, nMonth, nDay) & " " & Format(LTrim(str(nHour)), "00") & sTimeSep & Format(LTrim(str(nMinute)), "00")
      Exit Function
   
   End If
   ' Check hhmm/dd
   nDateSep = InStr(1, s, "/")
   If nDateSep = 0 Then
      MsgBox DateTimeHelpString, , "Date Entry Error"
      DateChecker = ""
      Exit Function
   ElseIf nDateSep <> 5 Then
      MsgBox "Time must be exactly four digits (example: 9 AM must be 0900).", , "Date Entry Error"
      DateChecker = ""
      Exit Function
   ElseIf Not IsNumeric(Left(s, 4)) Then
      MsgBox "First four characters must be the time of day (example: 2130/12 is 9:30 PM on the 12th of the month).", , "Date Entry Error"
      DateChecker = ""
      Exit Function
   ElseIf CInt(Left(s, 4)) > 2359 Then
      MsgBox "First four characters must be between 0000 and 2359 (which is the time of day).", , "Date Entry Error"
      DateChecker = ""
      Exit Function
   ElseIf Not IsNumeric(mID(s, 6)) Then
      MsgBox "Last characters must be the day of month (example: 2130/12 is 9:30 PM on the 12th of the month).", , "Date Entry Error"
      DateChecker = ""
      Exit Function
   ElseIf mID(s, 3, 2) > 59 Then
      MsgBox "Minutes cannot be greater than 59.", , "Date Entry Error"
      DateChecker = ""
      Exit Function
   End If
   nHour = Left(s, 2)
   nMinute = mID(s, 3, 2)
   
   ' Calculate the month...
   nDay = mID(s, 6)
   nNowDay = CInt(Format(Now(), "dd"))
   If nDay = nNowDay Then
      nMonth = CInt(Format(Now(), "mm"))
      nYear = CInt(Format(Now(), "yyyy"))
   ElseIf (nDay < nNowDay) And ((nNowDay - nDay) > 15) Then
      ' must be next month
      nMonth = Format(DateAdd("m", 1, Now()), "mm")
      ' next year ?
      If nMonth < CInt(Format(Now(), "mm")) Then
         nYear = Format(DateAdd("yyyy", 1, Now()), "yyyy")           '
      Else
         nYear = CInt(Format(Now(), "yyyy"))
      End If
   ElseIf (nDay > nNowDay) And ((nDay - nNowDay) > 15) Then
      ' must be last month
      nMonth = Format(DateAdd("m", -1, Now()), "mm")
      ' last year ?
      If nMonth = 12 Then
         nYear = Format(DateAdd("yyyy", -1, Now()), "yyyy")          '
      Else
         nYear = CInt(Format(Now(), "yyyy"))
      End If
   Else
      ' assume this month
      nMonth = CInt(Format(Now(), "mm"))
      nYear = CInt(Format(Now(), "yyyy"))
   End If
   '
   ' Put it all together...
   '
   DateChecker = DateSerial(nYear, nMonth, nDay) & " " & Format(LTrim(str(nHour)), "00") & sTimeSep & Format(LTrim(str(nMinute)), "00")
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwAPI.SetDateTimeHelpString: ", Err.Number, Err.Description
   DateChecker = ""
End Function


Private Function DateTimeHelpString() As String
   Dim sDF As String
   Dim iDay As Integer
   Dim iMonth As Integer
   Dim iYear As Integer
   Dim sSep As String
   Dim sHelp As String
   '
   ' Shit Meatball regional difference conversion tools...
   '
   sDF = goSession.API.GetLocaleString(LOCALE_SHORT_DATE) & " HH:MM"
   iDay = InStr(1, sDF, "d")
   iMonth = InStr(1, sDF, "M")
   iYear = InStr(1, sDF, "y")
   sSep = GetLocaleString(LOCALE_DATE_SEPARATOR)
   
   If iDay > iMonth Then
      ' Date is MM/DD
      sHelp = "mm" & sSep & "dd hh:mm"
   Else
      'Date is DD/MM
      sHelp = "dd" & sSep & "mm hh:mm"
   End If
   DateTimeHelpString = "Enter date as either hhmm/dd or " & UCase(sHelp)
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwAPI.DateTimeHelpString: ", Err.Number, Err.Description
End Function

Public Function MemoryStatusString() As String
   Dim lsMS As MEMORYSTATUS
   On Error GoTo FunctionError
   GlobalMemoryStatus lsMS
   MemoryStatusString = "% In Use: " & lsMS.dwMemoryLoad & ", Free Phy: " & _
    Format(lsMS.dwAvailPhys, "#,###") & ", Free User Bytes: " & Format(lsMS.dwAvailVirtual, "#,###")
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwAPI.MemoryStatusString: ", Err.Number, Err.Description
End Function


Public Function ShowVbFormHelp(VbFormName As String, Optional HelpContextID As String)
   Dim sFile As String
   Dim sAltFile As String
   Dim fso As New FileSystemObject
   Dim hRet As Long
   Dim loRs As Recordset
   Dim sSQL As String
   Dim CustomFolder As String
   Dim sHelpFileName As String
   Dim nHelpContextId As Long
  
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
  
  
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwcHelpContext where VbFormName='" & VbFormName & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount > 0 Then
      If Not IsNull(loRs!HelpFileName) Then
         sHelpFileName = loRs!HelpFileName
      Else
         sHelpFileName = VbFormName
      End If
   Else
      sHelpFileName = VbFormName
   End If
   '
   ' Does file exist ?
   '
   ' First see if there is a mwcThisSite.CustomHelpFolder value. If there is then
   ' look there first.
   '
   CustomFolder = goSession.ThisSite.GetExtendedProperty("CustomHelpFolder")
  
   If CustomFolder <> "" Then
      sFile = CustomFolder & "\" & sHelpFileName & ".chm"
      If Not fso.FileExists(sFile) Then
         sFile = CustomFolder & "\" & sHelpFileName & ".doc"
         If Not fso.FileExists(sFile) Then
            sFile = CustomFolder & "\" & sHelpFileName & ".docX"
            If Not fso.FileExists(sFile) Then
               sFile = CustomFolder & "\" & sHelpFileName & ".docM"
               If Not fso.FileExists(sFile) Then
                  sFile = CustomFolder & "\" & sHelpFileName & ".pdf"
                  If Not fso.FileExists(sFile) Then
                     sFile = goSession.GetAppPath & "\HelpManuals\" & sHelpFileName & ".chm"
                  End If
               End If
            End If
         End If
      End If
   Else
      sFile = goSession.GetAppPath & "\HelpManuals\" & sHelpFileName & ".chm"
   End If
  
   If Not fso.FileExists(sFile) Then
      If CustomFolder <> "" Then
         sAltFile = CustomFolder & "\mwUser.chm"
         If Not fso.FileExists(sAltFile) Then
            sAltFile = CustomFolder & "\mwUser.doc"
            If Not fso.FileExists(sAltFile) Then
               sAltFile = CustomFolder & "\mwUser.docX"
               If Not fso.FileExists(sAltFile) Then
                  sAltFile = CustomFolder & "\mwUser.docM"
                  If Not fso.FileExists(sAltFile) Then
                     sAltFile = CustomFolder & "\mwUser.pdf"
                     If Not fso.FileExists(sAltFile) Then
                        sAltFile = goSession.GetAppPath & "\HelpManuals\mwUser.chm"
                     End If
                  End If
               End If
            End If
         End If
      Else
         sAltFile = goSession.GetAppPath & "\HelpManuals\mwUser.chm"
      End If
      If fso.FileExists(sAltFile) Then
         goSession.LogIt mwl_Workstation, mwl_Information, "Could not find Help file : " & sFile & " Displaying " & sAltFile & " instead."
      
         hRet = RunShellExecute(sAltFile)
         '
         ' Empirical - 2=badboy, 42=goodboy ?
         '
         If hRet = 2 Then
            moParent.RaiseError "Error in mwAPI.ShowVbFormHelp, opening Help File: " & sAltFile
            ShowVbFormHelp = -1
         End If
      End If
   Else
      '
      ' OK...
      '
      If Trim(HelpContextID) = "" And (loRs.RecordCount = 0 Or IsNull(loRs!HelpContextID)) Then
   '      hRet = RunCreateProcess(sFile)
         goSession.LogIt mwl_Workstation, mwl_Information, "Displaying Help file : " & sFile
         hRet = RunShellExecute(sFile)
         '
         ' Empirical - 2=badboy, 42=goodboy ?
         '
         If hRet = 2 Then
            moParent.RaiseError "Error in mwAPI.ShowVbFormHelp, opening Help File: " & sFile
            ShowVbFormHelp = -1
         End If
      Else
         '
         ' Chm File... do context help call in...
         '
         If Trim(HelpContextID) = "" Then
            '
            'Microsoft HtmlHelp does not properly open a CHM when the loRs!HelpContextID is passed directly as an argument.
            'You must assign a variable as a long and pass the variable for HtmlHelp to correctly open the CHM using a HelpContextId.
            '
            If IsRecordLoaded(loRs) Then
               nHelpContextId = loRs!HelpContextID
            End If
         Else
            nHelpContextId = HelpContextID
         End If
         
         hRet = HtmlHelp(0, sFile, HH_HELP_CONTEXT, nHelpContextId)
         If hRet = 0 Then
            
            goSession.LogIt mwl_Workstation, mwl_Information, "Displaying Help file : " & sFile & " the HelpContextID " & nHelpContextId & " can not be found in that manual."
            hRet = RunShellExecute(sFile)
            '
            ' Empirical - 2=badboy, 42=goodboy ?
            '
            If hRet = 2 Then
               moParent.RaiseError "Error in mwAPI.ShowVbFormHelp, opening Help File: " & sFile
               ShowVbFormHelp = -1
            End If
         End If
      End If
   End If
   ShowVbFormHelp = 0
   CloseRecordset loRs
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.ShowVbFormHelp. ", Err.Number, Err.Description
   ShowVbFormHelp = -1
   CloseRecordset loRs
End Function


Public Function GetMachName() As String
   Dim sSystemName As String * 255   ' Name of this computer system.
   Dim lSystemNameLength As Long     ' Length of string that contains the name of this system.
   Dim lResult As Long
   
   On Error GoTo FunctionError
   
   lSystemNameLength = 255
   sSystemName = Space(lSystemNameLength)
  
   lResult = GetComputerName(sSystemName, lSystemNameLength)

   ' Return value of zero means the call failed; test for this before
   ' continuing.
   If (lResult = 0) Then
      goSession.RaiseError "Error: Unable to Retrieve the Current Computer Name"
      GetMachName = ""
      Exit Function
   End If
   
   GetMachName = Trim(sSystemName)
   GetMachName = Left(GetMachName, Len(GetMachName) - 1)    ' Strip off the silly C++ trailing "00"
   Exit Function
FunctionError:
   goSession.RaiseError "Error in GetMachName: " & Err.Number & "-" & Err.Description
   GetMachName = ""
End Function

' Walks all the process heaps and determines the total size of committed memory , uncommitted
' memory, and allocated memory elements

Public Function MemoryStats() As Long
   Dim hHeaps() As Long
   Dim lHeapCount As Long
   Dim iHeap As Integer
   Dim tDetails As PROCESS_HEAP_ENTRY

   Dim lCommitted As Long
   Dim lUncommitted As Long
   Dim lMemInUse As Long
   
   ' Initialise the returned counts
   lCommitted = 0
   lUncommitted = 0
   lMemInUse = 0

   ' Get count of existing process heaps
   ReDim hHeaps(0)
   lHeapCount = GetProcessHeaps(0, ByVal VarPtr(hHeaps(0)))
   ' Read an array of handles for the heaps
   ReDim hHeaps(0 To lHeapCount - 1)
   lHeapCount = GetProcessHeaps(lHeapCount, ByVal VarPtr(hHeaps(0)))

   ' Visit each heap to get its memory details
   For iHeap = 0 To lHeapCount - 1
      tDetails.lpData = 0
      Do While HeapWalk(hHeaps(iHeap), tDetails) <> 0
         If tDetails.wFlags And PROCESS_HEAP_REGION Then
            lCommitted = lCommitted + tDetails.dwCommittedSize
            lUncommitted = lUncommitted + tDetails.dwUnCommittedSize
         ElseIf tDetails.wFlags And PROCESS_HEAP_ENTRY_BUSY Then
            lMemInUse = lMemInUse + tDetails.cbData + tDetails.cbOverhead
         End If
      Loop
   Next iHeap
    
   MemoryStats = lMemInUse
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.MemoryStats. ", Err.Number, Err.Description
   MemoryStats = 0
End Function

Public Function GetOSUserName() As String
   Dim sUserName         As String
   Dim nlength          As Long
   Dim nretval           As Long

   On Error GoTo FunctionError
   
   nlength = 255
   sUserName = Space$(nlength)

   nretval = GetUserName(sUserName, nlength)
   sUserName = Left$(sUserName, nlength - 1)
   GetOSUserName = sUserName
   
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwAPI.GetOSUserName: ", Err.Number, Err.Description
   Exit Function
End Function

Public Function GetLongDomainName() As String
   Dim loSystemInfo As Object
   On Error Resume Next
   
   Set loSystemInfo = CreateObject("AdSystemInfo")
   If Not loSystemInfo Is Nothing Then
      GetLongDomainName = loSystemInfo.DomainDNSName
   End If
   Set loSystemInfo = Nothing
End Function

Public Function FleetAppRegSettingGet(KeyName As String) As String
   Dim loReg As Registry
   Set loReg = New Registry
   loReg.BaseRegistry = FLEET_APP_BASE_REGISTRY
   FleetAppRegSettingGet = loReg.GetLMReg(KeyName)
   Set loReg = Nothing
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwAPI.FleetAppRegSettingGet. ", Err.Number, Err.Description

End Function

