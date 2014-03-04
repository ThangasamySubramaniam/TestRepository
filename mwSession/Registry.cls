VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"regErrors"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' WorkFlow Desktop Registry Class - Support get and set registry
' 11/23/2000 ms
'

Option Explicit


   '********** Registry WIN32 API Staging  *******************************
   Private Const REG_SZ As Long = 1
   Private Const REG_DWORD As Long = 4

   Private Const HKEY_CLASSES_ROOT = &H80000000
   Private Const HKEY_CURRENT_USER = &H80000001
   Private Const HKEY_LOCAL_MACHINE = &H80000002
   Private Const HKEY_USERS = &H80000003

   Private Const ERROR_NONE = 0
   Private Const ERROR_BADDB = 1
   Private Const ERROR_BADKEY = 2
   Private Const ERROR_CANTOPEN = 3
   Private Const ERROR_CANTREAD = 4
   Private Const ERROR_CANTWRITE = 5
   Private Const ERROR_OUTOFMEMORY = 6
   Private Const ERROR_ARENA_TRASHED = 7
   Private Const ERROR_ACCESS_DENIED = 8
   Private Const ERROR_INVALID_PARAMETERS = 87
   Private Const ERROR_NO_MORE_ITEMS = 259

   Private Const KEY_ALL_ACCESS = &H3F

   Private Const REG_OPTION_NON_VOLATILE = 0

   Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
   Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
   Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
   Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
   Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
   Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
   
   '********* End of Staging *******************************

   Private mBaseRegistry As String
   Private moErrors As mwErrors


'******* Public Interface *******************************

Public Property Let BaseRegistry(ByVal Data As String)
    mBaseRegistry = Data
End Property

Public Property Get BaseRegistry() As String
    BaseRegistry = mBaseRegistry
End Property

'Public Property Set Errors(vData As mwErrors)
'    Set moErrors = vData
'End Property

Public Property Get Errors() As mwErrors
    If moErrors Is Nothing Then
        Set moErrors = New mwErrors
    End If

    Set Errors = moErrors
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

Public Function GetRegInt(Key As String, DefaultValue As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(Key)
   If Temp = "" Then
      GetRegInt = SetReg(Key, DefaultValue)
   Else
      GetRegInt = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.GetRegInt. ", Err.Number, Err.Description
End Function
Public Function GetRegLong(Key As String, DefaultValue As Long) As Long
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(Key)
   If Temp = "" Then
      GetRegLong = SetReg(Key, DefaultValue)
   Else
      GetRegLong = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.GetRegLong. ", Err.Number, Err.Description
End Function
Public Function GetRegStr(Key As String, DefaultValue As String) As String
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(Key)
   If Temp = "" Then
      GetRegStr = SetReg(Key, DefaultValue)
   Else
      GetRegStr = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.GetRegStr. ", Err.Number, Err.Description
End Function

Public Function GetReg(Key As String) As String
   Dim oErr As mwError
   
   GetReg = GetRegistryString(HKEY_CURRENT_USER, mBaseRegistry, Key)
   If Not CheckRegistrySetting(GetReg) Then
      GetReg = ""
      'Set oErr = moErrors.Add()
      'oErr.mwErrorCode = mwe_GENERAL_ERROR
      'oErr.mwErrorDescription = "Key Missing: " & key
   End If
End Function

Public Function SetReg(ByVal Key As String, ByVal value As String) As String
   SetKeyValue mBaseRegistry, Key, value, REG_SZ
   SetReg = value
End Function


Public Function SetHKLMRegKey(BaseRegistry As String, ByVal Key As String, ByVal value As String) As Boolean
   Dim lRetVal As Long         'result of the SetValueEx function
   Dim hKey As Long            'handle of open key
   On Error GoTo FunctionError
    
   SetHklmKeyValue BaseRegistry, Key, value, REG_SZ
   SetHKLMRegKey = True
   Exit Function
FunctionError:
   SetHKLMRegKey = False
End Function

' Functions for getting & setting grid column widths & positions.

Public Function ugGetWidth(GridName As String, FieldNum As Integer, DefaultWidth As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(GridName & "." & FieldNum & ".Width")
   If Temp = "" Then
      ugGetWidth = ugSetWidth(GridName, FieldNum, DefaultWidth)
   Else
      ugGetWidth = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugGetWidth. ", Err.Number, Err.Description

End Function

Public Function ugSetWidth(GridName As String, FieldNum As Integer, Width As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   SetReg GridName & "." & FieldNum & ".Width", Width
   
   ugSetWidth = Width
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugSetWidth. ", Err.Number, Err.Description

End Function

Public Function ugGetPosition(GridName As String, FieldNum As Integer, DefaultPos As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(GridName & "." & FieldNum & ".Position")
   If Temp = "" Then
      ugGetPosition = ugSetPosition(GridName, FieldNum, DefaultPos)
   Else
      ugGetPosition = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugGetPosition. ", Err.Number, Err.Description

End Function

Public Function ugSetPosition(GridName As String, FieldNum As Integer, Position As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   SetReg GridName & "." & FieldNum & ".Position", Position
   
   ugSetPosition = Position
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugSetPosition. ", Err.Number, Err.Description

End Function

Public Function ugGetRowHeight(GridName As String, BandNum As Integer, DefaultHeight As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   Temp = GetReg(GridName & "." & BandNum & ".Height")
   If Temp = "" Then
      ugGetRowHeight = ugSetRowHeight(GridName, BandNum, DefaultHeight)
   Else
      ugGetRowHeight = Temp
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugGetRowHeight. ", Err.Number, Err.Description

End Function

Public Function ugSetRowHeight(GridName As String, BandNum As Integer, Height As Integer) As Integer
   Dim Temp As String
   
   On Error GoTo FunctionError
   SetReg GridName & "." & BandNum & ".Height", Height
   
   ugSetRowHeight = Height
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General error in mwSession.Registry.ugSetRowHeight. ", Err.Number, Err.Description

End Function






'************ End of Public Interface ***********************

'Public Property Set Errors(vData As mwErrors)
'    Set moErrors = vData
'End Property

Private Sub Class_Initialize()
    Set moErrors = New mwErrors
End Sub

Private Sub Class_Terminate()
    Set moErrors = Nothing
End Sub


Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
             sValue = String(cch, 0)
             lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
             If lrc = ERROR_NONE Then
                 vValue = Left$(sValue, cch - 1)
             Else
                 vValue = Empty
             End If
        ' For DWORDS
        Case REG_DWORD:
             lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
             If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
             lrc = -1
    End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function
QueryValueExError:
       Resume QueryValueExExit
End Function
   
Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub

Private Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
   Dim lRetVal As Long         'result of the SetValueEx function
   Dim hKey As Long            'handle of open key

   'open the specified key
   CreateNewKey sKeyName, HKEY_CURRENT_USER
    
   lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
   lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
   RegCloseKey (hKey)
End Sub

Private Sub SetHklmKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
   Dim lRetVal As Long         'result of the SetValueEx function
   Dim hKey As Long            'handle of open key

   'open the specified key
   CreateNewKey sKeyName, HKEY_LOCAL_MACHINE
   
   lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
   lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
   RegCloseKey (hKey)
End Sub





Private Sub QueryValue(sKeyName As String, sValueName As String)
    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long            'handle of opened key
    Dim vValue As Variant       'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    MsgBox vValue
    RegCloseKey (hKey)
End Sub

Private Function GetRegistryString(hKey As Long, strKeyName As String, strValueName As String) As String
    Dim lRetVal As Long
    Dim phkResult As Long
    Dim vValue As Variant
    
    lRetVal = RegOpenKeyEx(hKey, strKeyName, 0, KEY_ALL_ACCESS, phkResult)
    lRetVal = QueryValueEx(phkResult, strValueName, vValue)
    RegCloseKey (hKey)
    
    GetRegistryString = vValue
End Function

Private Function CheckRegistrySetting(ByVal strValue As String) As Boolean
    Dim bResult As Boolean
    
    bResult = True
    If IsNull(strValue) Or strValue = "" Then bResult = False
    
    CheckRegistrySetting = bResult
End Function


Public Function GetLMReg(Key As String) As String
   Dim oErr As mwError
   
   GetLMReg = GetRegistryString(HKEY_LOCAL_MACHINE, mBaseRegistry, Key)
   If Not CheckRegistrySetting(GetLMReg) Then
      GetLMReg = ""
   End If
End Function
