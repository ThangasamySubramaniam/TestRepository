VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ConfigGroups"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private moConfigGroups As ConfigGroups
Private moErrors As mwErrors

Public Property Get ConfigGroups() As ConfigGroups
   On Error GoTo PropertyError
    Set ConfigGroups = moConfigGroups
    Exit Property
PropertyError:
   'MsgBox "here"
End Property


Public Property Set ConfigGroups(vData As ConfigGroups)
    Set moConfigGroups = vData
End Property

Public Property Get Errors() As mwErrors
    Set Errors = moErrors
End Property


Public Property Set Errors(vData As mwErrors)
    Set moErrors = vData
End Property

Public Function LoadINIConfig(strFile As String) As Boolean
   Dim i As Integer
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strBuffer As String
   Dim strElements() As String
   Dim strGroup As String
   
   On Error GoTo LoadINIConfig_error
   If moConfigGroups Is Nothing Then
      Set moConfigGroups = New ConfigGroups
   End If
   If Not fso.FileExists(strFile) Then
      RaiseErrorL "INI File Missing: " & strFile, err
      LoadINIConfig = False
      Exit Function
   End If
   Set ts = fso.OpenTextFile(strFile)
   Do While Not ts.AtEndOfStream
      strBuffer = ts.ReadLine
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
         Set oGroup = moConfigGroups.Add(strGroup, strGroup)
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
   Set ts = Nothing
   Set fso = Nothing
   LoadINIConfig = True
   Exit Function
   
   
LoadINIConfig_error:
   RaiseErrorL "Error processing Site INI File: " & strFile, err
   LoadINIConfig = False
   If Not ts Is Nothing Then Set ts = Nothing
   If Not fso Is Nothing Then Set fso = Nothing
End Function

' 2/24/2001 - Write Site INI information to specified file...
Public Function WriteINIConfig(strFile As String) As Boolean
   Dim i As Integer
   Dim fso As New FileSystemObject
   Dim ts As TextStream
   Dim oGroup As ConfigGroup
   Dim oKey As ConfigKey
   Dim strBuffer As String
   Dim strElements() As String
   Dim strGroup As String
   Dim strBakFile As String
   
   On Error GoTo WriteINIConfig_error
   If moConfigGroups Is Nothing Then
      WriteINIConfig = False
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
   For Each oGroup In moConfigGroups
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
   WriteINIConfig = True
   Exit Function
WriteINIConfig_error:
   RaiseErrorL "Error Writing Site INI File: " & strFile, err
   WriteINIConfig = False
   If Not fso Is Nothing Then Set fso = Nothing
   If Not ts Is Nothing Then Set ts = Nothing
End Function


Friend Function RaiseErrorL(Msg As String, ByRef vbErr As ErrObject)
   Dim oErr As mwError
   If moErrors Is Nothing Then
      Set moErrors = New mwErrors
   End If
   Set oErr = moErrors.Add
  oErr.mwErrorCode = mwe_GENERAL_ERROR
   oErr.mwErrorDescription = Msg
   oErr.mwErrorSource = "mwSession.Session"
   oErr.VbErrorNumber = vbErr.Number
   oErr.VbErrorDescription = vbErr.Description
   Set oErr = Nothing

End Function

Private Function MassageKey(ByVal str As String) As String
   str = LTrim(str)
   str = RTrim(str)
   MassageKey = str
End Function


Private Sub Class_Initialize()
   Set moConfigGroups = New ConfigGroups
End Sub

Private Sub Class_Terminate()
   Set moConfigGroups = Nothing

End Sub

Public Function GetAppKey(Group As String, Key As String) As String
   On Error GoTo GetAppKey_error
   GetAppKey = moConfigGroups(UCase(Group)).ConfigKeys(UCase(Key)).KeyValue
   Exit Function
GetAppKey_error:
   RaiseErrorL "Key Missing. Group=" & Group & ", Key=" & Key, err
   GetAppKey = ""
End Function

Public Function SetAppKey(Group As String, Key As String, NewKey As String) As Boolean
   Dim oErr As mwError
   On Error GoTo SetAppKey_error
   moConfigGroups(UCase(Group)).ConfigKeys(UCase(Key)).KeyValue = NewKey
   SetAppKey = True
   Exit Function
SetAppKey_error:
   RaiseErrorL "Key Missing. Group=" & Group & ", Key=" & Key, err
   SetAppKey = ""
End Function


