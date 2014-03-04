VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LogItems"
' Logger - Class allows maintenance of a collection of LogItems, with tools to do stuff.
' 11/28/2000 ms
'



Option Explicit
' Logger Class - For logging to persistent log file...

' Logger File Objects
Private fso As FileSystemObject
Private ts As TextStream
Private mLogMsg As String
Private mLogFile As String
Private moLogItems As mwLogItems
Private mUpdateMethod As mwl_UpdateMethod
Private moError As mwError
Private moParent As Session

Public Enum mwl_UpdateMethod
   mwl_AutomaticFileUpdate
   mwl_ManualFileUpdate
End Enum
   
Public Enum mwl_LogMessageType
   mwl_Workstation = 1
   mwl_Transport = 2
   mwl_User_Defined = 3
End Enum

Public Enum mwl_LogMessageCriticality
   mwl_Information = 1
   mwl_Warning = 2
   mwl_Critical = 3
End Enum

Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Get LogItems() As mwLogItems
    Set LogItems = moLogItems
End Property

Public Property Set LogItems(vData As mwLogItems)
    Set moLogItems = vData
End Property

Private Sub Class_Initialize()
   On Error GoTo SubError
   Set moLogItems = New mwLogItems
   Set fso = New FileSystemObject
   mUpdateMethod = mwl_AutomaticFileUpdate
   Exit Sub
SubError:
   MsgBox " General Error in mwSession.logger.Class_Initialize: " & Err.Number & " - " & Err.Description
End Sub

Private Sub Class_Terminate()
    Set moLogItems = Nothing
    Set fso = Nothing
End Sub

Public Function LogIt(msgType As mwl_LogMessageType, Criticality As mwl_LogMessageCriticality, _
  Msg As String, Optional oErr As mwError) As Boolean
   On Error GoTo FunctionError
   If oErr Is Nothing Then
      moLogItems.Add msgType, Criticality, Msg
   Else
      moLogItems.Add msgType, Criticality, Msg, oErr
   End If
   'Whether to write to disk...
   If mUpdateMethod = mwl_AutomaticFileUpdate Then
      If Not UpdateLogFileItem() Then
         LogIt = False
      Else
         LogIt = True
      End If
   End If
   Exit Function
FunctionError:
   MsgBox " General Error in mwSession.logger.logit: " & Err.Number & " - " & Err.Description
   LogIt = False
   
End Function


Private Function UpdateLogFileItem() As Boolean
   Dim strCriticality
   Dim i As Long
   Dim fso As FileSystemObject
   Dim strMsg As String
   Dim s As String
   Dim s1 As String
   Dim s2 As String
   Dim s3 As String
   Dim Stage As Integer
   
   On Error GoTo UpdateLogFileItem_error
Stage = 1
   Set fso = New FileSystemObject
Stage = 2
   i = moLogItems.Count
Stage = 3
   If Trim(mLogFile) = "" Then
      If moParent Is Nothing Or moParent.User Is Nothing Then
         mLogFile = "c:\MaritimeWorkstation.Log"
         If Trim(Environ("APPDATA")) <> "" Then
            mLogFile = Trim(Environ("APPDATA")) & "\MaritimeWorkstation.Log"
         End If
      Else
         If Trim(moParent.User.LogFile) = "" Then
            mLogFile = "c:\MaritimeWorkstation.Log"
            If Trim(Environ("APPDATA")) <> "" Then
               mLogFile = Trim(Environ("APPDATA")) & "\MaritimeWorkstation.Log"
            End If
         Else
            mLogFile = moParent.User.LogFile
         End If
      End If
'      s2 = fso.GetExtensionName(mLogFile)
'      s1 = Left(mLogFile, (Len(mLogFile) - Len(s2) - 1))
'      mLogFile = s1 & "_" & Format(Now(), "yyyymmdd") & "." & s2
   End If
Stage = 4
   If Not goSession Is Nothing Then
      If goSession.IsWFAProcessNow = True Then
         If PathIsDirectory(Parent.SiteRoot) = 0 Then
            
            mLogFile = "c:\MaritimeWorkstation.Log"
            If Trim(Environ("APPDATA")) <> "" Then
               mLogFile = Trim(Environ("APPDATA")) & "\MaritimeWorkstation.Log"
            End If
            
            moParent.KillObject ts
            Set ts = fso.OpenTextFile(mLogFile, ForAppending, True)
            If Not ts Is Nothing Then
               With moLogItems(i)
                  strMsg = "ShipNet Fleet, Version " & App.Major & "." & _
                           App.Minor & "." & Left(Format(App.Revision, "0000"), 1) & " Build " & CStr(Val(mID(Format(App.Revision, "0000"), 2, 2))) & _
                           ", User: " & goSession.User.UserID
                  s = (.MessageType & "|C|" & Format(.DateTime, "yyyy-mm-dd hh:mm:ss") & "|" & moParent.Site.SiteID & "|" & strMsg & "|")
                  ts.WriteLine s
                  
                  strMsg = SM(.Message)
                  s = (.MessageType & "|C|" & Format(.DateTime, "yyyy-mm-dd hh:mm:ss") & "|" & moParent.Site.SiteID & "|" & strMsg & "|")
                  ts.WriteLine s
                  
                  strMsg = "Shutting down Fleet Application due to network issue(s)."
                  s = (.MessageType & "|C|" & Format(.DateTime, "yyyy-mm-dd hh:mm:ss") & "|" & moParent.Site.SiteID & "|" & strMsg & "|")
                  ts.WriteLine s
                   
                  strMsg = "ShipNet Fleet Abnormal Termination for " & goSession.User.UserID & " - " & Now()
                  s = (.MessageType & "|C|" & Format(.DateTime, "yyyy-mm-dd hh:mm:ss") & "|" & moParent.Site.SiteID & "|" & strMsg & "|")
                  ts.WriteLine s
               End With
            End If
            
            moParent.KillObject ts
            goSession.RequestExitProgram
         End If
      End If
   End If
Stage = 5
   If fso.FileExists(mLogFile) Then
      If (fso.GetFile(mLogFile).Attributes And ReadOnly) Then
         MsgBox "mwSession.Logger.UpdateLogFileItem: Log file " & mLogFile & " exists but is ReadOnly and can not be opened for writing."
         moParent.KillObject fso
         UpdateLogFileItem = True
         Exit Function
      End If
   End If
Stage = 6
   
   On Error Resume Next
   Set ts = fso.OpenTextFile(mLogFile, ForAppending, True)
   
   If ts Is Nothing Then
      MsgBox "mwSession.Logger.UpdateLogFileItem: Log file " & mLogFile & " could not be opened for writing: " & Err.Number & "-" & Err.Description
      moParent.KillObject fso
      UpdateLogFileItem = True
      Exit Function
   End If
   
   On Error GoTo UpdateLogFileItem_error
   
   With moLogItems(i)
Stage = 7
      Select Case .Criticality
         Case Is = mwl_Information
            strCriticality = "I"
         Case Is = mwl_Warning
            strCriticality = "W"
         Case Is = mwl_Critical
            strCriticality = "C"
      End Select
Stage = 8
      strMsg = SM(.Message)
Stage = 9
      s = (.MessageType & "|" & strCriticality & "|" & Format(.DateTime, _
        "yyyy-mm-dd hh:mm:ss") & "|" & moParent.Site.SiteID & "|" & strMsg & "|")
Stage = 10
      ts.WriteLine s
Stage = 11
      If .Error.mwErrorCode <> mwe_NO_ERROR Then
Stage = 12
         ts.WriteLine "E|" & Format(.Error.ErrorDateTime, _
           "yyyy-mm-dd hh:mm:ss") & "|" & .Error.mwErrorCode & "|" _
           & SM(.Error.mwErrorDescription) & "|" & SM(.Error.mwErrorSource) _
           & "|" & SM(.Error.VbErrorDescription) & "|" & .Error.VbErrorNumber _
           & "|" & SM(.Error.VbErrorSource) & "|"
         .IsLoggedToFile = True
      End If
    End With
Stage = 13
    ts.Close
Stage = 14
    moParent.KillObject ts
Stage = 15
    moParent.KillObject fso
Stage = 16
    UpdateLogFileItem = True
Stage = 17
    Exit Function
UpdateLogFileItem_error:
   ' Don't RaiseError - creates infinite loop if drive isn't ready...
   If Err.Number = 430 Or Err.Number = 429 Then
      MsgBox "Unlogged Error in mwSession.Logger.UpdateLogFileItem, SCRRUN.DLL may be old or missing. " & vbCrLf & _
      "We recommend version 5.1.0.5010. Contact Shipnet, or get it from another computer.  Stage = " & Stage, vbCritical
   ElseIf Err.Number = 76 Then
      MsgBox "Unlogged Error in mwSession.Logger.UpdateLogFileItem, Path Not Found: " & mLogFile
   ElseIf Err.Number = 5 Then
      ' Bogus non-Ascii values from Oracle (Invalid procedure Call or Arguement).
      On Error Resume Next
      ts.WriteLine "Error in mwSession.Logger.UpdateLogFileItem, Bad Ascii description logging. Error number: " & Err.Number
      ts.Close
      moParent.KillObject ts
      moParent.KillObject fso
      UpdateLogFileItem = True
      Exit Function
      
   Else
      MsgBox "Unlogged General Error in mwSession.Logger.UpdateLogFileItem, Stage = " & Stage & " : " & Err.Number & "-" & Err.Description
   End If
   UpdateLogFileItem = False
   moParent.KillObject ts
   moParent.KillObject fso
End Function

Public Property Get LogFile() As String
    LogFile = mLogFile
End Property

Public Property Let LogFile(ByVal NewValue As String)
    mLogFile = NewValue
End Property

' SM - Strip Message of bad chars...
Private Function SM(Msg As String) As String
   ' Don't strip msg to support extended message format...
   'msg = Replace(msg, "|", "-")
   SM = Replace(Msg, vbCrLf, "-")
End Function

