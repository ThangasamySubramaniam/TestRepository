VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwWorkflowItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwWorkflowItem - utility class to help process a workflow item
'
' 12 Jan 2002 ms
'
Option Explicit
Private mFileName As String
Private mIsInbound As Boolean
Private mIsPackage As Boolean
Private mIsSchemaUpdate As Boolean
Private mIsRouteToFolderExclusive As Boolean
Private mRouteToFolder As String
Private mIsRouteToEmailAddressExclusive As Boolean
Private mRouteToEmailAddress As String
Private mFormID As String
Private mTemplateID As String
Private moKeys As ConfigKeys
Private mCheckExcelWord As Boolean
Private mSaveXmlMessages As Boolean



Private moParent As mwWorkflow



Public Property Set Parent(oParent As mwWorkflow)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As mwWorkflow
   Set Parent = moParent
End Property

Public Property Set Session(oSession As mwSession.Session)
End Property

Public Property Get Session() As mwSession.Session
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal vData As String)
    mFileName = vData
End Property

Public Property Get RouteToFolder() As String
    RouteToFolder = mRouteToFolder
End Property

Public Property Let RouteToFolder(ByVal vData As String)
    mRouteToFolder = vData
End Property

Public Property Let FormID(ByVal vData As String)
    mFormID = vData
End Property

Public Property Get FormID() As String
    FormID = mFormID
End Property

Public Property Let RouteToEmailAddress(ByVal vData As String)
    mRouteToEmailAddress = vData
End Property

Public Property Get RouteToEmailAddress() As String
    RouteToEmailAddress = mRouteToEmailAddress
End Property

Public Property Let TemplateID(ByVal vData As String)
    mTemplateID = vData
End Property

Public Property Get TemplateID() As String
    TemplateID = mTemplateID
End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

Public Property Get IsInbound() As Boolean
    IsInbound = mIsInbound
End Property

Public Property Let IsInbound(ByVal vData As Boolean)
    mIsInbound = vData
End Property

Public Property Get IsPackage() As Boolean
    IsPackage = mIsPackage
End Property

Public Property Let IsPackage(ByVal vData As Boolean)
    mIsPackage = vData
End Property

Public Property Get IsSchemaUpdate() As Boolean
    IsSchemaUpdate = mIsSchemaUpdate
End Property

Public Property Let IsSchemaUpdate(ByVal vData As Boolean)
    mIsSchemaUpdate = vData
End Property

Public Property Get IsRouteToFolderExclusive() As Boolean
    IsRouteToFolderExclusive = mIsRouteToFolderExclusive
End Property

Public Property Let IsRouteToFolderExclusive(ByVal vData As Boolean)
    mIsRouteToFolderExclusive = vData
End Property

Public Property Get IsRouteToEmailAddressExclusive() As Boolean
    IsRouteToEmailAddressExclusive = mIsRouteToEmailAddressExclusive
End Property

Public Property Let IsRouteToEmailAddressExclusive(ByVal vData As Boolean)
    mIsRouteToEmailAddressExclusive = vData
End Property

'Public Property Get b() As boolean
'    b = mb
'End Property

'Public Property Let b(ByVal vData As boolean)
'    mb = vData
'End Property

'Public Property Get b() As boolean
'    b = mb
'End Property

'Public Property Let b(ByVal vData As boolean)
'    mb = vData
'End Property

Public Function MaybeProcessWorkflowItem(Optional FileName As String)
   Dim loZip As Zipper
   Dim loform As mwForm
   Dim fso As FileSystemObject
   Dim strTemp As String
   Dim IsZip As Boolean
   On Error GoTo FunctionError
      
   '
   ' 1/6/2002 - Fixed file pattern, go ahead save it...
   '
   
   If Trim(FileName) <> "" Then
      mFileName = FileName
   End If
   If UCase(Right(mFileName, 4)) = ".DGM" Then
      '
      '  Datagram File...
      '
      mIsInbound = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".PKG" Then
      '
      '  Package File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".FMB" Then
      '
      '  Package File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".FCB" Then
      '
      '  Package File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".DAB" Then
      '
      '  Package File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".ZIZ" Then
      '
      '  ZipInZip Package File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 4)) = ".XML" And mSaveXmlMessages Then
      '
      '  XML File...
      '
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf moKeys.GetBoolKeyValue("ProcessZipFiles") = True And UCase(Right(mFileName, 4)) = ".ZIP" Then
      mIsInbound = True
      mIsPackage = True
      Exit Function
   ElseIf UCase(Right(mFileName, 7)) = ".SCHEMA" Then
      mIsInbound = True
      mIsSchemaUpdate = True
   ElseIf UCase(Right(mFileName, 23)) = "DB_SCHEMA_UPDATE_MW.SQL" Then
      mIsInbound = True
      mIsSchemaUpdate = True
   ElseIf IsZipEncryptFile(mFileName) Then
      '
      ' Is a zipped form, must unpack to get essentials...
      '
      Set loZip = New Zipper
      If goSession.Workflow.IsZipEncryptWfFiles Then
         loZip.IsEncypt = True
         loZip.Password = ENCRYPT_PASSWORD
      End If
      loZip.OpenZip mFileName
      If Not loZip.IsOpen Then
         mIsInbound = False
         goSession.RaiseWorkflowStatus "Error in mwWorkflowItem, file received but could not be opened: " & mFileName
         loZip.TerminateClass
         Set loZip = Nothing
         Exit Function
      End If
      Set fso = New FileSystemObject
      loZip.GetFromZip "*.*", goSession.Workflow.WorkingSessionFolder
      Select Case UCase(fso.GetExtensionName(mFileName))
         Case Is = "DGZ"
            mFileName = fso.GetParentFolderName(mFileName) & "\" & fso.GetBaseName(mFileName) & ".dgm"
            IsZip = True
         Case Is = "PKZ"
            mFileName = fso.GetParentFolderName(mFileName) & "\" & fso.GetBaseName(mFileName) & ".pkg"
            IsZip = True
         Case Is = "FCZ"
            mFileName = fso.GetParentFolderName(mFileName) & "\" & fso.GetBaseName(mFileName) & ".fcb"
            IsZip = True
         Case Is = "FMZ"
            mFileName = fso.GetParentFolderName(mFileName) & "\" & fso.GetBaseName(mFileName) & ".fmb"
            IsZip = True
         Case Is = "DAZ"
            mFileName = fso.GetParentFolderName(mFileName) & "\" & fso.GetBaseName(mFileName) & ".dab"
            IsZip = True
         Case Else
            goSession.RaiseWorkflowStatus "Error in mwWorkflowItem, invalid file extension encountered: " & mFileName
            mIsInbound = False
            loZip.TerminateClass
            Set loZip = Nothing
            Exit Function
      End Select
      loZip.TerminateClass
      Set loZip = Nothing
   ElseIf Not mCheckExcelWord Then
      '
      ' 17 April 2006... exclude Excel/Word Docs from processing
      ' 02 Sept 2006... by flag to support new Excel2Snaps...
      '
      mIsInbound = False
      Exit Function
   End If
   Set fso = Nothing
   '
   ' OK, attempt to open form
   '
   '
   Set loform = goSession.OpenNewForm()
   loform.SuppressOpenError = True
   
   If IsZip Then
      mIsInbound = True
      Exit Function
   ElseIf Not loform.OpenForm(mFileName) Then
      '
      ' not one of ours
      '
      mIsInbound = False
      goSession.TearDownForm loform
      Exit Function
   End If
   '
   ' One of ours, set the distribution properties...
   '
   mIsInbound = True
   mIsRouteToEmailAddressExclusive = loform.WorkflowProperties.IsRouteToMailboxExclusive
   mIsRouteToFolderExclusive = loform.WorkflowProperties.IsRouteToFolderExclusive
   mRouteToEmailAddress = loform.WorkflowProperties.RouteToEmailAddress
   mRouteToFolder = loform.WorkflowProperties.RouteToFolder
   mFormID = loform.FormProperties.FormID
   mTemplateID = loform.TemplateProperties.TemplateID
   goSession.TearDownForm loform, False
   goSession.KillObject loform
   goSession.API.SleepApp 4000
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwMailbox.MaybeProcessWorkflowItem. ", Err.Number, Err.Description
   mIsInbound = False

End Function

Public Function CopyOrMoveItemToFolder() As Boolean
   Dim strTargetFile As String
   Dim strTargetFile2 As String
   Dim fso As FileSystemObject
   Dim i As Integer
   Dim strErrorFile As String
   Dim sStrTarget As String
   On Error GoTo FunctionError
   '
   If Not mIsInbound Then
      CopyOrMoveItemToFolder = False
      Exit Function
   End If
   '
   ' Timing issue bug on Word being closed ?
   '
   'goSession.API.SleepApp 24000
   '
   Set fso = New FileSystemObject
   '
   If Trim(mRouteToFolder) = "" Then
      '
      ' Move the file to polling Folder only...
      '
      strTargetFile = goSession.Workflow.PollingFolder & "\" & fso.GetFileName(mFileName)
      
      If fso.FileExists(strTargetFile) Then
         strErrorFile = goSession.Workflow.ArchiveErrorZipFile & "\" & fso.GetFileName(mFileName)
         goSession.RaiseWorkflowStatus "Incoming Form Already exists and will be moved to ArchivedError folder: " & strErrorFile
         goSession.Workflow.RenameFileNowExtensionNew strTargetFile, sStrTarget
         strErrorFile = goSession.Workflow.ArchiveErrorZipFile & "\" & fso.GetFileName(sStrTarget)
         fso.MoveFile sStrTarget, strErrorFile
      End If
      fso.MoveFile mFileName, strTargetFile
   ElseIf mIsRouteToFolderExclusive Then
      strTargetFile = mRouteToFolder & "\" & fso.GetFileName(mFileName)
      If fso.FileExists(strTargetFile) Then
         goSession.Workflow.RenameFileNowExtension strTargetFile
         goSession.RaiseWorkflowStatus "Save to Folder Warning: " & strTargetFile & " exists and is being renamed."
      End If
      fso.MoveFile mFileName, strTargetFile
   Else
      '
      ' Both places
      '
      strTargetFile = goSession.Workflow.PollingFolder & "\" & fso.GetFileName(mFileName)
      If fso.FileExists(strTargetFile) Then
         goSession.Workflow.RenameFileNowExtension strTargetFile
         goSession.RaiseWorkflowStatus "Save to Folder Warning: " & strTargetFile & " exists and is being renamed."
      End If
      fso.CopyFile mFileName, strTargetFile
      strTargetFile2 = mRouteToFolder & "\" & fso.GetFileName(mFileName)
      If fso.FileExists(strTargetFile2) Then
         goSession.Workflow.RenameFileNowExtension strTargetFile2
         goSession.RaiseWorkflowStatus "Save to Folder Warning: " & strTargetFile2 & " exists and is being renamed."
      End If
      
      fso.MoveFile mFileName, strTargetFile2
   End If
   '
   ' Log it to WorkflowStatus
   '
   If mIsPackage Then
      goSession.RaiseWorkflowStatus "Received Package, saved to: " & strTargetFile
   Else
      goSession.RaiseWorkflowStatus "Received Form: " & mFormID & _
         ", saved to: " & strTargetFile
      If Trim(strTargetFile2) <> "" Then
         goSession.RaiseWorkflowStatus "Received Form: " & mFormID & _
            ", save to folder option in use: " & strTargetFile2
      End If
   End If
   CopyOrMoveItemToFolder = True
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwWorkflowItem.CopyOrMoveToFolder, Source: " & mFileName & _
     ", Target: " & strTargetFile, _
     Err.Number, Err.Description
   CopyOrMoveItemToFolder = False
   goSession.KillObject fso
End Function

Public Function IsZippedWorkflowFileExtension(FileName As String) As Boolean
   goSession.RaiseError "General Error in mwWorkflowItem.IsZippedWorkflowFileExtension. Function Retired"
End Function

Public Function IsZipEncryptFile(FileName As String) As Boolean
   Dim fso As FileSystemObject
   Dim sExt As String
   Set fso = New FileSystemObject
   sExt = UCase(fso.GetExtensionName(FileName))
   If sExt <> "DGZ" And sExt <> "PKZ" And sExt <> "FCZ" And sExt <> "FMZ" And sExt <> "DAZ" Then
      IsZipEncryptFile = False
   Else
      IsZipEncryptFile = True
   End If
   Set fso = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWorkflowItem.IsZipEncryptFile.", Err.Number, Err.Description
   IsZipEncryptFile = False
End Function

Private Sub Class_Initialize()
   On Error GoTo SubError
   Set moKeys = goSession.GetEventSecurityKeys(MW_EVENT_WORKFLOW_AGENT)
   
   mCheckExcelWord = moKeys.GetBoolKeyValue("ParseWordExcelInbound")
   
'   If moKeys.GetKeyValue("ParseWordExcelInbound") = "1" Then
'      mCheckExcelWord = True
'   Else
'      mCheckExcelWord = False
'   End If

   mSaveXmlMessages = moKeys.GetBoolKeyValue("SaveXmlMessages")

'   If moKeys.GetKeyValue("SaveXmlMessages") = "1" Then
'      mSaveXmlMessages = True
'   Else
'      mSaveXmlMessages = False
'   End If
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwWorkflowItem.Class_Initialize.", Err.Number, Err.Description
End Sub
