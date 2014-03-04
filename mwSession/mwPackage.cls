VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Session"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwPackage - Package Installer Class
' 8/2001 - ms Maritime Systems Inc
'
Option Explicit

Private m_ID As Long
Private mOriginatingSiteID As String
Private mFullFilename As String
Private mInstallLog As String
Private mIsOpen As Boolean
Private mIsInstalled As Boolean
Private mIsFatalError As Boolean
Private mLocalPackageReceiptID As Long
Private mWorkingFolder As String
Private mInstallItemLogMessage As String
Private mPackageExtension As String

' Objects
Private moCon As Connection
Private moRsRP As Recordset
Private moRsRpItem As Recordset
Private moIni As mwSession.IniConfig
Private moParent As mwWorkflow
Private moSess As Session

   Dim moTools As Object

Const pkg_PACKAGE_UPDATE_INI_FILE = "pkg.ini"
Const MW_EVENT_MANUAL_UPDATE = 22
Const MW_EVENT_SYSTEM_UPDATE = 23

Const CENTRALPKG_RECEIPT = 1
Const INSTALLEDPKG_RECEIPT = 2


Friend Property Set Parent(oParent As mwWorkflow)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Friend Property Get Parent() As mwWorkflow
   Set Parent = moParent
End Property

Public Property Set Session(oSess As mwSession.Session)
   If moSess Is Nothing Then
      Set moSess = oSess
      Set moCon = goCon
   End If
End Property

Public Property Get Session() As Session
   'If moSess Is Nothing Then
   '   Set moSess = New Session
   'End If
   Set Session = moSess
   'Set moCon = goCon
End Property

Public Property Let ID(ByVal Data As Long)
    m_ID = Data
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let OriginatingSiteID(ByVal Data As String)
    mOriginatingSiteID = Data
End Property

Public Property Get OriginatingSiteID() As String
    OriginatingSiteID = mOriginatingSiteID
End Property

Public Property Let FullFilename(ByVal Data As String)
    mFullFilename = Data
End Property

Public Property Get FullFilename() As String
    FullFilename = mFullFilename
End Property


Public Property Let WorkingFolder(ByVal Data As String)
    mWorkingFolder = Data
End Property

Public Property Get WorkingFolder() As String
    WorkingFolder = mWorkingFolder
End Property

Public Property Let IsOpen(ByVal Data As Boolean)
    mIsOpen = Data
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

Public Property Get InstallLog() As String
    InstallLog = mInstallLog
End Property

Public Property Get LocalPackageReceiptID() As Long
    LocalPackageReceiptID = mLocalPackageReceiptID
End Property

Public Property Get IsInstalled() As Boolean
    IsInstalled = mIsInstalled
End Property



Public Property Get PackageExtension() As String
   Dim loRs As Recordset
   Dim sSQL As String
   If moParent.IsSilentErrors Then
      On Error GoTo PropertyErrorSilent
   Else
      On Error GoTo PropertyError
   End If
   If Trim(mPackageExtension) <> "" Then
      PackageExtension = mPackageExtension
      Exit Property
   End If
   '
   ' Look it up
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwwfFormTypes where ID=4"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount = 0 Then
      If moParent.IsSilentErrors Then
         goSession.RaiseWorkflowStatus "Error in mwPackage.PackageExtension, mwwfFormType Package record 4 missing !"
      Else
         goSession.RaiseError "Error in mwPackage.PackageExtension, mwwfFormType Package record 4 missing !"
      End If
      PackageExtension = ""
      Exit Property
      goSession.CloseRecordset loRs
   End If
   mPackageExtension = loRs!FormFilePattern
   PackageExtension = mPackageExtension
   goSession.CloseRecordset loRs
   Exit Property
PropertyError:
   goSession.RaiseError "General error in mwPackage.PackageExtension. ", Err.Number, Err.Description
   PackageExtension = ""
   Exit Property
PropertyErrorSilent:
   goSession.RaiseWorkflowStatus "General error in mwPackage.PackageExtension. ", Err.Number, Err.Description
   PackageExtension = ""
   
End Property






'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

'
' Containers are defined as zip files in the polling folder...
'


Public Function InstallPackage() As Boolean
   
   Dim fso As FileSystemObject
   Dim strZipFile As String
   Dim strIniFile As String
   Dim strTargetFolder As String
   Dim mWorkingFolder As String
   Dim strTemp1 As String
   Dim strTemp2 As String
   Dim sSQL As String
   'Dim i As Integer
   On Error GoTo FunctionError
   
   mIsFatalError = False
   
   ' Must be open...
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwSession.mwPackage.InstallPackage, Package Not Open."
      If Not moParent.ArchiveWorkflowItem(mFullFilename, ArchiveError, True) Then
         moParent.RenameFileNowExtension (mFullFilename)
      End If
      InstallPackage = False
      Exit Function
   End If
   ' ... and not installed
   If mIsInstalled Then
      ' If package installed, archive it, log it, and wrap it up...
      goSession.RaiseWorkflowStatus "mwSession.mwPackage.InstallPackage, Package Already Installed: " & m_ID
      If Not moParent.ArchiveWorkflowItem(mFullFilename, ArchivedOK, True) Then
         moParent.RenameFileNowExtension (mFullFilename)
      End If
      InstallPackage = True
      Exit Function
   End If
   
   If moTools Is Nothing Then
      Set moTools = CreateObject("mwUtility.mwTools")
      moTools.InitSession goSession
   End If
   
   '
   ' moRsRp - Received Package Recordset
   '
   Set moRsRP = New Recordset
   moRsRP.CursorLocation = adUseClient
   sSQL = "select * from pkgReceivedPackage where ID=" & moIni.GetAppKey("package_header", "PackageID")
   moRsRP.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   '
   ' Has package been partially installed ?
   '
   'moRsRP.Filter = "ID = " & moIni.GetAppKey("package_header", "PackageID")
  If moRsRP.RecordCount < 1 Then
      '
      ' Add the Received Package Record
      '
      'moRsRP.Filter = adFilterNone
      If Not InsertReceivedPackageRecord() Then
         goSession.CloseRecordset moRsRP
         goSession.CloseRecordset moRsRpItem
         InstallPackage = False
         KillObject moTools
         Exit Function
      End If
   Else
      '
      ' Yes, partially installed...
      '
      goSession.RaiseWorkflowStatus "Package partially Installed, will try to complete the Installation of Package " & moRsRP!ID
      mInstallLog = "Package partially Installed, will try to complete the Installation of Package " & moRsRP!ID
   End If
   '
   '  Process Package Contents
   '
   ' Install the package contents...
   If InstallPackageContents() Then
      moRsRP!IsInstalled = True
      InstallPackage = True
      mIsInstalled = True
   Else
      InstallPackage = False
      moRsRP!IsInstalled = False
      mIsInstalled = False
   End If
   ' Create Receipt records, whether pass or fail...
   
   If CreateLocalPackageReceipt = True Then
      If Not moRsRP!IsdeveloperPackage Then
         If CreateCentralPackageReceipt Then
            UpdateLocalPackageReceipt mLocalPackageReceiptID, CENTRALPKG_RECEIPT
         End If
      Else
         moParent.RaiseStatusMessage "Received Developer Package - Bypassing Central Receipt."
      End If
   End If
   
   If mIsInstalled Then
      If Not moRsRP!IsdeveloperPackage Then
         If CreateCentralInstalledPackage Then
            UpdateLocalPackageReceipt mLocalPackageReceiptID, INSTALLEDPKG_RECEIPT
         End If
      Else
          moParent.RaiseStatusMessage "Received Developer Package - Bypassing Installation Package Receipt."
      End If
   End If
      
   ' Close up shop
   moRsRP.Update
   '
   ' Message Event for Workflow...
   '
   If InstallPackage Then
      If Not moParent.ArchiveWorkflowItem(mFullFilename, ArchivedOK, True) Then
         moParent.RenameFileNowExtension (mFullFilename)
      End If
      moParent.RaiseStatusMessage "Successfully installed package: " & moRsRP!ID
      If goSession.Workflow.GetExtendedProperty("IsNotifyMsgBoxOnPackageInstall") = "True" Then
         MsgBox "Workflow Agent has successfully installed package: " & moRsRP!ID, vbInformation, "Workflow Agent Notification"
      End If
   Else
      If mIsFatalError = True Then
         If Not moParent.ArchiveWorkflowItem(mFullFilename, ArchiveError, True) Then
            moParent.RenameFileNowExtension (mFullFilename)
         End If
      End If
      
      moParent.RaiseStatusMessage "Error in mwPackage.InstallPackage. Error encountered installing package: " & moRsRP!ID
      If goSession.Workflow.GetExtendedProperty("IsNotifyMsgBoxOnPackageInstall") = "True" Then
         MsgBox "Workflow Agent encountered an error installing package: " & moRsRP!ID, vbCritical, "Workflow Agent Notification"
      End If
   End If
   moRsRP.Close
   CloseRecordset moRsRP
   CloseRecordset moRsRpItem
   KillObject moIni
   KillObject fso
   KillObject moTools
   Exit Function
FunctionError:
   InstallPackage = False
   goSession.RaiseError "General Error in mwSession.mwPackage.InstallPackage.", Err.Number, Err.Description
   CloseRecordset moRsRP
   CloseRecordset moRsRpItem
   KillObject moIni
   KillObject fso
   KillObject moTools
   'If Not moParent.ArchiveWorkflowItem(mFullFilename, ArchivedOK, True) Then
   '   moParent.RenameFileNowExtension (mFullFilename)
   'End If
   
End Function

Private Function InstallPackageContents() As Boolean
   Dim sSQL As String
   
   Dim fso As FileSystemObject
   Dim mInstallItemLogMessage As String
   Dim isErrorOccurred As Boolean
   Dim i As Integer
   Dim iRecnum As Integer
   On Error GoTo FunctionError
   '
   ' moRsRpItem - Received Package Item Recordset
   '
   Set moRsRpItem = New Recordset
   moRsRpItem.CursorLocation = adUseClient
   moRsRpItem.Open "pkgReceivedPackageItem", moCon, adOpenDynamic, adLockOptimistic, adCmdTable

   For i = 2 To moIni.ConfigGroups.Count
      '
      ' See if Package Item record exists
      '
      'sSQL = "ID=" & moIni.ConfigGroups(i).ConfigKeys("ID").KeyValue & _
      '  " and PackageItemID=" & moIni.ConfigGroups(i).ConfigKeys("PackageItemID").KeyValue
      sSQL = "PkgReceivedPackageKey=" & moIni.ConfigGroups(1).ConfigKeys.GetKeyValue("PACKAGEID") & _
        " and ID=" & moIni.ConfigGroups(i).ConfigKeys.GetKeyValue("PackageItemID")
      moRsRpItem.Filter = sSQL
      If moRsRpItem.RecordCount < 1 Then
         'Create Package Item Record...
         If Not InsertPackageItemRecord(i) Then
            'Error has occurred
            InstallPackageContents = False
            mInstallItemLogMessage = "Error installing: " & moRsRpItem!PackageItemID
            goSession.CloseRecordset moRsRpItem
            Exit Function
         End If
      End If
      'iRecnum = moRsRpItem!ID
      ' Install File Required ?
      If (moRsRpItem!IsInstalled = False) Or IsNull(moRsRpItem!IsInstalled) Then
         '
         ' Install the Package Item file...
         If InstallPackageitemFile() Then
            moRsRpItem!IsInstalled = True
            moRsRpItem!InstallLogString = mInstallItemLogMessage
            moRsRpItem!ItemInstallDate = Now()
            moRsRpItem.Update
         Else
            InstallPackageContents = False
            mInstallItemLogMessage = "Error installing: " & moRsRpItem!ID
            goSession.CloseRecordset moRsRpItem
            Exit Function
         End If
      Else
         ' Record found, item already installed...
         mInstallItemLogMessage = "Already installed, skipping Package Item: " & moRsRpItem!ID
      End If
      moRsRpItem.Filter = adFilterNone
      ' Concatenate to Package Install message
      mInstallLog = mInstallLog & mInstallItemLogMessage & vbCrLf
   Next i
   InstallPackageContents = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.InstallPackageContents.", Err.Number, Err.Description
   InstallPackageContents = False
   mIsFatalError = True
End Function

Private Function InsertPackageItemRecord(ConfigOffset As Integer) As Boolean

   On Error GoTo FunctionError
   '
   ' Insert the Package Item record
   '
   With moIni.ConfigGroups(ConfigOffset)
      mInstallItemLogMessage = "Installing Package Item: " & .ConfigKeys("PackageItemID").KeyValue & vbCrLf
      moRsRpItem.AddNew
      'moRsRpItem!PackageItemID = .ConfigKeys("PackageItemID").KeyValue
      'moRsRpItem!ID = .ConfigKeys("ID").KeyValue
      moRsRpItem!ID = .ConfigKeys.GetKeyValue("PackageItemID")
      moRsRpItem!PkgReceivedPackageKey = moIni.ConfigGroups(1).ConfigKeys.GetKeyValue("PackageID")
      moRsRpItem!SourceFileLocation = .ConfigKeys.GetKeyValue("SourceFileLocation")
      moRsRpItem!TargetLocation = .ConfigKeys.GetKeyValue("TargetLocation")
      If Not .ConfigKeys.GetKeyValue("Reference") = "" Then
         moRsRpItem!Reference = .ConfigKeys.GetKeyValue("Reference")
      End If
      If Not .ConfigKeys.GetKeyValue("ShortDescription") = "" Then
         moRsRpItem!ShortDescription = .ConfigKeys.GetKeyValue("ShortDescription")
      End If
      If Not .ConfigKeys.GetKeyValue("LongDescription") = "" Then
         moRsRpItem!LongDescription = ReplaceVbCrLf(.ConfigKeys.GetKeyValue("LongDescription"))
      End If
      If Not .ConfigKeys.GetKeyValue("Hyperlink") = "" Then
         moRsRpItem!Hyperlink = .ConfigKeys.GetKeyValue("Hyperlink")
      End If
      
      If Not .ConfigKeys.GetKeyValue("ManualChapterKey") = "" Then
         moRsRpItem!ManualChapterKey = .ConfigKeys.GetKeyValue("ManualChapterKey")
         If moParent.IsNotifyManualRevisionUpdate Then
            moRsRpItem!IsAlertNotificationSent = False
         End If
      End If
      If Not .ConfigKeys.GetKeyValue("ManualRevisionNo") = "" Then
         moRsRpItem!ManualRevisionNo = .ConfigKeys.GetKeyValue("ManualRevisionNo")
      End If
      If Not .ConfigKeys.GetKeyValue("ManualRevisionDate") = "" Then
         moRsRpItem!ManualRevisionDate = Iso8601DateStringToVBDate(.ConfigKeys.GetKeyValue("ManualRevisionDate"))
      End If
      If Not .ConfigKeys.GetKeyValue("mwrChangeTableKey") = "" Then
         moRsRpItem!mwrChangeTableKey = .ConfigKeys.GetKeyValue("mwrChangeTableKey")
      End If
      
      If Not .ConfigKeys.GetKeyValue("ItemKey") = "" Then
         moRsRpItem!ItemKey = .ConfigKeys.GetKeyValue("ItemKey")
         If moParent.IsNotifyManualRevisionUpdate Then
            moRsRpItem!IsAlertNotificationSent = False
         End If
      End If
      If Not .ConfigKeys.GetKeyValue("ItemRevisionNo") = "" Then
         moRsRpItem!ItemRevisionNo = .ConfigKeys.GetKeyValue("ItemRevisionNo")
      End If
      If Not .ConfigKeys.GetKeyValue("ItemRevisionDate") = "" Then
         moRsRpItem!ItemRevisionDate = Iso8601DateStringToVBDate(.ConfigKeys.GetKeyValue("ItemRevisionDate"))
      End If
      
      
      moRsRpItem!IsInstalled = False
      moRsRpItem.Update
      InsertPackageItemRecord = True
   End With
   Exit Function
FunctionError:
      InsertPackageItemRecord = False
      mInstallItemLogMessage = "Failed Installing Package Item: " _
        & moIni.ConfigGroups(ConfigOffset).ConfigKeys.GetKeyValue("PackageItemID") & vbCrLf
      goSession.RaiseError "General Error in mwSession.mwPackage.InsertPackageItemRecord.", Err.Number, Err.Description
      mIsFatalError = True
      
End Function

Private Function InstallPackageitemFile() As Boolean
   Dim fso As FileSystemObject
   
   ' working variables..
   Dim strTargetFolder As String
   Dim strTemp1 As String
   Dim strTemp2 As String
   Dim loRs As Recordset
   Dim sSQL As String
   Dim StageNo As Integer
   Dim nAlertCount As Integer
   
   'Dim i As Integer
   'Dim iRecnum As Integer
   On Error GoTo FunctionError
   
   StageNo = 1
   Set fso = New FileSystemObject
   InstallPackageitemFile = True
   
   'strTargetFolder = goSession.GetPlace(moRsRpItem!TargetLocation, mw_DriveLetterPath)
   strTargetFolder = goSession.GetPlace2(moRsRpItem!TargetLocation, mw_DriveLetterPath) 'PBT-1671
   
   If strTargetFolder = "" Then
      
      '
      ' Can't find Place
      '
      InstallPackageitemFile = False
      mInstallItemLogMessage = "TargetFolder is missing from Database:" & moRsRpItem!TargetLocation
      goSession.RaiseWorkflowStatus mInstallItemLogMessage
   Else
      StageNo = 2
      strTemp1 = strTargetFolder & "\" & fso.GetFileName(moRsRpItem!SourceFileLocation)
      If Not goSession.Workflow.RenameFileNowExtension(strTemp1) Then
         '
         'Error Renaming file (File not found is considered ok, returns True)
         '
         InstallPackageitemFile = False
         mInstallItemLogMessage = "Error in InstallPackageItemFile renaming file to be replaced."
         goSession.RaiseWorkflowStatus mInstallItemLogMessage

      Else
         '
         'copy new one in place
         '
         StageNo = 3
         strTemp1 = mWorkingFolder & "\" & fso.GetFileName(moRsRpItem!SourceFileLocation)
         strTemp2 = strTargetFolder & "\" & fso.GetFileName(moRsRpItem!SourceFileLocation)
         
         moTools.MakeFolderExist strTemp2
         
         fso.CopyFile strTemp1, strTemp2, False
         StageNo = 4
         
         If ZeroNull(moRsRpItem!ManualChapterKey) > 0 And ZeroNull(moRsRpItem!ManualRevisionNo) > 0 Then
         
            ' Update the smManualChapter record but DON'T replicate the change!
            
            If ZeroNull(moRsRpItem!mwrChangeTableKey) = 0 Or ZeroNull(moRsRpItem!mwrChangeTableKey) = MWRT_smManualChapter Then
            
               Set loRs = New Recordset
               loRs.CursorLocation = adUseClient
               
               sSQL = "select * from smManualChapter where ID= " & moRsRpItem!ManualChapterKey
               loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
               
               If loRs.RecordCount = 1 Then
                  If ZeroNull(loRs!DistribRevisionNo) <= moRsRpItem!ManualRevisionNo Then
                     loRs!DistribRevisionNo = moRsRpItem!ManualRevisionNo
                     loRs!DistribRevisionDate = moRsRpItem!ManualRevisionDate
                     loRs.Update
                  End If
               Else
                  loRs.AddNew
                  loRs!ID = moRsRpItem!ManualChapterKey
                  loRs!DistribRevisionNo = moRsRpItem!ManualRevisionNo
                  loRs!DistribRevisionDate = moRsRpItem!ManualRevisionDate
                  loRs!IsTempRecordInsert = True
                  loRs.Update
                  
'                  InstallPackageitemFile = False
                  mInstallItemLogMessage = "Manual Chapter record " & CStr(moRsRpItem!ManualChapterKey) & " has not arrived yet. A Temporary record has been added to hold its place."
                  goSession.RaiseWorkflowStatus mInstallItemLogMessage
               End If
               '---PBT 1538
               If moParent.IsNotifyManualRevisionUpdate Then
                  nAlertCount = 0
                  If SendManualUpdateAlert(moRsRpItem!ManualChapterKey, nAlertCount) Then
                     If nAlertCount > 0 Then
                        moRsRpItem!IsAlertNotificationSent = True
                        goSession.RaiseWorkflowStatus "Alert message has been sent to the members which are linked to the Manual Chapter record " & CStr(moRsRpItem!ManualChapterKey)
                    Else
                        moRsRpItem!IsAlertNotificationSent = False
                        goSession.RaiseWorkflowStatus "No Responsibility viewer Assignments found. Alert message has NOT been sent to the members which are linked to the Manual Chapter record " & CStr(moRsRpItem!ManualChapterKey)
                    End If
                    moRsRpItem.Update
                  End If
               End If
               '---
               CloseRecordset loRs
            End If
         ElseIf ZeroNull(moRsRpItem!ItemKey) > 0 And ZeroNull(moRsRpItem!ItemRevisionNo) > 0 Then
            If ZeroNull(moRsRpItem!mwrChangeTableKey) = 0 Or ZeroNull(moRsRpItem!mwrChangeTableKey) = MWRT_smManualChapter Then
            
               Set loRs = New Recordset
               loRs.CursorLocation = adUseClient
               
               sSQL = "select * from smManualChapter where ID= " & moRsRpItem!ItemKey
               loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
               
               If loRs.RecordCount = 1 Then
                  If ZeroNull(loRs!DistribRevisionNo) <= moRsRpItem!ItemRevisionNo Then
                     loRs!DistribRevisionNo = moRsRpItem!ItemRevisionNo
                     loRs!DistribRevisionDate = moRsRpItem!ItemRevisionDate
                     loRs.Update
                  End If
               Else
                  loRs.AddNew
                  loRs!ID = moRsRpItem!ItemKey
                  loRs!DistribRevisionNo = moRsRpItem!ItemRevisionNo
                  loRs!DistribRevisionDate = moRsRpItem!ItemRevisionDate
                  loRs.Update
               
'                  InstallPackageitemFile = False
                  mInstallItemLogMessage = "Manual Chapter record " & CStr(moRsRpItem!ItemKey) & " has not arrived yet. A Temporary record has been added to hold its place."
                  goSession.RaiseWorkflowStatus mInstallItemLogMessage
               End If
               '---PBT 1538
               If moParent.IsNotifyManualRevisionUpdate Then
                  nAlertCount = 0
                  If SendManualUpdateAlert(moRsRpItem!ItemKey, nAlertCount) Then
                     If nAlertCount > 0 Then
                        moRsRpItem!IsAlertNotificationSent = True
                        goSession.RaiseWorkflowStatus "Alert message has been sent to the members which are linked to the Manual Chapter record " & CStr(moRsRpItem!ItemKey)
                     Else
                        moRsRpItem!IsAlertNotificationSent = False
                        goSession.RaiseWorkflowStatus "No Responsibility viewer Assignments found. Alert message has NOT been sent to the members which are linked to the Manual Chapter record " & CStr(moRsRpItem!ItemKey)
                     End If
                  End If
                  moRsRpItem.Update
               End If
               '---
               CloseRecordset loRs
            ElseIf ZeroNull(moRsRpItem!mwrChangeTableKey) = MWRT_mwwfFormTemplate Then
            
               Set loRs = New Recordset
               loRs.CursorLocation = adUseClient
               
               sSQL = "select * from mwwfFormTemplate where ID= " & moRsRpItem!ItemKey
               loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
               
               If loRs.RecordCount = 1 Then
                  If ZeroNull(loRs!DistributedRevisionNo) <= moRsRpItem!ItemRevisionNo Then
                     loRs!DistributedRevisionNo = moRsRpItem!ItemRevisionNo
                     loRs!DistributedRevisionDate = moRsRpItem!ItemRevisionDate
                     loRs.Update
                  End If
               Else
                  InstallPackageitemFile = False
                   mInstallItemLogMessage = "Could not complete the installation of the package at this time, the new Form Template record " & CStr(moRsRpItem!ItemKey) & " has not arrived yet. Package will remain in queue until it can be completed."
                  goSession.RaiseWorkflowStatus mInstallItemLogMessage
               End If
               
               CloseRecordset loRs
            
               If InstallPackageitemFile = True Then
                  Set loRs = New Recordset
                  loRs.CursorLocation = adUseClient
                  
                  sSQL = "select * from mwFormTemplateRev WHERE mwwfFormTemplateKey= " & moRsRpItem!ItemKey & _
                         " AND RevisionNo = " & moRsRpItem!ItemRevisionNo
                  loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
                  
                  If loRs.RecordCount = 1 Then
                     loRs!DistributedRevisionNo = moRsRpItem!ItemRevisionNo
                     loRs!DistributedRevisionDate = moRsRpItem!ItemRevisionDate
                     loRs.Update
                  Else
                     InstallPackageitemFile = False
                     mInstallItemLogMessage = "Could not complete the installation of the package at this time, the Form Template Revision record for Template " & CStr(moRsRpItem!ItemKey) & " has not arrived yet. Package will remain in queue until it can be completed."
                     goSession.RaiseWorkflowStatus mInstallItemLogMessage
                  End If
                  
                  CloseRecordset loRs
               End If
            End If
            
         End If
         '
         ' yippee, we did it. Maybe.
         '
         If InstallPackageitemFile = True Then
            mInstallItemLogMessage = strTemp2 & " Installed OK"
            moParent.RaiseStatusMessage strTemp2 & " Installed OK"
         End If
      End If
   End If
   Exit Function
FunctionError:
   If StageNo = 2 Then
      goSession.RaiseError "Error renaming file to be replace: " & strTemp1 & " in mwSession.mwPackage.InstallPackageItemFile.", Err.Number, Err.Description
   ElseIf StageNo = 3 Then
      goSession.RaiseError "Error copying source file " & strTemp1 & " to target file " & strTemp2 & " in mwSession.mwPackage.InstallPackageItemFile.", Err.Number, Err.Description
   Else
      goSession.RaiseError "General Error in mwSession.mwPackage.InstallPackageItemFile.", Err.Number, Err.Description
   End If
   
   goSession.KillObject fso
   InstallPackageitemFile = False
   mIsFatalError = True
   
End Function


Private Function InsertReceivedPackageRecord() As Boolean
   Dim iEvent As Long
   Dim loWork As mwEventWork
   On Error GoTo FunctionError
   '
   ' Create Package Record
   '
   moCon.BeginTrans
   With moIni
      moRsRP.AddNew
      If IsNumeric(.GetAppKey("package_header", "PackageID")) Then
         moRsRP!ID = .GetAppKey("package_header", "PackageID")
      Else
         goSession.RaiseError "Error in mwSession.mwPackage.InsertReceivedPackageRecord, PackageID is not numeric."
         moCon.RollbackTrans
         InsertReceivedPackageRecord = False
         Exit Function
      End If
      moRsRP!pkgPackageTypeKey = .GetAppKey("package_header", "PackageType")
      moRsRP!PackageName = .GetAppKey("package_header", "PackageName")
      If .GetAppKey("package_header", "SpecialInstructions") <> "" Then
         moRsRP!SpecialInstructions = .GetAppKey("package_header", "SpecialInstructions")
      End If
      If IsDate(.GetAppKey("package_header", "DateCreated")) Then
         moRsRP!DateCreated = .GetAppKey("package_header", "DateCreated")
'         moRsRP!DateCreated = Iso8601DateStringToVBDate(.GetAppKey("package_header", "DateCreated"))
      End If
      If IsDate(.GetAppKey("package_header", "DateTransmitted")) Then
         moRsRP!DateTransmitted = .GetAppKey("package_header", "DateTransmitted")
'         moRsRP!DateTransmitted = Iso8601DateStringToVBDate(.GetAppKey("package_header", "DateTransmitted"))
      End If
      
      If IsDate(.GetAppKey("package_header", "ISODateTransmitted")) Then
         moRsRP!DateTransmitted = Iso8601DateStringToVBDate(.GetAppKey("package_header", "ISODateTransmitted"))
      End If
      
      moRsRP!DateReceived = Now()
      If IsNumeric(.GetAppKey("package_header", "mwcRoleTypeKey")) Then
         moRsRP!mwcRoleTypekey = .GetAppKey("package_header", "mwcRoleTypeKey")
      End If
      moRsRP!Report_DateTime = Now()
      If goSession.Site.CurrentVoyage <> "" Then
         moRsRP!Voyage_NO = goSession.Site.CurrentVoyage
      End If
      If goSession.Site.SiteID <> "" Then
         moRsRP!Vessel_ID = goSession.Site.SiteID
      End If
      If .GetAppKey("package_header", "IsDeveloperPackage") = "True" Then
         moRsRP!IsdeveloperPackage = True
      Else
         moRsRP!IsdeveloperPackage = False
      End If
      moRsRP!sent = True
      moRsRP!IsInstalled = 0
      mInstallLog = "New Package found: " & mFullFilename & ", beginning installation at: " & Now() & vbCrLf
      moRsRP.Update
      ' Goofy, goofy - in Access, record doesn't appear without resynch.
      moRsRP.Resync
   End With
   '
   ' Create Event Record
   '
   Set loWork = New mwEventWork
   If moRsRP!pkgPackageTypeKey = 1 Then
      iEvent = loWork.AddEvent(MW_EVENT_MANUAL_UPDATE, moRsRP)
      goSession.RaiseWorkflowStatus "Inserted Manual Update Event into database for: " & iEvent
   ElseIf moRsRP!pkgPackageTypeKey = 2 Then
      iEvent = loWork.AddEvent(MW_EVENT_SYSTEM_UPDATE, moRsRP)
      goSession.RaiseWorkflowStatus "Inserted System Update Event into database for: " & iEvent
   Else
      goSession.RaiseError "Error in mwSession.mwPackage.InsertReceivedPackageRecord, invalid PackageType."
      goSession.KillObject loWork
      moCon.RollbackTrans
      InsertReceivedPackageRecord = False
   End If
   moCon.CommitTrans
   goSession.KillObject loWork
   InsertReceivedPackageRecord = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.InsertReceivedPackageRecord.", Err.Number, Err.Description
   On Error Resume Next
   moCon.RollbackTrans
   InsertReceivedPackageRecord = False
End Function



'
' Create Local Database receipt (pkgLocalPackageReceipt)...
' and message to central database (pkgCentralPackageReceipt)
'
Private Function CreateCentralPackageReceipt(Optional Comments As String) As Boolean
   Dim oRsRPR As Recordset
   Dim loform As mwForm
   Dim oRec As ConfigKeys
   Dim TargetSiteID As String
   
   On Error GoTo FunctionError
   If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
      ' Ship to Shore
      TargetSiteID = goSession.Site.TargetReplicateSiteID
   Else
      ' Shore to itself
      TargetSiteID = goSession.Site.SiteID
   End If
   
   '
   ' Create Remote record
   '
   Set loform = goSession.OpenNewForm()
   If Not loform.MakeNewForm("PKG_RECEIPT", , , , TargetSiteID) Then
      CreateCentralPackageReceipt = False
      goSession.KillObject loform
      Exit Function
   End If
   
   
   Set oRec = loform.DataForm.AddRecord
   With oRec
      .Add "Table", "pkgCentralPackageReceipt"
      .Add "Operation", "AMPK"
      
      .Add "PkgPackageKey", moRsRP!ID
      .Add "mwcSitesKey", goSession.Site.SiteKey
      .Add "mwcUsersKey", goSession.User.UserKey
      .Add "EmployeeID", goSession.User.EmployeeID
      .Add "mwcRoleTypeKey", goSession.User.RoleTypeKey
      .Add "ReceiptDate", Format(goSession.API.GetGmtNow, DATAGRAM_DATE_FORMAT)
      If Comments <> "" Then
         .Add "Comments", ReplaceVbCrLf(Comments)
      End If
      .Add "pkgLocalPackageReceiptKey", str(mLocalPackageReceiptID)
   End With
   '
   ' Write the message...
   '
   If Not loform.DataForm.WriteDataForm() Then
      CreateCentralPackageReceipt = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   '
   ' Transmit the message
   '
   loform.DeleteAfterSend = True
   If Not loform.Send() Then
      ' Error sending email
      CreateCentralPackageReceipt = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   goSession.KillObject loform
   goSession.KillObject oRec
   CreateCentralPackageReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwPackage.CreateCentralPackageReceipt.", Err.Number, Err.Description
   CreateCentralPackageReceipt = False

End Function


Public Function OpenPackage(FullFilename As String) As Boolean

   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim strIniFile As String
   Dim sSQL As String
   Dim strTargetSite As String
   Dim loZip As Zipper
   
   On Error GoTo FunctionError
   If mIsOpen = True Then
      goSession.RaiseError "Error in mwSession.mwPackage, Package Already Open."
      OpenPackage = False
   End If
   
   Set loZip = New Zipper
   '
   ' Flush out Working Temp Folder - must be empty to start...
   '
   '
   ' Unzip to Temp Folder
   '
   If goSession.Workflow.IsZipEncryptWfFiles Then
      loZip.IsEncypt = True
   End If
   If Not loZip.OpenZip(FullFilename) Then
      goSession.RaiseError ("Error in mwSession.mwPackage.OpenPackage, unable to open Zip File: " _
        & FullFilename & ". Error Code: " & loZip.ErrorCode)
      OpenPackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   mWorkingFolder = moParent.WorkingSessionFolder
   If Trim(mWorkingFolder) = "" Then
      goSession.RaiseError "Error in mwSession.mwPackage.OpenPackage, Folder missing: " & mWorkingFolder
      OpenPackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   If Not loZip.UnzipToTempFolder(mWorkingFolder) Then
      goSession.RaiseError ("Error in mwSession.mwPackage.OpenPackage, unable to UnZip File: " & FullFilename)
      OpenPackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   
   loZip.TerminateClass
   Set loZip = Nothing
   
   '
   ' Check TargetSite.dat - must be destined for this site...
   '
   Set fso = New FileSystemObject
   strIniFile = mWorkingFolder & "\TargetSite.dat"
   If Not fso.FileExists(strIniFile) Then
      '
      ' Error: TargetSite.dat Control File Missing
      '
      goSession.RaiseError "Error in mwSession.mwPackage.OpenPackage, TargetSite.dat file missing from " & strIniFile
      OpenPackage = False
      goSession.KillObject fso
      Exit Function
   End If
   Set ts = fso.OpenTextFile(strIniFile)
   strTargetSite = goSession.API.StripNulls(ts.ReadLine)
   If strTargetSite <> goSession.Site.SiteID Then
      '
      ' Error: TargetSite.dat Control File Missing
      '
      goSession.RaiseError "Error in mwSession.mwPackage.OpenPackage, Package not destined for this site."
      OpenPackage = False
      ts.Close
      Set ts = Nothing
      goSession.KillObject fso
      Exit Function
   End If
   ts.Close
   Set ts = Nothing
   
   '
   ' Load up the Ini File
   '
   strIniFile = mWorkingFolder & "\" & pkg_PACKAGE_UPDATE_INI_FILE
   If Not fso.FileExists(strIniFile) Then
      '
      ' Error: INI Control File Missing
      '
      goSession.RaiseError "Error in mwSession.mwPackage.OpenPackage, INI file missing from " & strIniFile
      OpenPackage = False
      goSession.KillObject fso
      Exit Function
   End If
   Set moIni = New mwSession.IniConfig
   If Not moIni.LoadINIConfig(strIniFile) Then
      '
      ' Error: Package INI Control File is damaged
      '
      goSession.RaiseError "Error in mwSession.mwPackage.OpenPackage, INI file is damaged in Package: " & strIniFile
      goSession.KillObject fso
      goSession.KillObject moIni
      OpenPackage = False
      Exit Function
   End If
   '
   ' Test if already processed...
   '
   mFullFilename = FullFilename
   Set moRsRP = New Recordset
   moRsRP.CursorLocation = adUseClient
   sSQL = "select * from pkgReceivedPackage where ID=" & moIni.GetAppKey("package_header", "PackageID")
   moRsRP.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   If moRsRP.RecordCount > 0 Then
      If IsNull(moRsRP!IsInstalled) Then
         moRsRP!IsInstalled = False
      End If
      mIsInstalled = moRsRP!IsInstalled
   Else
      mIsInstalled = False
   End If
   goSession.CloseRecordset moRsRP
   '
   mIsOpen = True
   OpenPackage = True
   moParent.RaiseStatusMessage "Package Opened: " & FullFilename
   goSession.KillObject fso
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.OpenPackage.", Err.Number, Err.Description
   OpenPackage = False
   goSession.KillObject fso
   goSession.KillObject moIni
   mIsOpen = False
End Function


Private Function CreateLocalPackageReceipt(Optional Comments As String) As Boolean
   Dim loRsLPR As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   '
   ' Create local record...
   '
   If Not IsRecordLoaded(moRsRP) Then
      Exit Function
   End If
   Set loRsLPR = New Recordset
   loRsLPR.CursorLocation = adUseClient
   sSQL = "select * from pkgLocalPackageReceipt where PkgReceivedPackageKey = " & moRsRP!ID
   loRsLPR.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   If loRsLPR.RecordCount = 0 Then
      loRsLPR.AddNew
      loRsLPR!ID = goSession.MakePK("pkgLocalPackageReceipt")
      loRsLPR!PkgReceivedPackageKey = moRsRP!ID
      loRsLPR!mwcSitesKey = goSession.Site.SiteKey
      loRsLPR!mwcUsersKey = goSession.User.UserKey
      loRsLPR!EmployeeID = goSession.User.EmployeeID
      loRsLPR!mwcRoleTypekey = goSession.User.RoleTypeKey
      If Comments <> "" Then
         loRsLPR!Comments = Comments
      End If
      loRsLPR!ReceiptDate = Now()
      If Not moRsRP!IsdeveloperPackage Then
         loRsLPR!IsCentralPackageReceiptSent = 0
         loRsLPR!IsInstalledPackageReceiptSent = 0
      End If
      loRsLPR.Update
      mLocalPackageReceiptID = loRsLPR!ID
      CreateLocalPackageReceipt = True
   End If
   goSession.CloseRecordset loRsLPR
   ' We be done...
   'CreateLocalPackageReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.CreateLocalPackageReceipt.", Err.Number, Err.Description
   CreateLocalPackageReceipt = False
End Function


Private Function ReplaceVbCrLf(str As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   strTemp = Replace(str, "|", " ")
   strTemp = Replace(strTemp, "|", vbCrLf)
   If Len(Trim(strTemp)) = 0 Then
      strTemp = " "
   End If
   ReplaceVbCrLf = strTemp
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.ReplaceVbCrLf: ", Err.Number, Err.Description
   ReplaceVbCrLf = " "
End Function

Private Function Iso8601DateStringToVBDate(ByVal sISO8601 As String) As Date

    'format of the date is: 1997-07-16 19:20:30
    
    'formatting breakdown
    '                1
    '       1234567890123456789
    '       1997-07-16 19:20:30

    
   Dim sYear As String
   Dim sMonth As String

   Dim sDay As String
   Dim sHour As String
   Dim sMinute As String

   Dim sSecond As String

   On Error GoTo FunctionError
   If Len(sISO8601) = 19 Then
      sYear = Left$(sISO8601, 4)
      sMonth = mID$(sISO8601, 6, 2)
      sDay = mID$(sISO8601, 9, 2)
      sHour = mID$(sISO8601, 12, 2)
      sMinute = mID$(sISO8601, 15, 2)
      sSecond = mID$(sISO8601, 18, 2)
    
      Iso8601DateStringToVBDate = CDate(DateSerial(sYear, sMonth, sDay) & " " & _
                                        TimeSerial(sHour, sMinute, sSecond))
   Else
      Iso8601DateStringToVBDate = sISO8601
   End If
   Exit Function
FunctionError:
'      goSession.RaiseError "General Error in mwSession.mwPackage.Iso8601DateStringToVBDate, Input Date = [" & sISO8601 & "]", Err.Number, Err.Description
      Iso8601DateStringToVBDate = sISO8601
    
End Function

Private Function SendManualUpdateAlert(ByVal ManualChapterKey As Long, nAlertCount As Integer) As Boolean
   'PBT-1538/SF-8459 - Alert user when there is revision update on manual
   Dim sSQL As String
   Dim loRs As Recordset
   Dim lomwAlertWork As Object
   Dim sAlertTitle As String
   Dim sAlertDetails As String
   Dim AlertTargetSiteKey As Long
   Dim AlertTargetUsersKey As Long
   Dim AlertSourceRoleTypeKey As Long
   Dim AlertUsersKeySource As Long

   On Error GoTo FunctionError
   
   Set lomwAlertWork = CreateObject("mwSession.mwAlertWork")
   AlertSourceRoleTypeKey = goSession.User.RoleTypeKey
   AlertUsersKeySource = goSession.User.UserKey
   sAlertTitle = "Manual Chapter Revision Update Alert"
   AlertTargetSiteKey = goSession.Site.SiteKey
   
   sSQL = "SELECT smManualRoles.RoleTypeName ,smResponsibilities.Description ,smManualRoleUsers.mwcUsersKey FROM smResponsibilities INNER JOIN (smManualRoles INNER JOIN smManualRoleUsers ON smManualRoles.ID = smManualRoleUsers.smManualRolesKey ) ON smResponsibilities.smManualRolesKey = smManualRoles.ID " & _
          " WHERE smResponsibilities.smManualChapterKey =" & ManualChapterKey
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   Do While Not loRs.EOF
      
      sAlertDetails = "Manual Chapter Revision update alert has been received." & vbCrLf & vbCrLf
      sAlertDetails = sAlertDetails & "Manual Chapter Name: " & loRs!Description & vbCrLf
      sAlertDetails = sAlertDetails & "Group Name: " & loRs!RoleTypeName
      
      AlertTargetUsersKey = loRs!mwcUsersKey
      lomwAlertWork.CreateAlert AlertTargetSiteKey, _
                              sAlertTitle, _
                              0, _
                              sAlertDetails, _
                              MW_ALERT_STATUS_SENT, _
                              MW_ALERT_TYPE_USER, _
                              AlertSourceRoleTypeKey, _
                              "", _
                              AlertUsersKeySource, _
                              AlertTargetUsersKey
     
      nAlertCount = nAlertCount + 1
      loRs.MoveNext
   Loop

   KillObject lomwAlertWork
   CloseRecordset loRs
   SendManualUpdateAlert = True
   
Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPackage.SendManualUpdateAlert. ", Err.Number, Err.Description
   CloseRecordset loRs
   KillObject lomwAlertWork
   SendManualUpdateAlert = False
End Function

Private Function CreateCentralInstalledPackage() As Boolean
   Dim loform As mwForm
   Dim oRec As ConfigKeys
   Dim sTargetSiteID As String
   
   On Error GoTo FunctionError
   If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
      ' Ship to Shore
      sTargetSiteID = goSession.Site.TargetReplicateSiteID
   Else
      ' Shore to itself
      sTargetSiteID = goSession.Site.SiteID
   End If
   
   '
   ' Create Remote record
   '
   Set loform = goSession.OpenNewForm()
   If Not loform.MakeNewForm("PKG_INSTALLED", , , , sTargetSiteID) Then
      CreateCentralInstalledPackage = False
      goSession.KillObject loform
      Exit Function
   End If
   
   Set oRec = loform.DataForm.AddRecord
   With oRec
      .Add "Table", "pkgCentralInstalledPackage"
      .Add "Operation", "AMPK"
      
      .Add "PkgPackageKey", moRsRP!ID
      .Add "mwcSitesKey", goSession.Site.SiteKey
      .Add "mwcUsersKey", goSession.User.UserKey
      .Add "IsInstalled", 1
      .Add "InstalledDate", Format(goSession.API.GetGmtNow, DATAGRAM_DATE_FORMAT)
      .Add "mwcRoleTypeKey", goSession.User.RoleTypeKey
   End With
   
   '
   ' Write the message...
   '
   If Not loform.DataForm.WriteDataForm() Then
      CreateCentralInstalledPackage = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   '
   ' Transmit the message
   '
   loform.DeleteAfterSend = True
   If Not loform.Send() Then
      ' Error sending email
      CreateCentralInstalledPackage = False
      goSession.KillObject loform
      goSession.KillObject oRec
      Exit Function
   End If
   goSession.KillObject loform
   goSession.KillObject oRec
   CreateCentralInstalledPackage = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.CreateCentralInstalledPackage.", Err.Number, Err.Description
   CreateCentralInstalledPackage = False

End Function

Public Function ProcessPendingAlerts() As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim nChapterKey As Long
   Dim nAlertCount As Integer
   
   On Error GoTo FunctionError
    
   sSQL = "Select * from pkgReceivedPackageItem where IsAlertNotificationSent = 0 "
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic

   If Not loRs.BOF Then
      loRs.MoveFirst
   End If
   
   Do While loRs.EOF = False
      nAlertCount = 0
      nChapterKey = 0
      If Not IsNull(loRs!ManualChapterKey) Then
         nChapterKey = loRs!ManualChapterKey
      ElseIf Not IsNull(loRs!ItemKey) Then
         nChapterKey = loRs!ItemKey
      End If
      If nChapterKey > 0 Then
         If SendManualUpdateAlert(nChapterKey, nAlertCount) Then
            If nAlertCount > 0 Then
               loRs!IsAlertNotificationSent = True
               loRs.Update
            End If
         End If
      End If
      loRs.MoveNext
   Loop
   
   goSession.CloseRecordset loRs
   
   ProcessPendingAlerts = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.ProcessPendingAlerts.", Err.Number, Err.Description
   ProcessPendingAlerts = False
   mIsFatalError = True
End Function

Public Function ProcessPendingCentralPackageReceipt() As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loform As mwForm
   Dim oRec As ConfigKeys
   Dim sTargetSiteID As String
   
   On Error GoTo FunctionError
    
   sSQL = "Select * from pkglocalpackagereceipt where IsCentralPackageReceiptSent = 0 "
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic

   If Not loRs.BOF Then
      loRs.MoveFirst
   End If
   
   If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
      ' Ship to Shore
      sTargetSiteID = goSession.Site.TargetReplicateSiteID
   Else
      ' Shore to itself
      sTargetSiteID = goSession.Site.SiteID
   End If
      
   Do While loRs.EOF = False
      '
      ' Create Remote record
      '
      Set loform = goSession.OpenNewForm()
      If Not loform.MakeNewForm("PKG_RECEIPT", , , , sTargetSiteID) Then
         ProcessPendingCentralPackageReceipt = False
         goSession.KillObject loform
         CloseRecordset loRs
         Exit Function
      End If
   
   
      Set oRec = loform.DataForm.AddRecord
      With oRec
         .Add "Table", "pkgCentralPackageReceipt"
         .Add "Operation", "AMPK"
         
         .Add "PkgPackageKey", loRs!PkgReceivedPackageKey
         .Add "mwcSitesKey", loRs!mwcSitesKey
         .Add "mwcUsersKey", loRs!mwcUsersKey
         .Add "EmployeeID", BlankNull(loRs!EmployeeID)
         .Add "mwcRoleTypeKey", loRs!mwcRoleTypekey
         .Add "ReceiptDate", Format(loRs!ReceiptDate, DATAGRAM_DATE_FORMAT)
         .Add "pkgLocalPackageReceiptKey", str(loRs!ID)
      End With
      '
      ' Write the message...
      '
      If Not loform.DataForm.WriteDataForm() Then
         ProcessPendingCentralPackageReceipt = False
         goSession.KillObject loform
         goSession.KillObject oRec
         Exit Function
      End If
      '
      ' Transmit the message
      '
      loform.DeleteAfterSend = True
      If Not loform.Send() Then
         ' Error sending email
         ProcessPendingCentralPackageReceipt = False
         goSession.KillObject loform
         goSession.KillObject oRec
         CloseRecordset loRs
         Exit Function
      End If
   
      loRs!IsCentralPackageReceiptSent = 1
      loRs.Update

      goSession.KillObject loform
      goSession.KillObject oRec
      
      loRs.MoveNext
      
   Loop
   
   
   CloseRecordset loRs
   
   ProcessPendingCentralPackageReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.ProcessPendingCentralPackageReceipt.", Err.Number, Err.Description
   ProcessPendingCentralPackageReceipt = False
   mIsFatalError = True
End Function

Public Function ProcessPendingInstalledPackageReceipt() As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loform As mwForm
   Dim oRec As ConfigKeys
   Dim sTargetSiteID As String
   
   On Error GoTo FunctionError
    
   sSQL = "Select * from pkglocalpackagereceipt where IsInstalledPackageReceiptSent = 0 "
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, moCon, adOpenDynamic, adLockOptimistic

   If Not loRs.BOF Then
      loRs.MoveFirst
   End If
   
   If goSession.Site.SiteType <> SITE_TYPE_SHORE Then
      ' Ship to Shore
      sTargetSiteID = goSession.Site.TargetReplicateSiteID
   Else
      ' Shore to itself
      sTargetSiteID = goSession.Site.SiteID
   End If
      
   Do While loRs.EOF = False
   
      '
      ' Create Remote record
      '
      Set loform = goSession.OpenNewForm()
      If Not loform.MakeNewForm("PKG_INSTALLED", , , , sTargetSiteID) Then
         ProcessPendingInstalledPackageReceipt = False
         goSession.KillObject loform
         CloseRecordset loRs
         Exit Function
      End If
    
      Set oRec = loform.DataForm.AddRecord
      With oRec
         .Add "Table", "pkgCentralInstalledPackage"
         .Add "Operation", "AMPK"
         
         .Add "PkgPackageKey", loRs!PkgReceivedPackageKey
         .Add "mwcSitesKey", loRs!mwcSitesKey
         .Add "mwcUsersKey", loRs!mwcUsersKey
         .Add "IsInstalled", 1
         .Add "InstalledDate", Format(loRs!ReceiptDate, DATAGRAM_DATE_FORMAT)
         .Add "mwcRoleTypeKey", loRs!mwcRoleTypekey
      End With
   
      '
      ' Write the message...
      '
      If Not loform.DataForm.WriteDataForm() Then
         ProcessPendingInstalledPackageReceipt = False
         goSession.KillObject loform
         goSession.KillObject oRec
         Exit Function
      End If
      '
      ' Transmit the message
      '
      loform.DeleteAfterSend = True
      If Not loform.Send() Then
         ' Error sending email
         ProcessPendingInstalledPackageReceipt = False
         goSession.KillObject loform
         goSession.KillObject oRec
         CloseRecordset loRs
         Exit Function
      End If
   
      loRs!IsInstalledPackageReceiptSent = 1
      loRs.Update
      
      goSession.KillObject loform
      goSession.KillObject oRec

      loRs.MoveNext
      
   Loop
   

   CloseRecordset loRs
   
   ProcessPendingInstalledPackageReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.ProcessPendingInstalledPackageReceipt.", Err.Number, Err.Description
   ProcessPendingInstalledPackageReceipt = False
   mIsFatalError = True
End Function

Private Function UpdateLocalPackageReceipt(npkgLocalPackageReceiptKey As Long, nPackageReceiptType As Integer) As Boolean
   Dim loRsLPR As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError

   If npkgLocalPackageReceiptKey < 1 Then
      Exit Function
   End If
   
   Set loRsLPR = New Recordset
   loRsLPR.CursorLocation = adUseClient
   sSQL = "select * from pkgLocalPackageReceipt where ID = " & npkgLocalPackageReceiptKey
   loRsLPR.Open sSQL, moCon, adOpenDynamic, adLockOptimistic
   
   If IsRecordLoaded(loRsLPR) Then
      If nPackageReceiptType = CENTRALPKG_RECEIPT Then
         loRsLPR!IsCentralPackageReceiptSent = 1
      ElseIf nPackageReceiptType = INSTALLEDPKG_RECEIPT Then
         loRsLPR!IsInstalledPackageReceiptSent = 1
      End If
      loRsLPR.Update
   End If
   
   CloseRecordset loRsLPR
   UpdateLocalPackageReceipt = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwPackage.UpdateLocalPackageReceipt.", Err.Number, Err.Description
End Function
