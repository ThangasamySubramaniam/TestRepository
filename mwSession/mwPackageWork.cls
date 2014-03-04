VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwPackageWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const PROGRESS_CREATE_PACKAGE_INCREMENT = 20
Const PROGRESS_SEND_PACKAGE_INCREMENT = 80


Const PKG_PACKAGE_FILE_EXTENSION = ".ini"
Const PKG_PACKAGE_PREFIX = "P"
Const SEND_CDROM_FALSE = False
Const SEND_CDROM_TRUE = True



Public Function AddPackage(PackageType As Long, Optional EventTypeKey As Long, _
 Optional EventGroupKey As Long, Optional NoCreateEvent As Boolean) As Long
   Dim loRs As Recordset
   Dim loRsSites As Recordset
   Dim loRsPkgSite As Recordset
   Dim loEvWork As mwEventWork
   Dim loEhWork As mwEventHistoryWork
   Dim sSQL As String
   Dim nEventKey As Long
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from pkgPackage where ID=-1"
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Move to higher level function
   'goCon.BeginTrans
      loRs.AddNew
      loRs!ID = goSession.MakePK("pkgPackage")
      loRs!sent = False
      loRs!DateCreated = Now()
      loRs!Report_DateTime = Now()
      loRs!pkgPackageTypeKey = PackageType
      loRs!mwcUsersKey = goSession.User.UserKey
      loRs!mwcRoleTypekey = goSession.User.RoleTypeKey
      loRs!mwcSitesKeySource = goSession.Site.SiteKey
      
      If Not NoCreateEvent Then
         Set loEvWork = CreateObject("mwSession.mwEventWork")
         nEventKey = loEvWork.AddEvent(EventTypeKey, loRs, , EventGroupKey, , goSession.Site.SiteKey)
         
         If nEventKey < 1 Then
            ' Cancel Operation, no event created...
            goSession.RaiseError "Error in mwPackageWork.AddPackage. ", Err.Number, Err.Description
            loRs.CancelUpdate
            CloseRecordset loRs
            KillObject loEvWork
            AddPackage = -1
            'goCon.RollbackTrans
            Exit Function
         End If
         loRs!mwEventLogKey = nEventKey
         KillObject loEvWork
      End If
      loRs.Update
      AddPackage = loRs!ID
      '
      ' Add Site Records...
      '
      Set loRsSites = New Recordset
      loRsSites.CursorLocation = adUseClient
      sSQL = "select * from mwcSites where IsReplicateSite Is Not Null and IsReplicateSite<>0 order by SiteName"
      loRsSites.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
      If loRsSites.RecordCount < 1 Then
         goSession.RaiseError "Error in mwPackageWork.AddPackage, no distribution/replication sites defined. "
         CloseRecordset loRs
         CloseRecordset loRsSites
         KillObject loEvWork
         AddPackage = -1
         'goCon.RollbackTrans
         Exit Function
      End If
      
      
      Set loRsPkgSite = New Recordset
      loRsPkgSite.CursorLocation = adUseClient
      sSQL = "select * from pkgPackageSite where mwcSitesKey=-1"
      loRsPkgSite.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      loRsSites.MoveFirst
      Do While Not loRsSites.EOF
         loRsPkgSite.AddNew
         loRsPkgSite!pkgPackageKey = loRs!ID
         loRsPkgSite!mwcSitesKey = loRsSites!ID
         loRsPkgSite!SiteName = loRsSites!SiteName
         loRsPkgSite!NoTimesSent = 0
         loRsPkgSite!SendUpdate = 0
         loRsPkgSite!ConfirmedReceipt = 0
         loRsPkgSite.Update
         loRsSites.MoveNext
      Loop
      If Not NoCreateEvent Then
         Set loEhWork = CreateObject("mwsession.mwEventHistoryWork")
         loEhWork.AddSystemRecord EventTypeKey, loRs.Fields("ID").value, _
          "PACKAGE UPDATE", "Package has been created.", , , , goSession.Site.SiteKey
         KillObject loEhWork
      End If
   'goCon.CommitTrans
   CloseRecordset loRs
   CloseRecordset loRsSites
   CloseRecordset loRsPkgSite
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwPackageWork.AddPackage. ", Err.Number, Err.Description
   AddPackage = -1
   On Error Resume Next
   'goCon.RollbackTrans
   CloseRecordset loRs
   CloseRecordset loRsSites
   CloseRecordset loRsPkgSite
   KillObject loEvWork
   KillObject loEhWork
End Function

Public Function AddPackageItem(PackageKey As Long, SourceFileLocation As String, _
 Optional PlaceID As String, Optional EventType As Long, Optional EventDetail As Long, _
 Optional ShortDescription As String, Optional Reference As String, _
 Optional LongDescription As String, Optional ManualChapterKey As Long, Optional ManualRevisionNo As Long, _
 Optional ManualRevisionDate As Date, Optional mwrChangeTableKey As Long, Optional SourceBLOBKey As Long) As Long
'Public Function AddPackageItem(PackageKey As Long, SourceFileLocation As String, _
' Optional PlaceID As String, Optional EventType As Long, Optional EventDetail As Long, _
' Optional ShortDescription As String, Optional Reference As String, _
' Optional LongDescription As String) As Long
   Dim loRsPI As Recordset
   Dim loEhWork As mwEventHistoryWork
   Dim sSQL As String
   On Error GoTo FunctionError:
   CloseRecordset loRsPI
   Set loRsPI = New Recordset
   loRsPI.CursorLocation = adUseClient
   sSQL = "select * from pkgPackageItem where ID=-1"
   loRsPI.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRsPI.AddNew
   loRsPI!ID = goSession.MakePK("pkgPackageItem")
   loRsPI!pkgPackageKey = PackageKey
   loRsPI!SourceFileLocation = SourceFileLocation
   If Not IsNull(PlaceID) Then
      loRsPI!TargetLocation = PlaceID
   End If
   If Not IsNull(Reference) Then
      loRsPI!Reference = Reference
   End If
    If Not IsNull(ShortDescription) Then
      loRsPI!ShortDescription = Left$(ShortDescription, 50)
      SetPackageName PackageKey, ShortDescription
   End If
   If Not IsNull(LongDescription) Then
      loRsPI!LongDescription = LongDescription
   End If
   
   loRsPI!mwrChangeTableKey = mwrChangeTableKey
   
   If mwrChangeTableKey = 0 Or mwrChangeTableKey = MWRT_smManualChapter Then
      If ManualChapterKey > 0 Then
         loRsPI!ManualChapterKey = ManualChapterKey
      End If
      If ManualRevisionNo > 0 Then
         loRsPI!ManualRevisionNo = ManualRevisionNo
      End If
      If ManualRevisionDate > 0 Then
         loRsPI!ManualRevisionDate = ManualRevisionDate
      End If
   Else
      If ManualChapterKey > 0 Then
         loRsPI!ItemKey = ManualChapterKey
      End If
      If ManualRevisionNo > 0 Then
         loRsPI!ItemRevisionNo = ManualRevisionNo
      End If
      If ManualRevisionDate > 0 Then
         loRsPI!ItemRevisionDate = ManualRevisionDate
      End If
   End If
   
   If SourceBLOBKey > 0 Then
      loRsPI!SourceBLOBKey = SourceBLOBKey
   End If
   
   
   'loRsPI!TargetLocation = PACKAGE_ITEM_TARGET_LOCATION_MANUALS
   loRsPI.Update
   
   ' Validate - Event History Record should be updated if parameters are passed.
   AddPackageItem = loRsPI!ID
   If EventDetail > 0 Then
      Set loEhWork = CreateObject("mwsession.mwEventHistoryWork")
      loEhWork.AddSystemRecord EventType, EventDetail, "PACKAGE ITEM UPDATE", _
       "Package Item has been created.", SourceFileLocation
      KillObject loEhWork
      Exit Function
   End If
   CloseRecordset loRsPI
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwPackageWork.AddPackage. ", Err.Number, Err.Description
   CloseRecordset loRsPI
   AddPackageItem = -1
   KillObject loEhWork
End Function

Private Function FetchUnsuspendedSiteKeys() As String
   On Error GoTo FunctionError
   

   FetchUnsuspendedSiteKeys = "select ID from mwcSites where (IsSuspendTransmit Is Null OR IsSuspendTransmit=0)"

   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwPackageWork.FetchUnsuspendedSiteKeys. ", Err.Number, Err.Description
End Function

Public Function SendPackage(PackageKey As Long, IsSendByMedia As Boolean, _
Optional ResendSiteKey As Long, Optional objProgressBar As Object, Optional objLblProgress As Object) As Boolean
   Dim loZip As mwSession.Zipper
   Dim sTargetSiteID As String
   Dim sPkgFile As String
   Dim IsEmailSend As Boolean
   Dim fso As FileSystemObject
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loRsPS As Recordset
   Dim UnsuspendedSites As String
   Dim i  As Integer
   Dim j As Integer
   Dim nProgIncr As Integer
   Dim nProgIncrComplete As Integer
   Dim sPkgOriFile As String
   Dim loFso As FileSystemObject
   
   On Error GoTo FunctionError
   
   Set loZip = New Zipper
   
   sPkgFile = goSession.Workflow.ArchiveOkZipFile & "\" & _
    PKG_PACKAGE_PREFIX & goSession.Site.ShortSiteID & Right(PackageKey, 4) & goSession.Workflow.Package.PackageExtension
   
   If goSession.Workflow.IsZipEncryptWfFiles = True Then
      loZip.IsEncypt = True
   End If
   If Not loZip.OpenZip(sPkgFile) Then
      goSession.RaiseError "Error in mwPackageWork.SendPackage, unable to open zip file: " & sPkgFile
      SendPackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   
   UnsuspendedSites = FetchUnsuspendedSiteKeys()
   '
   ' Open two recordsets...
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from pkgPackage where ID=" & PackageKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   Set loRsPS = New Recordset
   loRsPS.CursorLocation = adUseClient
   If ResendSiteKey > 0 Then
      If UnsuspendedSites = "" Then
         sSQL = "select * from pkgPackageSite where pkgPackageKey=" & PackageKey & _
         " and mwcSitesKey=" & ResendSiteKey
      Else
         sSQL = "select * from pkgPackageSite where pkgPackageKey=" & PackageKey & _
         " and mwcSitesKey=" & ResendSiteKey & " AND mwcSitesKey IN (" & UnsuspendedSites & ")"
      End If
   Else
      If UnsuspendedSites = "" Then
         sSQL = "select * from pkgPackageSite where pkgPackageKey=" & PackageKey & _
          " and SendUpdate<>0 and FirstSentDate is Null"
      Else
         sSQL = "select * from pkgPackageSite where pkgPackageKey=" & PackageKey & _
          " and SendUpdate<>0 and FirstSentDate is Null AND mwcSitesKey IN (" & UnsuspendedSites & ")"
      End If
   End If
   loRsPS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not objProgressBar Is Nothing Then
      If loRsPS.RecordCount > 0 Then
         nProgIncr = PROGRESS_SEND_PACKAGE_INCREMENT / loRsPS.RecordCount
         nProgIncrComplete = PROGRESS_SEND_PACKAGE_INCREMENT Mod loRsPS.RecordCount
      End If
   End If
   
   sPkgOriFile = goSession.Workflow.ArchiveOkZipFile & "\" & _
    PKG_PACKAGE_PREFIX & goSession.Site.ShortSiteID & Right(PackageKey, 4) & goSession.Workflow.Package.PackageExtension
   
   Set loFso = New FileSystemObject
   '
   ' Loop through each site receiving package
   '
   Do While Not loRsPS.EOF
   
      'DEV-1913 - Added By Krishna Kumar - To Re-Create the Zip file with Current Site ID
      
      If Not objLblProgress Is Nothing Then
         If Not IsNull(loRsPS!SiteName) Then
            objLblProgress.Caption = "Sending Package to Site : " & Trim(loRsPS!SiteName)
         End If
      End If
      '
      ' Add TargetSiteID to file
      '
      sTargetSiteID = goSession.Site.GetSiteID(loRsPS!mwcSitesKey)
      If Trim(sTargetSiteID) = "" Then
         goSession.RaiseError "Error in mwPackageWork.SendPackage, SiteID not found for SiteKey: " & _
          loRsPS!mwcSitesKey
         CloseRecordset loRs
         CloseRecordset loRsPS
         SendPackage = False
         loZip.TerminateClass
         Set loZip = Nothing
         Exit Function
      End If
      
      sPkgFile = goSession.Workflow.ArchiveOkZipFile & "\" & _
       PKG_PACKAGE_PREFIX & goSession.Site.ShortSiteID & goSession.Site.GetShipProperty(sTargetSiteID, "ShortSiteID") & _
       Right(PackageKey, 4) & goSession.Workflow.Package.PackageExtension
   
      loFso.CopyFile sPkgOriFile, sPkgFile, True
      
      If Not loZip.OpenZip(sPkgFile) Then
         goSession.RaiseError "Error in mwPackageWork.SendPackage, unable to open zip file: " & sPkgFile
         SendPackage = False
         loZip.TerminateClass
         Set loZip = Nothing
         Exit Function
      End If
      
      loZip.ZipStringToFile sTargetSiteID, "TargetSite.dat"
      
      If IsSendByMedia Then
         '
         ' Transport Media
         '
         If Not goSession.Workflow.SendByMedia(sPkgFile, , , , sTargetSiteID, SEND_CDROM_TRUE, ZeroNull(loRsPS!mwcSitesKey)) Then
            goSession.RaiseWorkflowStatus "Package FAILED send by media to: " & sTargetSiteID
            SendPackage = False
            CloseRecordset loRs
            CloseRecordset loRsPS
'            KillObject fso
            loZip.TerminateClass
            Set loZip = Nothing
            Exit Function
         Else
            goSession.RaiseWorkflowStatus "Package sent by media to: " & sTargetSiteID
         End If
      Else
         '
         ' Default Transport
         '
         If goSession.User.DefaultTransport = mwt_Media Then
            If Not goSession.Workflow.SendByMedia(sPkgFile, , , , sTargetSiteID, SEND_CDROM_FALSE, ZeroNull(loRsPS!mwcSitesKey)) Then
               goSession.RaiseWorkflowStatus "Package FAILED send by media to: " & sTargetSiteID
               SendPackage = False
               CloseRecordset loRs
               CloseRecordset loRsPS
'               KillObject fso
               loZip.TerminateClass
               Set loZip = Nothing
               Exit Function
            Else
               goSession.RaiseWorkflowStatus "Package sent by media to: " & sTargetSiteID
            End If
         ElseIf goSession.User.DefaultTransport = mwt_ElectronicMail Then
            With goSession.Message
               ' 7/2002 fix ms
               .Carrier = goSession.User.DefaultEmailCarrier
               .Server = goSession.User.MailServerName
               '
               ' Standard Message Contents...
               '
               
               If goSession.Workflow.IsZipWfFiles Then
                  sPkgFile = goSession.Workflow.CreateZipEncryptFile(sPkgFile)
               End If
               
               If ZeroNull(loRsPS!mwcSitesKey) > 0 Then
                  .AttachmentFile = goSession.Workflow.MaybeCreateZipInZipFile(sPkgFile, ZeroNull(loRsPS!mwcSitesKey))
               Else
                  .AttachmentFile = sPkgFile
               End If
               
               If IsNull(loRs!SpecialInstructions) Then
                  .Body = " "
               Else
                  .Body = loRs!SpecialInstructions
               End If
               
               .SendToAddress = goSession.Site.GetShipWorkflowSendToAddress(, ZeroNull(loRsPS!mwcSitesKey))
               
               ' Send the emails...
               If .SendToAddress = "" Then
                  goSession.RaiseError "Error in mwSession.SendPackage, Ship Workflow Send To Address is empty."
                  CloseRecordset loRs
                  CloseRecordset loRsPS
                  SendPackage = False
                  loZip.TerminateClass
                  Set loZip = Nothing
                  Exit Function
               End If
'               .SendFromAddress = goSession.User.GetExtendedProperty("FromEmailAddress")
               .SendFromAddress = goSession.Site.WorkflowSendFromAddress
               .subject = " Package Update: " & loRs!ID & " - " _
                 & loRs!PackageName & " for the " & Trim(loRsPS!SiteName)
               If Not .Send Then
                  goSession.RaiseWorkflowStatus "Package Failed transmit to Site: " & loRsPS!mwcSitesKey
                  CloseRecordset loRs
                  CloseRecordset loRsPS
                  SendPackage = False
                  loZip.TerminateClass
                  Set loZip = Nothing
                  Exit Function
               End If
               goSession.RaiseWorkflowStatus "Package transmitted to Site: " & loRsPS!mwcSitesKey
               loRsPS!SentToAddress = .SendToAddress
            End With
         Else
            goSession.RaiseError "Error in mwPackageWork.SendPackage, no valid email transport defined."
            SendPackage = False
            CloseRecordset loRs
            CloseRecordset loRsPS
'            KillObject fso
            loZip.TerminateClass
            Set loZip = Nothing
            Exit Function
         End If
      End If
      '
      ' Update The Record...
      '
      If IsNull(loRsPS!FirstSentDate) Then
         loRsPS!FirstSentDate = Now()
      End If
      loRsPS!LastSentDate = Now()
      loRsPS!NoTimesSent = loRsPS!NoTimesSent + 1
      loRsPS!SentToAddress = BlankNull(goSession.Site.GetShipWorkflowSendToAddress(, loRsPS!mwcSitesKey))
      loRsPS.Update
      
      If Not objProgressBar Is Nothing Then
         objProgressBar.value = objProgressBar.value + nProgIncr
         DoEvents
      End If

      loRsPS.MoveNext
   Loop
   '
   ' Update Package Record
   '
   '
   ' All sites have been transmitted to...
   '
   loRs!DateTransmitted = Now()
   loRs!sent = True
   SendPackage = True
   CloseRecordset loRs
   CloseRecordset loRsPS
   
   If Not objProgressBar Is Nothing Then
      objProgressBar.value = objProgressBar.value + nProgIncrComplete
      DoEvents
   End If
   
   loZip.TerminateClass
   Set loZip = Nothing
   
   KillObject loFso
   
   ' Reexecute the query bound to the UG Control...
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwPackageWork.SendPackage. ", _
        Err.Number, Err.Description
   CloseRecordset loRs
   CloseRecordset loRsPS
End Function

Public Function CreatePackage(PackageKey As Long, isSilent As Boolean, Optional objProgressBar As Object, Optional objLblProgress As Object) As Boolean
#If LATE_BIND Then
   Dim loIni As Object
   Dim loZip As Object
#Else
   Dim loIni As mwSession.IniConfig
   Dim loZip As mwSession.Zipper
#End If
   Dim fso As FileSystemObject
   Dim loRs As Recordset
   Dim loRsItems As Recordset
   Dim sGroup As String
   Dim sSQL As String
   Dim sIniFile As String
   Dim sPkgFile As String
   Dim nProgIncr As Integer
   Dim nProgIncrComplete As Integer
   On Error GoTo FunctionError
   
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from pkgPackage where ID=" & PackageKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
'   If Not IsNull(loRs!DatePrepared) Then
'      If isSilent Then
'         goSession.LogIt mwl_Workstation, mwl_Warning, "Error in mwSession.mwPackageWork.CreatePackage, Package already created: " & loRs!ID
'      Else
'         goSession.RaiseError "Error in mwSession.mwPackageWork.CreatePackage, Package already created: " & loRs!ID
'      End If
'      CreatePackage = True
'      CloseRecordset loRs
'      Exit Function
'   End If

   '
   ' is Valid package ?
   '
   'stop here
   
   
   
   Set loRsItems = New Recordset
   loRsItems.CursorLocation = adUseClient
   sSQL = "select * from pkgPackageItem where PkgPackageKey=" & PackageKey
   loRsItems.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If Not objProgressBar Is Nothing Then
      If loRsItems.RecordCount > 0 Then
         nProgIncr = PROGRESS_CREATE_PACKAGE_INCREMENT / loRsItems.RecordCount
         nProgIncrComplete = PROGRESS_CREATE_PACKAGE_INCREMENT Mod loRsItems.RecordCount
      End If
   End If
   If Not objLblProgress Is Nothing Then
      objLblProgress.Caption = "Creating Package..."
   End If
   
   '
   ' Make sure Folders exist...
   '
   Set fso = New FileSystemObject
   '
   ' Create Package (Zip) File: Prefix (P) + ShortSiteID (99) + Package ID (9999)
   '
   sPkgFile = goSession.Workflow.WorkingSessionFolder & "\" & _
     PKG_PACKAGE_PREFIX & goSession.Site.ShortSiteID & Right(loRs!ID, 4) & goSession.Workflow.Package.PackageExtension
     
   sIniFile = goSession.Workflow.WorkingSessionFolder & "\pkg.ini"
   '
   ' Create Control File
   '
   If fso.FileExists(sIniFile) Then
      fso.DeleteFile (sIniFile)
   End If
   
   If fso.FileExists(sPkgFile) Then
      fso.DeleteFile sPkgFile
   End If
   '
   ' Step 1 Make the Control Config group...
   '
   Set loIni = CreateObject("mwSession.IniConfig")
   Set loZip = New Zipper
   ' Leave source doc alone
   If goSession.Workflow.IsZipEncryptWfFiles Then
      loZip.IsEncypt = True
   End If
   loZip.DeleteAfterZip = False
   
   loIni.ConfigGroups.Add "PACKAGE_HEADER", "PACKAGE_HEADER"
   With loIni.ConfigGroups("PACKAGE_HEADER").ConfigKeys

      .Add "PackageID", loRs!ID
      .Add "PackageType", loRs!pkgPackageTypeKey
      .Add "PackageName", loRs!PackageName
      .Add "DateCreated", Format(loRs!DateCreated, "yyyy-mm-dd Hh:Nn:Ss")
      .Add "mwcRoleTypeKey", loRs!mwcRoleTypekey
      .Add "DateTransmitted", Format(Now(), "yyyy-mm-dd Hh:Nn:Ss")
      .Add "ISODateTransmitted", Format(Now(), "yyyy-mm-dd Hh:Nn:Ss")
      If Not IsNull(loRs!SpecialInstructions) Then
         '.Add "SpecialInstructions",  StripVbCrLf(loRS!SpecialInstructions)
      End If
   End With
   '
   ' Step 2 Make the Zip
   '
   If Not loZip.OpenZip(sPkgFile) Then
      goSession.RaisePublicError "Error in mwSession.mwPackageWork.CreatePackage Creating package Zip File: " & _
       sPkgFile & "Zip Error: " & loZip.ErrorCode
      CreatePackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   '
   ' Loop through chm files...
   'INI file
   '
   ' Step 3 - Add files to zip...
   '
   loRsItems.MoveFirst
   Do While Not loRsItems.EOF
      ' Zip it
      
      If ZeroNull(loRsItems!SourceBLOBKey) > 0 Then
      
         Dim TempFile As String
         Dim loBlobWork As mwBlobWork
         
         Set loBlobWork = New mwBlobWork
         
         TempFile = loBlobWork.FetchToFileSystem(loRsItems!SourceBLOBKey, False, True, False)
         
         If Not loZip.AddToZip(TempFile) Then
            goSession.RaisePublicError "Error in mwSession.mwPackageWork.CreatePackage Zipping Document: " & _
              TempFile & ". Error Code: " & loZip.ErrorCode
            CreatePackage = False
            
            fso.DeleteFile TempFile
            
            loZip.TerminateClass
            Set loZip = Nothing
            KillObject loBlobWork
            KillObject fso
            Exit Function
         End If
         
         fso.DeleteFile TempFile
         
         KillObject loBlobWork
      
      Else
         If Not fso.FileExists(loRsItems!SourceFileLocation) Then
            goSession.RaisePublicError "Error in mwSession.mwPackageWork.CreatePackage: File not found: " & _
              loRsItems!SourceFileLocation
            CreatePackage = False
            loZip.TerminateClass
            Set loZip = Nothing
            Exit Function
         End If
   
         If Not loZip.AddToZip(loRsItems!SourceFileLocation) Then
            goSession.RaisePublicError "Error in mwSession.mwPackageWork.CreatePackage Zipping Document: " & _
              loRsItems!SourceFileLocation & ". Error Code: " & loZip.ErrorCode
            CreatePackage = False
            loZip.TerminateClass
            Set loZip = Nothing
            Exit Function
         End If
      End If
      ' Add to Control INI...
      sGroup = loRsItems!ID
      ' sGroup = fso.GetBaseName(loRSItems!SourceFileLocation)
      loIni.ConfigGroups.Add sGroup, sGroup
      With loIni.ConfigGroups(sGroup).ConfigKeys
         .Add "PackageItemID", _
           loRsItems!ID, "PackageItemID"
         .Add "TargetLocation", _
           loRsItems!TargetLocation, "TargetLocation"
         .Add "SourceFileLocation", _
           loRsItems!SourceFileLocation, "SourceFileLocation"
         If Not IsNull(loRsItems!Reference) Then .Add "Reference", _
           loRsItems!Reference, "Reference"
         If Not IsNull(loRsItems!ShortDescription) Then .Add "ShortDescription", _
           loRsItems!ShortDescription, "ShortDescription"
         If Not IsNull(loRsItems!LongDescription) Then .Add "LongDescription", _
           StripVbCrLf(loRsItems!LongDescription), "LongDescription"
         If Not IsNull(loRsItems!Hyperlink) Then
            .Add "Hyperlink", _
             loRsItems!Hyperlink, "Hyperlink"
         End If
         
         If ZeroNull(loRsItems!mwrChangeTableKey) > 0 Then .Add "mwrChangeTableKey", loRsItems!mwrChangeTableKey
         If ZeroNull(loRsItems!mwrChangeTableKey) = 0 Or ZeroNull(loRsItems!mwrChangeTableKey) = MWRT_smManualChapter Then
            If Not IsNull(loRsItems!ManualChapterKey) Then .Add "ManualChapterKey", loRsItems!ManualChapterKey
            If Not IsNull(loRsItems!ManualRevisionNo) Then .Add "ManualRevisionNo", loRsItems!ManualRevisionNo
            If Not IsNull(loRsItems!ManualRevisionDate) Then .Add "ManualRevisionDate", Format(loRsItems!ManualRevisionDate, "yyyy-mm-dd Hh:Nn:Ss")
         Else
            If Not IsNull(loRsItems!ItemKey) Then .Add "ItemKey", loRsItems!ItemKey
            If Not IsNull(loRsItems!ItemRevisionNo) Then .Add "ItemRevisionNo", loRsItems!ItemRevisionNo
            If Not IsNull(loRsItems!ItemRevisionDate) Then .Add "ItemRevisionDate", Format(loRsItems!ItemRevisionDate, "yyyy-mm-dd Hh:Nn:Ss")
         End If
         
         
      End With
      '
      If Not objProgressBar Is Nothing Then
         objProgressBar.value = objProgressBar.value + nProgIncr
         DoEvents
      End If
      loRsItems.MoveNext
   Loop
   loRsItems.MoveFirst
   '
   ' Step 3 Write the Control file to disk
   '
   loIni.WriteINIConfig sIniFile
   If Not loZip.AddToZip(sIniFile) Then
      goSession.RaisePublicError "Error in mwSession.frmRemotePackages.CreatePackage Zipping File failed: " & _
        sIniFile & ". Zip Error Code: " & loZip.ErrorCode
      CreatePackage = False
      loZip.TerminateClass
      Set loZip = Nothing
      Exit Function
   End If
   
   loZip.TerminateClass
   Set loZip = Nothing
   fso.DeleteFile sIniFile
   '
   ' Baseline Package (without TargetSite.dat) - Back it up...
   '
   If Not goSession.Workflow.ArchiveWorkflowItem(sPkgFile, ArchivedOK, False) Then
      CreatePackage = False
      Exit Function
   End If
   '
   ' Update Package Record...
   '
   goCon.BeginTrans
      If IsNull(loRs!DatePrepared) Then
         loRs!DatePrepared = Now()
         loRs.Update
         loRs.Move (0)
      End If
      '
      ' Clear out unused PackageSite Records...
      '
      sSQL = "delete from pkgPackageSite where pkgPackageKey=" & loRs!ID & " and SendUpdate=0"
      goCon.Execute sSQL
   goCon.CommitTrans
   ' close
   If Not objProgressBar Is Nothing Then
      objProgressBar.value = objProgressBar.value + nProgIncrComplete
      DoEvents
   End If
   
   CreatePackage = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwSession.mwPackageWork.CreatePackage. ", Err.Number, Err.Description
   Exit Function
End Function


Private Function IsValidPackage(morsItems As Recordset) As Boolean
'   On Error GoTo FunctionError
'   '
'   '
'   Dim strErrorInPackage As String
   MsgBox "IsValidPackage Stump Code"
   Exit Function
   
   
'   IsValidPackage = False
'   morsItems.MoveFirst
'   Do While Not morsItems.EOF
'      If IsNull(morsItems!TargetLocation) Then
'         strErrorInPackage = strErrorInPackage & "Target Location needs Value " & vbCrLf
'         IsValidPackage = True
'      End If
'      If IsNull(morsItems!SourceFileLocation) Then
'         strErrorInPackage = strErrorInPackage & "Source File Location needs Value " & vbCrLf
'         IsValidPackage = True
'      End If
'      If IsNull(moRS!PackageName) Then
'         strErrorInPackage = strErrorInPackage & "Package Name needs Value " & vbCrLf
'         IsValidPackage = True
'      End If
'      morsItems.MoveNext
'   Loop
   Exit Function
FunctionError:
   IsValidPackage = True
   goSession.RaisePublicError "General Error in mwSession.mwPackageWork.IsValidPackage: ", _
        Err.Number, Err.Description
End Function

Private Function StripVbCrLf(str As String) As String
   Dim strTemp As String
   On Error GoTo FunctionError
   strTemp = Replace(str, "|", " ")
   strTemp = Replace(strTemp, vbCrLf, "|")
   If Len(Trim(strTemp)) = 0 Then
      strTemp = " "
   End If
   StripVbCrLf = strTemp
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPackageWork.StripVbCrLf: ", Err.Number, Err.Description
   StripVbCrLf = " "
End Function

Private Sub SetPackageName(pkgPackageKey As Long, ShortDescription As String)
   Dim loRs As Recordset
   Dim sSQL As String
   
   On Error GoTo FunctionError
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from pkgPackage where ID=" & pkgPackageKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount = 1 Then
      If IsNull(loRs!PackageName) Then
         loRs!PackageName = Left(ShortDescription, 50)
         loRs.Update
      End If
   End If
   
   CloseRecordset loRs
   
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPackageWork.SetPackageName: ", _
        Err.Number, Err.Description
End Sub

Public Function CanTransmitPackage(nPackageKey As Long) As Boolean
   
   Dim sSQL As String
   Dim loRsItems As Recordset
   Dim loRSPackageSite As Recordset
   
   On Error GoTo FunctionError

   sSQL = "Select 1 from pkgPackageItem where pkgPackageKey=" & nPackageKey & " and TargetLocation is not null and SourceFileLocation is not null"
   
   Set loRsItems = New Recordset
   loRsItems.CursorLocation = adUseClient
   '
   ' Read Only Record set if Transmitted...
   loRsItems.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRsItems.RecordCount < 1 Then
      CloseRecordset loRsItems
      Exit Function
   End If
   CloseRecordset loRsItems
   
   
   sSQL = "SELECT pkgPackageSite.SendUpdate, pkgPackageSite.NoTimesSent,IsSuspendTransmit,mwcSites.IsSuspendTransmit FROM mwcSites RIGHT JOIN pkgPackageSite ON mwcSites.ID = pkgPackageSite.mwcSitesKey WHERE pkgPackageKey=" & nPackageKey
   
   Set loRSPackageSite = New Recordset
   loRSPackageSite.CursorLocation = adUseClient
   loRSPackageSite.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRSPackageSite.RecordCount = 0 Then
      CloseRecordset loRSPackageSite
      Exit Function
   End If
   
   loRSPackageSite.Filter = adFilterNone
   loRSPackageSite.Filter = "IsSuspendTransmit = null or IsSuspendTransmit=0"
   If loRSPackageSite.RecordCount < 1 Then
      CloseRecordset loRSPackageSite
      Exit Function
   End If
   loRSPackageSite.Filter = adFilterNone
   
   
   loRSPackageSite.Filter = "SendUpdate<>0 and NoTimesSent=0 "
   If loRSPackageSite.RecordCount < 1 Then
      CloseRecordset loRSPackageSite
      Exit Function
   End If
   loRSPackageSite.Filter = adFilterNone
   
   loRSPackageSite.Filter = "SendUpdate<>0 and IsSuspendTransmit=0 "
   If loRSPackageSite.RecordCount < 1 Then
      CloseRecordset loRSPackageSite
      Exit Function
   End If
   loRSPackageSite.Filter = adFilterNone
   
   CloseRecordset loRSPackageSite
   
   ' Make it
   CanTransmitPackage = True
   
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwPackageWork.CanTransmitPackage: ", Err.Number, Err.Description
   CloseRecordset loRsItems
   CloseRecordset loRSPackageSite
End Function

