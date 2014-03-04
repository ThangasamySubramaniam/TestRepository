VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMailbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mw_Mailbox - Process Mailbox "Inbox" class module
' 3/12/2001 ms
' 9/16/2001 Gatwick Overhaul
' 12/13/2001 MetroNorth Forwarding
Option Explicit

'Early binding constants
Const mw__OL_FOLDER_INBOX = 6
Const mw__CDO_DEFAULT_FOLDER_INBOX = 1


Private moParent As Session
Private mMailboxType As mw_EmailCarrier
Private mMailServer As String
Private mLoginID As String
Private mLoginPassword As String
Private mCurrentMessageCount As Integer
Private mMessageCount As Integer
'
'Outlook Objects
'
'Private moOutlook As Outlook.Application
'Private moOutlookInbox As Outlook.MAPIFolder
Private moOutlook As Object
Private moOutlookInbox As Object

'
' Exchange (CDO 1.2) Objects
'
'Private moSession As MAPI.Session
'Private moExchangeInbox As MAPI.Folder
Private moSession As Object
Private moExchangeInbox As Object

Const LOTUS_RICHTEXT = 1
Const LOTUS_EMBED_ATTACHMENT = 1454
Const LOTUS_SMTP = 1280
Const LOTUS_INBOX = "($Inbox)"

Const MW_EVENT_WORKFLOW_AGENT = 908
'
' Dart POP3 Objects...
'

#If NO_DART = 1 Then
   Private moPOP3 As Object
   Private moPOPMessages As Object
#Else
   Private moPOP3 As DartMail.Pop
   Private moPOPMessages As DartMail.Messages
#End If
'
' Novell Groupwise objects
'
Dim moGWApp As Object
Dim moGWAccount As Object
Const GW_INBOX_FOLDER = 2
Const GW_egwIncoming = 1
Const GW_egwOutgoing = 2

Dim mWasOutlookObject As Boolean



Private mSaveAttachmentFolder As String
Private mIsLoggedIn As Boolean

Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
      mMailboxType = moParent.User.DefaultEmailCarrier
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Let MailboxType(ByVal Data As mw_EmailCarrier)
    mMailboxType = Data
End Property

Public Property Get MailboxType() As mw_EmailCarrier
    MailboxType = mMailboxType
End Property

Public Property Let MailServer(ByVal Data As String)
    mMailServer = Data
End Property

Public Property Get MailServer() As String
    MailServer = mMailServer
End Property

Public Property Get LoginID() As String
    LoginID = mLoginID
End Property

Public Property Let LoginID(ByVal Data As String)
    mLoginID = Data
End Property

Public Property Let LoginPassword(ByVal Data As String)
    mLoginPassword = Data
End Property

Public Property Get LoginPassword() As String
    LoginPassword = mLoginPassword
End Property

Public Property Get SaveAttachmentFolder() As String
    SaveAttachmentFolder = mSaveAttachmentFolder
End Property

Public Property Let SaveAttachmentFolder(ByVal Data As String)
    mSaveAttachmentFolder = Data
End Property

Public Property Get CurrentMessageCount() As Integer
    CurrentMessageCount = mCurrentMessageCount
End Property

Public Property Get MessageCount() As Integer
    MessageCount = mMessageCount
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


Public Function Login() As Boolean
   
   On Error GoTo Login_error
   
   Login = False
   Select Case mMailboxType
      Case Is = mw_OUTLOOK, mw_OUTLOOK_97 'mw_OUTLOOK_97 Added By N.Angelakis On 02 Feb 2010 'DEV-1797 Sending Email with Outlook client 97
         If LoginOutlook Then
            If moParent.Workflow.IsManualMode Then moParent.RaiseWorkflowStatus "Logged into Outlook."
            Login = True
            mIsLoggedIn = True
         End If
      Case Is = mw_EXCHANGE
         If LoginExchange Then
            If moParent.Workflow.IsManualMode Then moParent.RaiseWorkflowStatus "Logged into Exchange Server."
            Login = True
            mIsLoggedIn = True
         End If
      Case Is = mw_SMTP
         If LoginPOP3 Then
            If moParent.Workflow.IsManualMode Then moParent.RaiseWorkflowStatus "Logged into POP3 Server."
            Login = True
            mIsLoggedIn = True
         End If
      Case Is = mw_OUTLOOK_2002
         If LoginOutlook Then
            If moParent.Workflow.IsManualMode Then moParent.RaiseWorkflowStatus "Logged into Outlook (2002)."
            Login = True
            mIsLoggedIn = True
         End If
      Case Is = mw_AMOS_MAIL
         'If moParent.Workflow.IsManualMode Then goSession.RaiseWorkflowStatus "Logging in mwSession.mwMailbox, AmosMail Receive not supported. Setting to True and exiting..."
         Login = True
      Case Is = mw_MAPI
         If moParent.Workflow.IsManualMode Then goSession.RaiseWorkflowStatus "Logging in mwSession.mwMailbox, MAPI Receive not supported. Setting to True and exiting..."
         Login = True
         
      Case Is = mw_NOTES
         'If moParent.Workflow.IsManualMode Then goSession.RaiseWorkflowStatus "Logging in mwSession.mwMailbox, NOTES Receive not supported. Setting to True and exiting..."
         Login = True
      Case Is = mw_GROUPWISE
         If LoginGW Then
            If moParent.Workflow.IsManualMode Then moParent.RaiseWorkflowStatus "Logged into Groupwise Server."
            Login = True
            mIsLoggedIn = True
         End If
      
      
   End Select
   Exit Function
Login_error:
   moParent.RaiseError "General Error in mwMailbox.Login. ", Err.Number, Err.Description
   Login = False
End Function

Private Function LoginOutlook() As Boolean
   mWasOutlookObject = True
   On Error GoTo No_Outlook_Object_Error
   'Late Binding...
   Set moOutlook = GetObject(, "Outlook.Application")
   moParent.RaiseWorkflowStatus "Accessed Existing Outlook Session."
   
No_GetObject_Resume:
   On Error GoTo GeneralOutlook_Error
   '
   'Set a pointer to the Inbox object...
   '
   'MsgBox "about to point to mapi folder"
   'crashes ShipNet Fleet Outlook 97... 4/2001 Oslo Test Machine
   Set moOutlookInbox = moOutlook.GetNamespace("MAPI").GetDefaultFolder(mw__OL_FOLDER_INBOX)
   
   LoginOutlook = True
Exit Function

No_Outlook_Object_Error:
   'moParent.RaiseError "Error in mwMailbox.LoginOutlook, Outlook must be running before login... "
   'LoginOutlook = False
   On Error GoTo GeneralOutlook_Error
   Set moOutlook = CreateObject("Outlook.Application")
   moParent.RaiseWorkflowStatus "Opened New Outlook Session."
   mWasOutlookObject = False
   Resume No_GetObject_Resume
   Exit Function
GeneralOutlook_Error:
   moParent.RaiseError "General Error in mwMailbox.LoginOutlook. ", Err.Number, Err.Description
   LoginOutlook = False
End Function

Private Function LoginExchange() As Boolean
      '
   ' Open the Email Session
   '
   On Error GoTo LoginExchange_error
   'Set moSession = New MAPI.Session
   Set moSession = CreateObject("MAPI.Session")
   moSession.Logon profileInfo:=mMailServer & vbLf & mLoginID, ShowDialog:=False
   '
   ' Set a pointer to the Inbox folder...
   '
   Set moExchangeInbox = moSession.GetDefaultFolder(mw__CDO_DEFAULT_FOLDER_INBOX)
   LoginExchange = True
   Exit Function
LoginExchange_error:
   LoginExchange = False
   moParent.RaiseError "General Error in mwMailbox.LoginExchange. ", Err.Number, Err.Description
End Function


Private Function LoginPOP3() As Boolean
   Dim sServer As String
   Dim sEmailPortNumber As String
   Dim nEmailPortNumber As Integer
   Dim nEmailSecurityProtocol As Integer
   On Error GoTo FunctionError
#If NO_DART Then
   goSession.RaiseError "Error in mwSession.mwMailbox.LoginPop3, Dart Mail Disabled in this compiled version of mwSession."
   LoginPOP3 = False
#Else
   sServer = goSession.User.GetExtendedProperty("Pop3MailServer")
   
   sEmailPortNumber = goSession.User.GetExtendedProperty("Pop3EmailPortNumber")
   If sEmailPortNumber <> "" Then
      nEmailPortNumber = Val(sEmailPortNumber)
   End If
   
   nEmailSecurityProtocol = -1
   If goSession.User.GetExtendedProperty("Pop3EmailSecurityProtocol") <> "" Then
      nEmailSecurityProtocol = Val(goSession.User.GetExtendedProperty("Pop3EmailSecurityProtocol"))
   End If
   
   If sServer = "" Then
      sServer = moParent.User.MailServerName
   End If
   Set moPOPMessages = New DartMail.Messages
   Set moPOP3 = New DartMail.Pop
   moPOP3.Timeout = 10000
   'moPOP3.Login mMailServer, mLoginID, mLoginPassword
   If nEmailSecurityProtocol >= 0 Then
      moPOP3.Privacy = nEmailSecurityProtocol
   End If
   If moParent.User.MailPassword <> "-1" Then
      If nEmailPortNumber > 0 Then
         moPOP3.Login sServer, moParent.User.MailUserID, moParent.User.MailPassword, nEmailPortNumber
      Else
         moPOP3.Login sServer, moParent.User.MailUserID, moParent.User.MailPassword
      End If
   Else
      If nEmailPortNumber > 0 Then
         moPOP3.Login sServer, moParent.User.MailUserID, "", nEmailPortNumber
      Else
         moPOP3.Login sServer, moParent.User.MailUserID, ""
      End If
   End If
   LoginPOP3 = True
#End If
   Exit Function
FunctionError:
   LoginPOP3 = False
   moParent.RaiseError "General Error in mwMailbox.LoginPop3. ", Err.Number, Err.Description
End Function

Private Function LoginGW() As Boolean
   On Error GoTo FunctionError
   
   Set moGWApp = CreateObject("NovellGroupWareSession")
   Set moGWAccount = moGWApp.Login(moParent.User.MailUserID, moParent.User.MailPassword)
   LoginGW = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwMailbox.LoginGW(Groupwise). ", Err.Number, Err.Description
   LoginGW = False
End Function


Public Function Logout() As Boolean
   On Error GoTo Logout_error
   '
   ' Doesn't do much right now...
   '
   
   Logout = True
   Exit Function
Logout_error:
   moParent.RaiseError "General Error in mwMailbox.LoginOut ", Err.Number, Err.Description
   Logout = False

End Function


Public Function ProcessInbox() As Boolean
   On Error GoTo ProcessInbox_error
   Select Case mMailboxType
      Case Is = mw_OUTLOOK
         ProcessInbox = ProcessInbox_Outlook()
      Case Is = mw_EXCHANGE
         ProcessInbox = ProcessInbox_Exchange()
      Case Is = mw_SMTP
         ProcessInbox = ProcessInbox_POP3()
      Case Is = mw_OUTLOOK_2002
         ProcessInbox = ProcessInbox_Outlook()
      Case Is = mw_AMOS_MAIL
         ProcessInbox = True
      Case Is = mw_GROUPWISE
         ProcessInbox = ProcessInbox_GroupWise()
      Case Is = mw_MAPI
         ProcessInbox = True
      Case Is = mw_NOTES
         ProcessInbox = ProcessInbox_Notes()
         
      'DEV-1797 Sending Email with Outlook client 97
      'Added By N.Angelakis On 02 Feb 2010
      Case Is = mw_OUTLOOK_97
         ProcessInbox = ProcessInbox_Outlook()
         
      Case Else
         moParent.RaiseError "ProcessInbox Not supported for current MailboxType: " & mMailboxType
         ProcessInbox = False
   End Select
   Exit Function
ProcessInbox_error:
   moParent.RaiseError "General Error in mwMailbox.ProcessInbox. ", Err.Number, Err.Description
   ProcessInbox = False
End Function

'
' 9/16/2001 New Procedures for Outlook:
' 1. For each attachment
'     1.2 Attempt to open as form
'        1.2.1 If false, next attachment
'        1.2.2 If True, move to "RouteFolder" or "Inbound" folder if that is blank
'  2.

Private Function ProcessInbox_Outlook() As Boolean
   'Early Binding for Testing
'   Dim loMail As MailItem
'   Dim loMailForward As MailItem
'   Dim loAttachments As Outlook.Attachments
'   Dim loAttachment As Outlook.Attachment
'   Dim loNamespace As NameSpace
'   Dim loFolder As MAPIFolder
   
   Dim loMail As Object
   Dim loMailForward As Object
   Dim loAttachments As Object
   Dim loAttachment As Object
   Dim loNamespace As Object
   Dim loFolder As Object
   
   Dim IsInbound As Boolean
   
   Dim strWorkingFolder As String
   Dim strTargetFile As String
   Dim strForwardEmailAddress As String
   Dim IsProcessedOne As Boolean
   Dim IsAttachedOne As Boolean
   Dim fso As FileSystemObject
   Dim sFileName As String
   Dim strRouteToFolder As String
   Dim loform As mwForm
   Dim loWfItem As mwWorkflowItem
   Dim IsMwsFolder As Boolean
   Dim loWork As Object
   Dim loKeys As Object
   Dim loMailItems As Object
   Dim loDeleteMail As Object
   
   On Error GoTo FunctionError
   If moOutlookInbox Is Nothing Then
      moParent.RaiseError "Error in mwMailbox.ProcessInbox_Outlook, apparently you must first login to Outlook."
      ProcessInbox_Outlook = False
      Exit Function
   End If
   Set fso = New FileSystemObject
   strWorkingFolder = moParent.Workflow.WorkingSessionFolder
   Set loKeys = goSession.GetEventSecurityKeys(MW_EVENT_WORKFLOW_AGENT)
   IsProcessedOne = True
   Do While IsProcessedOne
      'If item in collection is deleted (IsProcessedOne), it whacks out For Each collection
      ' so, restart for each to make sure we get them all in one iteration.
      IsProcessedOne = False
      
      goSession.RaiseWorkflowStatus "Examining Outlook Inbox with " & moOutlookInbox.Items.Count & " items."
      Set loMailItems = moOutlookInbox.Items
    
      Set loMail = loMailItems.GetFirst
      Do While Not loMail Is Nothing
      'For Each loMail In moOutlookInbox.Items
         IsInbound = False
         '
         ' parse first attachment only
         '
         IsAttachedOne = False
         For Each loAttachment In loMail.Attachments
            On Error Resume Next
            sFileName = strWorkingFolder & "\" & loAttachment.FileName
            ' Trap invalid attachment (seems to be only way)...
            If Not Err Then
               Set loWfItem = moParent.Workflow.OpenNewWorkflowItem
               If fso.FileExists(sFileName) Then
                  moParent.LogIt mwl_Transport, mwl_Warning, "Warning in mwMailbox.ProcessInbox_Outlook, Attachment exists and will be renamed: " _
                    & sFileName
                  moParent.Workflow.RenameFileNowExtension sFileName
               End If
               loAttachment.SaveAsFile sFileName
               ' Examine file...
               loWfItem.MaybeProcessWorkflowItem sFileName
               If loWfItem.IsInbound And loWfItem.IsRouteToEmailAddressExclusive = False Then
                  '
                  ' We have an incoming form...
                  '
                  loWfItem.CopyOrMoveItemToFolder
                  IsProcessedOne = True
                  IsAttachedOne = True
                  'Exit For
               Else
                  '
                  ' Not a workflow document, or item is to be forwarded as email only...
                  '
                  fso.DeleteFile sFileName
               End If
               If loWfItem.IsRouteToEmailAddressExclusive Then
                  moParent.RaiseWorkflowStatus "Form Received for Exclusive Route to Email Address. Template: " & loWfItem.TemplateID & _
                    ", Form ID: " & loWfItem.FormID
               End If
               '
               ' Final steps with email...
               '
               'Display the mail to get rid of icon in systray...dumb but works
            End If
         Next loAttachment
         If Not loWfItem Is Nothing Then
    '         If loWfItem.IsInbound Or IsProcessedOne Then
           If loWfItem.IsInbound Or IsAttachedOne Then
               IsInbound = True
               loMail.Display
               ' olDiscard = 1
               loMail.Close 1
               loMail.UnRead = False
              '
               ' Possibly Forward Email
               '
               If Trim(loWfItem.RouteToEmailAddress) <> "" Then
                  Set loMailForward = loMail.Forward
                  loMailForward.Recipients.Add loWfItem.RouteToEmailAddress
                  loMailForward.Send
                  moParent.RaiseWorkflowStatus "Inbound form forwarded to: " & loWfItem.RouteToEmailAddress
                  Set loMailForward = Nothing
               End If
            End If
         End If
         moParent.KillObject loWfItem
         If IsInbound Then
            'loMail.Delete
            Set loDeleteMail = loMail
         Else
            '
            ' Marine Assurance - Emergency Vetting Request Trap - NEEDS 908 Switch to activate.
            '
            If loKeys.GetKeyValue("LookForVettingEmail") = "1" Then
               If InStr(1, UCase(BlankNull(loMail.subject)), "VETTING") <> 0 And loMail.Attachments.Count < 1 Then
                  Set loWork = CreateObject("mwSafety4.smMarineAssWork")
                  loWork.InitSession goSession
                  If loWork.CreateWqItemFromEmail(BlankNull(loMail.Body), _
                                                  BlankNull(loMail.SenderEmailAddress), _
                                                  loKeys.GetKeyValue("VettingNotifyEmailAddress")) Then
                     goSession.RaiseWorkflowStatus "Created Marine Assurance Vetting Request from: " & BlankNull(loMail.SenderEmailAddress)
                  Else
                     goSession.RaiseWorkflowStatus "FAILED Processing Marine Assurance Vetting Request from: " & BlankNull(loMail.SenderEmailAddress)
                  End If
                  KillObject loWork
                  Set loDeleteMail = loMail
                  'loMail.Delete
               End If
            End If
         End If
          
InvalidAttachmentContinue:
      On Error GoTo FunctionError
      Set loMail = loMailItems.GetNext
      
      If Not loDeleteMail Is Nothing Then
         loDeleteMail.Delete
         Set loDeleteMail = Nothing
      End If
      'Next loMail
          
      Loop
   Loop
   ProcessInbox_Outlook = True
   moParent.KillObject fso
   moParent.KillObject loMail
   moParent.KillObject loAttachment
   Set loKeys = Nothing
   
   If Not mWasOutlookObject Then
      moOutlook.Quit
      moParent.KillObject moOutlook
   End If
   
   
   Exit Function
FunctionError:
   If Err.Number = -2147467259 Then
      On Error GoTo FunctionError
      GoTo InvalidAttachmentContinue
   Else
      moParent.RaiseError "General Error in mwMailbox.ProcessInbox_Outlook. ", Err.Number, Err.Description
   End If
   moParent.KillObject fso
   moParent.KillObject loMail
   moParent.KillObject loAttachment
   ProcessInbox_Outlook = False
End Function



Private Function ProcessInbox_Exchange() As Boolean
   Dim oMailItem As Object
   Dim oAttachments As Object
   Dim oAttachment As Object
   ' Early Binding - Requires CDO 1.2 Library Reference
   'Dim oMailItem As MAPI.Message
   'Dim oAttachments As MAPI.Attachments
   'Dim oAttachment As MAPI.Attachment
   Dim iAtt As Integer
   Dim iFor As Integer
   Dim fso As FileSystemObject
   Dim sFileName As String
   On Error GoTo ProcessInbox_Exchange_error
   Set fso = New FileSystemObject
   Do While moExchangeInbox.Messages.Count > 0
      Set oMailItem = moExchangeInbox.Messages.GetFirst
      iAtt = oMailItem.Attachments.Count
      For iFor = 1 To (iAtt)
         Set oAttachment = oMailItem.Attachments.Item(iFor)
         sFileName = mSaveAttachmentFolder & "\" & oAttachment.Name
         If Not fso.FileExists(sFileName) Then
            oAttachment.WriteToFile sFileName
            'moParent.LogIt mwf_DocumentReceived, mwl_Information, sFilename, _
            '  mwt_ElectronicMail, "Mailbox Agent Attachment Save"
            moParent.LogIt mwl_Transport, mwl_Information, "Mailbox Agent Attachment Saved: " & sFileName
            mCurrentMessageCount = mCurrentMessageCount + 1
         Else
            moParent.RaiseWorkflowStatus "Error in mwMailbox.ProcessInbox_Exchange, Attachment already exists: " & sFileName
         End If
      Next iFor
      oMailItem.Delete
   Loop
   ProcessInbox_Exchange = True
   If Not fso Is Nothing Then Set fso = Nothing
   If Not oMailItem Is Nothing Then Set oMailItem = Nothing
   If Not oAttachment Is Nothing Then Set oAttachment = Nothing
   Exit Function
ProcessInbox_Exchange_error:
   moParent.RaiseError "General Error in mwMailbox.ProcessInbox_Exchange", Err.Number, Err.Description
   moParent.KillObject fso
   moParent.KillObject oMailItem
   moParent.KillObject oAttachment
   ProcessInbox_Exchange = False

End Function

Private Function ProcessInbox_POP3() As Boolean
   #If NO_DART = 1 Then
      Dim oMail As Object
   #Else
      Dim oMail As DartMail.Message
   #End If
   Dim iAtt As Integer
   Dim iFor As Integer
   Dim iFor2 As Integer
   Dim strContent() As String
   On Error GoTo FunctionError
   If Trim(mSaveAttachmentFolder) = "" Then
      mSaveAttachmentFolder = goSession.Workflow.PollingFolder
   End If
   moPOP3.AttachmentDirectory = mSaveAttachmentFolder
   If moPOP3.Count > 0 Then
      moParent.RaiseWorkflowStatus "Receiving File from SMTP Mailbox."
      moPOP3.Get
      mCurrentMessageCount = mCurrentMessageCount + moPOP3.Messages.Count
      moPOP3.Delete
   End If
   moPOP3.Logout
   ProcessInbox_POP3 = True
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwMailbox.ProcessInbox_POP3. ", Err.Number, Err.Description
   ProcessInbox_POP3 = False
End Function

Private Function ProcessInbox_GroupWise() As Boolean
   ' Receive Mail
   Dim loWfItem As mwWorkflowItem
   Dim loMsgs As Object
   'Dim loMsg As Message
   Dim loMsg As Object
   Dim loMsgForward As Object
   'Dim loFolders As Folders
   Dim loFolders As Object
   Dim fso As FileSystemObject
   Dim sWorkingFolder As String
   Dim sFileName As String
   Dim nAttCount As Long
   On Error GoTo FunctionError
   '
   Set fso = New FileSystemObject
   sWorkingFolder = moParent.Workflow.WorkingSessionFolder
   
   ' Get Folders
   Set loFolders = moGWAccount.AllFolders
   Set loMsgs = loFolders(GW_INBOX_FOLDER).Messages
   For Each loMsg In loMsgs
      If loMsg.BoxType = GW_egwOutgoing Then
         loMsg.Delete
      ElseIf loMsg.BoxType = GW_egwIncoming Then
         Set loWfItem = moParent.Workflow.OpenNewWorkflowItem
         ' Sometimes Groupwise throughs the HTML mody in as the first attachment...
         ' so, always evelaute the last attachment.
         nAttCount = loMsg.Attachments.Count
         If nAttCount > 0 Then
            'MsgBox "before loMsg.Attachments(GW_ATTACHMENT_OFFSET).FileName"
            sFileName = sWorkingFolder & "\" & loMsg.Attachments(nAttCount).FileName
   
            If fso.FileExists(sFileName) Then
               moParent.LogIt mwl_Transport, mwl_Warning, "Warning in mwMailbox.ProcessInbox_GroupWise, Attachment exists and will be renamed: " _
                 & sFileName
               moParent.Workflow.RenameFileNowExtension sFileName
            End If
            'loWfItem.IsInbound = False
            'MsgBox "before .attachments.save"
            loMsg.Attachments(nAttCount).Save sFileName
            'MsgBox "after .attachments.save"
            loWfItem.MaybeProcessWorkflowItem sFileName
            If loWfItem.IsInbound And loWfItem.IsRouteToEmailAddressExclusive = False Then
               '
               ' We have an incoming form...
               '
               loWfItem.CopyOrMoveItemToFolder
            Else
               '
               ' Not a workflow document, or item is to be forwarded as email only...
               '
               fso.DeleteFile sFileName
            End If
            If loWfItem.IsRouteToEmailAddressExclusive Then
               moParent.RaiseWorkflowStatus "Form Received for Exclusive Route to Email Address. Template: " & _
                loWfItem.TemplateID & ", Form ID: " & loWfItem.FormID
            End If
            If loWfItem.IsInbound Then
               '
               ' Possibly Forward Email
               '
               If Trim(loWfItem.RouteToEmailAddress) <> "" Then
                  MsgBox "before loMsg.Forward"
                  Set loMsgForward = loMsg.Forward
                  loMsgForward.Recipients.Add loWfItem.RouteToEmailAddress
                  loMsgForward.Send
                  moParent.RaiseWorkflowStatus "Inbound form forwarded to: " & loWfItem.RouteToEmailAddress
                  Set loMsgForward = Nothing
               End If
               moParent.KillObject loWfItem
               '
               ' Delete the processed message
               '
               loMsg.Delete
            End If
         End If
      End If
   Next loMsg
   ProcessInbox_GroupWise = True
   Exit Function
FunctionError:
   moParent.RaiseError "General error in mwMailbox.ProcessInbox_GroupWise. ", Err.Number, Err.Description
   ProcessInbox_GroupWise = False
End Function


Private Function ProcessInbox_Notes() As Boolean
   'Dim loSession As NotesSession
   'Dim loDocItem As Object
   'Dim loView As NOTESVIEW
   'Dim loDB As NOTESDATABASE
   'Dim loDoc As NotesDocument
   'Dim o As Variant
   'Dim loItemA   As NotesItem
   Dim loSession As Object
   Dim loDocItem As Variant
   Dim loView As Variant
   Dim loDB As Variant
   Dim loDoc As Variant
   Dim loNextDoc As Variant
   Dim loAttachment As Variant
   Dim loItemA  As Variant
   Dim nLimit As Long
   Dim sExt As String
   Dim fso As FileSystemObject
   Dim IsRemoveFromFolder As Boolean
   Dim IsInDebugMode As Boolean
   
   
   On Error GoTo FunctionError
   
   If goSession.ThisSite.GetExtendedProperty("DataBaseType") = "D" Then
      moParent.RaiseWorkflowStatus "Notes InBox Debug Mode ON"
      IsInDebugMode = True
   Else
      IsInDebugMode = False
   End If
   
   
   Set loSession = CreateObject("Notes.NotesSession")
   'loSession.Initialize (goSession.User.MailPassword)
   Set loDB = loSession.GETDATABASE(goSession.User.MailServerName, goSession.User.MailUserID)
   Set loView = loDB.GetView(LOTUS_INBOX)
   Set loDoc = loView.GetFirstDocument
   goSession.RaiseWorkflowStatus "Processing Lotus Notes InBox Mailbox..."
   Set fso = New FileSystemObject
   nLimit = 0
   
   On Error Resume Next
   
   Do While Not loDoc Is Nothing
      
      ' First, save the current next document so we can increment at the end of this loop.
      ' We get the NextDocument here in case we mess around with the current document and lose the pointer.
      ' By grabbing the NextDocument up here we are sure that it is valid.
      
      Set loNextDoc = loView.GetNextDocument(loDoc)
      
      nLimit = nLimit + 1
      If nLimit > 10000 Then
         moParent.RaiseWorkflowStatus "Error in mwMailbox.ProcessInbox_Notes, There are too many messages for Lotus Notes Inbox, Maximum is 10000. Check Lotus Notes Configuration "
         Set loDoc = Nothing
         Set loDocItem = Nothing
         Set loView = Nothing
         Set loDB = Nothing
         Set loSession = Nothing
         ProcessInbox_Notes = False
         Exit Function
      End If
      If loDoc.HasEmbedded Then
         If IsInDebugMode Then
            moParent.RaiseWorkflowStatus "M01 Item " & nLimit & " HasEmbedded = True"
         End If
         
         Set loDocItem = loDoc.GetFirstItem("Body")
         ' Ship sends in SMTP objects, shore sends Rich Text objects...
         If loDocItem.Type = LOTUS_SMTP Then
            If IsInDebugMode Then
               moParent.RaiseWorkflowStatus "M02 Type = LOTUS_SMTP, Getting Attachment"
            End If
            Set loDocItem = loDoc.GetFirstItem("Attachment")
         End If
         
         If loDocItem.Type = LOTUS_RICHTEXT Then
            If IsInDebugMode Then
               moParent.RaiseWorkflowStatus "M03 Type = LOTUS_RICHTEXT"
            End If
         ElseIf loDocItem.Type = LOTUS_SMTP Then
            If IsInDebugMode Then
               moParent.RaiseWorkflowStatus "M04 Type = LOTUS_SMTP"
            End If
         Else
            If IsInDebugMode Then
               moParent.RaiseWorkflowStatus "M05 Type = " & loDocItem.Type
            End If
         End If
         
         If loDocItem.Type = LOTUS_RICHTEXT Or loDocItem.Type = LOTUS_SMTP Then
         
            If Err.Number = 0 Then
               
               IsRemoveFromFolder = False
               
               ' Loop over all attachments
               For Each loAttachment In loDocItem.EmbeddedObjects
                  If Err.Number = 0 Then
                     If loAttachment.Type = LOTUS_EMBED_ATTACHMENT Then
                        'If Not err Then
                        If Err.Number = 0 Then
                           sExt = UCase(fso.GetExtensionName(loAttachment.Name))
                           If InStr(1, "XML,PKG,PKZ,DGM,DGZ,FCB,FCZ,FMZ,DAB,DAZ,EZZ,ZIZ", sExt) And Trim(sExt) <> "" Then
                              loAttachment.ExtractFile goSession.Workflow.PollingFolder & "\" & loAttachment.Name
                              goSession.RaiseWorkflowStatus "Saved File: " & goSession.Workflow.PollingFolder & "\" & loAttachment.Name
                              
                              IsRemoveFromFolder = True
                           Else
                              If IsInDebugMode Then
                                 moParent.RaiseWorkflowStatus "M06 Will not Process attachment, Ext = " & sExt
                              End If
                           End If
                        Else
                           If IsInDebugMode Then
                              moParent.RaiseWorkflowStatus "M07 Will not Process attachment, err = " & Err.Number
                           End If
                        End If
                     Else
                        If IsInDebugMode Then
                           moParent.RaiseWorkflowStatus "M08 Will not process, Attachment.Type = " & loAttachment.Type
                        End If
                     End If
                  Else
                     If IsInDebugMode Then
                        moParent.RaiseWorkflowStatus "M09 Will not Process attachment, err = " & Err.Number
                     End If
                  End If
                  'Exit For
               Next loAttachment
               
               If IsRemoveFromFolder = True Then
                  If IsInDebugMode Then
                     moParent.RaiseWorkflowStatus "M10 Removing Item from Folder"
                  End If
                  loDoc.RemoveFromFolder LOTUS_INBOX
                  If Err Then
                     goSession.RaiseWorkflowStatus "Error in mwMailbox.ProcessInbox_Notes, after RemoveFromFolder.", Err.Number, Err.Description
                  End If
               End If
            Else
               If IsInDebugMode Then
                  moParent.RaiseWorkflowStatus "M11 Will not Process attachment, err = " & Err.Number
               End If
            End If
         Else
            If IsInDebugMode Then
               moParent.RaiseWorkflowStatus "M12 Will not process, Type = " & loDocItem.Type
            End If
         End If
      Else
         If IsInDebugMode Then
            moParent.RaiseWorkflowStatus "M13 Item " & nLimit & " HasEmbedded = False"
         End If
      End If
      
      ' Increment to the next document
      
      Set loDoc = loNextDoc
      loView.Refresh
      
   Loop
   If IsInDebugMode Then
      moParent.RaiseWorkflowStatus "M14 Finished processing inBox"
   End If
   Set loDoc = Nothing
   Set loDocItem = Nothing
   Set loView = Nothing
   Set loDB = Nothing
   Set loSession = Nothing
   ProcessInbox_Notes = True
   Exit Function
FunctionError:
'Resume Next
   moParent.RaiseError "General error in mwMailbox.ProcessInbox_Notes. ", Err.Number, Err.Description
End Function

'
' First Version...uses Notes View
'
'Private Function ProcessInbox_Notes() As Boolean
'   'Dim loSession As NotesSession
'   'Dim loDocItem As Object
'   'Dim loView As NOTESVIEW
'   'Dim loDB As NOTESDATABASE
'   'Dim loDoc As NotesDocument
'   'Dim o As Variant
'   'Dim loItemA   As NotesItem
'   Dim loSession As Object
'   Dim loDocItem As Variant
'   Dim loView As Variant
'   Dim loDB As Variant
'   Dim loDoc As Variant
'   Dim o As Variant
'   Dim loItemA  As Variant
'   Dim nLimit As Long
'   On Error GoTo FunctionError
'   Set loSession = CreateObject("Notes.NotesSession")
'   'loSession.Initialize (goSession.User.MailPassword)
'   Set loDB = loSession.GETDATABASE(goSession.User.MailServerName, goSession.User.MailUserID)
'   Set loView = loDB.GetView("SHIPNET")
'   Set loDoc = loView.GetFirstDocument
'   goSession.RaiseWorkflowStatus "Processing Lotus Notes SHIPNET View Mailbox..."
'   nLimit = 0
'   Do While Not loDoc Is Nothing
'      nLimit = nLimit + 1
'      If nLimit > 100 Then
'         moParent.RaiseWorkflowStatus "Error in mwMailbox.ProcessInbox_Notes, 100 Iterations limit. Check Lotus Notes Configuration "
'         Set loDoc = Nothing
'         Set loDocItem = Nothing
'         Set loView = Nothing
'         Set loDB = Nothing
'         Set loSession = Nothing
'         ProcessInbox_Notes = False
'         Exit Function
'      End If
'      If loDoc.HasEmbedded Then
'         Set loDocItem = loDoc.GetFirstItem("Body")
'         If loDocItem.Type = LOTUS_RICHTEXT Then
'            For Each o In loDocItem.EmbeddedObjects
'               If o.Type = LOTUS_EMBED_ATTACHMENT Then
'                  o.ExtractFile goSession.Workflow.PollingFolder & "\" & o.Name
'                  goSession.RaiseWorkflowStatus "Saved File: " & goSession.Workflow.PollingFolder & "\" & o.Name
'                  Set loItemA = loDoc.AppendItemValue("SHIPNET_Processed", "Y")
'                  Call loDoc.Save(True, False)
'               End If
'            Next o
'         Else
'            Set loItemA = loDoc.AppendItemValue("SHIPNET_Processed", "Y")
'            Call loDoc.Save(True, False)
'            goSession.RaiseWorkflowStatus "Flagged Non RICHTEXT Item as processed..."
'         End If
'      End If
'    '   loDoc.Remove True
'      loView.Refresh
'      Set loDoc = loView.GetFirstDocument
'   Loop
'   Set loDoc = Nothing
'   Set loDocItem = Nothing
'   Set loView = Nothing
'   Set loDB = Nothing
'   Set loSession = Nothing
'   ProcessInbox_Notes = True
'   Exit Function
'FunctionError:
'   moParent.RaiseError "General error in mwMailbox.ProcessInbox_Notes. ", Err.Number, Err.Description
'End Function
'


