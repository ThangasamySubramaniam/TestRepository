VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' TransportMessage - Class responsible for staging and sending a message
'                    containing a document or a Transport container file
' 6/28/2001 ms - Added AMOS Mail/Link support
' 12/2001 ms - Added Lotus Notes support
' 2/2002 ms - Add Silent Error handling
Option Explicit
Dim moKeys As ConfigKeys
Dim mLaunchOutlook As Boolean
Dim mIsFunnelToMessageQueue As Boolean



' Declaring the AMOS Mail function
Private Declare Function SendMail Lib "mapitif.dll" _
(ByVal lpMessage As String) As Long

Public Enum mw_Priority
   mw_LOW
   mw_MEDIUM
   mw_HIGH
End Enum

Public Enum mw_MessageFormat
   mw_PLAIN_TEXT = 1
   mw_RICH_TEXT = 2
   mw_HTML = 3
End Enum

Public Enum mw_AttachmentOptions
   mw_LEAVE_SOURCE = 1
   mw_DELETE_SOURCE = 2
End Enum


' AMOS_MAIL= Windows Version
' AMOS_LINK= Dos Version
Public Enum mw_EmailCarrier
   mw_SMTP = 0
   mw_OUTLOOK = 1
   mw_EXCHANGE = 2
   mw_NOTES = 3
   mw_AMOS_MAIL = 4
   mw_AMOS_LINK = 5
   mw_OUTLOOK_2002 = 6
   mw_GROUPWISE = 7
   mw_MAPI = 8
   'DEV-1797 Sending Email with Outlook client 97
   'Added By N.Angelakis On 02 Feb 2010
   mw_OUTLOOK_97 = 9
  
   'DEV-1809 Send By Media Error Notification
   'Added By N.Angelakis On 04 Feb 2010
   mw_SENDBYMEDIA_LINK = 10
End Enum

'
' Lose it
'
Public Enum mw_EmailProduct
   mw_GLOBE = 1
   mw_TELESIS = 2
End Enum


Private mSendToAddress As String
Private mSendFromAddress As String
Private mAttachmentFile As String
Private mAttachmentNameOnly As String
Private mDocID As String
Private mAttachmentOptions As mw_AttachmentOptions
Private mZipAttachment As Boolean
Private mCarrier As mw_EmailCarrier
Private mPriority As mw_Priority
Private mSubject As String
Private mBody As String
'Private mOutgoingTransportFolder As String
Private mSMTPPort As String
Private mServer As String
Private moParent As Session
Private mUserResolveAddress As Boolean
Private mUserResolveLogon As Boolean
' Holds error string returned by email transport...
Private mEmailErrorMessage As String
#If NO_DART = 1 Then
   Private moSMTP As Object
   Private moMail As Object
#Else
   Private WithEvents moSMTP As DartMail.Smtp
Attribute moSMTP.VB_VarHelpID = -1
   Private moMail As DartMail.Message
#End If
'Private mAddresses As Collection
'
' Amos Mail/Store & Forward Parameters
'
Dim mSFMessageHeaderFolder As String
Dim mSFMessageHeaderFilename As String
Dim mSFAttachmentFolder As String
Dim mSFShortSiteID As String

'
' Amos Link Constants - Message Header File
'
Private Const AL_HEADER_LINE_1 = "AL31"
Private Const AL_HEADER_LINE_2 = "20MCB-options:NYYNAYAYA"
Private Const AL_HEADER_LINE_3 = "Attachment:"
Private Const AL_HEADER_LINE_4 = "Attachment-name:"
Private Const AL_HEADER_LINE_5 = "To:-@smtp{"
Private Const AL_HEADER_LINE_6 = "From:"
Private Const AL_HEADER_LINE_7 = "Subject:"




Event Status(Msg As String, percent_complete As Integer, Err As Boolean)


Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
   '
   ' Set default message parameters
   '
   mSendToAddress = moParent.Site.WorkflowSendToAddress
   mCarrier = moParent.User.DefaultEmailCarrier
   'mOutgoingTransportFolder = moParent.GetPlace(moParent.Site.TransportPlaceID, mw_DriveLetterPath)
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Let SendToAddress(ByVal Data As String)
    mSendToAddress = Data
End Property

Public Property Get SendToAddress() As String
    SendToAddress = mSendToAddress
End Property

Public Property Let SendFromAddress(ByVal Data As String)
    mSendFromAddress = Data
End Property

Public Property Get SendFromAddress() As String
    SendFromAddress = mSendFromAddress
End Property

Public Property Let AttachmentFile(ByVal Data As String)
   Dim fso As FileSystemObject
   Set fso = New FileSystemObject
    mAttachmentFile = Data
    mAttachmentNameOnly = fso.GetFileName(mAttachmentFile)
   Set fso = Nothing
End Property

Public Property Get AttachmentFile() As String
    AttachmentFile = mAttachmentFile
End Property

Public Property Let AttachmentOptions(ByVal Data As mw_AttachmentOptions)
    mAttachmentOptions = Data
End Property

Public Property Get AttachmentOptions() As mw_AttachmentOptions
    AttachmentOptions = mAttachmentOptions
End Property

Public Property Let Carrier(ByVal Data As mw_EmailCarrier)
    mCarrier = Data
End Property

Public Property Get Carrier() As mw_EmailCarrier
    Carrier = mCarrier
End Property

Public Property Let Priority(ByVal Data As mw_Priority)
    mPriority = Data
End Property

Public Property Get Priority() As mw_Priority
    Priority = mPriority
End Property

Public Property Get Body() As String
    Body = mBody
End Property

Public Property Let Body(ByVal Data As String)
    mBody = Data
End Property

Public Property Let subject(ByVal Data As String)
    mSubject = Data
End Property

Public Property Get subject() As String
    subject = mSubject
End Property

Public Property Let Server(ByVal Data As String)
    mServer = Data
End Property

Public Property Get Server() As String
    Server = mServer
End Property

Public Property Let SMTPPort(ByVal Data As String)
    mSMTPPort = Data
End Property

Public Property Get SMTPPort() As String
    SMTPPort = mSMTPPort
End Property

Public Property Let ZipAttachment(ByVal Data As Boolean)
    mZipAttachment = Data
End Property

Public Property Get ZipAttachment() As Boolean
    ZipAttachment = mZipAttachment
End Property

Public Property Let docID(ByVal Data As String)
    mDocID = Data
End Property

Public Property Get docID() As String
    docID = mDocID
End Property

Public Property Let UserResolveAddress(ByVal Data As Boolean)
    mUserResolveAddress = Data
End Property

Public Property Get UserResolveAddress() As Boolean
    UserResolveAddress = mUserResolveAddress
End Property

Public Property Let UserResolveLogon(ByVal Data As Boolean)
    mUserResolveLogon = Data
End Property

Public Property Get UserResolveLogon() As Boolean
    UserResolveLogon = mUserResolveLogon
End Property

Public Property Let SFMessageHeaderFolder(ByVal Data As String)
    mSFMessageHeaderFolder = Data
End Property

Public Property Get SFMessageHeaderFolder() As String
    SFMessageHeaderFolder = mSFMessageHeaderFolder
End Property

Public Property Let SFMessageHeaderFilename(ByVal Data As String)
    mSFMessageHeaderFilename = Data
End Property

Public Property Get SFMessageHeaderFilename() As String
    SFMessageHeaderFilename = mSFMessageHeaderFilename
End Property

Public Property Get SFAttachmentFolder() As String
    SFAttachmentFolder = mSFAttachmentFolder
End Property

Public Property Let SFAttachmentFolder(ByVal Data As String)
    mSFAttachmentFolder = Data
End Property

Public Property Let SFShortSiteID(ByVal Data As String)
    mSFShortSiteID = Data
End Property

Public Property Get SFShortSiteID() As String
    SFShortSiteID = mSFShortSiteID
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

Private Sub Class_Initialize()
   Dim nUserKey As Long
   On Error GoTo SubError
   Set moKeys = New ConfigKeys
   Set moKeys = goSession.GetEventSecurityKeys(MW_EVENT_WORKFLOW_AGENT)
   mLaunchOutlook = moKeys.GetBoolKeyValue("AutoLaunchOutlookOnSendTran")
   ' 8/2010 ms BGLNG-15
   nUserKey = ZeroNull(goSession.Site.GetExtendedProperty("EmailSendThisUserKey"))
   If nUserKey <> 0 And nUserKey <> goSession.User.UserKey Then
      mIsFunnelToMessageQueue = True
   Else
      mIsFunnelToMessageQueue = False
   End If
   
   Exit Sub
SubError:
   moParent.RaiseError "General Error in mwMessage.Class_Initialize.", Err.Number, Err.Description
End Sub


' General Send Function for all messatge types.
'
Public Function Send() As Boolean
   Dim oLog As mwLogItem
   Dim oZip As Zipper
   Dim strZipFile As String
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   
   '
   ' Send By Email
   '
   If mZipAttachment And Trim(mAttachmentFile) <> "" Then
      Set fso = New FileSystemObject
      RaiseEvent Status("Zipping Document", 10, False)
      'zip the file...
      Set oZip = New Zipper
      ' 8/2010 ms BGLNG-15 File will be sent later... so stick it in SendByMedia
      If mIsFunnelToMessageQueue Then
         strZipFile = goSession.Workflow.SendByMediaFolder & "\" & fso.GetFileName(mAttachmentFile) & ".zip"
      Else
         'strZipFile = goSession.GetAppPath() & "\" & fso.GetFileName(mAttachmentFile) & ".zip"
         strZipFile = goSession.Workflow.WorkingSessionFolder & "\" & fso.GetFileName(mAttachmentFile) & ".zip"
      End If
      oZip.DeleteAfterZip = False
      oZip.OpenZip strZipFile
      oZip.AddToZip mAttachmentFile
      If oZip.ErrorCode <> 0 Then
         'Zip Error
         moParent.RaiseError "General Error in mwMessage.Send. Zip Error.", Err.Number, Err.Description
         Send = False
         oZip.TerminateClass
         moParent.KillObject oZip
         Exit Function
      End If
      mAttachmentFile = strZipFile
      mAttachmentNameOnly = fso.GetFileName(strZipFile)
      Set fso = Nothing
      oZip.TerminateClass
      moParent.KillObject oZip
   End If
   
   RaiseEvent Status("Building Message", 25, False)
   
   ' Standard Text Body...
   If Trim(mBody) = "" Then
      mBody = "WorkFlow Form Transmission" & vbCrLf & _
              "Form ID.....: " & mDocID & vbCrLf & _
              "Submitted By: " & moParent.User.UserID & " - " & moParent.Site.SiteName & vbCrLf & _
              "Time........: " & Format(Now(), "mm/dd/yyyy hh:mm") & vbCrLf & vbCrLf _
              & mBody
   End If
   
   '8/2010 ms BGLNG-10 - If site is configured to send EMAIL by single user (Workflow Agent), write to Queue
   '          exception is if Resolve Address is TRUE, then an email Draft will display... Queue is bypassed
   '          This is done in mwVRS6, mwEventLINK, and mwEventFORM modules.
   If mIsFunnelToMessageQueue And Not mUserResolveAddress And Not mCarrier = mw_SENDBYMEDIA_LINK Then
      If FunnelToMessageQueue(mSendToAddress, mAttachmentFile, mSubject, mBody) Then
         goSession.RaiseWorkflowStatus "mwMessage.Send: Message Queued for delivery by Workflow Agent..."
         Send = True
      Else
         goSession.RaiseWorkflowStatus "Error in mwMessage.Send: Message NOT Queued for delivery by Workflow Agent..."
         Send = False
      End If
   Else
        ' Different methods of sending a message...
      Select Case mCarrier
        Case Is = mw_SMTP
            If SendSMTP() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "SMTP Email Sent by:" & moParent.User.UserID
            Else
               Send = False
            End If
         Case Is = mw_OUTLOOK
            If SendOutlook(False) Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Outlook Email Sent by:" & moParent.User.UserID
            Else
               Send = False
            End If
         
         Case Is = mw_OUTLOOK_2002
             If SendOutlook(True) Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Outlook Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
         Case Is = mw_EXCHANGE
         
           If SendExchange() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Exchange Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
         
         
         Case Is = mw_AMOS_MAIL
            ' Windows Version
            If SendAmosMail() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Amos Mail Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
         
         Case Is = mw_AMOS_LINK
            ' DOS Version
            If SendAmosLink() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Amos Link Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
            
         Case Is = mw_NOTES
            
            If SendLotusNotes() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Lotus Notes Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
         Case Is = mw_GROUPWISE
            
            If SendGroupWise() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Groupwise Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
         Case Is = mw_MAPI
            
            If SendMAPI() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "MAPI Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
            
         'DEV-1797 Sending Email with Outlook client 97
         'Added By N.Angelakis On 02 Feb 2010
         Case Is = mw_OUTLOOK_97
            If SendOutlook97() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "Outlook 97 Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
            
         'DEV-1809 Send By Media Error Notification
         'Added By N.Angelakis On 04 Feb 2010
         Case Is = mw_SENDBYMEDIA_LINK
            If SendByMediaLink() Then
               Send = True
               moParent.LogIt mwl_Transport, mwl_Information, "File Link Email Sent by: " & moParent.User.UserID
            Else
               Send = False
            End If
            
         Case Else
            moParent.RaiseError "Error in mwMessage.Send, invalid Email Carrier: " & mCarrier
            Send = False
      End Select
      
      '
      ' Normal Exit
      '
      If mZipAttachment Then
         Set fso = New FileSystemObject
         fso.DeleteFile strZipFile
         Set fso = Nothing
      End If
   End If
   '
   ' Clear message stuff for next use...
   '
   mSendToAddress = ""
   mAttachmentFile = ""
   mAttachmentNameOnly = ""
   mDocID = ""
   mZipAttachment = False
   mSubject = ""
   mBody = ""
   Set oLog = Nothing
   Exit Function
   
FunctionError:
      moParent.RaiseError "General Error in mwMessage.Send .", Err.Number, Err.Description
      RaiseEvent Status("General Send Error !", 0, True)
      Send = False
      If mZipAttachment Then
         Set fso = New FileSystemObject
         If fso.FileExists(strZipFile) Then
            fso.DeleteFile strZipFile
         End If
         Set fso = Nothing
      End If
      moParent.KillObject oZip
End Function

'
' Code copied from downloaded Internet code...
'
Private Function SendLotusNotes() As Boolean
'Set up the objects required for Automation into lotus notes
   Dim Maildb As Object 'The mail database
   Dim username As String 'The current users notes name
   Dim MailDbName As String 'THe current users notes mail database name
   Dim MailDoc As Object 'The mail document itself
   Dim AttachME As Object 'The attachment richtextfile object
   Dim nSession As Object 'The notes nSession
   Dim EmbedObj As Object 'The embedded object (Attachment)
   Dim loRTItem As Object
   On Error GoTo FunctionError
   
   'Start a session to notes
   Set nSession = CreateObject("Notes.NotesSession")
   'Get the sessions username and then calculate the mail file name
   'You may or may not need this as for MailDBname with some systems you
   'can pass an empty string
   username = nSession.username
   MailDbName = Left$(username, 1) & Right$(username, (Len(username) - InStr(1, username, " "))) & ".nsf"
   'Open the mail database in notes
   Set Maildb = nSession.GETDATABASE("", MailDbName)
   If Maildb.IsOpen = True Then
      'Already open for mail
   Else
      Maildb.OPENMAIL
   End If
   'Set up the new mail document
   Set MailDoc = Maildb.CREATEDOCUMENT
   MailDoc.Form = "Memo"
   If mSendToAddress <> "" Then
      MailDoc.sendTo = mSendToAddress
   End If
   MailDoc.subject = mSubject
   
   '
   ' New Code 2/6/2008 for Wilson...
   '
   Set loRTItem = MailDoc.CREATERICHTEXTITEM("Body")
   loRTItem.APPENDTEXT mBody
   If Trim(mAttachmentFile) <> "" Then
      loRTItem.EmbedObject 1454, "", mAttachmentFile
   End If
   '
   ' Old Code
   '
'   MailDoc.Body = mBody
'   If Trim(mAttachmentFile) <> "" Then
'      Set AttachME = MailDoc.CREATERICHTEXTITEM("Attachment")
'      'Set EmbedObj = AttachME.EMBEDOBJECT(1454, "", mAttachmentFile, "Attachment")
'      Set EmbedObj = AttachME.EmbedObject(1454, "", mAttachmentFile)
'      ' Next line in 5.5, crashes 6.5
'      ' MailDoc.CREATERICHTEXTITEM ("Attachment")
'   End If
   '
   ' end code switch
   '
   
   MailDoc.SAVEMESSAGEONSEND = True
   'Send the document
   If mSendToAddress <> "" Then
      MailDoc.Send 0, mSendToAddress
      MailDoc.PostedDate = Now() 'Gets the mail to appear in the sent items folder
   Else
      MailDoc.Save 1, 0
      MsgBox "Email has been created and is in the Draft Folder of Lotus Notes", vbInformation, "Lotus Notes Email Notification"
   End If
   'Clean Up
   moParent.KillObject Maildb
   moParent.KillObject MailDoc
   moParent.KillObject AttachME
   moParent.KillObject nSession
   moParent.KillObject EmbedObj
   SendLotusNotes = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwMessage.SendLotusNotes. ", Err.Number, Err.Description
   SendLotusNotes = False
   Exit Function
End Function





Private Function SendSMTP() As Boolean
   Dim sMessageText As String
   Dim sEmailPortNumber As String
   Dim nEmailPortNumber As Integer
   Dim nEmailSecurityProtocol As Integer
   On Error GoTo FunctionError
#If NO_DART Then
   goSession.RaiseError "Error in mwMessage.SendSMTP, Dart Library disabled in this build."
   SendSMTP = False
#Else
   Set moSMTP = New DartMail.Smtp
   
   RaiseEvent Status("Sending SMTP Email...", 25, False)
   
   sEmailPortNumber = goSession.User.GetExtendedProperty("EmailPortNumber")
   If sEmailPortNumber <> "" Then
      nEmailPortNumber = Val(sEmailPortNumber)
   End If
   
   nEmailSecurityProtocol = -1
   If goSession.User.GetExtendedProperty("EmailSecurityProtocol") <> "" Then
      nEmailSecurityProtocol = Val(goSession.User.GetExtendedProperty("EmailSecurityProtocol"))
   End If

   
   ' Quick and Dirty...
   'moSMTP.QuickSend mSendToAddress, mSendFromAddress, mSubject, mBody, mAttachmentFile, mServer
   If nEmailSecurityProtocol >= 0 Then
      moSMTP.Privacy = nEmailSecurityProtocol
   End If
   If nEmailPortNumber > 0 Then
      moSMTP.Login mServer, nEmailPortNumber, moParent.User.MailUserID, moParent.User.MailPassword
   Else
      moSMTP.Login mServer, , moParent.User.MailUserID, moParent.User.MailPassword
   End If
   With moSMTP.Message
      If Trim(mAttachmentFile) <> "" Then
         .AddAttachment mAttachmentFile
      End If
      .To.Add mSendToAddress
       .subject = mSubject
      If Trim(mSendFromAddress) = "" Then
         mSendFromAddress = goSession.User.GetExtendedProperty("FromEmailAddress")
      End If
      If Trim(mSendFromAddress) <> "" Then
         .From = mSendFromAddress
      End If
      .AddText mBody
   End With
   moSMTP.Send
   moSMTP.Logout
#End If
   Set moSMTP = Nothing
   SendSMTP = True
   Exit Function

FunctionError:
   'raise error
   moParent.RaiseError "General Error in mwMessage.SendSMTP, Server: " & mServer & ", Address: " & mSendToAddress & ". ", Err.Number, Err.Description
   SendSMTP = False
   moParent.KillObject moSMTP
End Function


' Event raised by SMTP object
#If Not NO_DART = 1 Then
   Private Sub moSMTP_Error(ByVal Number As DartMail.MailErrorConstants, ByVal Description As String)
      On Error GoTo SubError
      mEmailErrorMessage = "SMTP Error: " & str(Number) & " - " & Description
      moParent.RaiseWorkflowStatus mEmailErrorMessage
      RaiseEvent Status(mEmailErrorMessage, 0, True)
SubError:
   End Sub

   Private Sub moSMTP_Progress(ByVal Method As DartMail.SmtpMethodConstants, _
      ByVal Status As DartMail.SmtpStatusConstants, ByVal Reply As String, _
      ByVal Count As Long, ByVal Size As Long)
      On Error GoTo SubError
      If Size <> 0 And Count <> 0 Then
         RaiseEvent Status("Sending SMTP Email...", (100 / (Size / Count)), False)
      End If
SubError:
   End Sub
#End If



'
' Requires CDO DLL be registered on computer...
'
Private Function SendExchange() As Boolean
   'Dim oSession As mapi.Session
   'Dim oMessage As Message
   'Dim oRecipient As Recipient
   Dim oSession As Object
   Dim oMessage As Object
   Dim oRecipient As Object
   Dim strRecipient As String
   On Error GoTo FunctionError
   '
   ' Open the Email Session
   '
   'Set oSession = New mapi.Session
   Set oSession = CreateObject("mapi.Session")
   If mUserResolveLogon Then
      oSession.Logon profileInfo:=mServer & vbLf & mSendFromAddress
   Else
      oSession.Logon profileInfo:=mServer & vbLf & mSendFromAddress, ShowDialog:=False
   End If
   '
   
   Set oMessage = oSession.Outbox.Messages.Add
   With oMessage
      .subject = mSubject
      .Text = mBody
      Set oRecipient = .Recipients.Add
      With oRecipient
         'CdoTo = 1
         .Type = 1
         .Name = mSendToAddress
         If mUserResolveAddress Then
            .Resolve
         Else
            .Resolve ShowDialog:=False
         End If
      End With
      ' Optional Attachment
      '.Attachments.Add mAttachmentFile, , , mAttachmentNameOnly
      If Trim(mAttachmentFile) <> "" Then
         .Attachments.Add mAttachmentNameOnly, , , mAttachmentFile
      End If
      If mUserResolveAddress Then
         .Send
      Else
         .Send ShowDialog:=False
      End If
   End With        '
   ' 12/11/2000 ms Destroy objects...
   moParent.KillObject oMessage
   moParent.KillObject oRecipient
   SendExchange = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwMessage.SendExchange.", Err.Number, Err.Description
   SendExchange = False
End Function

'
' July 2002 - Incorporate Outlook Redemption for virus warning bypass...
'
Private Function SendOutlook(Is2002 As Boolean) As Boolean
   Dim WasOutlookObject As Boolean
   Dim i As Integer
   'Dim oOut As Outlook.Application
   Dim oOut As Object
   Dim oNS As Object
   Dim oMail As Object
   Dim oSafeMail As Object
   Dim oMAPIUtils As Object
   Dim oBtn As Object
   'Dim oMail As MailItem
   WasOutlookObject = True
   On Error GoTo No_Outlook_Object_Error
   'Late Binding...
   Set oOut = GetObject(, "Outlook.Application")
   
No_GetObject_Resume:
   If moParent.Workflow.IsSilentErrors Then
      On Error GoTo SendOutlook_ErrorSilent
   Else
      On Error GoTo SendOutlook_Error
   End If
   'Set oOut = New Outlook.Application
   ' olMailItem = 0
   'oOut.ActiveInspector.EditorType = 1
   Set oMail = oOut.CreateItem(0)
   If Is2002 Then
      Set oSafeMail = CreateObject("mwUtilOut.SafeMailItem")
      Set oMAPIUtils = CreateObject("mwUtilOut.MAPIUtils")
      Set oSafeMail.Item = oMail
      With oSafeMail
         .subject = mSubject
         mBody = Replace(mBody, vbCrLf, " <br> ")
         .htmlbody = mBody
         If Trim(mAttachmentFile) <> "" Then
            oMail.Attachments.Add mAttachmentFile
         End If
         
         If mSendToAddress <> "" Then
            .To = mSendToAddress
         End If
         If mUserResolveAddress Then
            .Recipients.Add mSendToAddress
            .Display vbModal
         Else
            .Recipients.Add mSendToAddress
            .Send
         End If
      End With
      oMAPIUtils.DeliverNow
      'Outlook 2003..below raises object 429 error...
      On Error Resume Next
      If Not oOut.ActiveExplorer Is Nothing Then
'        Set oBtn = oOut.ActiveExplorer.CommandBars.FindControl(1, 5488)
         Set oBtn = oOut.ActiveExplorer.CommandBars.FindControl(1, 7095)
         oBtn.Execute
      End If
      If moParent.Workflow.IsSilentErrors Then
         On Error GoTo SendOutlook_ErrorSilent
      Else
         On Error GoTo SendOutlook_Error
      End If
   Else
   
      With oMail
         .subject = mSubject
         mBody = Replace(mBody, vbCrLf, " <br> ")
         ' 29 July 2011 ms - BGLNG uses Outlook97 mode for Citrix compatbility, but needs the .htmlbody field set for CR/LF's to display....
         .htmlbody = mBody
         If Trim(mAttachmentFile) <> "" Then
            .Attachments.Add mAttachmentFile
         End If
         '
         ' Allow user to send email...
         '
         If mSendToAddress <> "" Then
            .To = mSendToAddress
         End If
         If mUserResolveAddress Then
            .Display vbModal
         Else
            .Recipients.Add mSendToAddress
            .Send
         End If
      End With
   End If
   '
   ' Outlook 2002 workaround code from Outlook Redemption (renamed mwUtilOut)
   '
   '
   moParent.KillObject oMail
   moParent.KillObject oSafeMail
   moParent.KillObject oMAPIUtils
   moParent.KillObject oBtn
   If Not WasOutlookObject Then
      On Error Resume Next
      oOut.Quit
      moParent.KillObject oOut
      Set oNS = Nothing
   End If
   SendOutlook = True
   Exit Function
No_Outlook_Object_Error:
   'If moParent.Workflow.IsSilentErrors Then
   '   On Error GoTo SendOutlook_ErrorSilent
   'Else
   If mLaunchOutlook Then
      On Error GoTo SendOutlook_Error
      Set oOut = CreateObject("Outlook.Application")
      Set oNS = oOut.GetNamespace("MAPI")
      oNS.Logon
      WasOutlookObject = False
      Resume No_GetObject_Resume
      End If
   moParent.RaiseError "Error in mwSession.mwMessage.SendOutlook, Outlook Must Be Running Before Sending Data. Please start Outlook and try again."
   SendOutlook = False
   moParent.KillObject oMail
   moParent.KillObject oSafeMail
   moParent.KillObject oMAPIUtils
   moParent.KillObject oBtn
   Set oNS = Nothing
   Exit Function
   
SendOutlook_Error:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendOutlook.", Err.Number, Err.Description
   SendOutlook = False
   moParent.KillObject oMail
   moParent.KillObject oSafeMail
   moParent.KillObject oMAPIUtils
   moParent.KillObject oBtn
   Exit Function
SendOutlook_ErrorSilent:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendOutlook.", Err.Number, Err.Description
   SendOutlook = False
   moParent.KillObject oMail
   moParent.KillObject oSafeMail
   moParent.KillObject oMAPIUtils
   moParent.KillObject oBtn

End Function

Private Function SendAmosMail() As Boolean
   Dim strMessage As String
   Dim strAttachTemp As String
   Dim IsDatagram As Boolean
   Dim strAttachBak As String
   Dim fso As FileSystemObject
   Dim i As Long
   On Error GoTo FunctionError
   '
   '  Copy attachment to temp folder...Amosmail will delete it...
   '
   If Trim(mAttachmentFile) <> "" Then
      Set fso = New FileSystemObject
      strAttachTemp = moParent.Workflow.WorkingSessionFolder & "\" & fso.GetFileName(mAttachmentFile)
      'strAttachTemp = moParent.Workflow.SendByMediaFolder & "\" & fso.GetFileName(mAttachmentFile)
      fso.CopyFile mAttachmentFile, strAttachTemp
      '
      ' Determine if its a datagram...
      '
      If fso.GetParentFolderName(mAttachmentFile) = fso.GetParentFolderName(strAttachTemp) Then
         IsDatagram = True
         'AmosMail wipes file to send - in datagrams, this is a bad thing.
         strAttachBak = strAttachTemp & ".bak"
         fso.CopyFile mAttachmentFile, strAttachBak
      End If
      
   End If
   

   ' Building the header information
   strMessage = "To:   " & mSendToAddress & vbLf
   If Trim(mAttachmentFile) <> "" Then
      strMessage = strMessage & "Attach: " & strAttachTemp & vbLf
   End If
   'Using GUI doesn't send it - puts user in draft mode...
   If mUserResolveAddress Then
      strMessage = strMessage & "Gui:    Yes" & vbLf
   End If
   strMessage = strMessage & "Subject: " & mSubject & vbLf & vbLf
   strMessage = strMessage & mBody
   '
   ' Now send it...
   '
   If moParent.DebugMode Then
      MsgBox "Debug Mode: " & strMessage
   End If
   If moParent.Workflow.GetExtendedProperty("AmosMailSimulator") = "True" Then
      'Amos Mail Simulator
      fso.DeleteFile strAttachTemp
      MsgBox "AmosMail Simulator is on, mail flagged mail as sent...."
      i = 1
   Else
      '
      ' Call to MAPITIF.DLL
      '
      i = SendMail(strMessage)
   End If
   If i = 1 Then
      SendAmosMail = True
      ' If datagram, restore it so it can be archived...
      If IsDatagram Then
         fso.MoveFile strAttachBak, mAttachmentFile
      End If
   Else
      moParent.RaiseError "Error sending via AmosMAIL, note you must be logged into AmosMAIL in order to send a message."
      SendAmosMail = False
   End If
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendAmosMail.", Err.Number, Err.Description
   SendAmosMail = False
   Exit Function
End Function

Private Function SendGroupWise() As Boolean
   Dim objApp As Object
   Dim objAccount As Object
   Dim objDraftMsg As Object
   Dim objSentMsg As Object
   Dim objRecip As Object
   On Error GoTo FunctionError
   '
   ' Login (using default user name and password) - if req'ed ask for password
   '
   Set objApp = CreateObject("NovellGroupWareSession")
   'MsgBox "User: " & moParent.User.MailUserID & ", Pwd: " & moParent.User.MailPassword
   Set objAccount = objApp.Login(moParent.User.MailUserID, moParent.User.MailPassword)
   'Set objAccount = objApp.Login("vrs", "vrsndal")
   'MsgBox "after login"
   ' Get some objecte for later use
   'Set objRootFolder = objAccount.RootFolder
   'Set objAddressBook = objAccount.SystemAddressBook
   'Set objAddressBookEntries = objAddressBook.AddressBookEntries
   'Message & Header
   Set objDraftMsg = objAccount.WorkFolder.Messages.Add ' create draft in work folder
   'MsgBox "after msg create"
   objDraftMsg.subject = mSubject
   If mAttachmentFile <> "" Then
      objDraftMsg.Attachments.Add (mAttachmentFile)
   End If
   objDraftMsg.BodyText = mBody
   
   'Setup Recipients & Resolve
   'Set objRecip = objDraftMsg.Recipients.Add(objAccount.Owner.EmailAddress, egwTo)  ' to self
   ' MsgBox "mSendToAddress=" & mSendToAddress
   If mSendToAddress <> "" And Not mUserResolveAddress Then
      Set objRecip = objDraftMsg.Recipients.Add(mSendToAddress)
      Call objRecip.Resolve
      'If objRecip.Resolved <> egwResolved Then
      If objRecip.Resolved <> 3 Then
         goSession.RaiseError "Error in mwMessage.SendGroupWise, unable to resolve: " & mSendToAddress
         KillObject objRecip
         KillObject objDraftMsg
         KillObject objAccount
         KillObject objApp
         SendGroupWise = False
         Exit Function
         'Set objRecip = objDraftMsg.Recipients.AddByDisplayName(objAccount.Owner.DisplayName, egwTo)
         'Set objRecip = objDraftMsg.Recipients.AddByDisplayName(objAccount.Owner.DisplayName, 0)
      End If
      Set objSentMsg = objDraftMsg.Send
   End If
   ' Now, send the stuff
   'MsgBox "just before .send"
   'Set objSentMsg = objDraftMsg.Send
   'MsgBox "after .send"
   SendGroupWise = True
   KillObject objRecip
   KillObject objDraftMsg
   KillObject objAccount
   KillObject objApp
   Exit Function
FunctionError:
   If Err.Number = 429 Then
      moParent.RaiseError "Error in mwSession.mwMessage.SendGroupWise. Cannot create Groupwise object; maybe not installed ?"
   Else
      moParent.RaiseError "General Error in mwSession.mwMessage.SendGroupWise.", Err.Number, Err.Description
   End If
   KillObject objRecip
   KillObject objDraftMsg
   KillObject objAccount
   KillObject objApp
   SendGroupWise = False
   Exit Function
End Function

'
' AmosLINK is the DOS Version - an MHS like store/forward system
'  MSG goes to one folder,
'  the attachment to another...
'
Private Function SendAmosLink() As Boolean
   On Error GoTo FunctionError
   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim s As String
   Dim sMsgFolder As String
   Dim sMsgFullFilename As String
   Dim sRandomFile As String
   Dim sAttFolder As String
   Dim sAttFullFilename As String
   Dim sTemp As String
   On Error GoTo FunctionError
   Set fso = New FileSystemObject
   'sMsgFolder = moParent.GetPlace(moParent.Site.TransportPlaceID, mw_DriveLetterPath)
   ' Following line may not work, but no active AmosLink clients anyway...
   sMsgFolder = moParent.Workflow.SendByMediaFolder & "\snd"
   sAttFolder = moParent.Workflow.SendByMediaFolder & "\parcel"
   '
   ' Message Header Folder
   '
   If Not fso.FolderExists(sMsgFolder) Then
      moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Folder missing.: " & sMsgFolder
      SendAmosLink = False
      moParent.KillObject fso
      Exit Function
   End If
   If Not fso.FolderExists(sAttFolder) Then
      moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Folder missing.: " & sAttFolder
      SendAmosLink = False
      moParent.KillObject fso
      Exit Function
   End If
   '
   ' last four chars to alpha
   '
   s = fso.GetBaseName(mAttachmentFile)
   sRandomFile = UCase(Left(s, 4)) & _
    Chr(Asc(mID$(s, 5, 1)) + 17) & _
    Chr(Asc(mID$(s, 6, 1)) + 17) & _
    Chr(Asc(mID$(s, 7, 1)) + 17) & _
    Chr(Asc(mID$(s, 8, 1)) + 17)
   
   
   sMsgFullFilename = sMsgFolder & "\" & sRandomFile
   '
   ' Message Header Filename, if short enough (8.3) - make the same as attachment. Else, fabricate a name.
   '
   If fso.FileExists(sMsgFullFilename) Then
      moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Message Header Already Exists: " & sMsgFullFilename
      SendAmosLink = False
      moParent.KillObject ts
      moParent.KillObject fso
      Exit Function
   End If
   
   ' Attachment Header File
   If Trim(mAttachmentFile) <> "" Then
      sAttFullFilename = sAttFolder & "\" & sRandomFile & ".001"
      If fso.FileExists(sAttFullFilename) Then
         moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Attachment File Already Exists: " & sAttFullFilename
         SendAmosLink = False
         moParent.KillObject ts
         moParent.KillObject fso
         Exit Function
      End If
   End If
   '
   ' Let's do it...
   '
   Set ts = fso.CreateTextFile(sMsgFullFilename, False)
   ts.WriteLine AL_HEADER_LINE_1
   ts.WriteLine AL_HEADER_LINE_2
   If Trim(mAttachmentFile) = "" Then
      sTemp = ""
   Else
      sTemp = fso.GetFileName(mAttachmentFile)
   End If
   ts.WriteLine AL_HEADER_LINE_3 & sRandomFile & ".001"
   ts.WriteLine AL_HEADER_LINE_4 & sTemp
   ts.WriteLine AL_HEADER_LINE_5 & mSendToAddress & "}"
   ts.WriteLine AL_HEADER_LINE_6 & mSendFromAddress
   ts.WriteLine AL_HEADER_LINE_7 & Left(mSubject, 20)
   ts.WriteBlankLines 1
   ts.WriteLine mBody
   ts.Close
   Set ts = Nothing
   '
   ' Copy Attachment File...
   '
   If Trim(mAttachmentFile) <> "" Then
      fso.CopyFile mAttachmentFile, sAttFullFilename
   End If
   moParent.KillObject fso
   moParent.KillObject ts
   SendAmosLink = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendAmosLink.", Err.Number, Err.Description
   SendAmosLink = False
   moParent.KillObject ts
   moParent.KillObject fso
End Function



'Private Function GetRandomFile(Optional sFile As String) As String
'   Dim s As String
'   On Error GoTo FunctionError:
'
'
'
'   Exit Function
'FunctionError:
'   moParent.RaiseError "General Error in mwSession.mwMessage.SendAmosLink.", err.Number, err.Description
'End Function



Private Function SendMAPI() As Boolean
   On Error GoTo FunctionError
   goSession.RaiseError "MAPI Send Function currently under construction."
   SendMAPI = False
   
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendMAPI.", Err.Number, Err.Description
   SendMAPI = False
End Function


Private Function SendOutlook97() As Boolean
   Dim WasOutlookObject As Boolean
   Dim oOut As Object
   Dim oMail As Object
   Dim IsGetOutlookApplication As Boolean
   
   'DEV-1797 Sending Email with Outlook client 97
   'Added By N.Angelakis On 02 Feb 2010
   
   'this routine is the same as SendOutlook, the difference being that for
   'version 97 Getobject will fail and we need CreateObject
   'enabling mLaunchOutlook = moKeys.GetBoolKeyValue("AutoLaunchOutlookOnSendTran")
   'would have done it but being in switches is not obvious to anyone
   
   WasOutlookObject = True
   On Error GoTo ErrorHandler
   
   'Late Binding...
   IsGetOutlookApplication = True
   Set oOut = GetObject(, "Outlook.ApplicationX")
   
No_GetObject_Resume:
   IsGetOutlookApplication = False
   
   Set oMail = oOut.CreateItem(0)
   With oMail
      .subject = mSubject
      mBody = Replace(mBody, vbCrLf, " <br> ")

       .Body = mBody
       .htmlbody = mBody
      If Trim(mAttachmentFile) <> "" Then
         .Attachments.Add mAttachmentFile
      End If
      '
      ' Allow user to send email...
      '
      If mUserResolveAddress Then
         .Display vbModal
      Else
         .Recipients.Add mSendToAddress
         .Send
         
      End If
   End With
  
   moParent.KillObject oMail
  If Not WasOutlookObject Then
      oOut.Quit
      moParent.KillObject oOut
   End If
   
   SendOutlook97 = True
   Exit Function
   
ErrorHandler:
   If IsGetOutlookApplication = True Then
      IsGetOutlookApplication = False
      On Error GoTo GeneralOutlook_Error
      Set oOut = CreateObject("Outlook.Application") 'client outlook 97
      Resume No_GetObject_Resume
   Else
      moParent.RaiseError "Error in mwSession.mwMessage.SendOutlook97" & vbCrLf & Err.Number & ", " & Err.Description
      SendOutlook97 = False
      moParent.KillObject oMail
      Exit Function
   End If
GeneralOutlook_Error:
      moParent.RaiseError "Error in mwSession.mwMessage.SendOutlook97" & vbCrLf & Err.Number & ", " & Err.Description
      SendOutlook97 = False
End Function



Private Function SendByMediaLink() As Boolean

   Dim fso As FileSystemObject
   Dim ts As TextStream
   Dim s As String
   Dim sMsgFolder As String
   Dim sMsgFullFilename As String
   Dim sRandomFile As String
   Dim sAttFolder As String
   Dim sAttFullFilename As String
   Dim sTemp As String

   On Error GoTo FunctionError
   'DEV-1809 Send By Media Error Notification
   'Added By N.Angelakis On 04 Feb 2010

'   Set fso = New FileSystemObject

   sMsgFolder = moParent.Workflow.SendByMediaFolder & "\snd"
   sAttFolder = moParent.Workflow.SendByMediaFolder & "\parcel"
   '
   ' Message Header Folder
   '
'   If Not fso.FolderExists(sMsgFolder) Then
'      moParent.RaiseError "Error in mwSession.mwMessage.SendByMediaLink. Folder missing.: " & sMsgFolder
'      SendByMediaLink = False
'      moParent.KillObject fso
'      Exit Function
'   End If
'   If Not fso.FolderExists(sAttFolder) Then
'      moParent.RaiseError "Error in mwSession.mwMessage.SendByMediaLink. Folder missing.: " & sAttFolder
'      SendByMediaLink = False
'      moParent.KillObject fso
'      Exit Function
'   End If
''   '
''   ' last four chars to alpha
''   '
''   s = fso.GetBaseName(mAttachmentFile)
''   sRandomFile = UCase(Left(s, 4)) & _
''    Chr(Asc(Mid$(s, 5, 1)) + 17) & _
''    Chr(Asc(Mid$(s, 6, 1)) + 17) & _
''    Chr(Asc(Mid$(s, 7, 1)) + 17) & _
''    Chr(Asc(Mid$(s, 8, 1)) + 17)
''
''
''   sMsgFullFilename = sMsgFolder & "\" & sRandomFile
''   '
''   ' Message Header Filename, if short enough (8.3) - make the same as attachment. Else, fabricate a name.
''   '
''   If fso.FileExists(sMsgFullFilename) Then
''      moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Message Header Already Exists: " & sMsgFullFilename
''      SendAmosLink = False
''      moParent.KillObject ts
''      moParent.KillObject fso
''      Exit Function
''   End If
''
''   ' Attachment Header File
''   If Trim(mAttachmentFile) <> "" Then
''      sAttFullFilename = sAttFolder & "\" & sRandomFile & ".001"
''      If fso.FileExists(sAttFullFilename) Then
''         moParent.RaiseError "Error in mwSession.mwMessage.SendAmosLink. Attachment File Already Exists: " & sAttFullFilename
''         SendAmosLink = False
''         moParent.KillObject ts
''         moParent.KillObject fso
''         Exit Function
''      End If
''   End If
''   '
''   ' Let's do it...
''   '
''   Set ts = fso.CreateTextFile(sMsgFullFilename, False)
''   ts.WriteLine AL_HEADER_LINE_1
''   ts.WriteLine AL_HEADER_LINE_2
''   If Trim(mAttachmentFile) = "" Then
''      sTemp = ""
''   Else
''      sTemp = fso.GetFileName(mAttachmentFile)
''   End If
''   ts.WriteLine AL_HEADER_LINE_3 & sRandomFile & ".001"
''   ts.WriteLine AL_HEADER_LINE_4 & sTemp
''   ts.WriteLine AL_HEADER_LINE_5 & mSendToAddress & "}"
''   ts.WriteLine AL_HEADER_LINE_6 & mSendFromAddress
''   ts.WriteLine AL_HEADER_LINE_7 & Left(mSubject, 20)
''   ts.WriteBlankLines 1
''   ts.WriteLine mBody
''   ts.Close
''   Set ts = Nothing
''   '
''   ' Copy Attachment File...
''   '
''   If Trim(mAttachmentFile) <> "" Then
''      fso.CopyFile mAttachmentFile, sAttFullFilename
''   End If
''   moParent.KillObject fso
''   moParent.KillObject ts
   SendByMediaLink = True

   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwMessage.SendByMediaLink.", Err.Number, Err.Description
   SendByMediaLink = False
End Function


' 8/2010 ms BGLNG-15
Private Function FunnelToMessageQueue(SendToAddress As String, _
   AttachmentFullFileName As String, subject As String, TextBody As String) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   sSQL = "select * from mwEmailQueue where ID=-1"
   Set loRs = New Recordset
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.AddNew
   loRs!ID = goSession.MakePK("mwEmailQueue")
   loRs!DateTimeQueued = Now()
   loRs!OriginatorUserKey = goSession.User.UserKey
   loRs!SendToAddress = SendToAddress
   loRs!AttachmentFullFileName = AttachmentFullFileName
   loRs!subject = Left(subject, 255)
   loRs!TextBody = TextBody
   loRs.Update
   CloseRecordset loRs
   FunnelToMessageQueue = True
   Exit Function
FunctionError:
   moParent.RaiseError "General Error in mwSession.mwMessage.FunnelToMessageQueue.", Err.Number, Err.Description
   FunnelToMessageQueue = False
End Function


