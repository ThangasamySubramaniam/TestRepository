VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwTransportContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"TransportDocuments"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mw_Container - Transport Container class
' 11/26/2000 ms - Copyright 2000 Maritime Systems Inc - All Rights Reserved
' Transport Container is a zip file that contains a collection of documents/files...
' 1/18/2001 - Continue Development
'
Option Explicit

Private Const mw__CONTAINER_PREFIX = "TC"
Private Const mw__CONTAINER_SUFFIX = ".zip"

Private mFromSiteID As String
Private mToSiteID As String
Private mTransportZipFile As String
Private moErr As mwError
Private mIsOpen As Boolean
Private moDocuments As mwZipItems
Private moZip As Zipper
Private mDuplicateOverwrite As String
Private mTransportFolder As String

Private moParent As Session
Event Status(Msg As String, percent_complete As Integer, Err As Boolean)


Public Property Set Parent(oParent As Session)
   If moParent Is Nothing Then
      Set moParent = oParent
   End If
End Property

Public Property Get Parent() As Session
   Set Parent = moParent
End Property

Public Property Get Documents() As mwZipItems
    Set Documents = moDocuments
End Property

Public Property Let TransportFolder(ByVal Data As String)
    mTransportFolder = Data
End Property

Public Property Get TransportFolder() As String
    TransportFolder = mTransportFolder
End Property



'Public Property Set Documents(vData As mwZipItems)
'    Set moDocuments = vData
'End Property

Public Property Let FromSiteID(ByVal Data As String)
    mFromSiteID = Data
End Property

Public Property Get FromSiteID() As String
    FromSiteID = mFromSiteID
End Property

Public Property Let ToSiteID(ByVal Data As String)
    mToSiteID = Data
End Property

Public Property Get ToSiteID() As String
    ToSiteID = mToSiteID
End Property

Public Property Get TransportFilename() As String
    'TransportFilename = MakeFilename
End Property

Public Property Get IsContainerOpen() As Boolean
    IsContainerOpen = mIsOpen
End Property

Public Property Let DuplicateOverwrite(ByVal Data As String)
   If Data = "0" Or Data = "1" Or Data = "2" Then
      mDuplicateOverwrite = Data
   End If
End Property

Public Property Get DuplicateOverwrite() As String
    DuplicateOverwrite = mDuplicateOverwrite
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


Public Function OpenContainer() As Boolean
   Dim oErr As mwError
   Dim fso As FileSystemObject
   Dim strFile As String
   On Error GoTo OpenContainer_error
   If mIsOpen Then
      OpenContainer = True
      Exit Function
   End If
   '
   ' Error checking
   '
   Set fso = New FileSystemObject
   If Not fso.FolderExists(moParent.Container.TransportFolder) Then
      moParent.RaiseError "Error in mwSession.mwTransportContainer.OpenContainer, Transport Folder Does Not Exist: " & moParent.Container.TransportFolder, Err.Number, Err.Description
      OpenContainer = False
      mIsOpen = False
      moParent.KillObject fso
      Exit Function
   End If
   If Trim(mFromSiteID) = "" Or Trim(mToSiteID) = "" Then
      moParent.RaiseError "Error in mwSession.mwTransportContainer.OpenContainer, FromSiteID and ToSiteID are missing", Err.Number, Err.Description
      OpenContainer = False
      mIsOpen = False
      moParent.KillObject fso
      Exit Function
   End If
   '
   ' Set Zip File Name, ready to go...
   '
   mTransportZipFile = moParent.Container.TransportFolder & "\" & mw__CONTAINER_PREFIX & mToSiteID _
     & mFromSiteID & mw__CONTAINER_SUFFIX
   '
   ' Open the zip file
   '
   moZip.ZipFile = mTransportZipFile
   If Not moZip.OpenZip Then
      moParent.RaiseError "Error in mwSession.mwTransportContainer.OpenContainer, Zip File, " & mTransportZipFile _
        & " is Damaged, Error Code is: " & moZip.ErrorCode, Err.Number, Err.Description
      OpenContainer = False
      mIsOpen = False
      moParent.KillObject fso
      Exit Function
   End If
   mIsOpen = True
   OpenContainer = True
   '
   ' Set Documents List = Zip Items...
   '
   Set moDocuments = moZip.mwZipItems
   Exit Function
OpenContainer_error:
   moParent.RaiseError "General Error in mwSession.mwTransportContainer.OpenContainer.", Err.Number, Err.Description
   OpenContainer = False
   mIsOpen = False
   moParent.KillObject fso
End Function

Private Sub Class_Initialize()
   Set moDocuments = New mwZipItems
   Set moZip = New Zipper
   mDuplicateOverwrite = "0"
End Sub


Private Sub Class_Terminate()
   Set moDocuments = Nothing
   moZip.TerminateClass
   Set moZip = Nothing
End Sub


Public Function AddFile(DocFile As String) As Boolean
   Dim oErr As mwError
   Dim fso As FileSystemObject
   '
   ' Transport Container (Zip File)
   '
   If Not mIsOpen Then
      AddFile = False
      Exit Function
   End If
   Set fso = New FileSystemObject
   If Not fso.FileExists(DocFile) Then
      AddFile = False
      Set fso = Nothing
      Exit Function
   End If
   Set fso = Nothing
   RaiseEvent Status("Placing Document in Transport Container", 10, False)
   If Not moZip.AddToZip(DocFile) Then
      'Zip Error
      moParent.RaiseError "Error in mwTransportContainer, Add to Zip Error: " & moZip.ErrorCode
      Exit Function
   End If
   moParent.LogIt mwl_Transport, mwl_Information, "Form Posted to Transport Container by: " & moParent.User.UserID
   AddFile = True
End Function



Public Function RemoveDocument(DocName As String)
   If Not mIsOpen Then
      RemoveDocument = False
      Exit Function
   End If
   MsgBox "RemoveDocument - Not implemented."

End Function

