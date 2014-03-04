VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Zipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"mwZipItems"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Zipper Class - Zip and Unzip functions for WF Maintenance
' 12/05/2000 - ms
' 12/28/2004 - ms Upgrade to Secure Max Dynazip
' Uses Dynazip Active X controls...
Option Explicit
' Dynazip objects that do the work...

#If NO_DYNAZIP Then
   Private moZip As Object
   Private moUnzip As Object
   Private mZipAction As Object
   Private mUnzipAction As Object
#Else
   'Private moZip As dzactxctrl
   'Private moUnzip As duzactxctrl
   Private moZip As dzsactxctrl
   Private moUnzip As duzsactxctrl
   Private mZipAction As ZipAction
   Private mUnzipAction As UnZipAction
#End If

Private mIsOpen As Boolean
Private mErrorCode As Integer

' 5/2004 Encryption Support
Private mIsEncypt As Boolean
Private mPassword As String

Public Enum ZipAction
   ZIP_FRESHEN = 1
   ZIP_DELETE = 2
   ZIP_UPDATE = 3
   ZIP_ADD = 4
   ZIP_MEMTOFILE = 5
End Enum

Public Enum UnZipAction
   UNZIP_COUNTALLZIPMEMBERS = 1
   UNZIP_GETNEXTZIPINFO = 2
   UNZIP_COUNTNAMEDZIPMEMBERS = 3
   UNZIP_GETNEXTNAMEDZIPINFO = 4
   UNZIP_GETCOMMENTSIZE = 5
   UNZIP_GETCOMMENT = 6
   UNZIP_GETINDEXEDZIPINFO = 7
   UNZIP_EXTRACT = 8
   UNZIP_FILETOMEM = 9
End Enum
Private momwZipItems As mwZipItems

Friend Property Set Parent(oParent As mwSession.Session)
End Property

Friend Property Get Parent() As mwSession.Session
End Property


Friend Property Get mwZipItems() As mwZipItems
    Set mwZipItems = momwZipItems
End Property


Friend Property Set mwZipItems(vData As mwZipItems)
    Set momwZipItems = vData
End Property



' Zip Errors
'ZE_MISS = -1             'used by procname(), zipbare()
'ZE_OK = 0                'success
'ZE_EOF = 2               'unexpected end of zip file
'ZE_FORM = 3              'zip file structure error
'ZE_MEM = 4               'out of memory
'ZE_LOGIC = 5             'internal logic error
'ZE_BIG = 6               'entry too large to split
'ZE_NOTE = 7              'invalid comment format
'ZE_TEST = 8              'zip test (-T) failed or out of memory
'ZE_ABORT = 9             'user interrupt or termination
'ZE_TEMP = 10             'error using a temp file
'ZE_READ = 11             'read or seek error
'ZE_NONE = 12             'nothing to do
'ZE_NAME = 13             'missing or empty zip file
'ZE_WRITE = 14            'error writing to a file
'ZE_CREAT = 15            'couldn't open to write
'ZE_PARMS = 16            'bad command line
'ZE_INCOMPLETE = 17       'Could Not Complete Operation
'ZE_OPEN = 18             'could not open a specified file to read
'ZE_MEDIA = 19            'Media error or HW failure
'ZE_MVPARMS = 20          'Invalid combination of control parameters
'ZE_MVUSAGE = 21          'Improper use of a Multi-Volume Zip file
   

'Unzip errors
'UE_OK = 0                ' success
'UE_EOF = 2               ' unexpected end of zip file
'UE_STRUCT = 3            ' structure error in zip file
'UE_MEM1 = 4              ' out of memory
'UE_MEM2 = 5              ' out of memory
'UE_NOFILE = 9            ' file not found error
'UE_BORED = 11            ' nothing to do
'UE_SAMEVOL = 12          ' same volume
'UE_INDEX = 25            ' index out of bounds
'UE_OUTPUT = 28           ' error creating output file
'UE_OPEN = 29             ' error opening output file
'UE_BADCRC = 39           ' crc error
'UE_ABORT = 40            ' Application canceled operation
'UE_ENCRYPT = 41          ' file skipped, encrypted
'UE_UNKNOWN = 42          ' unknown compression method
'UE_BADKEY = 44           ' Bad or missing decrypt code
'UE_BUSY = 45         ' Busy, can't enter now
'UE_NOVOL = 46            ' can't unzip a volume item
'UE_CMDERR = 47           ' bad command structure
'UE_CANCEL = 48           ' user cancelled this operation
'UE_SKIP = 49             ' user skipped this operation
'UE_DISKFULL = 50         ' disk full


Public Property Get Zip() As dzsactxctrl
   Set Zip = moZip
End Property

Public Property Get Unzip() As duzsactxctrl
   Set Unzip = moUnzip
End Property

Public Property Let DeleteAfterZip(ByVal Data As Boolean)
    moZip.DeleteOriginalFlag = Data
End Property

Public Property Get DeleteAfterZip() As Boolean
    DeleteAfterZip = moZip.DeleteOriginalFlag
End Property

Public Property Let ZipFile(ByVal Data As String)
   moZip.ZipFile = Data
   moUnzip.ZipFile = Data
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = mIsOpen
End Property

Public Property Get ErrorCode() As Integer
    ErrorCode = mErrorCode
End Property

Public Property Let IsEncypt(ByVal Data As Boolean)
   On Error GoTo PropertyError
   If Data Then
      mIsEncypt = Data
      moZip.EncryptFlag = True
      moZip.EncryptCode = "9763p00dlep00p"
      moUnzip.DecryptFlag = True
      moUnzip.DecryptCode = "9763p00dlep00p"
   End If
   Exit Property
PropertyError:
   goSession.RaiseError "General Error in Zipper.IsEncypt. ", Err.Number, Err.Description
End Property

Public Property Get IsEncypt() As Boolean
    IsEncypt = mIsEncypt
End Property

Public Property Let Password(ByVal Data As String)
    mPassword = Data
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property

Public Function OpenZip(Optional file As String) As Boolean
   ' Open zip file, create collection of zip items...
   Dim iCounter As Integer
   Dim iTotal As Integer
   
   Dim oItem As mwZipItem
   On Error GoTo FunctionError
   
   If Trim(file) <> "" Then
      moZip.ZipFile = file
      moUnzip.ZipFile = file
   End If
      
   mUnzipAction = UNZIP_COUNTALLZIPMEMBERS
   moUnzip.ActionDZ = mUnzipAction
   If moUnzip.ErrorCode = 3 Then
      ' Not a zip file
      OpenZip = False
      mIsOpen = False
      Exit Function
   End If
   '
   ' Clear old momwZip Items...
   '
   Do While momwZipItems.Count > 0
      momwZipItems.Remove 1
   Loop
   iTotal = moUnzip.ReturnCount
   For iCounter = 0 To (iTotal - 1)
      mUnzipAction = UNZIP_GETNEXTZIPINFO
      moUnzip.ActionDZ = mUnzipAction
      If IsDate(moUnzip.zi_DateTime) Then
         momwZipItems.Add moUnzip.zi_FileName, _
           moUnzip.zi_oSize, moUnzip.zi_DateTime, _
           moUnzip.zi_cPathType, moUnzip.zi_FileName
      Else
         momwZipItems.Add moUnzip.zi_FileName, _
           moUnzip.zi_oSize, Now(), _
           moUnzip.zi_cPathType, moUnzip.zi_FileName
      End If
   Next iCounter
   OpenZip = True
   mIsOpen = True
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in Zipper.OpenZip. ", Err.Number, Err.Description
   OpenZip = False
End Function



Public Function AddToZip(FileName As String) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   If Not mIsOpen Then
      AddToZip = False
      Exit Function
   End If
   If Trim(FileName) = "" Then
      goSession.RaiseError "Error in Zipper.AddToZip, no FileName parameter specified."
      AddToZip = False
      Exit Function
   End If
   If moZip.DeleteOriginalFlag Then
      Set fso = New FileSystemObject
      fso.GetFile(FileName).Attributes = Normal
   End If
   'RemoveFromZip FileName
   moZip.ItemList = Chr(34) & FileName & Chr(34)
   
   ' Execute the Zip Command
   mZipAction = ZIP_ADD
   moZip.ActionDZ = mZipAction
   mErrorCode = moZip.ErrorCode
    ' Error 12 means "nothing to do"
   'If moZip.ErrorCode <> 0 And moZip.ErrorCode <> 12 Then
   If moZip.ErrorCode <> 0 Then
      AddToZip = False
      mErrorCode = moZip.ErrorCode
   Else
      AddToZip = True
      momwZipItems.Add FileName, "-1", "1/1/1900", False, FileName
      If moZip.DeleteOriginalFlag Then
         If fso.FileExists(FileName) Then
            fso.DeleteFile FileName
         End If
         Set fso = Nothing
      End If
   End If
   Exit Function
FunctionError:
   mErrorCode = Err.Number
   If Err.Number = 70 Then
      goSession.RaiseError "Error in Zipper.AddToZip, File is locked and cannot be deleted now: " & FileName
   Else
      goSession.RaiseError "General Error in Zipper.AddToZip. ", Err.Number, Err.Description
   End If
   AddToZip = False
End Function


Public Function GetFromZip(FileList As String, TargetFolder As String) As Boolean
    ' Function Call Settings...
   On Error GoTo FunctionError
   If Not mIsOpen Then
      GetFromZip = False
      Exit Function
   End If
   moUnzip.Filespec = FileList
   'moUnzip.ZipFile = ZipFile
   moUnzip.Destination = TargetFolder
   ' doit
   moUnzip.ActionDZ = 8
   mErrorCode = moUnzip.ErrorCode
   If moUnzip.ErrorCode <> 0 Then
      GetFromZip = False
      goSession.LogIt mwl_Workstation, mwl_Warning, "Error " & moUnzip.ErrorCode & " from mwSession.Zipper.GetFromZip"
   Else
      GetFromZip = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.Zipper.GetFromZip.", Err.Number, Err.Description
   GetFromZip = False
End Function

Public Function RemoveFromZip(FileName As String) As Boolean
   If Not mIsOpen Then
      RemoveFromZip = False
      Exit Function
   End If
   'moZip.ItemList = goSession.FSO.GetFileName(FileName)
   ' Execute the Zip Command
   moZip.ActionDZ = 2
   mErrorCode = moUnzip.ErrorCode
    ' Error 12 means "nothing to do"
   If moZip.ErrorCode <> 0 And moZip.ErrorCode <> 12 Then
      RemoveFromZip = False
      mErrorCode = moZip.ErrorCode
   Else
      RemoveFromZip = True
      ' Remove from collection
      momwZipItems.Remove FileName
   End If
End Function

Private Sub Class_Initialize()
   'On Error GoTo SubError
   ' Stage the zip control
   Set moZip = New dzsactxctrl
   moZip.AllQuiet = True
   moZip.MessageCallbackFlag = True
   moZip.NoDirectoryEntriesFlag = True
   moZip.NoDirectoryNamesFlag = True
   moZip.BackgroundProcessFlag = True
   moZip.IncludeOnlyFollowingFlag = False
   
   'Stage the unzip control
   Set moUnzip = New duzsactxctrl
   moUnzip.AllQuiet = True
   moUnzip.MessageCallbackFlag = True
   moUnzip.BackgroundProcessFlag = True
   moUnzip.NoDirectoryItemsFlag = True
   moUnzip.NoDirectoryNamesFlag = True
   'Temporary - overwrite is yes...
'    moUnzip.OverwriteFlag = False
   moUnzip.OverwriteFlag = True
   
   Set momwZipItems = New mwZipItems
   Set momwZipItems.Parent = Me
   mIsOpen = False
   
End Sub
Public Sub TerminateClass()
   Set momwZipItems = Nothing
   Set moZip = Nothing
   Set moUnzip = Nothing
End Sub

Private Sub Class_Terminate()
   Set momwZipItems = Nothing
   Set moZip = Nothing
   Set moUnzip = Nothing
End Sub


Public Function UnzipToTempFolder(TempFolder As String) As Boolean
   Dim fso As FileSystemObject
   Dim oFolder As Folder
   Dim oFile As file
   Dim sAction As String
   
   On Error GoTo FunctionError
   ' Zip must be open...
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwSession.Zipper.UnzipToTempFolder, Zip File Not Open."
      UnzipToTempFolder = False
      Exit Function
   End If
   ' Destroy and remake folder
   Set fso = New FileSystemObject
   If fso.FolderExists(TempFolder) Then
      Set oFolder = fso.GetFolder(TempFolder)
      For Each oFile In oFolder.Files
         oFile.Attributes = Normal
      Next
      sAction = "DeleteFolder"
      fso.DeleteFolder TempFolder
   End If
   sAction = "CreateFolder"
   fso.CreateFolder TempFolder
   '
   ' Unzip file to working folder...
   '
   GetFromZip "*.*", TempFolder
   If mErrorCode <> 0 Then
      goSession.RaiseError "Error in mwSession.Zipper.UnzipToTempFolder, Zip Error Code: " & mErrorCode
      UnzipToTempFolder = False
   Else
      UnzipToTempFolder = True
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.Zipper.UnzipToTempFolder. Action = " & sAction & " Folder = " & TempFolder, Err.Number, Err.Description
   UnzipToTempFolder = False
End Function

Public Function ZipStringToFile(strData As String, FileName As String) As Boolean
   On Error GoTo FunctionError
   moZip.ItemList = FileName
   mZipAction = ZIP_DELETE
   moZip.ActionDZ = mZipAction
   moZip.ItemList = FileName
   moZip.ZIPString = strData
   mZipAction = ZIP_MEMTOFILE
   moZip.ActionDZ = mZipAction
   If moZip.ErrorCode = 0 Then
      ZipStringToFile = True
   Else
      ZipStringToFile = False
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.Zipper.ZipStringToFile.", Err.Number, Err.Description
End Function

Private Sub moZip_ZipRenameCallback(ItemName As String, iDate As Long, iTime As Long, lAttrib As Long, OrigItemName As String, rc As Long)
   
'  Set the RenameCallbackFlag to True to activate
'  Change the file extension from .ext to .SNOext when it's put into the ZIP and rename it back to .ext when it is unzipped
'   moUnzip.RenameCallbackFlag = True
   
   MsgBox ("Renaming this item: [" & ItemName & "]")
   
   ItemName = ItemName & "new"

End Sub

Private Sub moUnzip_UnZipRenameCallback(ItemName As String, iDate As Long, iTime As Long, lAttrib As Long, OrigItemName As String, rc As Long)
   
'  Set the RenameCallbackFlag to True to activate
'  Change the file extension from .ext to .SNOext when it's put into the ZIP and rename it back to .ext when it is unzipped
'   moUnzip.RenameCallbackFlag = True
   
   MsgBox ("Renaming this item: [" & ItemName & "]")
   
   ItemName = ItemName & "new"
End Sub
