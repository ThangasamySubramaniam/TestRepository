VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwBlobWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1

Public Function AddRecord(FullFilename As String, BriefDescription As String, _
 BlobTypeKey As Long, EventTypeKey As Long, EventDetailKey As Long, _
 Optional FileTypeDescription As String) As Long
   Dim loStream As ADODB.Stream
   Dim fso As FileSystemObject
   Dim sSQL As String
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   If Not fso.FileExists(FullFilename) Then
      goSession.RaisePublicError "Error in mwBlobWork.AddRecord, File to load not found: " & FullFilename
      KillObject fso
      AddRecord = -1
   End If
   
   Set loStream = New ADODB.Stream
   Set moRS = New Recordset
   
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwBlobFile where ID=-1"
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.LoadFromFile FullFilename
                      
      With moRS
         .AddNew
         .Fields("ID").value = goSession.MakePK("mwBlobFile")
         .Update
         .Fields("FullFileName").value = FullFilename
         .Fields("BriefDescription").value = BriefDescription
         .Fields("FileBlob").value = loStream.Read
         .Fields("mwBlobFileTypeKey").value = BlobTypeKey
         .Fields("mwEventTypeKey").value = EventTypeKey
         .Fields("mwEventDetailKey").value = EventDetailKey
         If FileTypeDescription <> "" Then
            .Fields("FileTypeDescription").value = FileTypeDescription
         End If
         .Fields("BlobCreated").value = Now()
         'Replication
         .Update
         
         AddRecord = moRS.Fields("ID").value
              
      End With
   
    
   CloseRecordset moRS
  
   KillObject fso
   KillObject loStream
     
   
   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.AddRecord. ", err.Number, err.Description
   CloseRecordset moRS
   KillObject fso
   KillObject loStream
   AddRecord = -1
End Function
Public Function FetchToFileSystem(BlobFileKey As Long, Optional IsSaveInOriginalFolder As Boolean, _
   Optional IsOverwrite As Boolean, Optional IsSetToReadOnly As Boolean, Optional SaveToFolder As String, Optional SaveToFileName As String) As String
      
   Dim loStream As ADODB.Stream
   Dim fso As FileSystemObject
   Dim strSQL As String
   Dim strFullFileName As String
   Dim strFilename As String
   Dim strFolder As String
   
   On Error GoTo FunctionError
   If IsSaveInOriginalFolder And SaveToFolder <> "" Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem,conflicting parameters on where to save the source file."
      Exit Function
   End If
   
   strSQL = "SELECT * FROM mwBlobFile WHERE ID=" & BlobFileKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If moRS.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, no source files to fetch."
      FetchToFileSystem = Null
      CloseRecordset moRS
      Exit Function
   End If
   Set fso = New FileSystemObject
   
   If SaveToFileName <> "" Then
      strFullFileName = SaveToFileName
   Else
      strFullFileName = moRS!FullFilename
   End If
   
   Set loStream = New ADODB.Stream
      
   If IsSaveInOriginalFolder Then
      If Not fso.FolderExists(fso.GetParentFolderName(strFullFileName)) Then
         strFolder = fso.GetParentFolderName(strFullFileName)
         goSession.GUI.ImprovedMsgBox "Folder not found: " & strFolder & vbCrLf & vbCrLf & "Creating it now.", vbInformation, "FetchToFileSystem Folder Not Found!"
         fso.CreateFolder (strFolder)
      Else
'         If fso.FileExists(strFullFileName) Then
'            If IsOverwrite Then
'               fso.GetFile(strFilename).Attributes = Normal
'               fso.DeleteFile (strFullFileName)
'            Else
'               goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, file exists and is not flagged to be overwritten."
'               CloseRecordset moRS
'               KillObject fso
'               KillObject loStream
'               FetchToFileSystem = Null
'               Exit Function
'            End If
'         End If
      End If
      strFilename = strFullFileName
   ElseIf SaveToFolder <> "" Then
      strFilename = SaveToFolder & "\" & fso.GetFileName(strFullFileName)
   Else
      strFilename = goSession.Workflow.WorkingSessionFolder & "\" & fso.GetFileName(strFullFileName)
'      If fso.FileExists(strFilename) Then
'         If Not IsOverwrite Then
'            MsgBox "The existing file is not flagged to be overwritten."
'            CloseRecordset moRS
'            KillObject fso
'            KillObject loStream
'            FetchToFileSystem = ""  ''Null wouldnt work?
'            Exit Function
'         Else
'            fso.GetFile(strFilename).Attributes = Normal
'            fso.DeleteFile (strFilename)
'         End If
'      End If
   End If
   
   If Not fso.FolderExists(fso.GetParentFolderName(strFilename)) Then
      strFolder = fso.GetParentFolderName(strFilename)
      goSession.GUI.ImprovedMsgBox "Folder not found: " & strFolder & vbCrLf & vbCrLf & "Creating it now.", vbInformation, "FetchToFileSystem Folder Not Found!"
      fso.CreateFolder (strFolder)
   End If
   
   If fso.FileExists(strFilename) Then
      If IsOverwrite Then
         fso.GetFile(strFilename).Attributes = Normal
         fso.DeleteFile (strFilename)
      Else
         goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, file exists and is not flagged to be overwritten."
         CloseRecordset moRS
         KillObject fso
         KillObject loStream
         FetchToFileSystem = Null
         Exit Function
      End If
   End If
   
   loStream.Type = adTypeBinary
   loStream.Open
   If IsNull(moRS.Fields("FileBlob").value) Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, file blob is Null !"
   Else
      
      loStream.Write moRS.Fields("FileBlob").value
      loStream.SaveToFile strFilename, adSaveCreateOverWrite
      If IsSetToReadOnly Then
         fso.GetFile(strFilename).Attributes = ReadOnly
      End If
      
      FetchToFileSystem = strFilename
   End If
   CloseRecordset moRS
   KillObject fso
   KillObject loStream
   Exit Function
   
FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.FetchToFileSystem", err.Number, err.Description
   CloseRecordset moRS
   KillObject fso
   KillObject loStream
   'FetchToFileSystem = Null

End Function
Public Function FetchToStream(BlobFileKey As Long) As ADODB.Stream
      
   Dim loStream As ADODB.Stream
   Dim strSQL As String
   Dim strFullFileName As String
      
   On Error GoTo FunctionError
   
   strSQL = "SELECT * FROM mwBlobFile WHERE ID=" & BlobFileKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If moRS.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToStream, no blobs to fetch."
      FetchToStream = Null
      CloseRecordset moRS
      Exit Function
   End If
   
   strFullFileName = moRS!FullFilename
   
   Set loStream = New ADODB.Stream
       
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.Write moRS.Fields("FileBlob").value
                  
   Set FetchToStream = loStream
   
   CloseRecordset moRS
   KillObject loStream
   
   Exit Function
   
FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.FetchToStream", err.Number, err.Description
   CloseRecordset moRS
   KillObject loStream
   FetchToStream = Null
End Function

Public Function UpdateRecord(BlobFileKey As Long, FullFilename As String, Optional IsReplaceFullFilename) As Boolean

   Dim loStream As ADODB.Stream
   
   Dim fso As FileSystemObject
   Dim strSQL As String
      
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   strSQL = "SELECT * FROM mwBlobFile WHERE ID=" & BlobFileKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If moRS.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, no blobs exist."
      UpdateRecord = False
      CloseRecordset moRS
      KillObject fso
      Exit Function
   End If
   If Not fso.FileExists(FullFilename) Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, no blobs exist."
      UpdateRecord = False
      CloseRecordset moRS
      KillObject fso
      Exit Function
   End If
   
   Set loStream = New ADODB.Stream
      
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.LoadFromFile FullFilename
                     
      With moRS
         
         .Fields("FileBlob").value = loStream.Read
         If IsReplaceFullFilename Then
         .Fields("FullFileName").value = FullFilename
         End If
         .Fields("BlobUpdated").value = Now()
         .Update
         
         UpdateRecord = True
                 
      End With
            
      CloseRecordset moRS
      KillObject fso
   
   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.UpdateRecord", err.Number, err.Description
   CloseRecordset moRS
   KillObject fso
   UpdateRecord = False


End Function

Public Function DeleteRecord(BlobFileKey As Long) As Boolean

   Dim fso As FileSystemObject
   Dim strSQL As String
   
   On Error GoTo FunctionError
   
   strSQL = "SELECT * FROM mwBlobFile WHERE ID=" & BlobFileKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open strSQL, goCon, adOpenDynamic, adLockOptimistic

   If moRS.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwBlobWork.FetchToFileSystem, no blobs exist."
      DeleteRecord = False
      CloseRecordset moRS
      KillObject fso
      Exit Function
   End If
      
   moRS.Delete
      
   CloseRecordset moRS
   KillObject fso
      
   Exit Function
   
FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.DeleteRecord", err.Number, err.Description
   CloseRecordset moRS
   KillObject fso
   DeleteRecord = False

End Function

Public Function GetFullFilename(BlobFileKey As Long) As String
   Dim strSQL As String
          
   On Error GoTo FunctionError
       
   strSQL = "SELECT * FROM mwBlobFile WHERE ID=" & BlobFileKey
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If moRS.RecordCount < 1 Then
      goSession.RaisePublicError "Error in smBlobWork.GetFullFileName, No Chapter Source found to update."
      CloseRecordset moRS
      Exit Function
   End If
      
   GetFullFilename = moRS!FullFilename
         
   CloseRecordset moRS
            
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwBlobWork.GetFullFileName, Err.Number, Err.Description"
   CloseRecordset moRS
End Function


Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   If loWork Is Nothing Then
      Set loWork = CreateObject("mwSession.mwReplicateWillChange")
      
      If Not loWork.Initialize("mwBlobFile") Then
         Set loWork = Nothing
         Exit Sub
      End If
   End If
   loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset

   Exit Sub
SubError:
   goSession.RaisePublicError "General error in mwBlobWork.moRS_WillChangeRecord. ", err.Number, err.Description
End Sub

