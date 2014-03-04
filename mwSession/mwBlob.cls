VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwBlob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CreateDocumentBlob(FullFilename As String, BriefDescription As String) As Long
   Dim loStream As ADODB.Stream
   Dim loRsBlobFile As Recordset
   Dim fso As FileSystemObject
   
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   If Not fso.FileExists(FullFilename) Then
      goSession.RaiseError "General Error in mwBlob.CreateDocumentBlob, Blob File doesnt exist: ", err.Number, err.Description
      KillObject fso
      CreateDocumentBlob = -1
   End If
   
   Set loStream = New ADODB.Stream
   Set loRsBlobFile = New Recordset
   
   loRsBlobFile.CursorLocation = adUseClient
   loRsBlobFile.Open "smBlobFile", goSession.DBConnection, adOpenDynamic, adLockOptimistic, adCmdTable
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.LoadFromFile FullFilename
                      
      With loRsBlobFile
         .AddNew
         .Fields("ID").value = goSession.MakePK("smBlobFile")
         .Fields("FullFileName").value = FullFilename
         .Fields("BriefDescription").value = BriefDescription
         .Fields("FileBlob").value = loStream.Read
         .Update
         
         CreateDocumentBlob = loRsBlobFile.Fields("ID").value
      
      End With
   
      CloseRecordset loRsBlobFile
      KillObject fso
   
   Exit Function

FunctionError:
   goSession.RaiseError "General Error in mwBlob.CreateDocumentBlob, Blob File Record already exists: ", err.Number, err.Description
   CloseRecordset loRsBlobFile
   KillObject fso
   CreateDocumentBlob = -1
   
   
End Function


Public Function FetchDocumentBlob(BlobFileKey As Long, Optional IsSaveInOriginalFolder As Boolean, _
  Optional IsOverwrite As Boolean) As String
      
   Dim loStream As ADODB.Stream
   Dim loRsBlobFile As Recordset
   Dim fso As FileSystemObject
   Dim strSQL As String
   Dim strFullFileName As String
   Dim strFilename As String
   
   On Error GoTo FunctionError
   
   strSQL = "SELECT * FROM smBlobFile WHERE ID=" & BlobFileKey
   Set loRsBlobFile = New Recordset
   loRsBlobFile.CursorLocation = adUseClient
   loRsBlobFile.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   
   If loRsBlobFile.RecordCount < 1 Then
      goSession.RaiseError "Error in mwBlob.FetchDocumentBlob, no blobs to fetch."
      FetchDocumentBlob = Null
      CloseRecordset loRsBlobFile
      Exit Function
   End If
   
   strFullFileName = loRsBlobFile!FullFilename
   
   Set fso = New FileSystemObject
   Set loStream = New ADODB.Stream
      
   
   If IsSaveInOriginalFolder Then
      If Not fso.FolderExists(fso.GetParentFolderName(strFullFileName)) Then
         fso.CreateFolder (fso.GetParentFolderName(strFullFileName))
      Else
         If fso.FileExists(strFullFileName) And Not IsOverwrite Then
            goSession.RaiseError ""
            CloseRecordset loRsBlobFile
            KillObject fso
            FetchDocumentBlob = Null
            Exit Function
         End If
      End If
      strFilename = strFullFileName
   Else
      strFilename = goSession.Workflow.WorkingSessionFolder & "\" & fso.GetFileName(strFullFileName)
      If fso.FileExists(strFilename) And Not IsOverwrite Then
            goSession.RaiseError ""
            CloseRecordset loRsBlobFile
            KillObject fso
            FetchDocumentBlob = Null
            Exit Function
      End If
   End If
   
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.Write loRsBlobFile.Fields("FileBlob").value
   loStream.SaveToFile strFilename
         
   FetchDocumentBlob = strFilename  'Is this correct??
         
   Exit Function
   
FunctionError:
   goSession.RaiseError "General Error in mwBlob.FetchDocumentBlob", err.Number, err.Description
   CloseRecordset loRsBlobFile
   KillObject fso
   FetchDocumentBlob = Null
End Function




Public Function UpdateDocumentBlob(BlobFileKey As Long, FullFilename As String, Optional IsReplaceFullFilename) As Boolean

   Dim loStream As ADODB.Stream
   Dim loRsBlobFile As Recordset
   Dim fso As FileSystemObject
   Dim strSQL As String
      
   On Error GoTo FunctionError
   
   Set fso = New FileSystemObject
   strSQL = "SELECT * FROM smBlobFile WHERE ID=" & BlobFileKey
   Set loRsBlobFile = New Recordset
   loRsBlobFile.CursorLocation = adUseClient
   loRsBlobFile.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic
   
   If loRsBlobFile.RecordCount < 1 Then
      goSession.RaiseError "Error in mwBlob.FetchBlobRecord, no blobs exist."
      UpdateDocumentBlob = False
      CloseRecordset loRsBlobFile
      Exit Function
   End If
   If Not fso.FileExists(FullFilename) Then
      goSession.RaiseError "Error in mwBlob.FetchBlobRecord, no blobs exist."
      UpdateDocumentBlob = False
      CloseRecordset loRsBlobFile
      Exit Function
   End If
   
   Set loStream = New ADODB.Stream
      
   loStream.Type = adTypeBinary
   loStream.Open
   loStream.LoadFromFile FullFilename
                     
      With loRsBlobFile
         
         .Fields("FileBlob").value = loStream.Read
         If IsReplaceFullFilename Then
            .Fields("FullFileName").value = FullFilename
         End If
         .Update
         
         UpdateDocumentBlob = True
                 
      End With
            
      CloseRecordset loRsBlobFile
      KillObject fso
   
   Exit Function

FunctionError:
   goSession.RaiseError "General Error in mwBlob.UpdateDocumentBlob", err.Number, err.Description
   CloseRecordset loRsBlobFile
   KillObject fso
   UpdateDocumentBlob = False


End Function


Public Function DeleteDocumentBlob(BlobFileKey As Long) As Boolean

   Dim loRsBlobFile As Recordset
   Dim fso As FileSystemObject
   Dim strSQL As String
   
   On Error GoTo FunctionError
   
   strSQL = "SELECT * FROM smBlobFile WHERE ID=" & BlobFileKey
   Set loRsBlobFile = New Recordset
   loRsBlobFile.CursorLocation = adUseClient
   loRsBlobFile.Open strSQL, goSession.DBConnection, adOpenDynamic, adLockOptimistic

   If loRsBlobFile.RecordCount < 1 Then
      goSession.RaiseError "Error in mwBlob.FetchBlobRecord, no blobs exist."
      DeleteDocumentBlob = False
      CloseRecordset loRsBlobFile
      Exit Function
   End If
      
   goSession.DBConnection.Execute "delete from smBlobFile WHERE ID=" & BlobFileKey
      
   Exit Function
   
FunctionError:
   goSession.RaiseError "General Error in mwBlob.DeleteDocumentBlob", err.Number, err.Description
   CloseRecordset loRsBlobFile
   KillObject fso
   DeleteDocumentBlob = False


End Function

