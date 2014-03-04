VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mEncryptionEnabled As Boolean
Dim mEncryptionKey As String
Dim moEncrypt As New eneEncryption
Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"


Public Function EnableEncryption(ActivationPassword As String) As Boolean
' Allow encryption calls from external DLLs
   On Error GoTo FunctionError
   
   If ActivationPassword = ENCRYPT_PSWD Then
      mEncryptionEnabled = True
      CreateKey
   Else
      mEncryptionEnabled = False
   End If
   EnableEncryption = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.EnableEncryption. ", err.Number, err.Description
   EnableEncryption = False
End Function

Private Sub CreateKey()
   ' Maritime Workstation
   ' NoItAtSkRoW EmItIrAm
   Dim ProductName As String
   Dim ThisChar As String
   Dim xx As Integer

   On Error GoTo FunctionError
   
   ProductName = "Maritime Workstation"
   mEncryptionKey = ""
   
   For xx = Len(ProductName) To 1 Step -1
      ThisChar = mID(ProductName, xx, 1)
      If xx Mod 2 = 1 Then
         ThisChar = UCase(ThisChar)
      End If
      
      mEncryptionKey = mEncryptionKey & ThisChar
      
   Next xx
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.CreateKey. ", err.Number, err.Description

End Sub

Public Function ChangePassword() As Boolean
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   ChangePassword = ChangePasswordX
End Function

Public Function ChangePasswordX(Optional ByVal blnForceChangeNoCancel As Boolean) As Boolean
' Display the Change Password form
   Dim loform As frmPassword
   
   On Error GoTo FunctionError
   
   Set loform = New frmPassword
   'By N.Angelakis On 22 April 2009
   'DEV-1174 Advance Password Settings
   If blnForceChangeNoCancel = True Then
      loform.ForcePasswordChange True
   End If
   
   loform.Show vbModal
   
   If loform.IsCanceled Then
      ChangePasswordX = False
   Else
      ChangePasswordX = True '
   End If
   
   Unload loform
   KillObject loform
   
   DoEvents
   Exit Function

FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.ChangePasswordX. ", err.Number, err.Description
   KillObject loform
   ChangePasswordX = False
End Function

Public Function EncryptString(InputString As String) As String
   On Error GoTo FunctionError

   If mEncryptionEnabled Then
      EncryptString = moEncrypt.Encryption(InputString, mEncryptionKey, encStrong)
   Else
      goSession.RaisePublicError "Encryption functions not enabled"
      EncryptString = ""
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.EncryptString. ", err.Number, err.Description
   EncryptString = ""

End Function

Public Function DecryptString(InputString As String) As String
' Decrypt a string
   On Error GoTo FunctionError
   
   If mEncryptionEnabled Then
      DecryptString = moEncrypt.Decryption(InputString, mEncryptionKey, encStrong)
   Else
      goSession.RaisePublicError "Encryption functions not enabled"
      DecryptString = ""
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.DecryptString. ", err.Number, err.Description
   DecryptString = ""

End Function

Public Function EncryptINI(INIFileName As String) As Boolean
' Encrypt the INI file

   Dim fso As FileSystemObject
   Dim OutStream As TextStream
   Dim InStream As TextStream
   Dim strOutFileName As String
   Dim WorkString As String
   Dim FirstRecord As Boolean
   Dim strTemp As String
   
   On Error GoTo FunctionError
   EncryptINI = True
   
   Set fso = New FileSystemObject
   
   strOutFileName = INIFileName & ".TEMP"
   
   If fso.FileExists(INIFileName) Then
   
      Set InStream = fso.OpenTextFile(INIFileName, ForReading)
      Set OutStream = fso.CreateTextFile(strOutFileName, True)
      
      FirstRecord = True
      
      While Not InStream.AtEndOfStream
         WorkString = InStream.ReadLine
         
         ' Make sure the file isn't encrypted already.
         
         If FirstRecord = True Then
            FirstRecord = False
            strTemp = moEncrypt.Decryption(WorkString, mEncryptionKey, encStrong)
            If strTemp = "ENCRYPTED" Then
               InStream.Close
               OutStream.Close
               If fso.FileExists(strOutFileName) Then
                  fso.DeleteFile strOutFileName, True
               End If
               Exit Function
            Else
               OutStream.WriteLine (moEncrypt.Encryption("ENCRYPTED", mEncryptionKey, encStrong))
            End If
         End If
         If Len(WorkString) > 0 Then
            OutStream.WriteLine (moEncrypt.Encryption(WorkString, mEncryptionKey, encStrong))
         Else
            OutStream.WriteLine WorkString
         End If
      Wend
      
      InStream.Close
      OutStream.Close
      
      If fso.FileExists(INIFileName & ".SAVE") Then
         fso.DeleteFile INIFileName & ".SAVE", True
      End If
      
      fso.MoveFile INIFileName, INIFileName & ".SAVE"
'      fso.DeleteFile INIFileName, True
      fso.MoveFile strOutFileName, INIFileName
      fso.DeleteFile INIFileName & ".SAVE", True
'      fso.DeleteFile strOutFileName, True
   Else
      goSession.RaisePublicError INIFileName & " Does not exist."
      EncryptINI = False
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.EncryptINI. ", err.Number, err.Description
   EncryptINI = False

End Function

Public Function DecryptINI(INIFileName As String) As Boolean
' Decrypt the INI file
   Dim fso As FileSystemObject
   Dim OutStream As TextStream
   Dim InStream As TextStream
   Dim strOutFileName As String
   Dim WorkString As String

   On Error GoTo FunctionError
   
   DecryptINI = True
   
   Set fso = New FileSystemObject
   
   strOutFileName = INIFileName & ".TEMP"
   
   If fso.FileExists(INIFileName) Then
   
      Set InStream = fso.OpenTextFile(INIFileName, ForReading)
      Set OutStream = fso.CreateTextFile(strOutFileName, True)
      
      WorkString = InStream.ReadLine
      If Len(WorkString) > 0 Then
         WorkString = moEncrypt.Decryption(WorkString, mEncryptionKey, encStrong)
      End If
      
      If WorkString = "ENCRYPTED" Then
         While Not InStream.AtEndOfStream
            WorkString = InStream.ReadLine
            If Len(WorkString) > 0 Then
               OutStream.WriteLine (moEncrypt.Decryption(WorkString, mEncryptionKey, encStrong))
            Else
               OutStream.WriteLine WorkString
            End If
         Wend
         
         InStream.Close
         OutStream.Close
         
         If fso.FileExists(INIFileName & ".SAVE") Then
            fso.DeleteFile INIFileName & ".SAVE", True
         End If
         
         fso.MoveFile INIFileName, INIFileName & ".SAVE"
'         fso.DeleteFile INIFileName, True
         fso.MoveFile strOutFileName, INIFileName
         fso.DeleteFile INIFileName & ".SAVE", True
'         fso.DeleteFile strOutFileName, True
      Else
         goSession.RaisePublicError INIFileName & " Is not encrypted."
         DecryptINI = False
      End If
   Else
      goSession.RaisePublicError INIFileName & " Does not exist."
      DecryptINI = False
   End If
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEncrypt.DecryptINI. ", err.Number, err.Description
   DecryptINI = False

End Function
