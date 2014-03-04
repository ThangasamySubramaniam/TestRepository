VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwWordWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' CloseDocument - Utility Class for manipulating Word documents
' 21 May 2003 ms

Option Explicit


' Microsoft Word BuiltinDocumentProperties constants...
Const mw_wdPropertyCategory = 18
Const mw_wdPropertyTitle = 1
Const mw_wdPropertyAuthor = 3
Const mw_wdPropertyTimeCreated = 11
Const mw_wdPropertyTimeLastSaved = 12
Const mw_wdPropertySubject = 2
Const mw_wdPropertyManager = 20
Const mw_wdFieldFormCheckBox = 71

'Public Enum mwWordProtectOffsets
'   mw_wdAllowOnlyRevisions = 0
'   mw_wdAllowOnlyComments = 1
'   mw_wdAllowOnlyFormFields = 2
'End Enum
   

' Word Window States
Const MW_wdWindowStateMaximize = 1
Const MW_wdWindowStateMinimize = 2
Const MW_wdWindowStateNormal = 0
Const MW_wdDoNotSaveChanges = 0
Const MW_wdSaveChanges = -1

Private moWord As Object
Private moDoc As Object
'Private moWord As Word.Application
'Private moDoc As Word.Document

Private mIsOpen As Boolean
Private mIsApplicationOpen As Boolean
Private mIsLeaveApplicationOpen As Boolean

'#If LATE_BIND Then
   Public Property Set WordApp(ByVal Data As Object)
       Set moWord = Data
   End Property
   
   Public Property Get WordApp() As Object
      Set WordApp = moWord
   End Property
   
   Public Property Set WordDoc(ByVal Data As Object)
       Set moDoc = Data
   End Property

   Public Property Get WordDoc() As Object
      Set WordDoc = moDoc
   End Property

'#Else
'   Public Property Set WordApp(ByVal Data As Word.Application)
'       Set moWord = Data
'   End Property
'
'   Public Property Set WordDoc(ByVal Data As Word.Document)
'       Set moDoc = Data
'   End Property
'#End If




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

'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property


Public Function CloseDocument(Optional DoNotSave As Boolean) As Boolean
   On Error GoTo FunctionError
   If mIsOpen Then
      If Not moDoc Is Nothing Then
         If DoNotSave Then
            moDoc.Close MW_wdDoNotSaveChanges
         Else
            moDoc.Close MW_wdSaveChanges
         End If
      End If
      If Not mIsLeaveApplicationOpen Then
         
         If DoNotSave Then
            moWord.Quit MW_wdDoNotSaveChanges
         Else
            moWord.Quit MW_wdSaveChanges
         End If
         KillObject moWord
         mIsApplicationOpen = False
      End If
   End If
   CloseDocument = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordWork.CloseDocument.", Err.Number, Err.Description
   CloseDocument = False
End Function


Public Function OpenDocument(FileName As String, Optional LeaveApplicationOpen As Boolean) As Boolean
   Dim fso As FileSystemObject
   On Error GoTo FunctionError
   If mIsOpen Then
      goSession.RaiseError "Error in mwWordWork.OpenDocument, Word Document already Open."
      OpenDocument = False
      Exit Function
   End If
   Set fso = New FileSystemObject
   
   '
   ' File exists ?
   '
   If Not fso.FileExists(FileName) Then
      goSession.RaiseError "Error in mwSession.mwWordWork.OpenDocument, File not found: " & FileName
      OpenDocument = False
      goSession.KillObject fso
      Exit Function
   End If
      
   If Not mIsApplicationOpen Then
      Set moWord = CreateObject("Word.Application")
      mIsApplicationOpen = True
      If LeaveApplicationOpen Then
         mIsLeaveApplicationOpen = True
      End If
   End If
   moWord.Documents.Open FileName
   Set moDoc = moWord.ActiveDocument
   mIsOpen = True
   OpenDocument = True
   goSession.KillObject fso
   Exit Function
FunctionError:
   If Err.Number = 429 Then
      goSession.RaiseError "Error in mwSession.mwWordWork.OpenDocument, could not open Word Application object."
   Else
      goSession.RaiseError "General Error in mwSession.mwWordWork.OpenDocument.", Err.Number, Err.Description
   End If
   goSession.KillObject fso
End Function

Public Function EditDocument() As Boolean
   On Error GoTo FunctionError
   '
   ' If form isn't open already, open it...
   '
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwWordWork.EditDocument, no document is open to edit."
      EditDocument = False
      Exit Function
   End If
   '
   ' Enter edit Mode in Word...
   '
   If goSession.User.StartOfficeMaximized Then
      moWord.WindowState = MW_wdWindowStateMaximize
   End If
   moWord.Visible = True
   EditDocument = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwSession.mwWordWork.EditDocument.", Err.Number, Err.Description
   EditDocument = False
End Function

Public Function SaveDocument() As Boolean
   On Error GoTo FunctionError
   If Not mIsOpen Then
      goSession.RaiseError "Error in mwWordWork.SaveDocument, no document is open to save."
      SaveDocument = False
      Exit Function
   End If
   
   moDoc.Save
   SaveDocument = True
   Exit Function
FunctionError:
   SaveDocument = False
End Function

Public Function TearDown(Optional LeaveAppOpen As Boolean)
   On Error GoTo FunctionError
   If Not moWord Is Nothing Then
      If Not LeaveAppOpen Then
         On Error Resume Next
         moWord.Quit MW_wdDoNotSaveChanges
      End If
   End If
   If Not moDoc Is Nothing Then
      Set moDoc = Nothing
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordWork.TearDown.", Err.Number, Err.Description

End Function

Public Function SetBookMark(BookmarkName As String, value As String) As Boolean
   On Error GoTo FunctionError
   moDoc.Bookmarks(BookmarkName).Range.Text = value
   SetBookMark = True
   Exit Function
FunctionError:
   SetBookMark = False
End Function

Public Function SetFormField(FieldName As String, value As String) As Boolean
   On Error GoTo FunctionError
   'MsgBox "debug code: Doc: " & moDoc.Name
   'MsgBox "form fields count: " & moDoc.formfields.Count
   If value <> "" Then
'      moDoc.FormFields(CVar(FieldName)).Result = value
      If moDoc.FormFields(CVar(FieldName)).Type = mw_wdFieldFormCheckBox Then
         If value = "1" Then
            moDoc.FormFields(CVar(FieldName)).CheckBox.value = True
         Else
            moDoc.FormFields(CVar(FieldName)).CheckBox.value = False
         End If
      Else
         moDoc.FormFields(CVar(FieldName)).Range.Fields(1).Result.Text = value
'         moDoc.FormFields(CVar(FieldName)).Result = Left(value, 250)
'         moDoc.FormFields(CVar(FieldName)).Result = value
         
'         If Len(value) < 255 Then
'            moDoc.FormFields(CVar(FieldName)).Result = value
'         Else
'            moDoc.FormFields(CVar(FieldName)).Result = "||||"
'
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            moWord.Selection.Collapse
'            moWord.Selection.MoveRight wdCharacter, 1
'            moWord.Selection.TypeText value
'            moWord.Selection.GoTo what:=wdGoToBookmark, Name:=FieldName
'            With moWord.Selection.Find
'               .Execute FindText:="|", replacewith:="", Replace:=wdReplaceAll
'            End With
'         End If
         
      End If
   End If
   
   SetFormField = True
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaiseError "Error in mwWordWork.SetFormField, Field Missing: " & FieldName
   Else
      goSession.RaiseError "General Error in mwWordWork.SetFormField, on field: " & FieldName, Err.Number, Err.Description
   End If
   SetFormField = False
End Function

'Private Sub WorkAround255Limit()
'      ' Set Text1 form field to a unique string.
'      ActiveDocument.FormFields("text1").Result = "****"
'      If ActiveDocument.ProtectionType <> wdNoProtection Then
'          ActiveDocument.Unprotect
'      End If
'      Selection.GoTo what:=wdGoToBookmark, Name:="Text1"
'      Selection.Collapse
'      Selection.MoveRight wdCharacter, 1
'      Selection.TypeText (String(256, "W"))
'      Selection.GoTo what:=wdGoToBookmark, Name:="Text1"
'      ' Remove unique characters from Text1 form field.
'      With Selection.Find
'         .Execute FindText:="*", replacewith:="", Replace:=wdReplaceAll
'      End With
'      ActiveDocument.Protect Password:="", NoReset:=True, Type:= _
'         wdAllowOnlyFormFields
'End Sub

Public Function GetBookMark(BookmarkName As String) As String
   On Error GoTo FunctionError
   GetBookMark = moDoc.Bookmarks(BookmarkName).Range.Text
   Exit Function
FunctionError:
   GetBookMark = ""
End Function

Public Function GetFormField(FieldName As String) As String
   On Error GoTo FunctionError
   GetFormField = moDoc.Fields(FieldName).Result
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaiseError "Error in mwWordWork.GetFormField, Field Missing: " & FieldName
   Else
      goSession.RaiseError "General Error in mwWordWork.GetFormField, on field: " & FieldName, Err.Number, Err.Description
   End If
   GetFormField = ""
End Function

Public Function OpenApplication() As Boolean
   On Error GoTo FunctionError
   If mIsApplicationOpen Then
      OpenApplication = True
      Exit Function
   End If
   Set moWord = CreateObject("Word.Application")
   mIsLeaveApplicationOpen = True
   mIsApplicationOpen = True
   OpenApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordWork.OpenApplication. ", Err.Number, Err.Description
   OpenApplication = False
End Function

Public Function CloseApplication() As Boolean
   On Error GoTo FunctionError
   If Not mIsApplicationOpen Then
      CloseApplication = True
      Exit Function
   End If
   
   If Not moDoc Is Nothing Then
      Set moDoc = Nothing
   End If
   If Not moWord Is Nothing Then
      moWord.Quit MW_wdDoNotSaveChanges
      Set moWord = Nothing
   End If
   mIsLeaveApplicationOpen = False
   mIsApplicationOpen = False
   CloseApplication = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordWork.CloseApplication. ", Err.Number, Err.Description
   CloseApplication = False

End Function

Private Sub Class_Terminate()
   On Error Resume Next
   If Not moDoc Is Nothing Then
      moDoc.Close MW_wdDoNotSaveChanges
      Set moDoc = Nothing
   End If
   If Not moWord Is Nothing Then
      moWord.Quit MW_wdDoNotSaveChanges
      Set moWord = Nothing
   End If
End Sub


Public Function PrintForm() As Boolean
   On Error GoTo FunctionError
   moDoc.PrintOut
   PrintForm = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwWordWork.PrintForm. ", Err.Number, Err.Description
   PrintForm = False

End Function

Public Function LoadPicture(FileName As String, Bookmark As String) As Boolean
   On Error GoTo FunctionError
    
      'If moDoc.InlineShapes.Count < 1 Then
      moDoc.InlineShapes.AddPicture FileName, True, True, moDoc.Bookmarks(Bookmark).Range
      'End If
   LoadPicture = True
   Exit Function
FunctionError:
   If Err.Number = 5941 Then
      goSession.RaisePublicError "Error in mwSession.mwWordWork.LoadPicture, Bookmark (" _
        & Bookmark & ") missing."
   Else
      goSession.RaisePublicError "General error in mwSession.mwWordWork.LoadPicture", Err.Number, Err.Description
   End If
   LoadPicture = False
End Function

Public Function ProtectDocument(ProtectionType As mwWordProtectOffsets, Password As String) As Boolean
   Dim strPassword As String
   Dim strPW As String
   On Error GoTo FunctionError
   strPW = Password
   moDoc.Protect ProtectionType, , strPW
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwWordWork.ProtectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwWordWork.ProtectWorkbook. ", Err.Number, Err.Description
   End If
End Function

Public Function UnProtectDocument(Password As String)
   Dim strPW As String
   Dim iType As Long
   On Error GoTo FunctionError
   strPW = Password
   moDoc.Unprotect strPW
   Exit Function
FunctionError:
   If Err.Number = 1004 Then
      goSession.RaiseError "Error in mwWordWork.ProtectWorkbook, password in configuration file is incorrect."
   Else
      goSession.RaiseError "General Error in mwWordWork.ProtectWorkbook. ", Err.Number, Err.Description
   End If
End Function




