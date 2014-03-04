VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwDataWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' mwDataWork - Support ADO operations against database connections listed in mwcDataConnection


Option Explicit

Private moCon As Connection
Private moShapeCon As Connection
Private mIsConnectionOpen As Boolean
Private mIsRecordsetOpen As Boolean
Private mIsUpdateAllowed As Boolean
Private mLeaveOpenOnDestroy As Boolean
Private mShapeConnectString As String
Private mDgGeneralDateFormat As String
Private mDgShortDateFormat As String
Private mDgDecimalSeparator As String
Private mEncryptionPswd As String


Private Const ENCRYPT_PSWD = "Gray" & "bar" & "327"

Private moRS As Recordset

Friend Property Set Parent(oParent As mwSession.Session)
End Property

Friend Property Get Parent() As mwSession.Session
   'If goSession Is Nothing Then
   '   Set goSession = oParent
   'End If
End Property

Public Property Get DBConnection() As Connection
   If moCon Is Nothing Then
      goSession.RaiseError "Error in mwDataWork.DbConnection_Get, Connection not open."
   Else
      Set DBConnection = moCon
   End If
End Property


Public Property Get DgGeneralDateFormat() As String
    DgGeneralDateFormat = mDgGeneralDateFormat
End Property

Public Property Get DgShortDateFormat() As String
    DgShortDateFormat = mDgShortDateFormat
End Property

Public Property Get DgDecimalSeparator() As String
    DgDecimalSeparator = mDgDecimalSeparator
End Property


'Public Property Get A() As String
'    A = mA
'End Property

'Public Property Let A(ByVal vData As String)
'    mA = vData
'End Property



' Object Template
'Public Property Get A() As Object
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set A = moA
'End Property

'Public Property Set A(ByVal vData As Object)
'   If moA Is Nothing Then
'      Set moA = New aa
'   End If
'   Set moA = vData
'End Property

Public Property Get IsConnectionOpen() As Boolean
    IsConnectionOpen = mIsConnectionOpen
End Property

Public Property Get IsRecordsetOpen() As Boolean
    IsRecordsetOpen = mIsRecordsetOpen
End Property

Public Property Get IsUpdateAllowed() As Boolean
    IsUpdateAllowed = mIsUpdateAllowed
End Property

Public Property Let IsUpdateAllowed(ByVal vData As Boolean)
    mIsUpdateAllowed = vData
End Property

Public Property Get LeaveOpenOnDestroy() As Boolean
    LeaveOpenOnDestroy = mLeaveOpenOnDestroy
End Property

Public Property Let LeaveOpenOnDestroy(ByVal vData As Boolean)
    mLeaveOpenOnDestroy = vData
End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property

'Public Property Get A() As string
'    A = mA
'End Property

'Public Property Let A(ByVal vData As string)
'    mA = vData
'End Property
Public Sub SetEncryptionPswd(Pswd As String)
   mEncryptionPswd = Pswd
End Sub

Public Function OpenConnection(ConnectionID As String) As Connection
   Dim loRsDB As Recordset
   Dim sSQL As String
   Dim strConnectString As String
   Dim loEncrypt As New mwEncrypt
   
   On Error GoTo FunctionError
   If Not moCon Is Nothing Then
      goSession.RaiseError "Error in mwDataWork.OpenConnection, Connection already open: " & ConnectionID
      Exit Function
   End If
   Set loRsDB = New Recordset
   loRsDB.CursorLocation = adUseClient
   sSQL = "select DbConnectString, ShapeConnectString, Encrypted from mwcDbConnection where ConnectionID='" & ConnectionID & "'"
   loRsDB.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsDB.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataWork.OpenConnection, Connection record not found: " & ConnectionID
      goSession.CloseRecordset loRsDB
   End If
   
   If IsNull(loRsDB!DBConnectString) Then
      strConnectString = goSession.GetDecryptedDBConnectString(ENCRYPT_PSWD)
      mShapeConnectString = goSession.GetDecryptedDbShapeConnectString(ENCRYPT_PSWD)
   Else
      If loRsDB!Encrypted = -1 Then
         loEncrypt.EnableEncryption (ENCRYPT_PSWD)
         strConnectString = loEncrypt.DecryptString(loRsDB!DBConnectString)
      Else
         strConnectString = loRsDB!DBConnectString
      End If
      
      If Not IsNull(loRsDB!ShapeConnectString) Then
         If loRsDB!Encrypted = -1 Then
            mShapeConnectString = loEncrypt.DecryptString(loRsDB!ShapeConnectString)
         Else
            mShapeConnectString = loRsDB!ShapeConnectString
         End If
      End If
   End If
   goSession.CloseRecordset loRsDB
   '
   ' Open Connection
   '
   Set moCon = New Connection
   moCon.Open strConnectString
   Set OpenConnection = moCon
   mIsConnectionOpen = True
   
   If IsDBSqlServer(moCon) Then
      goCon.CommandTimeout = 3600
      goSession.LogIt mwl_Workstation, mwl_Information, "SQL Server database connection CommandTimeout set to 1 hour on connection " & ConnectionID
   End If
   
   '
   ' Open Shape Connection, if specified
   '
   If Trim(mShapeConnectString) <> "" Then
      '
      Set moShapeCon = New ADODB.Connection
      moShapeCon.Provider = "MSDataShape"
      moShapeCon.Open mShapeConnectString
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataWork.OpenConnection. ", Err.Number, Err.Description
   goSession.CloseConnection moCon
End Function

Public Function GetConnection(ConnectionID As String) As Connection
   Set GetConnection = OpenConnection(ConnectionID)
End Function


Public Function OpenRecordSet(SqlString As String, Optional IsOpenAsTable As Boolean) As Recordset
   On Error GoTo FunctionError
   If moCon Is Nothing Then
      goSession.RaiseError "Error in mwDataWork.OpenRecordset, Connection not open."
      Exit Function
   End If
   If Not moRS Is Nothing Then
      goSession.RaiseError "Error in mwDataWork.OpenRecordset, record set already open."
      Exit Function
   End If
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   If IsOpenAsTable Then
      If mIsUpdateAllowed Then
         moRS.Open SqlString, moCon, adOpenDynamic, adLockOptimistic, adCmdTable
      Else
         moRS.Open SqlString, moCon, adOpenForwardOnly, adLockReadOnly, adCmdTable
      End If
   Else
      If mIsUpdateAllowed Then
         moRS.Open SqlString, moCon, adOpenDynamic, adLockOptimistic
      Else
         moRS.Open SqlString, moCon, adOpenForwardOnly, adLockReadOnly
      End If
   End If
   'If moRS.RecordCount > 0 Then
   Set OpenRecordSet = moRS
   mIsRecordsetOpen = True
   'Else
   '   goSession.RaiseWorkflowStatus "mwDataWork.OpenRecordset, record set is empty. SQL: " & SqlString
   '   goSession.CloseRecordset moRS
   '   mIsRecordsetOpen = False
   'End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataWork.OpenRecordSet. ", Err.Number, Err.Description
   goSession.CloseRecordset moRS
   mIsRecordsetOpen = False
End Function

Public Function SetRecordsetFilter(FilterString As String) As Boolean
   On Error GoTo FunctionError
   
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataWork.SetRecordsetFilter. ", Err.Number, Err.Description
   goSession.CloseRecordset moRS
   
End Function

Private Sub Class_Terminate()
   On Error GoTo SubError
   If Not LeaveOpenOnDestroy Then
      ' moRs
      If Not moRS Is Nothing Then
         If moRS.State = adStateOpen Then
            If moRS.LockType <> adLockReadOnly Then
               If Not (moRS.EOF Or moRS.BOF) Then
                  'moRs.Update
                  moRS.Move (0)
               End If
            End If
            moRS.Close
         End If
         Set moRS = Nothing
      End If
      ' moCon
      If Not moCon Is Nothing Then
         If moCon.State = adStateOpen Then
            moCon.Close
         End If
         Set moCon = Nothing
      End If
      ' moShapeCon
      If Not moShapeCon Is Nothing Then
         If moShapeCon.State = adStateOpen Then
            moShapeCon.Close
         End If
         Set moShapeCon = Nothing
      End If
   End If
   Exit Sub
SubError:
   goSession.RaiseError "General Error in mwDataWork.Class_Terminate. ", Err.Number, Err.Description
End Sub

Public Function GetField(FieldName) As String
   On Error GoTo FunctionError
   If moRS.RecordCount < 1 Then
      GetField = ""
   ElseIf IsNull(moRS.Fields(FieldName).value) Then
      GetField = ""
   Else
      GetField = moRS.Fields(FieldName).value
   End If
   Exit Function

FunctionError:
   If Err.Number = 3265 Then
      goSession.RaiseError "Error in mwDataWork.GetField, Column Missing: " & FieldName
   Else
      goSession.RaiseError "General Error in mwDataWork.GetField: " & FieldName, Err.Number, Err.Description
   End If
   GetField = ""
End Function

Public Function CloseRecordset() As Boolean
   If mIsRecordsetOpen Then
      goSession.CloseRecordset moRS
   End If
   CloseRecordset = True
   mIsRecordsetOpen = False
End Function

Public Function GetShapeConnection() As Connection
   ' Removed by T. McSally, 08/04/2004
   goSession.RaiseError "mwDataWork.GetShapeConnection is a deprecated function."
   GetShapeConnection = False

'   If Not moShapeCon Is Nothing Then
'      Set GetShapeConnection = moShapeCon
'   End If
End Function

Public Function GetDbConnectionStrings(ConnectionID As String, ByRef DbString As String, ByRef DbShapeString As String) As Boolean

   ' Removed by T. McSally, 08/04/2004
   goSession.RaiseError "mwDataWork.GetDbConnectionStrings is a deprecated function."
   GetDbConnectionStrings = False
End Function


Public Function DebugShowRsProperties(ByRef loRs As Recordset) As String
   On Error GoTo FunctionError
   Dim i As Integer
   Dim s As String
   On Error GoTo FunctionError
   For i = 0 To loRs.Properties.Count - 1
      s = s & loRs.Properties(i).Name & " = " & loRs.Properties(i).value & vbCrLf
   Next i
   DebugShowRsProperties = s
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwDataWork.DebugShowRsProperties. " & Err.Number & Err.Description
   DebugShowRsProperties = s
End Function

Public Function DebugShowFieldTypeProperties(ByRef loRs As Recordset) As String
   On Error GoTo FunctionError
   Dim loField As field
   Dim i As Integer
   Dim s As String
   On Error GoTo FunctionError
   For Each loField In loRs.Fields
      s = s & loField.Name & " = Type: " & loField.Type & vbCrLf
   Next loField
   DebugShowFieldTypeProperties = s
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwDataWork.DebugShowFieldTypeProperties. " & Err.Number & Err.Description
   DebugShowFieldTypeProperties = s
End Function


Public Function DebugShowRsFields(ByRef loRs As Recordset) As String
   On Error GoTo FunctionError
   Dim i As Integer
   Dim s As String
   On Error GoTo FunctionError
   For i = 0 To loRs.Fields.Count - 1
      If IsNull(loRs.Fields(i).value) Then
         s = s & loRs.Fields(i).Name & " = <null>" & vbCrLf
      Else
         s = s & loRs.Fields(i).Name & " = " & loRs.Fields(i).value & vbCrLf
      End If
   Next i
   DebugShowRsFields = s
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwDataWork.DebugShowRsFields. " & Err.Number & Err.Description
   DebugShowRsFields = s
End Function


Public Function MergeRecordsets(ByRef loRsSource1 As Recordset, ByRef loRsSource2 As Recordset) As Recordset
   Dim loRs As Recordset
   Dim loSkipList As Collection
   Dim loField As field
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Staging/Validation
   If loRsSource1.State = adStateClosed Then
      goSession.RaiseError "Error in mwDataWork.MergeRecordsets, Recordset is closed."
      Exit Function
   End If
   ' Any records in there ?
   If loRsSource1.RecordCount < 1 Or loRsSource2.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataWork.MergeRecordsets, no records in Recordset."
      Exit Function
   End If
   '
   ' Create new recordset...
   '
   Set loSkipList = New Collection
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ' Add columns to fabricated recordset
   For Each loField In loRsSource1.Fields
      If goSession.IsOracle And loField.Type = adNumeric Then
         loRs.Fields.Append loField.Name, adInteger, 4
      Else
         loRs.Fields.Append loField.Name, loField.Type, loField.DefinedSize
      End If
   Next loField
   ' Add expanded columns to recordset...
   On Error Resume Next
   For Each loField In loRsSource2.Fields
      If goSession.IsOracle And loField.Type = adNumeric Then
         loRs.Fields.Append loField.Name, adInteger, 4
      Else
         loRs.Fields.Append loField.Name, loField.Type, loField.DefinedSize
      End If
      If Err Then
         ' Error 3367-already in collection... 3001-Shape RS pointer to child RS, invalid...
         If Err.Number = 3367 Or Err.Number = 3001 Then
            loSkipList.Add Item:=loField.Name, Key:=loField.Name
         Else
            goSession.RaiseError "General Error in mwDataWork.MergeRecordsets. ", Err.Number, Err.Description
            goSession.CloseRecordset loRs
            Set MergeRecordsets = loRsSource1
         End If
         Err.Clear
      End If
   Next loField
   On Error GoTo FunctionError
   '
   ' Append data to fabricated recordset...
   '
   loRs.Open
   loRs.AddNew
   ' Event Detail...
   For Each loField In loRsSource1.Fields
      With loField
         If Not IsNull(.value) And Not IsEmpty(.value) Then
            loRs(.Name).value = .value
         End If
      End With
   Next loField
   '
   ' Add Activity record fields to flat fabricated recordset
   '
   For Each loField In loRsSource2.Fields
      With loField
         If Not IsNull(.value) And Not IsEmpty(.value) And Not IsInCollection(loSkipList, .Name) Then
            loRs(.Name).value = .value
         End If
      End With
   Next loField
   ' That's all folks...
   Set MergeRecordsets = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaisePublicError "General Error in mwDataWork.MergeRecordsets. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   ' revert to original recordset
   Set MergeRecordsets = loRsSource1
End Function


Public Function ConvertRecordToInsertStatement(ByRef loRs As Recordset, TableName As String) As String
   ' Note, assumes recordset pointer is on record to convert...
   Dim strFields As String
   Dim strData As String
   Dim sSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   strFields = "Insert into " & TableName & " ("
   strData = " values ("
   For Each loField In loRs.Fields
      If (Not IsNull(loField.value)) Then
         ' concatenate the field name...
         strFields = strFields & loField.Name & ", "
         ' concatenate field value
         Select Case loField.Type
            Case Is = adInteger, adDouble, adNumeric
               strData = strData & loField.value & ", "
               
            Case Else
               strData = strData & "'" & loField.value & "', "
         End Select
      End If
   Next loField
   ' strip last "," and put it together...
   strFields = Left(strFields, Len(strFields) - 2)
   strData = Left(strData, Len(strData) - 2)
   ConvertRecordToInsertStatement = strFields & ") " & strData & ")"
   Exit Function
FunctionError:
   goSession.RaiseWorkflowStatus "General Error in mwDataForm.ConvertRecordToInsertStatement.", Err.Number, Err.Description
End Function


Public Function GetCrystalLogin(ConnectionID As String, ByRef CrystalServerName As String, _
  ByRef CrystalDatabaseName As String, CrystalUser As String, CrystalPW As String) As Boolean

'   Dim loRsDB As Recordset
'   Dim sSQL As String
'   Dim strConnectString As String
'   On Error GoTo FunctionError
'   If Not moCon Is Nothing Then
'      goSession.RaiseError "Error in mwDataWork.OpenConnection, Connection already open: " & ConnectionID
'      Exit Function
'   End If
'   Set loRsDB = New Recordset
'   loRsDB.CursorLocation = adUseClient
'   sSQL = "select * from mwcDbConnection where ConnectionID='" & ConnectionID & "'"
'   loRsDB.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
'   If loRsDB.RecordCount > 0 Then
'      CrystalServerName = loRsDB!CrystalServer & ""
'      CrystalDatabaseName = loRsDB!CrystalDatabase & ""
'      CrystalUser = loRsDB!CrystalUserID & ""
'      CrystalPW = loRsDB!CrystalPassword & ""
'      GetCrystalLogin = True
'   Else
'      goSession.RaiseError "Error in mwDataWork.OpenConnection, Connection record not found: " & ConnectionID
'      GetCrystalLogin = False
'   End If
'   goSession.CloseRecordset loRsDB
'   Exit Function
'FunctionError:
'   goSession.RaiseError "General error in mwDataWork.GetCrystalLogin. " & err.Number & err.Description
   goSession.RaiseError "mwDataWork.GetCrystalLogin is a deprecated function."
   GetCrystalLogin = False

End Function

Public Function ConvertToSQLSyntax(SqlStatement As String) As String
   Dim strArray() As String
   
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   'comment out previous declarations below no longer used
   'Added By N.Angelakis On 06 Mar 2009
   'DEV-252 Changes required for schema Access referenced date statements
   'Dim strDateConvertedQuery As String
   'Dim arrDateReferences() As String
   'Dim intCount As Integer
   
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   Dim nDateStartPos As Long
   Dim nDateEndPos As Long
   Dim sDatePart As String
   Dim sDateFormat As String
   
   On Error GoTo FunctionError
   
   SqlStatement = Trim(SqlStatement) & " "
   
   ' Convert multiple blanks to single blanks
   SqlStatement = Replace(SqlStatement, "    ", " ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "   ", " ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "  ", " ", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " autoincrement ", " int IDENTITY (1,1) ", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " long ", " int ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long;", " int;", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long,", " int,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long)", " int)", , , vbTextCompare)
'   If UCase(Right(SqlStatement, 4)) = "LONG" Then
'      SqlStatement = Replace(SqlStatement, " long", " int", , , vbTextCompare)
'   End If

   SqlStatement = Replace(SqlStatement, " double ", " float ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " double;", " float;", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " double,", " float,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " double)", " float)", , , vbTextCompare)
'   If UCase(Right(SqlStatement, 6)) = "DOUBLE" Then
'      SqlStatement = Replace(SqlStatement, " double", " float", , , vbTextCompare)
'   End If
   
   SqlStatement = Replace(SqlStatement, " memo ", " ntext ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " memo;", " ntext;", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " memo,", " ntext,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " memo)", " ntext)", , , vbTextCompare)
'   If UCase(Right(SqlStatement, 4)) = "MEMO" Then
'      SqlStatement = Replace(SqlStatement, " memo", " ntext", , , vbTextCompare)
'   End If
   SqlStatement = Replace(SqlStatement, " LONGBINARY ", " varbinary ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " LONGBINARY;", " varbinary;", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " LONGBINARY,", " varbinary,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " LONGBINARY)", " varbinary)", , , vbTextCompare)
'   If UCase(Right(SqlStatement, 10)) = "LONGBINARY" Then
'      SqlStatement = Replace(SqlStatement, " LONGBINARY", " varbinary", , , vbTextCompare)
'   End If

   If InStr(1, SqlStatement, "ALTER TABLE ", vbTextCompare) > 0 Then
      SqlStatement = Replace(SqlStatement, " ADD COLUMN ", " ADD ", , , vbTextCompare)
      
   ElseIf InStr(1, SqlStatement, "DROP INDEX ", vbTextCompare) > 0 Then
      strArray = Split(SqlStatement, " ", , vbTextCompare)
      
      SqlStatement = "DROP INDEX " & strArray(4) & "." & strArray(2)
   End If
   
   
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   'comment out code below, no longer used
   'Added By N.Angelakis On 06 Mar 2009
   'DEV-252 Changes required for schema Access referenced date statements
'   If InStr(SqlStatement, "#") > 0 Then
'      arrDateReferences = Split(SqlStatement, "#", , vbTextCompare)
'      If UBound(arrDateReferences) > 0 Then
'         For intCount = 0 To UBound(arrDateReferences)
'            If IsDate(arrDateReferences(intCount)) Then
'               If Hour(arrDateReferences(intCount)) = 0 And Minute(arrDateReferences(intCount)) = 0 And Second(arrDateReferences(intCount)) = 0 Then
'                     'syntax used as declared in mwUtility.modUtility.DateLiteral
'                     arrDateReferences(intCount) = "CONVERT(DATETIME, '" & Format(arrDateReferences(intCount), "yyyy.mm.dd") & "', 102)"
'                  Else
'                     arrDateReferences(intCount) = "CONVERT(DATETIME, '" & Format(arrDateReferences(intCount), "yyyy-mm-dd hh:mm:ss") & "', 20)"
'               End If
'            End If
'            'Build query addind modified date syntax where necessary
'            strDateConvertedQuery = strDateConvertedQuery & " " & arrDateReferences(intCount)
'         Next
'         If strDateConvertedQuery <> "" Then
'             SqlStatement = strDateConvertedQuery
'         End If
'      End If
'      If UBound(arrDateReferences) > 0 Then Erase arrDateReferences
'   End If

   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   nDateStartPos = InStr(SqlStatement, "#")
   nDateEndPos = InStr(nDateStartPos + 1, SqlStatement, "#")
   Do While nDateStartPos > 0 And nDateEndPos > 0
      'reset variables
      sDatePart = ""
      sDateFormat = ""
      
      sDatePart = mID(SqlStatement, nDateStartPos + 1, InStr(nDateStartPos + 1, SqlStatement, "#") - nDateStartPos - 1)
      If IsDate(sDatePart) Then
         If Hour(sDatePart) = 0 And Minute(sDatePart) = 0 And Second(sDatePart) = 0 Then
            'syntax used as declared in mwUtility.modUtility.DateLiteral
            sDateFormat = "CONVERT(DATETIME, '" & Format(sDatePart, "yyyy.mm.dd") & "', 102)"
         Else
            sDateFormat = "CONVERT(DATETIME, '" & Format(sDatePart, "yyyy-mm-dd hh:mm:ss") & "', 20)"
         End If
         SqlStatement = Replace(SqlStatement, "#" + sDatePart + "#", sDateFormat, , vbTextCompare)
      End If
      nDateStartPos = InStr(nDateStartPos + 1, SqlStatement, "#")
      nDateEndPos = InStr(nDateStartPos + 1, SqlStatement, "#")
   Loop
   
   
   ConvertToSQLSyntax = SqlStatement
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.ConvertToSQLSyntax.", Err.Number, Err.Description
   ConvertToSQLSyntax = SqlStatement
End Function
Public Function ConvertToOracleSyntax(SqlStatement As String, ByRef SqlStatement2 As String, ByRef SqlStatement3 As String) As String
   On Error GoTo FunctionError
   Dim GotAutoInc As Long
   Dim SplitRslt() As String
   Dim SeqName As String
   Dim TriggerName As String
   Dim TableName As String
   Dim xx As Integer
   
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   'comment out previous declarations below no longer used
   'Added By N.Angelakis On 06 Mar 2009
   'DEV-252 Changes required for schema Access referenced date statements
'   Dim strDateConvertedQuery As String
'   Dim arrDateReferences() As String
'   Dim intCount As Integer
      
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   Dim nDateStartPos As Long
   Dim nDateEndPos As Long
   Dim sDatePart As String
   Dim sDateFormat As String
   
      
   SqlStatement2 = ""
   SqlStatement3 = ""
   
   SqlStatement = Trim(SqlStatement) & " "
   
   ' Convert multiple blanks to single blanks
   SqlStatement = Replace(SqlStatement, "    ", " ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "   ", " ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "  ", " ", , , vbTextCompare)
   
'      SqlStatement = Replace(SqlStatement, "autoincrement", " int IDENTITY (1,1)", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long ", " NUMBER(10) ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long,", " NUMBER(10),", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " long)", " NUMBER(10))", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " integer ", " NUMBER(10) ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " integer,", " NUMBER(10),", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " integer)", " NUMBER(10))", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " int ", " NUMBER(10) ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " int,", " NUMBER(10),", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " int)", " NUMBER(10))", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " smallint ", " NUMBER(6) ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " smallint,", " NUMBER(6),", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " smallint)", " NUMBER(6))", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " varchar ", " VARCHAR2 ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " varchar(", " VARCHAR2(", , , vbTextCompare)
        
   SqlStatement = Replace(SqlStatement, " double ", " NUMBER ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " double,", " NUMBER,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " double)", " NUMBER)", , , vbTextCompare)
    
   SqlStatement = Replace(SqlStatement, " float ", " NUMBER ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " float,", " NUMBER,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " float)", " NUMBER)", , , vbTextCompare)
    
'   If Right(SqlStatement, 5) = " memo" Then
'      SqlStatement = Replace(SqlStatement, " memo", " CLOB", , , vbTextCompare)
'   End If
    
   SqlStatement = Replace(SqlStatement, " memo ", " CLOB ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " memo,", " CLOB,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " memo)", " CLOB)", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " ntext ", " NCLOB ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " ntext,", " NCLOB,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " ntext)", " NCLOB)", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " datetime ", " date ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " datetime,", " date,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " datetime)", " date)", , , vbTextCompare)
    
   SqlStatement = Replace(SqlStatement, " LONGBINARY ", " BLOB ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " LONGBINARY,", " BLOB,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " LONGBINARY)", " BLOB)", , , vbTextCompare)
    
   SqlStatement = Replace(SqlStatement, " bit" & vbCrLf, " NUMBER(1)" & vbCrLf, , , vbBinaryCompare)
   If Right(SqlStatement, 4) = " bit" Then
      SqlStatement = Replace(SqlStatement, " bit", " NUMBER(1)", , , vbTextCompare)
   End If
    
   SqlStatement = Replace(SqlStatement, " bit ", " NUMBER(1) ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " bit,", " NUMBER(1),", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " bit)", " NUMBER(1))", , , vbTextCompare)
    
   SqlStatement = Replace(SqlStatement, "[ZONE]", " ZONE ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "[SECTION]", " SECTION ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, "[Code]", " Code ", , , vbTextCompare)
    
   SqlStatement = Replace(SqlStatement, " default true ", " default -1 ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " default true,", " default -1,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " default true)", " default -1)", , , vbTextCompare)
   
   SqlStatement = Replace(SqlStatement, " default false ", " default 0 ", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " default false,", " default 0,", , , vbTextCompare)
   SqlStatement = Replace(SqlStatement, " default false)", " default 0)", , , vbTextCompare)
    
    ' Replace AUTOINCREMENT with a combination of a SEQUENCE and a TRIGGER
    '
    '   CREATE SEQUENCE tablename_SEQUENCE
    '                 START WITH 1 INCREMENT BY 1
    '                 NOMAXVALUE NOCACHE ORDER
    '
    '   CREATE OR REPLACE TRIGGER tablename_AUTOINC
    '                 BEFORE INSERT ON tablename
    '                 FOR EACH ROW
    '   BEGIN
    '         IF :new.ID IS NULL THEN
    '              SELECT tablename_SEQUENCE.NextVal INTO :new.ID FROM DUAL;
    '         ENDIF;
    '   END;
    
   If InStr(1, SqlStatement, "alter table", vbTextCompare) > 0 Then
      If InStr(1, SqlStatement, " add column ", vbTextCompare) > 0 Then
         SqlStatement = Replace(SqlStatement, " add column ", " add ( ", , , vbTextCompare)
         SqlStatement = SqlStatement & ")"
      ElseIf InStr(1, SqlStatement, " add ", vbTextCompare) > 0 Then
         SqlStatement = Replace(SqlStatement, " add ", " add ( ", , , vbTextCompare)
         SqlStatement = SqlStatement & ")"
      End If
       
      If InStr(1, SqlStatement, " drop column ", vbTextCompare) > 0 Then
   '     SqlStatement = Replace(SqlStatement, " add ", " add ( ", , , vbTextCompare)
   '     SqlStatement = SqlStatement & ")"
      ElseIf InStr(1, SqlStatement, " drop constraint ", vbTextCompare) > 0 Then
   '     SqlStatement = Replace(SqlStatement, " add ", " add ( ", , , vbTextCompare)
   '     SqlStatement = SqlStatement & ")"
      ElseIf InStr(1, SqlStatement, " drop ", vbTextCompare) > 0 Then
         SqlStatement = Replace(SqlStatement, " drop ", " drop ( ", , , vbTextCompare)
         SqlStatement = SqlStatement & ")"
      ElseIf InStr(1, SqlStatement, " alter column ", vbTextCompare) > 0 Then
         SqlStatement = Replace(SqlStatement, " alter column ", " modify ( ", , , vbTextCompare)
         SqlStatement = SqlStatement & ")"
      End If
   ElseIf InStr(1, SqlStatement, "drop index ", vbTextCompare) > 0 Then
      xx = InStr(1, SqlStatement, " on ", vbTextCompare)
      If xx > 0 Then
         SqlStatement = Left$(SqlStatement, xx - 1)
      End If
   End If
    
    
   If InStr(1, SqlStatement, " autoincrement ", vbTextCompare) > 0 Then
      '   make a sequence & trigger
      SqlStatement = Replace(SqlStatement, " autoincrement ", " NUMBER(10) ", , , vbTextCompare)
      SplitRslt() = Split(SqlStatement, " ", 4, vbTextCompare)
       
      TableName = UCase(Replace(SplitRslt(2), ")", ""))
      SeqName = TableName & "_SEQ"
      SqlStatement2 = " CREATE SEQUENCE " & SeqName & _
             " START WITH 1 INCREMENT BY 1 NOMAXVALUE NOCACHE ORDER"
             
      TriggerName = TableName & "_AUTOINC"
      SqlStatement3 = " CREATE OR REPLACE TRIGGER " & SeqName & _
             " BEFORE INSERT ON " & TableName & " FOR EACH ROW BEGIN IF :new.ID IS NULL THEN " & _
             "SELECT " & SeqName & ".NextVal INTO :new.ID FROM DUAL; END IF; END;"
       
   End If
      
      
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   'comment out code below, no longer used
   'Added By N.Angelakis On 06 Mar 2009
   'DEV-252 Changes required for schema Access referenced date statements
'   If InStr(SqlStatement, "#") > 0 Then
'      arrDateReferences = Split(SqlStatement, "#", , vbTextCompare)
'      If UBound(arrDateReferences) > 0 Then
'         For intCount = 0 To UBound(arrDateReferences)
'            If IsDate(arrDateReferences(intCount)) Then
'               If Hour(arrDateReferences(intCount)) = 0 And Minute(arrDateReferences(intCount)) = 0 And Second(arrDateReferences(intCount)) = 0 Then
'                     arrDateReferences(intCount) = "TO_DATE('" & Format(arrDateReferences(intCount), "yyyy-mm-dd") & "','YYYY-MM-DD')"
'                  Else
'                     'syntax used as declared in mwUtility.modUtility.DateLiteral
'                     arrDateReferences(intCount) = "TO_DATE('" & Format(arrDateReferences(intCount), "yyyy-mm-dd hh:nn:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'               End If
'            End If
'            'Build query addind modified date syntax where necessary
'            strDateConvertedQuery = strDateConvertedQuery & " " & arrDateReferences(intCount)
'         Next
'         If strDateConvertedQuery <> "" Then
'             SqlStatement = strDateConvertedQuery
'         End If
'      End If
'      If UBound(arrDateReferences) > 0 Then Erase arrDateReferences
'   End If
      
   'DEV-2059 Allow Schema to Accommodate Hash Symbol (#) as Part of a Text String
   'By N.Angelakis 09 March 2011
   nDateStartPos = InStr(SqlStatement, "#")
   nDateEndPos = InStr(nDateStartPos + 1, SqlStatement, "#")
   Do While nDateStartPos > 0 And nDateEndPos > 0
      'reset variables
      sDatePart = ""
      sDateFormat = ""
      
      sDatePart = mID(SqlStatement, nDateStartPos + 1, InStr(nDateStartPos + 1, SqlStatement, "#") - nDateStartPos - 1)
      If IsDate(sDatePart) Then
         If Hour(sDatePart) = 0 And Minute(sDatePart) = 0 And Second(sDatePart) = 0 Then
            'syntax used as declared in mwUtility.modUtility.DateLiteral
            sDateFormat = "TO_DATE('" & Format(sDatePart, "yyyy-mm-dd") & "','YYYY-MM-DD')"
            
         Else
            'syntax used as declared in mwUtility.modUtility.DateLiteral
            sDateFormat = "TO_DATE('" & Format(sDatePart, "yyyy-mm-dd hh:nn:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            
         End If
         SqlStatement = Replace(SqlStatement, "#" + sDatePart + "#", sDateFormat, , vbTextCompare)
      End If
      nDateStartPos = InStr(nDateStartPos + 1, SqlStatement, "#")
      nDateEndPos = InStr(nDateStartPos + 1, SqlStatement, "#")
   Loop
      
      
   ConvertToOracleSyntax = SqlStatement
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.ConvertToOracleSyntax.", Err.Number, Err.Description
   ConvertToOracleSyntax = SqlStatement
End Function
Public Function IsSqlServer() As Boolean
   On Error GoTo FunctionError
   
   IsSqlServer = IsDBSqlServer(goCon)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsSqlServer.", Err.Number, Err.Description
End Function
Public Function IsAccess() As Boolean
   On Error GoTo FunctionError
   
   IsAccess = IsDBAccess(goCon)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsAccess.", Err.Number, Err.Description
End Function

Public Function IsOracle() As Boolean

   On Error GoTo FunctionError
   
   IsOracle = IsDBOracle(goCon)
   
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsOracle.", Err.Number, Err.Description
End Function

Private Function IsDBSqlServer(DBConn As Connection) As Boolean
   Dim s As String
   
   s = UCase(DBConn.ConnectionString)
   On Error GoTo FunctionError
   
   If InStr(s, "SQLOLEDB") Then
      IsDBSqlServer = True
   ElseIf InStr(s, "{SQL Server}") Then
      IsDBSqlServer = True
   ElseIf InStr(s, "SQLNCLI") Then
      IsDBSqlServer = True ' is sqlExpress (SQLNCLI.1) or sqlexpress2008r2 (SQLNCLI10)
   ElseIf InStr(s, "MSDASQL") Then
      '
      ' DSN Connection - tough to discern provider...
      ' Access: Provider=MSDASQL.1;Extended Properties="DSN=mws;DBQ=C:\mw_data.vela\database\mwDatabase.mdb;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
      ' SQL Server: Provider=MSDASQL.1;Extended Properties="DSN=mws_sql;UID=sa;PWD=;APP=Visual Basic;WSID=CHOCTAW;DATABASE=mwDataDEMO"
      '
      If InStr(s, "FIL=MS Access") Then
         IsDBSqlServer = False
      Else
         IsDBSqlServer = True
      End If
   Else
      IsDBSqlServer = False
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsDBSqlServer.", Err.Number, Err.Description
End Function
Private Function IsDBAccess(DBConn As Connection) As Boolean
   Dim s As String
   s = UCase(DBConn.ConnectionString)
   On Error GoTo FunctionError
   
   If InStr(s, "MICROSOFT.JET.OLEDB") Then
      IsDBAccess = True
   ElseIf InStr(s, "MICROSOFT.ACE.OLEDB") Then
      IsDBAccess = True
   ElseIf InStr(s, "MSDASQL") Then
      '
      ' DSN Connection - tough to discern provider...
      ' Access: Provider=MSDASQL.1;Extended Properties="DSN=mws;DBQ=C:\mw_data.vela\database\mwDatabase.mdb;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;UID=admin;"
      ' SQL Server: Provider=MSDASQL.1;Extended Properties="DSN=mws_sql;UID=sa;PWD=;APP=Visual Basic;WSID=CHOCTAW;DATABASE=mwDataDEMO"
      '
      If InStr(s, "FIL=MS Access") Then
         IsDBAccess = True
      Else
         IsDBAccess = False
      End If
   Else
      IsDBAccess = False
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsDBAccess.", Err.Number, Err.Description
End Function

Private Function IsDBOracle(DBConn As Connection) As Boolean
   Dim s As String
   s = UCase(DBConn.ConnectionString)
   On Error GoTo FunctionError
   
   If InStr(s, "ORAOLEDB") Then
      IsDBOracle = True
   Else
      IsDBOracle = False
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwDataWork.IsDBOracle.", Err.Number, Err.Description
End Function


Public Function Get_SqlServerDATEFORMAT(ConnectionID As String) As String

   Dim loRsDB As Recordset
   Dim sSQL As String
   Dim strConnectString As String
   On Error GoTo FunctionError
   Set loRsDB = New Recordset
   loRsDB.CursorLocation = adUseClient
   sSQL = "select * from mwcDbConnection where ConnectionID='" & ConnectionID & "'"
   loRsDB.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsDB.RecordCount > 0 Then
      If IsNull(loRsDB!SqlServerDATEFORMAT) Then
         Get_SqlServerDATEFORMAT = ""
      Else
         Get_SqlServerDATEFORMAT = loRsDB!SqlServerDATEFORMAT
      End If
   Else
      Get_SqlServerDATEFORMAT = ""
   End If
   goSession.CloseRecordset loRsDB
   Exit Function
FunctionError:
   goSession.RaiseError "General error in mwDataWork.Get_SqlServerDATEFORMAT. " & Err.Number & Err.Description
   Get_SqlServerDATEFORMAT = ""
End Function


Public Function FetchTargetRegionalSettings(ConnectionID As String) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select * from mwcDbConnection where ConnectionID='" & ConnectionID & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwDataForm.FetchTargetRegionalSettings, Connection record not found: " & ConnectionID
      goSession.CloseRecordset loRs
      FetchTargetRegionalSettings = False
      Exit Function
   End If
   If Not IsNull(loRs!DgGeneralDateFormat) Then
      mDgGeneralDateFormat = DATAGRAM_DATE_FORMAT
      'mDgGeneralDateFormat = loRs!DgGeneralDateFormat
   End If
   If Not IsNull(loRs!DgShortDateFormat) Then
      mDgShortDateFormat = DATAGRAM_SHORT_DATE_FORMAT
      'mDgShortDateFormat = loRs!DgShortDateFormat
   End If
   If Not IsNull(loRs!DgDecimalSeparator) Then
      mDgDecimalSeparator = loRs!DgDecimalSeparator
   Else
      mDgDecimalSeparator = DATAGRAM_DECIMAL_FORMAT
   End If
   goSession.CloseRecordset loRs
   FetchTargetRegionalSettings = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataWork.FetchTargetRegionalSettings. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   FetchTargetRegionalSettings = False
End Function

Public Function FetchTargetRegionalSettingsByBatch(BatchKey As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select ConnectionID from mwrBatchType where ID=" & BatchKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaisePublicError "Error in mwDataWork.FetchTargetRegionalSettingsByBatch, BatchType record missing for: " & BatchKey
      FetchTargetRegionalSettingsByBatch = False
   ElseIf IsNull(loRs!ConnectionID) Then
      goSession.RaisePublicError "Error in mwDataWork.FetchTargetRegionalSettingsByBatch, ConnectionID missing in BatchType record: " & BatchKey
      FetchTargetRegionalSettingsByBatch = False
   Else
      FetchTargetRegionalSettingsByBatch = FetchTargetRegionalSettings(loRs!ConnectionID)
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataWork.FetchTargetRegionalSettings. ", Err.Number, Err.Description
   goSession.CloseRecordset loRs
   FetchTargetRegionalSettingsByBatch = False
End Function

