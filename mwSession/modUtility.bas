Attribute VB_Name = "modUtility"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
      ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, _
      ByVal CY As Long, ByVal wFlags As Long) As Long
      
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
       (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
       
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Global Const GW_HWNDNEXT = 2
      
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Function CloseRecordset(ByRef oRs As Recordset)
   On Error GoTo FunctionError
   If Not oRs Is Nothing Then
      If oRs.State = adStateOpen Then
         If oRs.LockType <> adLockReadOnly Then
            If Not (oRs.EOF Or oRs.BOF) Then
               oRs.Move (0)
               'oRS.Update
            End If
         End If
         oRs.Close
      End If
      Set oRs = Nothing
   End If
   Exit Function
FunctionError:
   If Err.Number = -2147467259 Then
      oRs.Cancel
      Set oRs = Nothing
   Else
      goSession.RaisePublicError "General Error in mwSession.modUtility.CloseRecordset ", Err.Number, Err.Description
   End If
End Function

Public Function HideUltragridColumns(ByRef ug1 As SSUltraGrid, band As Integer)
   Dim i As Integer
   On Error GoTo FunctionError
   If ug1.Bands.Count = 0 Then Exit Function
   For i = 0 To (ug1.Bands(band).Columns.Count - 1)
      ug1.Bands(band).Columns(i).Hidden = True
   Next i
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.modUtility.HideUltraGridColumns. ", Err.Number, Err.Description
End Function


Public Function IsInCollection(loCollection As Collection, Key) As Boolean
   Dim s As String
   On Error GoTo FunctionError
   s = loCollection(Key)
   IsInCollection = True
   Exit Function
FunctionError:
   IsInCollection = False
End Function


Public Function IsColumnInTable(ByRef loRs As Recordset, ColumnName As String) As Boolean
   Dim strTemp As String
   On Error GoTo FunctionError
   strTemp = loRs.Fields(ColumnName).Name
   IsColumnInTable = True
   Exit Function
FunctionError:
   If Err.Number = 3265 Then
      IsColumnInTable = False
   Else
      goSession.RaiseError "General Error in mwSession.modUtility.IsColumnIn Table. ", Err.Number, Err.Description
   End If
End Function


Public Function KillObject(obj As Object)
   If Not obj Is Nothing Then
      Set obj = Nothing
   End If
End Function

Public Function IsBlank(str As Variant) As Boolean
   
   If IsNull(str) Then
      IsBlank = True
   ElseIf IsEmpty(str) Then
      IsBlank = True
   ElseIf Len(str) <= 0 Then
      IsBlank = True
   ElseIf Len(Trim$(str)) <= 0 Then
      IsBlank = True
   Else
      IsBlank = False
   End If
   
End Function
Public Function NullTrim(InVar As Variant, Optional OptLen As Integer = 0) As Variant
   
   On Error GoTo Err
   
   If IsNull(InVar) Then
      NullTrim = InVar
   ElseIf IsEmpty(InVar) Then
      NullTrim = InVar
   ElseIf Len(InVar) <= 0 Then
      NullTrim = InVar
   Else
      If OptLen > 0 Then
         NullTrim = Left$(InVar, OptLen)
      Else
         NullTrim = InVar
      End If
   End If
   
'   If InVar.Type = adNumeric And Len(NullTrim) <= 0 Then
'      NullTrim = "0"
'   ElseIf (InVar.Type = adDBTimeStamp Or Str.Type = adDate) And Len(NullTrim) <= 0 Then
'      NullTrim = "0"
'   End If
Err:
   Exit Function
End Function

Public Function BlankNull(str As Variant) As String
   
   On Error GoTo Err
   
   If IsNull(str) Then
      BlankNull = ""
   ElseIf IsEmpty(str) Then
      BlankNull = ""
   ElseIf Len(str) <= 0 Then
      BlankNull = ""
   Else
      BlankNull = Trim$(str)
   End If
   
   If TypeName(str) = "Field" Then
   
      If str.Type = adNumeric And Len(BlankNull) <= 0 Then
         BlankNull = "0"
      ElseIf (str.Type = adDBTimeStamp Or str.Type = adDate) And Len(BlankNull) <= 0 Then
         BlankNull = "0"
      End If
   End If
Err:
   Exit Function
End Function
Public Function ZeroNull(str As Variant) As Long
   
   On Error GoTo Err
   
   If IsNull(str) Then
      ZeroNull = 0
   ElseIf IsEmpty(str) Then
      ZeroNull = 0
   ElseIf Len(str) <= 0 Then
      ZeroNull = 0
   Else
      ZeroNull = 0
      If TypeName(str) <> "Field" Then
         ZeroNull = CLng(Trim$(str))
      Else
      Select Case str.Type
         Case adBoolean
            If str = True Then
               ZeroNull = 1
            Else
               ZeroNull = 0
            End If
'         Case adChar
'         Case adDate, adDBDate, adDBTime, adDBTimeStamp
'         Case adDecimal
'         Case adDouble
'         Case adInteger
'         Case adLongVarChar
'         Case adNumeric
'         Case adSingle
'         Case adSmallInt
'         Case adTinyInt
'         Case adUnsignedBigInt
'         Case adUnsignedInt
'         Case adUnsignedSmallInt
'         Case adUnsignedTinyInt
'            ZeroNull = CLng(Str)
'         Case adVarChar
         
         Case Else
            If Len(str) <= 0 Then
               ZeroNull = 0
            Else
               ZeroNull = CLng(Trim$(str))
            End If
         
      End Select
      End If
   End If
   Exit Function
   
Err:
   ZeroNull = 0
   Exit Function
End Function

Public Function NullEmpty(InVal As Variant) As Variant
   
   On Error GoTo Err
   
   If IsNull(InVal) Then
      NullEmpty = Null
   ElseIf IsEmpty(InVal) Then
      NullEmpty = Null
   ElseIf Len(InVal) <= 0 Then
      NullEmpty = Null
   Else
      NullEmpty = Trim$(InVal)
   End If
   
Err:
   Exit Function
End Function
Public Function NullEmptyOrZero(InVal As Variant) As Variant
   
   On Error GoTo Err
   
   If IsNull(InVal) Then
      NullEmptyOrZero = Null
   ElseIf IsEmpty(InVal) Then
      NullEmptyOrZero = Null
   ElseIf Len(InVal) <= 0 Then
      NullEmptyOrZero = Null
   Else
      NullEmptyOrZero = Trim$(InVal)
      If NullEmptyOrZero = "0" Then
         NullEmptyOrZero = Null
      End If
   End If
   
   
Err:
   Exit Function
End Function

Public Sub SetNonNull(ByRef Rs As Recordset, DestVar As String, SrcVar As Variant)

   On Error GoTo Err
   
   If Not IsNull(SrcVar) And Not IsEmpty(SrcVar) Then
      Rs(DestVar) = Trim$(SrcVar)
   End If
   
   Exit Sub
   
Err:
   Exit Sub
End Sub
Public Sub SetNonBlank(ByRef Rs As Recordset, DestVar As String, SrcVar As Variant)

   On Error GoTo Err
   
   If Not IsBlank(SrcVar) Then
      Rs(DestVar) = Trim$(SrcVar)
   End If
   
   Exit Sub
   
Err:
   Exit Sub
End Sub

Public Function GetNonNull(ByRef Rs As Recordset, ByVal SrcVar As String) As Variant

   On Error GoTo Err
   
   If Rs.BOF = True Or Rs.EOF = True Then
      GetNonNull = Empty
   Else
      GetNonNull = Rs(SrcVar)
   End If
   
   Exit Function
   
Err:
   Exit Function
End Function
Public Sub HandleUG_KeyDown(SSug As SSUltraGrid, KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
   On Error GoTo SubError
      
   ' KEYUP,DOWN, TAB,BACKTAB, HOME,BACKHOME, END,BACKEND
   If KeyCode = vbKeyUp Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionAboveCell
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyDown Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionBelowCell
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyTab And Shift = 0 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionNextCellByTab
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyTab And Shift = 1 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionPrevCellByTab
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyHome And Shift = 0 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
'      SSug.ColScrollRegions(0).Scroll ssColScrollActionLeft
      SSug.PerformAction ssKeyActionFirstCellInRow
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyHome And Shift = vbCtrlMask Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionFirstCellInGrid
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyEnd And Shift = 0 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionLastCellInRow
      SSug.PerformAction ssKeyActionEnterEditMode
   
   ElseIf KeyCode = vbKeyEnd And Shift = vbCtrlMask Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionLastCellInGrid
      SSug.PerformAction ssKeyActionEnterEditMode
      
   ElseIf KeyCode = vbKeyPageDown And Shift = 0 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionPageDownCell
      SSug.PerformAction ssKeyActionEnterEditMode
      
   ElseIf KeyCode = vbKeyPageUp And Shift = 0 Then
      KeyCode = 0
      SSug.PerformAction ssKeyActionExitEditMode
      SSug.PerformAction ssKeyActionPageUpCell
      SSug.PerformAction ssKeyActionEnterEditMode
      
   ElseIf KeyCode = vbKeyDelete And Shift = vbCtrlMask Then
      KeyCode = 0
      
      If Not SSug.ActiveCell Is Nothing Then
         If SSug.ActiveCell.Activation = ssActivationAllowEdit Then
      
            SSug.PerformAction ssKeyActionExitEditMode
            SSug.ActiveCell.value = Null
         End If
      End If
'      SSug.ActiveCell.Appearance.Clear
   
   ElseIf KeyCode = vbKeyEscape Then
      KeyCode = 0
   Else
      If SSug.ActiveCell Is Nothing Then
         Exit Sub
      End If
      KeyCode = 0
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwUtility.HandleUG_KeyDown ", Err.Number, Err.Description
End Sub

Public Function IsRecordLoaded(ByRef oRs As Recordset) As Boolean
   On Error GoTo FunctionError
   If oRs Is Nothing Then
      IsRecordLoaded = False
   ElseIf oRs.State <> adStateOpen Then
      IsRecordLoaded = False
   ElseIf (oRs.EOF Or oRs.BOF) Then
      IsRecordLoaded = False
   ElseIf oRs.RecordCount = 0 Then
      IsRecordLoaded = False
   Else
      IsRecordLoaded = True
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwWorks5.modUtility.IsRecordLoaded ", Err.Number, Err.Description
End Function


Public Function TranslateBoolean(BooleanValue As Boolean) As String
   On Error GoTo FunctionError
   If BooleanValue = vbTrue Then
      TranslateBoolean = "1"
   Else
      TranslateBoolean = "0"
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwReplicateWillChange.TranslateBoolean. ", Err.Number, Err.Description
   TranslateBoolean = False
End Function

' Used to set the position for grid columns

Public Function IncrCounter(ByRef Counter As Long) As Long
   IncrCounter = Counter
   Counter = Counter + 1
End Function


Public Function BoolNull(str As Variant) As Boolean
   On Error GoTo Err
   
   If IsNull(str) Then
      BoolNull = False
   ElseIf IsEmpty(str) Then
      BoolNull = False
   ElseIf Len(str) <= 0 Then
      BoolNull = False
   Else
      BoolNull = False
      ' (is case sensitive) Boolean, Byte, Currency, Date, Decimal,Double, Empty, Error,Integer, Long, Nothing, Null, Single String, Unknown, Variant()
      Select Case TypeName(str)
         Case Is = "Boolean"
            If str = True Then
               BoolNull = True
            Else
               BoolNull = False
            End If
            
         Case Else
            If Len(str) <= 0 Then
               BoolNull = False

            ElseIf UCase(str) = "TRUE" Then 'By N.Angelakis 26 April 2010
               BoolNull = True
            ElseIf UCase(str) = "FALSE" Then 'By N.Angelakis 26 April 2010
               BoolNull = False

            Else
               BoolNull = CLng(Trim$(str))
            End If
      End Select
   End If
   Exit Function
   
Err:
   BoolNull = 0
   Exit Function
End Function



