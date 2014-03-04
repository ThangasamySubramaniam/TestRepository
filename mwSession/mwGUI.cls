VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwGUI - GUI tools migrated from modUtility at the application layer...
' 10/27/2003 ms
'

Option Explicit
Private mUG_ParentBandBackColorNoEdit As Long
Private mUG_ParentBand_BackColorWithEdit As Long
Private mUG_EditBackColor As Long
Private mUG_NoEdit_BackColor As Long
Private mIsGuiKeysLoaded As Boolean

Const UG_FORMAT_DATETIME_WIDTH = 2000
Const UG_FORMAT_EDIT_BACKGROUND = &HD7F9FD   ' gold background
Const UG_FORMAT_NOEDIT_BACKGROUND = &HF0F0F0 ' light grey background
Const UG_EDIT_BACKGROUND = &HFFFFFFFF         ' standard white
Const UG_FORMAT_PARENT_BAND = &HFFFFD7        ' light green background


Private moGuiKeys As ConfigKeys

' silent childRef check
Private msChildRefError As String
Private mIsSilentChildRefError As Boolean


Public Property Get UG_ParentBandBackColorNoEdit() As Long
    UG_ParentBandBackColorNoEdit = mUG_ParentBandBackColorNoEdit
End Property

Public Property Get UG_ParentBand_BackColorWithEdit() As Long
    UG_ParentBand_BackColorWithEdit = mUG_ParentBand_BackColorWithEdit
End Property

Public Property Get UG_EditBackColor() As Long
    UG_EditBackColor = mUG_EditBackColor
End Property

Public Property Get UG_NoEdit_BackColor() As Long
    UG_NoEdit_BackColor = mUG_NoEdit_BackColor
End Property

Public Property Get SilentChildRefError() As String
   SilentChildRefError = msChildRefError
End Property
Public Property Let IsSilentChildRefError(bValue As Boolean)
    mIsSilentChildRefError = bValue
End Property

'Public Property Get A() As long
'    A = mA
'End Property




'Public Property Get A() As long
'    A = mA
'End Property

'Public Property Let A(ByVal vData As long)
'    mA = vData
'End Property

Public Function SetCursorWait(StatusMessage As String) As Boolean
   Screen.MousePointer = vbHourglass
End Function
Public Function SetCursorNormal() As Boolean
   Screen.MousePointer = vbDefault
End Function


Public Function TraverseUgSearch(SearchString As String, ByRef ugControl As Object, IsStartBeginning As Boolean) As Boolean
   Dim loRow As SSRow
   Dim loChildRow As SSRow
   Dim loGrid As SSUltraGrid
   Dim loStartRow As SSRow
   Dim loCell As SSCell
   Dim IsDone As Boolean
   Dim FoundParent As Boolean
   Dim UCaseSearchString As String
   
   On Error GoTo FunctionError
   
   Set loGrid = ugControl
   
   If loGrid Is Nothing Then
      Exit Function
   ElseIf Not loGrid.HasRows Then
      Exit Function
   ElseIf IsNull(loGrid.ActiveRow) Or (loGrid.ActiveRow Is Nothing) Then
      'Exit Function
      Set loGrid.ActiveRow = loGrid.GetRow(ssChildRowFirst)
      If loGrid.ActiveRow Is Nothing Then
         Exit Function
      End If
   End If
   
   
   
   UCaseSearchString = UCase(Trim(Left(SearchString, 50)))

   IsDone = False
   
   Set loStartRow = loGrid.ActiveRow
   loGrid.Selected.ClearAll
   
   
   '
   ' Traverse down the child tree If not found yet ...
   '
   
   ' If Is StartBeginning then go to the first row in the grid and start there
   If IsStartBeginning Then
      Set loGrid.ActiveRow = loGrid.GetRow(ssChildRowFirst)
      Set loStartRow = loGrid.ActiveRow
      Set loRow = loGrid.ActiveRow
      
      IsDone = SearchUGChildren(UCaseSearchString, ugControl, loRow)
   
   Else
      Set loRow = loGrid.ActiveRow
      
      ' Skip the current row and start with the child rows.
      
      If loRow.HasChild Then
         ' check if child...
         Set loChildRow = loRow.GetChild(ssChildRowFirst)
         
         IsDone = SearchUGChildren(UCaseSearchString, ugControl, loChildRow)
      End If
   End If
   '
   
   '
   ' If no match found yet then check the siblings
   '
   If Not IsDone And loRow.HasNextSibling Then
      Set loRow = loRow.GetSibling(ssSiblingRowNext)
      IsDone = SearchUGChildren(UCaseSearchString, ugControl, loRow)
   End If
   
   '
   ' If still no match then check the siblings of my parents
   '
   Do While loRow.HasParent And Not IsDone
      
      '
      ' Get my parent
      '
      Set loRow = loRow.GetParent
      
      '
      ' Check if parent has a next sibling...
      '
      If loRow.HasNextSibling Then
         Set loRow = loRow.GetSibling(ssSiblingRowNext)
         IsDone = SearchUGChildren(UCaseSearchString, ugControl, loRow)
      End If
   Loop
   
   If Not IsDone Then
      ' loGrid traversed...complete but no match
      ImprovedMsgBox "'" & SearchString & "' not found."
      Set loGrid.ActiveRow = loStartRow
   End If
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwGUI.TraverseUgSearch. ", Err.Number, Err.Description
End Function

Private Function SearchUGChildren(SearchString As String, ByRef ugControl As Object, TopRow As SSRow) As Boolean
   Dim loRow As SSRow
   Dim loChildRow As SSRow
   Dim loGrid As SSUltraGrid
   Dim loCell As SSCell
   Dim IsDone As Boolean
   
   On Error GoTo FunctionError
   Set loGrid = ugControl
   If loGrid Is Nothing Then
      Exit Function
   ElseIf Not loGrid.HasRows Then
      Exit Function
   ElseIf IsNull(loGrid.ActiveRow) Or (loGrid.ActiveRow Is Nothing) Then
      Exit Function
   End If
   
   Set loRow = TopRow
   SearchUGChildren = False
   
   '

   IsDone = False
   ' Loop through all Child & Sibling rows...
   Do While Not IsDone
   
      If IsStringInRow(loRow, SearchString) Then
         IsDone = True
         Set loGrid.ActiveRow = loRow
         loRow.Selected = True
         SearchUGChildren = True
         
      End If
   
      '
      ' Traverse down the child tree If not found yet ...
      '
      If Not IsDone And loRow.HasChild Then
         ' check if child...
         Set loChildRow = loRow.GetChild(ssChildRowFirst)
         
         IsDone = SearchUGChildren(SearchString, ugControl, loChildRow)
      End If
      
      If Not IsDone And loRow.HasNextSibling Then
         ' has as a  brother or sister...
         Set loRow = loRow.GetSibling(ssSiblingRowNext)
      Else
         SearchUGChildren = IsDone
         Exit Function
      End If
   Loop
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwGUI.SearchUGChildren. ", Err.Number, Err.Description
End Function

Private Function IsStringInRow(ByRef oRow As SSRow, SearchString As String) As Boolean
   Dim loCell As SSCell
   On Error Resume Next
   For Each loCell In oRow.Cells
      If Not loCell.Column.Hidden Then
         If InStr(1, UCase(loCell.GetText), SearchString, vbTextCompare) > 0 Then
            If Err.Number = 0 Then
               IsStringInRow = True
               Exit Function
            End If
            Err.Clear
         End If
      End If
   Next loCell
   IsStringInRow = False
   Exit Function
End Function


Public Function HideUltragridColumns(ByRef ugControl As Object, band As Integer) As Boolean
   Dim i As Integer
   Dim loGrid As SSUltraGrid
   On Error GoTo FunctionError
   Set loGrid = ugControl
   If loGrid.Bands.Count = 0 Then Exit Function
   For i = 0 To (loGrid.Bands(band).Columns.Count - 1)
      loGrid.Bands(band).Columns(i).Hidden = True
   Next i
   HideUltragridColumns = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwGUI.HideUltraGridColumns. ", Err.Number, Err.Description
   HideUltragridColumns = False
End Function

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
   goSession.RaiseError "General Error in mwGUI.IsRecordLoaded ", Err.Number, Err.Description
End Function

Public Function CloseRecordset(ByRef oRs As Recordset)
   On Error GoTo FunctionError
   If Not oRs Is Nothing Then
      If oRs.State = adStateOpen Then
         If oRs.LockType <> adLockReadOnly Then
            If Not (oRs.EOF Or oRs.BOF Or oRs.RecordCount = 0) Then
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
   Else
      goSession.RaiseError "General Error in mwGUI.CloseRecordset ", Err.Number, Err.Description
   End If
   Set oRs = Nothing
End Function

Public Function UpdateRecordset(ByRef oRs As Recordset)
   On Error GoTo FunctionError
   If Not oRs Is Nothing Then
      If oRs.State = adStateOpen Then
         If oRs.LockType <> adLockReadOnly Then
            If Not (oRs.EOF Or oRs.BOF) Then
               'oRS.Move (0)
               oRs.Update
            End If
         End If
      End If
   End If
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwGUI.UpdateRecordset ", Err.Number, Err.Description
End Function

Public Function AddUgAppearances(ByRef ugControl As Object) As Boolean
   Dim loGrid As SSUltraGrid
   On Error GoTo FunctionError
   '
   ' Setup Appearances for highlighting rows...
   loGrid.Appearances.Add "HighlightRow"
   loGrid.Appearances("HighlightRow").BorderColor = vbBlue
   loGrid.Appearances("HighlightRow").ForeColor = vbRed
   loGrid.Appearances.Add "NormalRow"
   loGrid.Appearances("NormalRow").BorderColor = vbBlack
   loGrid.Appearances("NormalRow").ForeColor = vbBlack
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwGUI.UpdateRecordset ", Err.Number, Err.Description
End Function

Public Function LockForm(ByRef TheForm As Object)
   Dim oCtrl As Control
   For Each oCtrl In TheForm.Controls
      If Trim(oCtrl.TAG) <> "" Then
       oCtrl.Enabled = False
      End If
   Next oCtrl
   
End Function

Public Function UnLockForm(ByRef TheForm As Object)
   Dim oCtrl As Control
   For Each oCtrl In TheForm.Controls
      If Trim(oCtrl.TAG) <> "" Then
       oCtrl.Enabled = True
      End If
   Next oCtrl
End Function

Public Function SetFormOverrides(ByRef FormObject As Object, EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim strSQL As String
   Dim loform As Form
   On Error GoTo FunctionError
   '
   ' Are there Tweak Records...?
   '
   Set loform = FormObject
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwEventInterfaceTweak where IsActive <> 0 and mwEventTypeKey=" & EventType
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      CloseRecordset loRs
      SetFormOverrides = True
      Exit Function
   End If
   '
   ' Each record describes a control override
   '
   Do While Not loRs.EOF
      '
      ' Active Tab ?
      If loRs!IsTabControl Then
         ' Yes, Active Tab...
         If IsNull(loRs!ControlIndex) Then
            goSession.RaisePublicError "Error in mwSession.mwGUI.SetFormOverrides, no Index specified for Tab control, Record: " & loRs!ID
         Else
            '
            ' Active Tab - set caption/visibility of tabs...
            '
            If loRs!IsHidden Then
               'loform.Controls(loRs!ControlName).Tabs(loRs!ControlIndex + 1).Visible = False
               loform.Controls(loRs!ControlName).Tabs(loRs!ControlIndex).Visible = False
            Else
               If Not IsNull(loRs!NewCaption) Then
                  'loform.Controls(loRs!ControlName).Tabs(loRs!ControlIndex + 1).Caption = loRs!NewCaption
                  loform.Controls(loRs!ControlName).Tabs(loRs!ControlIndex).Caption = loRs!NewCaption
               End If
            End If
         End If
      Else
         ' No, regular control type
         '
         ' Caption change ?
         If Not IsNull(loRs!NewCaption) Then
            If Not IsNull(loRs!ControlIndex) Then
               loform.Controls(loRs!ControlName, loRs!ControlIndex).Caption = loRs!NewCaption
            Else
               loform.Controls(loRs!ControlName).Caption = loRs!NewCaption
            End If
         End If
         '
         ' Hide the control ?
         '
         If loRs!IsHidden Then
            If IsNull(loRs!ControlIndex) Then
               loform.Controls(loRs!ControlName).Visible = False
            Else
               loform.Controls(loRs!ControlName, loRs!ControlIndex).Visible = False
            End If
         End If
         '
         ' Is reposition the control ?
         '
         If Not IsNull(loRs!NewLeftCoordinate) Then
            If IsNull(loRs!ControlIndex) Then
               loform.Controls(loRs!ControlName).Left = loRs!NewLeftCoordinate
            Else
               loform.Controls(loRs!ControlName, loRs!ControlIndex).Left = loRs!NewLeftCoordinate
            End If
         End If
         If Not IsNull(loRs!NewTopCoordinate) Then
            If IsNull(loRs!ControlIndex) Then
               loform.Controls(loRs!ControlName).Top = loRs!NewTopCoordinate
            Else
               loform.Controls(loRs!ControlName, loRs!ControlIndex).Top = loRs!NewTopCoordinate
            End If
         End If
         
      End If
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   SetFormOverrides = True
   Exit Function
FunctionError:
   If Err.Number <> 31104 Then
      goSession.RaisePublicError "General Error in mwSession.mwGUI.SetFormOverrides ", Err.Number, Err.Description
   Else
      goSession.RaiseWorkflowStatus "Error in mwSession.mwGUI.SetFormOverrides, Tab control raises an error: Tab control must have at least one visible Tab"
   End If
   SetFormOverrides = False
End Function

Public Function SearchUltragridKey(ByRef ugGrid As Object, SearchBand As Integer, RecKey As Long, Optional SearchCol As Integer) As Boolean
   On Error GoTo FunctionError:
   Dim loGrid As SSUltraGrid
   Dim IsDone As Boolean
   Dim IsBingo As Boolean
   Dim nMaxCount As Long
   Dim loRow As SSRow
   Dim loCurrentRow As SSRow
   Set loGrid = ugGrid
   '
   ' validate inputs
   '
   
   If loGrid Is Nothing Then
      Exit Function
   ElseIf Not loGrid.HasRows Then
      Exit Function
   End If
   
   If (SearchBand < 0) Or (loGrid.Bands.Count < SearchBand) Then
      Exit Function
   End If
   If RecKey < 0 Then
      Exit Function
   End If
   '
   ' start at beginning of loGrid
   '
   If Not loGrid.ActiveRow Is Nothing Then
      Set loCurrentRow = loGrid.ActiveRow
   End If
   Set loRow = loGrid.GetRow(ssChildRowFirst)
   IsDone = False
   nMaxCount = 0
   Do While Not IsDone
      If loRow.Cells(SearchCol).value = RecKey And loRow.band.Index = SearchBand Then
         'Bingo
         IsBingo = True
         IsDone = True
      ElseIf loRow.band.Index = SearchBand Then
         If loRow.HasNextSibling Then
            Set loRow = loRow.GetSibling(ssSiblingRowNext)
         ElseIf loRow.HasParent Then
            Set loRow = loRow.GetParent
            If loRow.HasNextSibling Then
               Set loRow = loRow.GetSibling(ssSiblingRowNext)
            Else
               IsBingo = False
               IsDone = True
            End If
         Else
            IsBingo = False
            IsDone = True
         End If
      ElseIf loRow.HasChild Then
         Set loRow = loRow.GetChild(ssChildRowFirst)
      ElseIf loRow.HasParent Then
         Set loRow = loRow.GetParent
         If loRow.HasNextSibling Then
            Set loRow = loRow.GetSibling(ssSiblingRowNext)
         Else
            IsBingo = False
            IsDone = True
         End If
      ElseIf loRow.HasNextSibling Then
         Set loRow = loRow.GetSibling(ssSiblingRowNext)
      Else
         IsBingo = False
         IsDone = True
      End If
   Loop
   If IsBingo Then
      Set loGrid.ActiveRow = loRow
      SearchUltragridKey = True
   Else
      SearchUltragridKey = False
      If Not loCurrentRow Is Nothing Then
         Set loGrid.ActiveRow = loCurrentRow
      End If
   End If
   KillObject loRow
   KillObject loCurrentRow
   Exit Function
FunctionError:
   If Err.Number = 13 Then
      goSession.LogIt mwl_Workstation, mwl_Warning, "Type Mismatch Error in mwGUI.SearchUltragridKey, logging and leaving. "
   Else
      goSession.RaisePublicError "Error in mwGUI.SearchUltragridKey. ", Err.Number, Err.Description
   End If
   KillObject loRow
End Function

Public Function FetchSelectedUgKeys(ugGrid As Object) As Recordset
   Dim loGrid As SSUltraGrid
   Dim loRow As SSRow
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loGrid = ugGrid
   '
   ' Fabricate the RS
   '
   Set loRs = New Recordset
   loRs.Fields.Append "ID", adInteger
   loRs.Open
   For Each loRow In loGrid.Selected.Rows
      loRs.AddNew
      loRs!ID = loRow.Cells(0).value
      loRs.Update
   Next loRow
   If loRs.RecordCount > 0 Then
      loRs.MoveFirst
   End If
   Set FetchSelectedUgKeys = loRs
   KillObject loRow
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwGUI.FetchSelectedUgKeys. ", Err.Number, Err.Description
End Function

Public Function ShowEventReportsPopup(EventType As Long, WindowTitle As String, _
      CurrentSiteKey As Long, Optional SubType As Long, _
      Optional IsModal As Boolean, Optional DetailKey As Long, Optional DetailKey2 As Long, Optional DetailKey3 As Long) As Boolean
   Dim loform As frmEventReports
   On Error GoTo FunctionError
   Set loform = New frmEventReports
   If Not loform.InitForm(WindowTitle, EventType, CurrentSiteKey, SubType, True, DetailKey, DetailKey2, DetailKey3) Then
      Unload loform
      Set loform = Nothing
   Else
      
      If IsModal Then
         loform.Show vbModal
         Unload loform
         Set loform = Nothing
      Else
         loform.Show
      End If
   End If
   Exit Function
FunctionError:
   goSession.RaisePublicError "Error in mwGUI.ShowEventReportsPopup. ", Err.Number, Err.Description
End Function


Friend Function LoadGuiKeys() As Boolean
   Dim loRs As Recordset
   Dim strSiteID As String
   Dim strSQL As String
   Dim loField As field
   On Error GoTo FunctionError
   If mIsGuiKeysLoaded Then
      LoadGuiKeys = True
      Exit Function
   End If
   Set moGuiKeys = New ConfigKeys
   '
   '
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   '
   ' Fetch it...
   '
   strSQL = "select * from mwcGuiParameters"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount < 1 Then
      goSession.RaiseError "Error in mwSession.mwSite.LoadGuiKeys, mwcGuiParameters Record Missing.", Err.Number, Err.Description
      goSession.CloseRecordset loRs
      LoadGuiKeys = False
      Exit Function
   End If
   '
   ' OK, populate the collection...
   '
   For Each loField In loRs.Fields
      Select Case UCase(loField.Name)
         Case Is = "UG_PARENTBANDBACKCOLORNOEDIT"
            If Not IsNull(loField.value) And loField.value <> 0 Then
               mUG_ParentBandBackColorNoEdit = loField.value
            Else
               mUG_ParentBandBackColorNoEdit = UG_FORMAT_NOEDIT_BACKGROUND
            End If
         Case Is = "UG_PARENTBAND_BACKCOLOREDIT"
            If Not IsNull(loField.value) And loField.value <> 0 Then
               mUG_ParentBand_BackColorWithEdit = loField.value
            Else
               mUG_ParentBand_BackColorWithEdit = UG_FORMAT_EDIT_BACKGROUND
            End If
         Case Is = "UG_EDITBACKCOLOR"
            If Not IsNull(loField.value) And loField.value <> 0 Then
               mUG_EditBackColor = loField.value
            Else
               mUG_ParentBandBackColorNoEdit = UG_FORMAT_EDIT_BACKGROUND
            End If
         Case Is = "UG_NOEDIT_BACKCOLOR"
            If Not IsNull(loField.value) And loField.value <> 0 Then
               mUG_NoEdit_BackColor = loField.value
            Else
               mUG_NoEdit_BackColor = UG_FORMAT_NOEDIT_BACKGROUND
            End If
         Case Else
            '
            ' Otherwise, Add to the Collection of configkeys... (extended properties)
            '
            If IsNull(loField.value) Then
               moGuiKeys.Add UCase(loField.Name), "", UCase(loField.Name)
            Else
               moGuiKeys.Add UCase(loField.Name), loField.value, UCase(loField.Name)
            End If
      End Select
   Next
   CloseRecordset loRs
   mIsGuiKeysLoaded = True
   LoadGuiKeys = True
   Exit Function
FunctionError:
   goSession.RaiseError "Error in mwSession.mwGUI.LoadGuiKeys.", Err.Number, Err.Description
   LoadGuiKeys = False
End Function


Public Function GetExtendedProperty(Key As String) As String
   On Error GoTo FunctionError
   GetExtendedProperty = moGuiKeys.GetKeyValue(Key)
   Exit Function
FunctionError:
   GetExtendedProperty = ""
End Function

Private Sub Class_Initialize()
   LoadGuiKeys
End Sub



Public Function IsFkRecordsExist(TableName As String, FkColumnName As String, RecordKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   'validate
   If Len(TableName) < 1 Then
      goSession.RaisePublicError "Error In IsFkRecordsExist, no TableName Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   ElseIf Len(FkColumnName) < 1 Then
      goSession.RaisePublicError "Error In IsFkRecordsExist, no Fk ColumnName Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   ElseIf RecordKey < 1 Then
      goSession.RaisePublicError "Error In IsFkRecordsExist, no Record Key Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT Count(*) FROM " & TableName & " WHERE " & FkColumnName & " =" & RecordKey
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRs.Fields(0).value >= 1 Then
      IsFkRecordsExist = True
   Else
      IsFkRecordsExist = False
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwGUI.IsFkRecordsExist. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function IsStrFkRecordsExist(TableName As String, FkColumnName As String, strFKey As String) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   
   'validate
   If Len(TableName) < 1 Then
      goSession.RaisePublicError "Error In IsStrFkRecordsExist, no TableName Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   ElseIf Len(FkColumnName) < 1 Then
      goSession.RaisePublicError "Error In IsStrFkRecordsExist, no Fk ColumnName Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   ElseIf Len(strFKey) < 1 Then
      goSession.RaisePublicError "Error In IsStrFkRecordsExist, no Record Key Value", vbInformation, "Check If Foreign Key Records Exist"
      Exit Function
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT Count(*) FROM " & TableName & " WHERE " & FkColumnName & " = '" & strFKey & "'"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   If loRs.Fields(0).value >= 1 Then
      IsStrFkRecordsExist = True
   Else
      IsStrFkRecordsExist = False
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwGUI.IsStrFkRecordsExist. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function IsChildReferencesExist(TableName As String, PKColumnName As String, RecordKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim UpperTableName As String
   Dim TemplateID As String
   
   On Error GoTo FunctionError
   
   msChildRefError = ""             ' return silentChildRefError
   
   'validate
   If Len(TableName) < 1 Then
      If mIsSilentChildRefError Then
         msChildRefError = "Error In IsChildReferencesExist, no TableName Value " & " Check If Child References Exist"
      Else
         goSession.RaisePublicError "Error In IsChildReferencesExist, no TableName Value", vbInformation, "Check If Child References Exist"
      End If
      Exit Function
   ElseIf Len(PKColumnName) < 1 Then
      If mIsSilentChildRefError Then
         msChildRefError = "Error In IsChildReferencesExist, no PK ColumnName Value " & " Check If Child References Exist"
      Else
         goSession.RaisePublicError "Error In IsChildReferencesExist, no PK ColumnName Value", vbInformation, "Check If Child References Exist"
      End If
      Exit Function
   ElseIf RecordKey < 1 Then
      If mIsSilentChildRefError Then
         msChildRefError = "Error In IsChildReferencesExist, no Record Key Value " & " Check If Child References Exist"
      Else
         goSession.RaisePublicError "Error In IsChildReferencesExist, no Record Key Value", vbInformation, "Check If Child References Exist"
      End If
      Exit Function
   End If
      
   UpperTableName = UCase(TableName)
   
   If goSession.IsAccess Then
      sSQL = "SELECT szRelationship AS ConsName, szReferencedObject AS RefTab," & _
            " szReferencedColumn AS RefCol, szObject AS ChildTab, szColumn As ChildCol" & _
            " FROM MSysRelationships WHERE UCase(szReferencedObject)  = '" & UpperTableName & "'" & _
            " ORDER BY szObject, szColumn"
   ElseIf goSession.IsSqlServer Then
      sSQL = "SELECT   FO.name AS ConsName, SO2.name AS RefTab, SC2.name AS RefCol, " & _
             " SO1.name AS ChildTab, SC1.name AS ChildCol " & _
             " FROM dbo.sysforeignkeys FK " & _
             " INNER JOIN dbo.syscolumns SC1 ON FK.fkeyid = SC1.id AND FK.fkey = SC1.colid " & _
             " INNER JOIN dbo.syscolumns SC2 ON FK.rkeyid = SC2.id AND FK.rkey = SC2.colid " & _
             " INNER JOIN dbo.sysobjects SO1 ON FK.fkeyid = SO1.id " & _
             " INNER JOIN dbo.sysobjects SO2 ON FK.rkeyid = SO2.id " & _
             " INNER JOIN dbo.sysobjects FO ON FK.constid = FO.id " & _
             " WHERE Upper(SO2.name) = '" & UpperTableName & "' ORDER BY ChildTab, ChildCol "
   ElseIf goSession.IsOracle Then
      sSQL = "SELECT c.CONSTRAINT_NAME ConsName, c.TABLE_NAME ChildTab, cc.COLUMN_NAME ChildCol, " & _
             " r.TABLE_NAME RefTab, rc.COLUMN_NAME RefCol " & _
             " FROM dba_constraints c, dba_constraints r, dba_cons_columns cc, dba_cons_columns rc " & _
             " WHERE c.CONSTRAINT_TYPE = 'R' AND UPPER(R.TABLE_NAME) = '" & UpperTableName & "'" & _
             " AND c.R_OWNER = r.OWNER AND c.R_CONSTRAINT_NAME = r.CONSTRAINT_NAME " & _
             " AND c.CONSTRAINT_NAME = cc.CONSTRAINT_NAME AND c.OWNER = cc.OWNER " & _
             " AND r.CONSTRAINT_NAME = rc.CONSTRAINT_NAME AND r.OWNER = rc.OWNER " & _
             " AND cc.POSITION = rc.POSITION " & _
             " order by c.TABLE_NAME, cc.COLUMN_NAME"
   Else
      If mIsSilentChildRefError Then
         msChildRefError = "Error In IsChildReferencesExist, Unknown Database Engine " & " Check If Child References Exist"
      Else
         goSession.RaisePublicError "Error In IsChildReferencesExist, Unknown Database Engine", vbInformation, "Check If Child References Exist"
      End If
   End If
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   
   IsChildReferencesExist = False
   
   If UpperTableName = "MWWFFORMTEMPLATE" Then
      TemplateID = GetTemplateID(RecordKey)
   End If
   
   Do While loRs.BOF = False And loRs.EOF = False And IsChildReferencesExist = False
   
      If UpperTableName = "MWWFFORMTEMPLATE" And UCase(loRs!ChildCol) = "TEMPLATEID" Then
         IsChildReferencesExist = IsStrFkRecordsExist(loRs!ChildTab, loRs!ChildCol, TemplateID)
      Else
         IsChildReferencesExist = IsFkRecordsExist(loRs!ChildTab, loRs!ChildCol, RecordKey)
      End If
      
      If IsChildReferencesExist = True Then
         If mIsSilentChildRefError Then                           ' silent childRef
            msChildRefError = "Unable to Delete " & TableName & " because Foreign Key Reference exists in " & _
               Trim(loRs!ChildTab) & "." & Trim(loRs!ChildCol) & "Unable To Delete Record"
         Else
            ImprovedMsgBox "Unable to Delete " & TableName & " because Foreign Key Reference exists in " & _
               Trim(loRs!ChildTab) & "." & Trim(loRs!ChildCol), vbInformation, "Unable To Delete Record"
         End If

      End If
      loRs.MoveNext
   Loop
   
      
   
   CloseRecordset loRs
   Exit Function
FunctionError:

   If goSession.IsAccess Then
   
      IsChildReferencesExist = False
      CloseRecordset loRs
   
   Else
      goSession.RaisePublicError "General Error in mwSession.mwGUI.IsChildReferencesExist. ", Err.Number, Err.Description
      CloseRecordset loRs
      IsChildReferencesExist = True
   End If
End Function
Private Function GetTemplateID(mwwfFormTemplateKey As Long) As String
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
     
   Set loRs = New Recordset
   
   sSQL = "SELECT TemplateID FROM mwwfFormTemplate WHERE ID= " & mwwfFormTemplateKey
   
   loRs.CursorLocation = adUseClient
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount > 0 Then
      GetTemplateID = loRs!TemplateID
   Else
      GetTemplateID = ""
   End If
   
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwGUI.GetTemplateID ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Sub AccessDelayMsgBox(strPrompt As String, strTitle As String)
   
   On Error GoTo FunctionError
   
   If goSession.IsAccess Then
      DoEvents
      goSession.API.SleepApp 2000
   
      ImprovedMsgBox strPrompt, vbInformation, strTitle
      DoEvents
      
      goSession.API.SleepApp 1000
   End If
   Exit Sub
FunctionError:
   goSession.RaisePublicError "General Error in mwSession.mwGUI.AccessDelayMsgBox. ", Err.Number, Err.Description
End Sub

Public Function ImprovedMsgBox(Prompt As String, Optional Buttons As Integer, Optional Title As String) As Integer

'      Buttons
'      vbOKOnly                     0 Display OK button only.
'      vbOKCancel                   1 Display OK and Cancel buttons.
'      vbAbortRetryIgnore           2 Display Abort, Retry, and Ignore buttons.
'      vbYesNoCancel                3 Display Yes, No, and Cancel buttons.
'      vbYesNo                      4 Display Yes and No buttons.
'      vbRetryCancel                5 Display Retry and Cancel buttons.
'      vbCritical                  16 Display Critical Message icon.
'      vbQuestion                  32 Display Warning Query icon.
'      vbExclamation               48 Display Warning Message icon.
'      vbInformation               64 Display Information Message icon.
'      vbDefaultButton1             0 First button is default.
'      vbDefaultButton2           256 Second button is default.
'      vbDefaultButton3           512 Third button is default.
'      vbDefaultButton4           768 Fourth button is default.
'      vbApplicationModal           0 Application modal; the user must respond to the message box before continuing work in the current application.
'      vbSystemModal             4096 System modal; all applications are suspended until the user responds to the message box.
'      vbMsgBoxHelpButton       16384 Adds Help button to the message box
'      VbMsgBoxSetForeground    65536 Specifies the message box window as the foreground window
'      vbMsgBoxRight           524288 Text is right aligned
'      vbMsgBoxRtlReading     1048576 Specifies text should appear as right-to-left reading on Hebrew and Arabic systems

'      Return Values
'
'      vbOK       1 OK
'      vbCancel   2 Cancel
'      vbAbort    3 Abort
'      vbRetry    4 Retry
'      vbIgnore   5 Ignore
'      vbYes      6 Yes
'      vbNo       7 No


   Dim loform As frmImprovedMsgBox
   On Error GoTo FunctionError
   
   Set loform = New frmImprovedMsgBox
   loform.InitForm Prompt, Buttons, Title
   
   loform.Show vbModal
   ImprovedMsgBox = loform.WhichButton
   Unload loform
   KillObject loform
   Exit Function
FunctionError:
   MsgBox "General Error in mwSession.mwGUI.ImprovedMsgBox. " & " - " & Err.Number & " - " & Err.Description & _
          vbCrLf & vbCrLf & "Trying to display error message: " & Prompt
   
End Function
