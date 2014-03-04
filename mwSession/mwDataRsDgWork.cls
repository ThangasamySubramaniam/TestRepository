VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwDataRsDgWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' mwDataRsDgWork - Utility class to merge recordsets for datagrams.
' 1/2003 ms
'

Private moRsBaseline As Recordset
Private moRsFabricate As Recordset
Private mIsFabricatedRsOpen As Boolean

' Main Working Recordset

Public Property Set BaselineRS(ByVal oRs As Recordset)
   Dim loField As field
   On Error GoTo PropertyError
   If Not moRsBaseline Is Nothing Then
      goSession.RaiseError "Error in mwDataRsDgWork, BaselineRS already initialized."
      Exit Property
   End If
   If oRs.RecordCount <> 1 Then
      goSession.RaiseError "Error in mwDataRsDgWork, BaselineRS must be recordset with exactly one record."
      Exit Property
   End If
   Set moRsBaseline = oRs
   '
   ' Fabricated Recordset has all fields of the baseline
   '
   For Each loField In moRsBaseline.Fields
      moRsFabricate.Fields.Append loField.Name, loField.Type, loField.DefinedSize
   Next loField
   Exit Property
   '
PropertyError:
   goSession.RaiseError "General Error in mwDataRsDgWork.Property_Set_BaselineRS. ", err.Number, err.Description
End Property

Public Function AddColumn(ColumnName As String, Datatype As Long, DefinedSize As Long) As Boolean
   On Error GoTo FunctionError
   If mIsFabricatedRsOpen Then
      goSession.RaiseError "Error in mwDataRsDgWork.AddColumn, columns cannot be added once fabricated recordset is open. "
      AddColumn = False
      Exit Function
   End If
   moRsFabricate.Fields.Append ColumnName, Datatype, DefinedSize
   AddColumn = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataRsDgWork.AddColumn. ", err.Number, err.Description
End Function

Public Function AddColumnValue(ColumnName As String, ColumnValue As String) As Boolean
   On Error GoTo FunctionError
   If Not mIsFabricatedRsOpen Then
      goSession.RaiseError "Error in mwDataRsDgWork.AddColumnValue, Fabricated RS must be open first. "
      AddColumnValue = False
      Exit Function
   End If
   moRsFabricate.Fields(ColumnName).value = ColumnValue
   AddColumnValue = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataRsDgWork.AddColumnValue. ", err.Number, err.Description
End Function


Public Function OpenFabricatedRS() As Boolean
   Dim loField As field
   On Error GoTo FunctionError
   If mIsFabricatedRsOpen Then
      OpenFabricatedRS = True
      Exit Function
   End If
   moRsFabricate.Open
   moRsFabricate.AddNew
   For Each loField In moRsBaseline.Fields
      With loField
         If Not IsNull(.value) And Not IsEmpty(.value) Then
            moRsFabricate(.Name).value = .value
         End If
      End With
   Next loField
   mIsFabricatedRsOpen = True
   OpenFabricatedRS = True
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataRsDgWork.AddColumn. ", err.Number, err.Description
End Function

Public Function GetFabricatedRS() As Recordset
   On Error GoTo FunctionError
   If Not mIsFabricatedRsOpen Then
      goSession.RaiseError "Error in mwDataRsDgWork.GetFabricatedRS, Fabricated RS must be open first. "
      Exit Function
   End If
   Set GetFabricatedRS = moRsFabricate
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwDataRsDgWork.AddColumn. ", err.Number, err.Description
End Function



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



Private Sub Class_Initialize()
   Set moRsFabricate = New Recordset
   moRsFabricate.CursorLocation = adUseClient
End Sub

Private Sub Class_Terminate()
   'CloseRecordset moRsFabricate
End Sub

Public Function ClearAll() As Boolean
   CloseRecordset moRsFabricate
   Set moRsFabricate = New Recordset
   moRsFabricate.CursorLocation = adUseClient
   Set moRsBaseline = Nothing
   mIsFabricatedRsOpen = False
End Function
