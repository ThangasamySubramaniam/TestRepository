VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwCommercialOperatorWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Commercial Operator Work
' 7 May 2003 ms

Option Explicit


Public Function IsCommercialOperators() As Boolean
   Dim loRs As Recordset
   Dim strSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   strSQL = "select * from mwcCommercialOperator"
   loRs.Open strSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.RecordCount > 0 Then
      IsCommercialOperators = True
   Else
      IsCommercialOperators = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventLinkWork.IsCommercialOperators. ", err.Number, err.Description
   IsCommercialOperators = False
End Function


Public Function ChooseCommercialOperator() As Boolean
   Dim loform As frmSelector
   On Error GoTo FunctionError
   Set loform = New frmSelector
   If loform.InitForm("COMMERCIAL_OPERATOR") Then
      loform.Show vbModal
   End If
   Unload loform
   Set loform = Nothing
   ChooseCommercialOperator = True
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventLinkWork.IsCommercialOperators. ", err.Number, err.Description
   ChooseCommercialOperator = False
End Function
