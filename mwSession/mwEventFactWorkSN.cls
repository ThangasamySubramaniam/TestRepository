VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mwEventFactWorkSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' mwEventFactsSNWork - Support functions for managing Voyage Activities
' 12/27/03 ms
'

Option Explicit
Private mVoyNo As String
Private mSiteKey As Long
Dim WithEvents moRS As Recordset
Attribute moRS.VB_VarHelpID = -1
'
'Applies to ?
'
' Log Offsets
Const RS_ID = 0
Const RS_mwEventTypeKey = 1
Const RS_mwEventDetailKey = 2
Const RS_mwEventFactTypeSNKey = 3
Const RS_FactTitle = 4
Const RS_StartValue = 5
Const RS_EndValue = 6
Const RS_FactValue = 7
Const RS_RemarksValue = 8
Const RS_IsShoreCharge = 9
Const RS_DateSortValue = 10
Const RS_IsSent = 11
Const RS_WAITOFFHIRE = 12
Const RS_CON_MAINENGINE = 13
Const RS_CON_AUXENGINE = 14
Const RS_ZONE = 15
Const RS_StartValueDate = 16
Const RS_EndValueDate = 17
Const RS_mwcSitesKey = 18
Const RS_LinkedFactTypeSNKey = 19
Const RS_LinkedFactLogSNKey = 20

    Const MW_EVENT_TYPE_VOYAGE_PLAN = 2
    Const MW_EVENT_TYPE_POSASS6 = 3
    Const MW_EVENT_TYPE_PORT_CALL6 = 15
    Const MW_EVENT_TYPE_BERTH6 = 16
    Const MW_EVENT_TYPE_CARGO6 = 17
    Const MW_EVENT_TYPE_VRS_VOYEVENT = 20301
   
    Const MW_EVENT_TYPE_VOY_EXPLORER = 406


' Join with Type...

'Const RS_ID = 21
'Const RS_mwEventTypeKey = 22
'Const RS_mwEventFactCatSNKey = 23
Const RS_EventTypeSubcode = 24
Const RS_IsMandatory = 37
Const RS_IsStartAllowed = 39
Const RS_IsEndRequired = 43
Const RS_IsFactValueRequired = 44
Const RS_IsRemarkRequired = 45

Private mReducedFactCategoryFilter As String
Private mFactTypeFilter As Integer

Public Property Let VoyNo(ByVal Data As String)
    mVoyNo = Data
End Property

Public Property Get VoyNo() As String
    VoyNo = mVoyNo
End Property

Public Property Let ReducedFactCategoryFilter(sCategoryFilter As String)
    mReducedFactCategoryFilter = sCategoryFilter
End Property
Public Property Let FactTypeFilter(ByVal vData As Integer)
    mFactTypeFilter = vData
End Property

Public Property Get FactTypeFilter() As Integer
    FactTypeFilter = mFactTypeFilter
End Property

'Public Property Let A(ByVal Data As String)
'    mA = Data
'End Property

'Public Property Get A() As String
'    A = mA
'End Property


Public Function AutoCreate(EventType As Long, EventDetailKey As Long, _
 Optional SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loRsType As Recordset
   On Error GoTo FunctionError
   
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   Set loRsType = New Recordset
   loRsType.CursorLocation = adUseClient
   If SiteKey < 1 Then
      SiteKey = goSession.Site.SiteKey
   End If
   'sSQL = "SELECT * from mwEventFactTypeSN " & _
   '  " Where mwEventTypeKey=" & EventType & " And (IsSuggested <> 0 or IsMandatory <> 0) and IsActive <> 0" & _
   '  " ORDER BY DisplaySequence"
   sSQL = "SELECT mwEventFactTypeSN.* " & _
    " FROM mwcFleetSites, mwcFleets, mwEventFactTypeSN, mwEventFactTypeSnFleet where " & _
    " mwEventTypeKey=" & EventType & " And mwEventFactTypeSN.ID = mwEventFactTypeSnFleet.mwEventFactTypeSNKey and " & _
    " mwcFleets.ID = mwEventFactTypeSnFleet.mwcFleetsKey and " & _
    " mwcFleetSites.mwcFleetsKey = mwcFleets.ID and mwcFleetSites.mwcSitesKey=" & SiteKey & _
    " And (IsSuggested <> 0 or IsMandatory <> 0) and IsActive <> 0"
   
   If mFactTypeFilter = VRS_FACT_TYPE_CARGO_BOTH Then
      sSQL = sSQL & " AND (IsLoadPortFact <> 0 OR IsDiscPortFact <> 0)"
   ElseIf mFactTypeFilter = VRS_FACT_TYPE_CARGO_LOAD Then
      sSQL = sSQL & " AND IsLoadPortFact <> 0 "
   ElseIf mFactTypeFilter = VRS_FACT_TYPE_CARGO_DISCH Then
      sSQL = sSQL & " AND IsDiscPortFact <> 0"
   ElseIf (mFactTypeFilter = VRS_FACT_TYPE_CANAL Or mFactTypeFilter = VRS_FACT_TYPE_BUNKERING) Then
      sSQL = sSQL & " AND IsCanalBunkerFact <> 0"
   End If
   
   
   loRsType.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRsType.RecordCount < 1 Then
      CloseRecordset loRsType
      AutoCreate = True
      Exit Function
   End If
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLogSN where ID=-1"
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   '
   Do While Not loRsType.EOF
      '
      moRS.AddNew
      moRS!ID = goSession.MakePK("mwEventFactLogSN")
      moRS!mwEventTypeKey = EventType
      moRS!mwEventDetailKey = EventDetailKey
      moRS!mwEventFactTypeSNKey = loRsType!ID
      moRS!FactTitle = loRsType!FactTitle
      
      If ZeroNull(loRsType!LinkedFactTypeSNKey) > 0 Then
         moRS!LinkedFactTypeSNKey = loRsType!LinkedFactTypeSNKey
      End If
      
      '
      ' Set dummay date for Remarks...
      '
      If loRsType!IsStartAllowed = False And loRsType!IsEndAllowed = False Then
         moRS!DateSortValue = "01/01/2099"
      End If
      If SiteKey > 0 Then
         moRS!mwcSitesKey = SiteKey
      End If
      moRS.Update
      loRsType.MoveNext
   Loop
   CloseRecordset loRsType
   CloseRecordset moRS
   
   If goSession.IsAccess Then
      DoEvents
      goSession.API.SleepApp 3000
      DoEvents
   End If
   
   AutoCreate = True
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.AutoCreate. ", Err.Number, Err.Description
   AutoCreate = False
   CloseRecordset moRS
   mSiteKey = 0
End Function

Public Function AddRecord(EventType As Long, EventDetailKey As Long, Optional SiteKey As Long) As Long
   Dim loform As frmEventChooser
   Dim strFactTitle As String
   Dim sSQL As String
   Dim iPK As Long
   Dim iFactType As Long
   Dim iLinkedFactTypeSNKey As Long
   On Error GoTo FunctionError
   '
   ' Display list of possible activities in form...
   '
   If SiteKey > 0 Then
      mSiteKey = SiteKey
   End If
   Set loform = New frmEventChooser
   If SiteKey > 0 Then
      gAddEventFactSiteKey = SiteKey
   Else
      gAddEventFactSiteKey = goSession.Site.SiteKey
   End If
   
   ' reduced fact category filter
   loform.ReducedFactCategoryFilter = mReducedFactCategoryFilter
   loform.FactTypeFilter = mFactTypeFilter
   
   If Not loform.EventFactSNInitChooser(EventType, EventDetailKey) Then
      AddRecord = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   
   loform.Show vbModal
   If loform.IsCancelled Then
      AddRecord = -1
      Unload loform
      Set loform = Nothing
      Exit Function
   End If
   '
   CloseRecordset moRS
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open "Select * from mwEventFactLogSN where ID=-1", goCon, adOpenDynamic, adLockOptimistic
   iPK = goSession.MakePK("mwEventFactLogSN")
   
   Do While loform.IsSelectedRows
      iFactType = loform.FetchNextSelectedFactSN(strFactTitle, iLinkedFactTypeSNKey)
      If iFactType < 1 Then
         Exit Do
      End If
      moRS.AddNew
      moRS!ID = iPK
      moRS!mwEventTypeKey = EventType
      moRS!mwEventDetailKey = EventDetailKey
      moRS!mwEventFactTypeSNKey = iFactType
      moRS!FactTitle = strFactTitle
      
      If iLinkedFactTypeSNKey > 0 Then
         moRS!LinkedFactTypeSNKey = iLinkedFactTypeSNKey
      End If
      
      If SiteKey > 0 Then
         ' col 18          change from moRs to moRs & replicate
          moRS!mwcSitesKey = SiteKey
      End If
      
      If IsRemarkOnly(iFactType) Then
         moRS!DateSortValue = "01/01/2099"
      End If
      moRS.Update
      iPK = iPK + 1
   Loop
   If iPK > 0 Then
      goSession.UpdatePrimaryKeySequence "mwEventFactLogSN", iPK
   End If
   AddRecord = iPK
   Unload loform
   If goSession.IsAccess Then
      DoEvents
      goSession.API.SleepApp 4000
      DoEvents
   End If
   Set loform = Nothing
   
   CloseRecordset moRS
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.AddRecords. ", Err.Number, Err.Description
   AddRecord = -1
   CloseRecordset moRS
   mSiteKey = 0
End Function



Public Function FetchFactShape(EventType As Long, EventDetailKey As Long, IsDateOrder As Boolean) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   Dim strOrder As String
   Dim sRootSQL As String
   Dim sNodeSQL As String
   Dim sAppendSQL As String
   Dim nMaxLevel As Integer
   Dim nLevel As Integer
   Dim nCount As Integer

   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient

'   sSQL = "SHAPE {SELECT distinct mwEventFactCatSN.* FROM mwEventFactCatSN RIGHT JOIN " & _
'    " mwEventFactTypeSN ON mwEventFactCatSN.ID = mwEventFactTypeSN.mwEventFactCatSNKey " & _
'    " WHERE mwEventFactTypeSN.mwEventTypeKey=" & EventType & "}  AS Command1 " & _
'    "APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'    " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'    " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'    " ORDER BY "

'   If goSession.IsOracle Then
'      sSQL = "SHAPE {SELECT distinct mwEventFactCatSN.* " & _
'         " FROM mwEventFactCatSN , mwEventFactTypeSN" & _
'         " WHERE mwEventFactCatSN.ID(+) = mwEventFactTypeSN.mwEventFactCatSNKey AND " & _
'         " mwEventFactTypeSN.mwEventTypeKey = " & EventType & _
'         " and mwEventFactTypeSN.IsActive<>0  order by mwEventFactCatSN.DisplaySequence}  AS Command1 " & _
'         " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
'         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
'         " mwEventFactLogSN.mwEventTypeKey = " & EventType & " And " & _
'         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & _
'         " ORDER BY "
'   Else
'      sSQL = "SHAPE {SELECT distinct mwEventFactCatSN.* FROM mwEventFactCatSN RIGHT JOIN " & _
'      " mwEventFactTypeSN ON mwEventFactCatSN.ID = mwEventFactTypeSN.mwEventFactCatSNKey " & _
'      " WHERE mwEventFactTypeSN.mwEventTypeKey=" & EventType & _
'      " and mwEventFactTypeSN.IsActive<>0 order by mwEventFactCatSN.DisplaySequence}  AS Command1 " & _
'      " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'      " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'      " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'      " ORDER BY "
'   End If
'
'   If IsDateOrder Then
'      sSQL = sSQL & "mwEventFactLogSN.StartValue, mwEventFactLogSN.ID} " & _
'     " AS Command2 RELATE 'ID' TO 'mwEventFactCatSNKey') AS Command2"
'   Else
'      sSQL = sSQL & "mwEventFactTypeSN.DisplaySequence, mwEventFactLogSN.ID} " & _
'     " AS Command2 RELATE 'ID' TO 'mwEventFactCatSNKey') AS Command2"
'   End If

'NSB-108 Implementation of second hiearchy level in VRS Facts Tab

   If IsDateOrder Then
      strOrder = " mwEventFactLogSN.StartValue, mwEventFactLogSN.ID} #Command# "
   Else
      strOrder = " mwEventFactTypeSN.DisplaySequence, mwEventFactLogSN.ID} #Command# "
   End If

   If goSession.IsOracle Then
      sRootSQL = "SHAPE {SELECT mwEventFactCatSN.*,'' as FACTCATEGORYBAND " & _
         " FROM mwEventFactCatSN " & _
         " WHERE ParentEventFactCatSNKey is null order by mwEventFactCatSN.DisplaySequence} " & _
         " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey = " & EventType & " And " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & _
         " ORDER BY " & strOrder & _
         " RELATE 'ID' TO 'mwEventFactCatSNKey') "

      sNodeSQL = "(SHAPE {SELECT mwEventFactCatSN.*,'' as FACTCATEGORYBAND " & _
         " FROM mwEventFactCatSN " & _
         " WHERE ParentEventFactCatSNKey is not null  order by mwEventFactCatSN.DisplaySequence} " & _
         " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey = " & EventType & " And " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & _
         " ORDER BY " & strOrder & _
         " RELATE 'ID' TO 'mwEventFactCatSNKey')) "
   Else
      sRootSQL = "SHAPE {SELECT mwEventFactCatSN.*,'' as FACTCATEGORYBAND FROM mwEventFactCatSN " & _
      " WHERE ParentEventFactCatSNKey is null order by mwEventFactCatSN.DisplaySequence}  " & _
      " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
      " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
      " ORDER BY " & strOrder & _
      " RELATE 'ID' TO 'mwEventFactCatSNKey') "

      sNodeSQL = "(SHAPE {SELECT mwEventFactCatSN.*,'' as FACTCATEGORYBAND FROM mwEventFactCatSN " & _
      " WHERE ParentEventFactCatSNKey is not null order by mwEventFactCatSN.DisplaySequence}  " & _
      " APPEND ({SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
      " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
      " ORDER BY " & strOrder & _
      " RELATE 'ID' TO 'mwEventFactCatSNKey'))"
   End If

   nMaxLevel = 3     'Maxmimum depth devel

   sAppendSQL = vbNullString

   For nLevel = 1 To nMaxLevel
      If nLevel = nMaxLevel Then
         sAppendSQL = sAppendSQL & "#Command#" & " APPEND ( " & sNodeSQL
      Else
         sAppendSQL = sAppendSQL & "#Command#" & " APPEND (( SHAPE " & sNodeSQL
      End If
   Next nLevel

   For nLevel = 1 To nMaxLevel
      If nLevel = nMaxLevel Then
         sAppendSQL = sAppendSQL & " RELATE 'ID' TO 'ParentEventFactCatSNKey')"
      Else
         sAppendSQL = sAppendSQL & " RELATE 'ID' TO 'ParentEventFactCatSNKey'))"
      End If
   Next nLevel

   sSQL = "SHAPE ( " & sRootSQL & ") " & sAppendSQL

   nCount = 1
   Do While InStr(1, sSQL, "#Command#", vbTextCompare) > 0
      nCount = nCount + 1
      sSQL = Replace(sSQL, "#Command#", "as C" & CStr(nCount), 1, 1, vbTextCompare)
   Loop

   loRs.Open sSQL, goConShape, adOpenDynamic, adLockOptimistic
   If loRs.RecordCount > 0 Then
      loRs.Fields(loRs.Fields.Count - 1).value.Properties("Unique Table").value = "mwEventFactLogSN"
   End If
   Set FetchFactShape = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchFactShape. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function
Public Function FetchDelayFactsRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'     " and mwEventFactTypeSN.IsDelayActivity <> 0 " & _
'     " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID;"
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
     " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
     " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
     " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
     " mwEventFactTypeSN.IsDelayActivity <> 0 " & _
     " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchDelayFactsRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchDelayFactsRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchColumnNameRS(EventType As Long, EventDetailKey As Long, ColumnName As String) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " WHERE (mwEventFactLogSN.mwEventTypeKey=" & EventType & ") AND mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'     " AND (mwEventFactTypeSN.StartColumnName='" & ColumnName & "'  OR mwEventFactTypeSN.EndColumnName='" & ColumnName & _
'     "' OR mwEventFactTypeSN.RemarksColumnName='" & ColumnName & "')"
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
      " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
      " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
      " (mwEventFactTypeSN.StartColumnName = '" & ColumnName & "'  OR " & _
      "  mwEventFactTypeSN.EndColumnName = '" & ColumnName & "' OR " & _
      "  mwEventFactTypeSN.RemarksColumnName = '" & ColumnName & "')"
   
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchColumnNameRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchColumnNameRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchExpandedFieldList(EventType As Long, EventDetailKey As Long) As Collection
   Dim sSQL As String
   Dim loRs As Recordset
   Dim s As String
   Dim loCol As Collection
   On Error GoTo FunctionError
   Set loCol = New Collection
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT StartColumnName, EndColumnName, RemarksColumnName, IsDelayActivity " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " WHERE mwEventFactLogSN.mwEventTypeKey=" & EventType & " AND mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey
   
   sSQL = "SELECT StartColumnName, EndColumnName, RemarksColumnName, IsDelayActivity " & _
     " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
     " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
     " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey
   
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      On Error Resume Next
      If Not loRs!IsDelayActivity Then
         If Not IsNull(loRs!StartColumnName) Then
            If Not IsInCollection(loCol, loRs!StartColumnName) Then
               s = loRs!StartColumnName
               loCol.Add Item:=s, Key:=s
            End If
         End If
         If Not IsNull(loRs!EndColumnName) Then
            If Not IsInCollection(loCol, loRs!EndColumnName) Then
               s = loRs!EndColumnName
               loCol.Add Item:=s, Key:=s
            End If
         End If
         If Not IsNull(loRs!RemarksColumnName) Then
            If Not IsInCollection(loCol, loRs!RemarksColumnName) Then
               s = loRs!RemarksColumnName
               loCol.Add Item:=s, Key:=s
            End If
         End If
      End If
      On Error GoTo FunctionError
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   Set FetchExpandedFieldList = loCol
   Set loCol = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchExpandedFieldList. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function




Public Function Submit(EventType As Long, EventDetailKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   Dim loform As mwForm
   Dim strFile As String
   On Error GoTo FunctionError
   MsgBox "mwEventFactWorkSN.Submit stump code"
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.Submit. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function



Public Function DeleteUnusedFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   
   On Error GoTo FunctionError
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLogSN where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteUnusedFacts = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   ' delete facts not used...
   Do While Not moRS.EOF
      If IsNull(moRS.Fields(RS_StartValue).value) And IsNull(moRS.Fields(RS_EndValue).value) _
       And IsNull(moRS.Fields(RS_FactValue).value) _
       And IsNull(moRS.Fields(RS_RemarksValue).value) Then
         ' Nothing filled in, delete it...
         
         mSiteKey = ZeroNull(moRS!mwcSitesKey)
         moRS.Delete
      End If
      moRS.MoveNext
   Loop
   goSession.CloseRecordset moRS
   DeleteUnusedFacts = True
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.DeleteUnusedFacts. ", Err.Number, Err.Description
   DeleteUnusedFacts = False
   mSiteKey = 0
End Function

Public Function DeleteFact(FactKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   On Error GoTo FunctionError
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLogSN where ID=" & FactKey
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteFact = True
      goSession.CloseRecordset moRS
      Exit Function
   End If

   If BoolNull(moRS!IsSent) = False Then
      mSiteKey = ZeroNull(moRS!mwcSitesKey)
      moRS.Delete
      DeleteFact = True
   Else
      goSession.LogIt mwl_Workstation, mwl_Warning, "Logging Warning mwEventFactWorkSN, Fact can't be deleted because it has already been Submitted."
      DeleteFact = False
   End If
   goSession.CloseRecordset moRS
   
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.DeleteFact. ", Err.Number, Err.Description
   DeleteFact = False
   mSiteKey = 0
End Function



Public Function IsSubmittedFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   'sSQL = "select count(*) from mwEventFactLogSN where mwEventTypeKey=" & EventType & _
   '  " and mwEventDetailKey=" & EventKey & " and IsSubmitted=True"
   sSQL = "select count(*) from mwEventFactLogSN WHERE mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey & " AND (IsSubmitted is not null AND IsSubmitted <> 0)"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   ' Got Forms ?
   If loRs.Fields(0).value > 0 Then
      IsSubmittedFacts = True
   Else
      IsSubmittedFacts = False
   End If
   goSession.CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.IsSubmittedFacts. ", Err.Number, Err.Description
   IsSubmittedFacts = False
End Function



'
' Expand a recordset with all EventFacts fields for a particular event detail record
'
Public Function ExpandRecordset(ByRef oRsEventDetail As Recordset, EventTypeKey As Long, _
  EventDetailKey As Long) As Recordset
   Dim loRsFact As Recordset
   Dim loRs As Recordset
   Dim loSkipList As Collection
   Dim loField As field
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Staging/Validation
   If oRsEventDetail.State = adStateClosed Then
      goSession.RaiseError "Error in mwEventFactWorkSN.ExpandRecordset, Recordset is closed."
      Exit Function
   End If
   ' Any records in there ?
   If oRsEventDetail.RecordCount < 1 Then
      goSession.RaiseError "Error in mwEventFactWorkSN.ExpandRecordset, no records in Recordset."
      Exit Function
   End If
   '
   ' Create new recordset...
   '
   Set loSkipList = New Collection
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   ' Add columns to fabricated recordset
   For Each loField In oRsEventDetail.Fields
      ' Skip shape RS last column...
      If loField.Type <> adChapter Then
         'loRS.Fields.Append loField.Name, loField.Type, loField.DefinedSize
         If goSession.IsOracle And loField.Type = adNumeric Then
            loRs.Fields.Append loField.Name, adInteger, 4
         Else
            loRs.Fields.Append loField.Name, loField.Type, loField.DefinedSize
         End If
      End If
   Next loField
   ' Get the activities
   Set loRsFact = FetchFactRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   
   Do While Not loRsFact.EOF
      On Error Resume Next
      If Not loRsFact!IsDelayActivity Then
         ' Start Date/Time
         If Not IsNull(loRsFact!StartColumnName) Then
            'Debug.Print loRsFact.Fields("StartColumnName").value
            loRs.Fields.Append loRsFact.Fields("StartColumnName").value, loRsFact.Fields("StartValueDate").Type, loRsFact.Fields("StartValueDate").DefinedSize
            If Err Then
               If Err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("StartColumnName").value, Key:=loRsFact.Fields("StartColumnName").value
               Else
                  goSession.RaiseError "General Error in mwEventFactWorkSN.ExpandRecordset. ", Err.Number, Err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
         'End Date/Time
         If Not IsNull(loRsFact!EndColumnName) Then
            loRs.Fields.Append loRsFact.Fields("EndColumnName").value, loRsFact.Fields("EndValueDate").Type, loRsFact.Fields("EndValueDate").DefinedSize
            If Err Then
               If Err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("EndColumnName").value, Key:=loRsFact.Fields("EndColumnName").value
               Else
                  goSession.RaiseError "General Error in mwEventFactWorkSN.ExpandRecordset. ", Err.Number, Err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
         ' Fact Value column...
         If Not IsNull(loRsFact!FactValueColumnName) And Not IsNull(loRsFact!FactValueAdoType) _
                And Not IsNull(loRsFact.Fields("FactValueAdoDefinedSize").value) Then
            loRs.Fields.Append loRsFact.Fields("FactValueColumnName").value, loRsFact.Fields("FactValueAdoType").value, loRsFact.Fields("FactValueAdoDefinedSize").value
            If Err Then
               If Err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("FactValueColumnName").value, Key:=loRsFact.Fields("FactValueColumnName").value
               Else
                  goSession.RaiseError "General Error in mwEventFactWorkSN.ExpandRecordset. ", Err.Number, Err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
         ' Remarks/Other Information
         If Not IsNull(loRsFact!RemarksColumnName) Then
            loRs.Fields.Append loRsFact.Fields("RemarksColumnName").value, loRsFact.Fields("RemarksValue").Type, loRsFact.Fields("RemarksValue").DefinedSize
            If Err Then
               If Err.Number = 3367 Then
                  loSkipList.Add Item:=loRsFact.Fields("RemarksColumnName").value, Key:=loRsFact.Fields("RemarksColumnName").value
               Else
                  goSession.RaiseError "General Error in mwEventFactWorkSN.ExpandRecordset. ", Err.Number, Err.Description
                  CloseRecordset loRsFact
                  CloseRecordset loRs
                  Set ExpandRecordset = oRsEventDetail
               End If
            End If
         End If
      End If
      On Error GoTo FunctionError
      loRsFact.MoveNext
   Loop
   On Error GoTo FunctionError
   '
   ' Append data to fabricated recordset...
   '
   loRs.Open
   loRs.AddNew
   ' Event Detail...
   For Each loField In oRsEventDetail.Fields
      With loField
         ' Skip shape RS last column...
         If loField.Type <> adChapter Then
            If Not IsNull(.value) And Not IsEmpty(.value) Then
               loRs(.Name).value = .value
            End If
         End If
      End With
   Next loField
   
   loRs.Update
   '
   ' Add "Fact" record fields to flat fabricated recordset
   '
   If loRsFact.RecordCount > 0 Then
      loRsFact.MoveFirst
      Do While Not loRsFact.EOF
         If Not loRsFact!IsDelayActivity Then
'            ' Start Date/Time
'            If Not IsNull(loRsFact!StartColumnName) And Not IsNull(loRsFact!StartValueDate) _
'              And Not IsInCollection(loSkipList, loRsFact!StartColumnName) Then
'               loRs.Fields(loRsFact.Fields("StartColumnName").value).value = loRsFact!StartValueDate
'            End If
'            ' End Date/Time
'            If Not IsNull(loRsFact!EndColumnName) And Not IsNull(loRsFact!EndValueDate) _
'              And Not IsInCollection(loSkipList, loRsFact!EndColumnName) Then
'               loRs.Fields(loRsFact.Fields("EndColumnName").value).value = loRsFact!EndValueDate
'            End If
'            ' FactValue
'            If Not IsNull(loRsFact.Fields("FactValueAdoDefinedSize").value) Then
'               If Not IsNull(loRsFact!FactValueColumnName) And Not IsNull(loRsFact!FactValue) _
'                 And Not IsInCollection(loSkipList, loRsFact!FactValueColumnName) And Not IsNull(loRsFact!FactValueAdoType) Then
'                  If loRsFact!FactValueAdoType = adDouble Then
'                     loRs.Fields(loRsFact.Fields("FactValueColumnName").value).value = CDbl(loRsFact!FactValue)
'                  Else
'                     loRs.Fields(loRsFact.Fields("FactValueColumnName").value).value = loRsFact!FactValue
'                  End If
'               End If
'            End If
'            ' Remarks/Other Information
'            If Not IsNull(loRsFact!RemarksColumnName) And Not IsNull(loRsFact!RemarksValue) _
'              And Not IsInCollection(loSkipList, loRsFact!RemarksColumnName) Then
'               loRs.Fields(loRsFact.Fields("RemarksColumnName").value).value = loRsFact!RemarksValue
'            End If
'         End If
            ' Start Date/Time
            If Not IsNull(loRsFact!StartColumnName) And Not IsNull(loRsFact!StartValueDate) Then
               If loRs.Fields.Item(loRsFact.Fields("StartColumnName").value).Type = loRsFact.Fields("StartValueDate").Type Then
                  loRs.Fields(loRsFact.Fields("StartColumnName").value).value = loRsFact!StartValueDate
               End If
            End If
            ' End Date/Time
            If Not IsNull(loRsFact!EndColumnName) And Not IsNull(loRsFact!EndValueDate) Then
               If loRs.Fields.Item(loRsFact.Fields("EndColumnName").value).Type = loRsFact.Fields("EndValueDate").Type Then
                  loRs.Fields(loRsFact.Fields("EndColumnName").value).value = loRsFact!EndValueDate
               End If
            End If
            ' FactValue
            If Not IsNull(loRsFact.Fields("FactValueAdoDefinedSize").value) Then
               If Not IsNull(loRsFact!FactValueColumnName) And Not IsNull(loRsFact!FactValue) _
                 And Not IsNull(loRsFact!FactValueAdoType) Then
                  If loRs.Fields.Item(loRsFact.Fields("FactValueColumnName").value).Type = loRsFact.Fields("FactValueAdoType").value Then
                     If loRsFact!FactValueAdoType = adDouble Then
                        loRs.Fields(loRsFact.Fields("FactValueColumnName").value).value = CDbl(loRsFact!FactValue)
                     Else
                        loRs.Fields(loRsFact.Fields("FactValueColumnName").value).value = loRsFact!FactValue
                     End If
                  End If
               End If
            End If
            ' Remarks/Other Information
            If Not IsNull(loRsFact!RemarksColumnName) And Not IsNull(loRsFact!RemarksValue) Then
               If loRs.Fields.Item(loRsFact.Fields("RemarksColumnName").value).Type = loRsFact.Fields("RemarksValue").Type Then
                  loRs.Fields(loRsFact.Fields("RemarksColumnName").value).value = loRsFact!RemarksValue
               End If
            End If
         End If

         loRsFact.MoveNext
      Loop
      loRs.Update
   End If
   CloseRecordset loRsFact
   ' That's all folks...
   Set ExpandRecordset = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   'Resume Next
   goSession.RaisePublicError "General Error in mwEventFactWorkSN.ExpandRecordset. ", Err.Number, Err.Description
   goSession.RaiseWorkflowStatus "Reverting to original recordset, no merge taking place..."
   CloseRecordset loRsFact
   CloseRecordset loRs
   ' revert to original recordset
   Set ExpandRecordset = oRsEventDetail
End Function



Public Function LockEventFacts(EventTypeKey As Long, _
  EventDetailKey As Long, Optional EventTypeSubCode As String) As Boolean
   Dim loRsFact As Recordset
   Dim sSQL As String
   Dim strLockCode
   On Error GoTo FunctionError
   ' Get the activities
   Set loRsFact = FetchFactRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   Do While Not loRsFact.EOF
      If EventTypeSubCode = "" Or EventTypeSubCode = loRsFact!EventTypeSubCode Then
         loRsFact!IsSent = True
         loRsFact.Update
      End If
      loRsFact.MoveNext
   Loop
   CloseRecordset loRsFact
   LockEventFacts = True
   ' That's all folks...
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactWorkSN.LockEventFacts. ", Err.Number, Err.Description
   LockEventFacts = False
End Function

Public Function UnLockEventFacts(EventTypeKey As Long, _
  EventDetailKey As Long) As Boolean
   Dim loRsFact As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   ' Get the activities
   Set loRsFact = FetchFactRS(EventTypeKey, EventDetailKey)
   ' Add expanded columns to recordset...
   Do While Not loRsFact.EOF
      loRsFact!IsSent = False
      loRsFact.Update
      loRsFact.MoveNext
   Loop
   CloseRecordset loRsFact
   UnLockEventFacts = True
   ' That's all folks...
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactWorkSN.UnLockEventFacts. ", Err.Number, Err.Description
   UnLockEventFacts = False
End Function


Public Function IsEventFactTypesExist(EventType As Long) As Boolean
   Dim loRs As Recordset
   Dim sSQL As String
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "select count(*) from mwEventFactTypeSN where mwEventTypeKey=" & EventType & " and IsActive<>0"
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs.Fields(0).value > 0 Then
      IsEventFactTypesExist = True
   Else
      IsEventFactTypesExist = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFactWorkSN.IsEventFactTypesExist. ", Err.Number, Err.Description
End Function



Public Function FetchShipnetActivityRS(EventType As Long, EventDetailKey As Long, _
 Optional EventTypeSubCode As String, Optional IsMatchFactDispSeq As Boolean) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If Trim(EventTypeSubCode) <> "" Then
      sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey  =" & EventType & " AND " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
         " not mwEventFactTypeSN.SN_ActivityTypeCode is Null AND " & _
         " mwEventFactTypeSN.SN_ActivityTypeCode <> 'L' AND " & _
         " mwEventFactTypeSN.EventTypeSubcode ='" & EventTypeSubCode & "'"
      If IsMatchFactDispSeq Then
         sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      Else
         sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      End If
   Else
      sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
         " not mwEventFactTypeSN.SN_ActivityTypeCode is Null AND " & _
         " mwEventFactTypeSN.SN_ActivityTypeCode <> 'L' "
      If IsMatchFactDispSeq Then
         sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      Else
         sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      End If
   End If

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchShipnetActivityRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchShipnetActivityRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchShipnetLaytimeRS(EventType As Long, EventDetailKey As Long, _
 Optional EventTypeSubCode As String, Optional IsMatchFactDispSeq As Boolean) As Recordset
 
   'By N.Angelakis on 20 april 2010
   'DEV 1608 - VRS6: Add Laytime Facts to Cargo Facts in POSPORT datagram from Ship to Shore
   'Current SN Activity Types: 'At Sea' (S), Port (P), Cargo (C), Laytime (L).
   'Added new  SN Activity Type option (X) 'Laytime & Cargo".
   'modified query to search for mwEventFactTypeSN.SN_ActivityTypeCode = 'X'
   ' " (mwEventFactTypeSN.SN_ActivityTypeCode = 'L' OR mwEventFactTypeSN.SN_ActivityTypeCode = 'X') AND " & _

 
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   If Trim(EventTypeSubCode) <> "" Then
      sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey  =" & EventType & " AND " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
         " (mwEventFactTypeSN.SN_ActivityTypeCode = 'L' OR mwEventFactTypeSN.SN_ActivityTypeCode = 'X') AND " & _
         " mwEventFactTypeSN.EventTypeSubcode ='" & EventTypeSubCode & "'"
      If IsMatchFactDispSeq Then
         sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      Else
         sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      End If
   Else
      sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
         " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
         " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
         " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
         " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
         "  " & _
         " (mwEventFactTypeSN.SN_ActivityTypeCode = 'L' OR mwEventFactTypeSN.SN_ActivityTypeCode = 'X')"
      If IsMatchFactDispSeq Then
         sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      Else
         sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
      End If
   End If

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchShipnetLaytimeRS = loRs
   Set loRs = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchShipnetLaytimeRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function




Private Function IsRemarkOnly(iFactType As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * from mwEventFactTypeSN where ID=" & str(iFactType)
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If loRs!IsStartAllowed = False And loRs!IsEndAllowed = False And loRs!IsFactAllowed = False Then
      IsRemarkOnly = True
   Else
      IsRemarkOnly = False
   End If
   CloseRecordset loRs
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.IsRemarkOnly. ", Err.Number, Err.Description
   CloseRecordset loRs
   IsRemarkOnly = False
End Function

Public Function IsValidFacts(EventType As Long, EventDetailKey As Long, _
  ByRef ErrorString As String, Optional EventTypeSubCode As String) As Boolean
   Dim lIsError As Boolean
   Dim loRs As Recordset
   On Error GoTo FunctionError
   ' Loop through all records, performing validation...
   lIsError = False
   Set loRs = FetchFactRS(EventType, EventDetailKey)
   If loRs.RecordCount < 1 Then
      IsValidFacts = True
      CloseRecordset loRs
      Exit Function
   End If
   Do While Not loRs.EOF
      
      ' Arrival Messages require EventTypeSubcode to decide which records to check...
      If EventTypeSubCode = "" Or EventTypeSubCode = loRs.Fields(RS_EventTypeSubcode).value Then
         If loRs.Fields(RS_IsMandatory).value Then
            ' Start Date...
            If loRs.Fields(RS_IsStartAllowed).value And _
              IsNull(loRs.Fields(RS_StartValue).value) Then
                ErrorString = ErrorString & _
                  "Mandatory Start Date/Time Missing: " & loRs.Fields(RS_FactTitle).value & vbCrLf
                lIsError = True
            End If
            '
         
         End If
         ' Is end date/time required ?
         If loRs.Fields(RS_IsEndRequired).value Then
            '
            If IsNull(loRs.Fields(RS_EndValue).value) Then
               ErrorString = ErrorString & _
                 "Required End Date/Time Missing: " & loRs.Fields(RS_FactTitle).value & vbCrLf
               lIsError = True
            End If
         End If
         
         
         ' Is Fact Value ?
         If loRs.Fields(RS_IsFactValueRequired).value = True Then
            '
            If IsNull(loRs.Fields(RS_FactValue).value) Then
               ErrorString = ErrorString & _
                 "Required Entry Missing: " & loRs.Fields(RS_FactTitle).value & vbCrLf
               lIsError = True
            End If
         End If
         
         ' Is remark required ?
         If loRs.Fields(RS_IsRemarkRequired).value = True Then
            '
            If IsNull(loRs.Fields(RS_RemarksValue).value) Then
               ErrorString = ErrorString & _
                 "Required Entry Missing: " & loRs.Fields(RS_FactTitle).value & vbCrLf
               lIsError = True
            End If
         End If
      End If
      loRs.MoveNext
   Loop
   CloseRecordset loRs
   IsValidFacts = Not lIsError
   Exit Function
FunctionError:
   goSession.RaisePublicError "General Error in mwEventFacts.IsValidFacts. ", Err.Number, Err.Description
   IsValidFacts = False
End Function


Public Function FetchFactLogRecordRS(EventTypeKey As Long, EventDetailKey As Long, FactTypeKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " Where mwEventFactLogSN.mwEventTypeKey=" & EventTypeKey & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'     " and  mwEventFactLogSN.mwEventFactTypeSNKey=" & FactTypeKey
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
      " mwEventFactLogSN.mwEventTypeKey = " & EventTypeKey & " AND " & _
      " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & " AND " & _
      " mwEventFactLogSN.mwEventFactTypeSNKey = " & FactTypeKey

   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchFactLogRecordRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchFactLogRecordRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function DeleteByFilter(EventType As Long, EventDetail As Long, FilterNo As Integer, FilterCode As String) As Boolean
   Dim sSQL As String
   
   On Error GoTo FunctionError
   
'   sSQL = "SELECT mwEventFactLogSN.* FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN " & _
'     " ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " WHERE mwEventFactLogSN.mwEventTypeKey=" & EventType & " AND mwEventFactLogSN.mwEventDetailKey=" & EventDetail
   
   sSQL = "SELECT mwEventFactLogSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
      " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
      " mwEventFactLogSN.mwEventDetailKey = " & EventDetail
   
   Select Case FilterNo
      Case 1
         sSQL = sSQL & " AND Filter1='" & FilterCode & "'"
      Case 2
         sSQL = sSQL & " AND Filter2='" & FilterCode & "'"
      Case 3
         sSQL = sSQL & " AND Filter3='" & FilterCode & "'"
      Case Else
         goSession.RaiseError "General Error in mwEventFactWorkSN.ReduceFactLogList, invalid filter number. "
         DeleteByFilter = False
         Exit Function
   End Select
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   moRS.Properties("Unique Table").value = "mwEventFactLogSN"
   Do While Not moRS.EOF
      mSiteKey = ZeroNull(moRS!mwcSitesKey)
      moRS.Delete
      moRS.MoveNext
   Loop
   CloseRecordset moRS
   DeleteByFilter = True
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.DeleteByFilter. ", Err.Number, Err.Description
   mSiteKey = 0
End Function

Public Function FetchFactLogRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   sSQL = "SELECT * from mwEventFactLogSN Where mwEventTypeKey=" & EventType & _
    " And mwEventDetailKey=" & EventDetailKey
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   Set FetchFactLogRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchFactLogRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function



Public Function DeleteAllFacts(EventType As Long, EventKey As Long) As Boolean
   On Error GoTo FunctionError
   Dim sSQL As String
   
   On Error GoTo FunctionError
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLogSN where mwEventTypeKey=" & EventType & _
     " and mwEventDetailKey=" & EventKey
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got Forms ?
   If moRS.RecordCount < 1 Then
      DeleteAllFacts = True
      goSession.CloseRecordset moRS
      Exit Function
   End If
   ' delete facts not used...
   Do While Not moRS.EOF
      mSiteKey = ZeroNull(moRS!mwcSitesKey)
      moRS.Delete
      moRS.MoveNext
   Loop
   goSession.CloseRecordset moRS
   DeleteAllFacts = True
   mSiteKey = 0
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.DeleteAllFacts. ", Err.Number, Err.Description
   DeleteAllFacts = False
   mSiteKey = 0
End Function




Public Function FetchFactRS(EventType As Long, EventDetailKey As Long) As Recordset
   Dim strSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   strSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & EventDetailKey & _
'     " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID;"
   
   strSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
     " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
     " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
     " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
     " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & _
     " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   
   loRs.Open strSQL, goCon, adOpenDynamic, adLockOptimistic
   loRs.Properties("Unique Table").value = "mwEventFactLogSN"
   Set FetchFactRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchFactRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function


Public Function FetchFactsReadOnly(EventType As Long, EventDetailKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
'   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
'     " FROM mwEventFactTypeSN INNER JOIN mwEventFactLogSN ON mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey " & _
'     " Where mwEventFactLogSN.mwEventTypeKey=" & EventType & " And mwEventFactLogSN.mwEventDetailKey=" & _
'     EventDetailKey
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND " & _
      " mwEventFactLogSN.mwEventTypeKey = " & EventType & " AND " & _
      " mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey
   
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchFactsReadOnly = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchFactsReadOnly. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function IsEventFactCatHavingValidFacts(EventType As Long, EventFactCatKey As Long, SiteKey As Long) As Boolean
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT mwEventFactTypeSN.* " & _
    " FROM mwcFleetSites, mwcFleets, mwEventFactTypeSN, mwEventFactTypeSnFleet where " & _
    " mwEventTypeKey=" & EventType & " and  mwEventFactTypeSN.ID = mwEventFactTypeSnFleet.mwEventFactTypeSNKey and " & _
    " ISACTIVE <> 0 AND mwEventFactTypeSN.mwEventFactCatSNKey = " & EventFactCatKey & " AND " & _
    " mwcFleets.ID = mwEventFactTypeSnFleet.mwcFleetsKey and " & _
    " mwcFleetSites.mwcFleetsKey = mwcFleets.ID and mwcFleetSites.mwcSitesKey=" & SiteKey
    
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   If Not (loRs.BOF And loRs.EOF) Then
      IsEventFactCatHavingValidFacts = True
   End If
   CloseRecordset loRs
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.IsEventFactCatHavingValidFacts. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Sub UpdateLinkedFacts(EventType As Long, EventDetailKey As Long)
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo ErrorHandler
   
   Set moRS = New Recordset
   moRS.CursorLocation = adUseClient
   sSQL = "select * from mwEventFactLogSN " & _
          " where mwEventTypeKey=" & EventType & _
          " and mwEventDetailKey=" & EventDetailKey & _
          " AND LinkedFactTypeSNKey Is Not Null"
   moRS.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   ' Got any Linked Facts ?
   
   If moRS.RecordCount > 0 Then
      Do While Not moRS.EOF
      
         Set loRs = FindLinkedFact(EventType, EventDetailKey, ZeroNull(moRS!LinkedFactTypeSNKey), ZeroNull(moRS!LinkedFactLogSNKey))
         
         If loRs.RecordCount > 0 Then
         
            moRS!StartValue = loRs!StartValue
            moRS!EndValue = loRs!EndValue
            moRS!FactValue = loRs!FactValue
            moRS!RemarksValue = loRs!RemarksValue
            moRS!StartValueDate = loRs!StartValueDate
            moRS!EndValueDate = loRs!EndValueDate
            
            moRS!LinkedFactLogSNKey = loRs!ID
            moRS.Update
            
         ElseIf ZeroNull(moRS!LinkedFactLogSNKey) > 0 Then
            moRS!StartValue = Null
            moRS!EndValue = Null
            moRS!FactValue = Null
            moRS!RemarksValue = Null
            moRS!StartValueDate = Null
            moRS!EndValueDate = Null
            
            moRS!LinkedFactLogSNKey = Null
            moRS.Update
         End If
         goSession.CloseRecordset loRs
         
         moRS.MoveNext
      Loop
   End If
   goSession.CloseRecordset moRS
   Exit Sub
ErrorHandler:
   goSession.RaiseError "General Error in mwEventFactWorkSN.UpdateLinkedFacts. ", Err.Number, Err.Description
End Sub

Private Function FindLinkedFact(EventType As Long, EventDetailKey As Long, LinkedFactTypeSNKey As Long, LinkedFactLogSNKey As Long) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo ErrorHandler
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   If LinkedFactTypeSNKey > 0 And LinkedFactLogSNKey > 0 Then
      sSQL = "select * from mwEventFactLogSN " & _
             " where ID=" & LinkedFactLogSNKey & _
             " AND mwEventFactTypeSNKey=" & LinkedFactTypeSNKey
             
      loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
      
      If loRs.RecordCount > 0 Then
         Set FindLinkedFact = loRs
         Exit Function
      End If
      
      goSession.CloseRecordset loRs
      Set loRs = New Recordset
      loRs.CursorLocation = adUseClient
      
   End If
   
   If EventType = MW_EVENT_TYPE_CARGO6 Then
      sSQL = "select * from mwEventFactLogSN " & _
             " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
             " AND (" & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_BERTH6 & " AND mwEventDetailKey=" & _
               "(SELECT BerthDetailKey FROM vrsCargoDetail WHERE ID = " & EventDetailKey & ")" & _
             " ) OR " & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_PORT_CALL6 & " AND mwEventDetailKey=" & _
               "(SELECT PortCallKey FROM vrsBerthDetail WHERE ID = " & _
                  "(SELECT BerthDetailKey FROM vrsCargoDetail WHERE ID = " & EventDetailKey & "))" & _
             " ) OR " & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_VOYAGE_PLAN & " AND mwEventDetailKey=" & _
               "(SELECT vrsVoyageKey FROM vrsPortCall WHERE ID = " & _
                  "(SELECT PortCallKey FROM vrsBerthDetail WHERE ID = " & _
                     "(SELECT BerthDetailKey FROM vrsCargoDetail WHERE ID = " & EventDetailKey & ")))" & _
             " )) ORDER BY mwEventTypeKey"
   ElseIf EventType = MW_EVENT_TYPE_BERTH6 Then
      sSQL = "select * from mwEventFactLogSN " & _
             " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
             " AND (" & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_PORT_CALL6 & " AND mwEventDetailKey=" & _
               "(SELECT PortCallKey FROM vrsBerthDetail WHERE ID = " & EventDetailKey & ")" & _
             " ) OR " & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_VOYAGE_PLAN & " AND mwEventDetailKey=" & _
               "(SELECT vrsVoyageKey FROM vrsPortCall WHERE ID = " & _
                  "(SELECT PortCallKey FROM vrsBerthDetail WHERE ID = " & EventDetailKey & "))" & _
             " )) ORDER BY mwEventTypeKey"
   ElseIf EventType = MW_EVENT_TYPE_PORT_CALL6 Then
      sSQL = "select * from mwEventFactLogSN " & _
             " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
             " AND (" & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_VOYAGE_PLAN & " AND mwEventDetailKey=" & _
               "(SELECT vrsVoyageKey FROM vrsPortCall WHERE ID = " & EventDetailKey & ")" & _
             " )) ORDER BY mwEventTypeKey"
   ElseIf EventType = MW_EVENT_TYPE_POSASS6 Then
      sSQL = "select * from mwEventFactLogSN " & _
             " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
             " AND (" & _
             " (mwEventTypeKey=" & MW_EVENT_TYPE_VOYAGE_PLAN & " AND mwEventDetailKey=" & _
               "(SELECT vrsVoyageKey FROM vrsPosAss WHERE ID = " & EventDetailKey & ")" & _
             " )) ORDER BY mwEventTypeKey"
   Else
      If GetEventTableName(EventType) = UCase("vrsVoyageEvent") Then
         sSQL = "select * from mwEventFactLogSN " & _
                " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
                " AND (" & _
                " (mwEventTypeKey=" & MW_EVENT_TYPE_VOYAGE_PLAN & " AND mwEventDetailKey=" & _
                  "(SELECT vrsVoyageKey FROM vrsVoyageEvent WHERE ID = " & EventDetailKey & ")" & _
                " )) ORDER BY mwEventTypeKey"
      Else
         sSQL = "select * from mwEventFactLogSN " & _
                " WHERE mwEventFactTypeSNKey=" & LinkedFactTypeSNKey & _
                " AND mwEventDetailKey=" & EventDetailKey & _
                " ORDER BY mwEventTypeKey"
      End If
   End If
   
   
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic

   Set FindLinkedFact = loRs
   
   Exit Function
ErrorHandler:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FindLinkedFact. ", Err.Number, Err.Description
End Function

Private Function GetEventTableName(EventType As Long) As String
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo ErrorHandler
   
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "select TableName from mwEventType where ID=" & EventType
             
   loRs.Open sSQL, goCon, adOpenDynamic, adLockOptimistic
   
   If loRs.RecordCount > 0 Then
      GetEventTableName = UCase(BlankNull(loRs!TableName))
   End If
   
   goSession.CloseRecordset loRs
   Exit Function
ErrorHandler:
   goSession.RaiseError "General Error in mwEventFactWorkSN.GetEventTableName. ", Err.Number, Err.Description
End Function

Private Sub moRs_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   Static loWork As Object
   On Error GoTo SubError
   
   If mSiteKey > 0 Then
      If loWork Is Nothing Then
         Set loWork = CreateObject("mwSession.mwReplicateWillChange")
         
         If Not loWork.Initialize("mwEventFactLogSN") Then
            Set loWork = Nothing
            Exit Sub
         End If
      End If
      loWork.WillChangeRecord adReason, cRecords, adStatus, pRecordset
   End If
   
   Exit Sub
SubError:
   goSession.RaisePublicError "General Error in mwSession.mwEventFactWorkSN.moRs_WillChangeRecord. ", Err.Number, Err.Description
End Sub



Public Function IsFactCatRestricted(mwEventFactCatSNkey As Long, ByVal mwcRoleTypekey As Long, ByRef IsRoleRestricted As Boolean, ByRef IsViewAllowed As Boolean, ByRef IsModifyAllowed As Boolean) As Boolean
   'NSB-118 Changes to Port Call Event
   'Added By N.Angelakis 1 November 2010
   
   Dim sSQL As String
   Dim loRs As Recordset
   
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT    id, mwcRoleTypekey, IsViewAllowed, IsModifyAllowed " & _
            " FROM   mwEventFactCatSNrole " & _
            " WHERE  mwEventFactCatSNkey = " & mwEventFactCatSNkey & _
            " AND    mwcRoleTypekey > 0"
    
   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Do While Not loRs.EOF
      'if rows were returned it indicates that fact cat is role restricted
      IsFactCatRestricted = True
      IsRoleRestricted = True
      
      If mwcRoleTypekey = ZeroNull(loRs!mwcRoleTypekey) Then
         'role is assigned to fact category so get permissions
         IsRoleRestricted = False 'role was found in list of excluded
         'get permissions for role
         IsViewAllowed = BoolNull(loRs!IsViewAllowed)
         IsModifyAllowed = BoolNull(loRs!IsModifyAllowed)
         Exit Do
      End If

      loRs.MoveNext
   Loop
   CloseRecordset loRs
   
   
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.IsFactCatRestricted. ", Err.Number, Err.Description
   CloseRecordset loRs
   IsFactCatRestricted = True
End Function
Public Function FetchConsolidatedShipnetActivityRS(EventType As Long, EventDetailKey As Long, _
    Optional EventType2 As Long, Optional EventDetailKey2 As Long, _
    Optional EventType3 As Long, Optional EventDetailKey3 As Long, _
    Optional EventTypeSubCode As String, Optional IsMatchFactDispSeq As Boolean) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND (" & _
      " (mwEventFactLogSN.mwEventTypeKey  =" & EventType & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & ")"
   If EventType2 > 0 And EventDetailKey2 > 0 Then
      sSQL = sSQL & " OR (mwEventFactLogSN.mwEventTypeKey  =" & EventType2 & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey2 & ")"
   End If
   If EventType3 > 0 And EventDetailKey3 > 0 Then
      sSQL = sSQL & " OR (mwEventFactLogSN.mwEventTypeKey  =" & EventType3 & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey3 & ")"
   End If
   sSQL = sSQL & " ) AND not mwEventFactTypeSN.SN_ActivityTypeCode is Null AND " & _
      " mwEventFactTypeSN.SN_ActivityTypeCode <> 'L' "
   If Trim(EventTypeSubCode) <> "" Then
      sSQL = sSQL & " AND mwEventFactTypeSN.EventTypeSubcode ='" & EventTypeSubCode & "'"
   End If
   If IsMatchFactDispSeq Then
      sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   Else
      sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   End If

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchConsolidatedShipnetActivityRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchConsolidatedShipnetActivityRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

Public Function FetchConsolidatedShipnetLaytimeRS(EventType As Long, EventDetailKey As Long, _
    Optional EventType2 As Long, Optional EventDetailKey2 As Long, _
    Optional EventType3 As Long, Optional EventDetailKey3 As Long, _
    Optional EventTypeSubCode As String, Optional IsMatchFactDispSeq As Boolean) As Recordset
   Dim sSQL As String
   Dim loRs As Recordset
   On Error GoTo FunctionError
   Set loRs = New Recordset
   loRs.CursorLocation = adUseClient
   
   sSQL = "SELECT mwEventFactLogSN.*, mwEventFactTypeSN.* " & _
      " FROM mwEventFactTypeSN , mwEventFactLogSN " & _
      " WHERE mwEventFactTypeSN.ID = mwEventFactLogSN.mwEventFactTypeSNKey AND (" & _
      " (mwEventFactLogSN.mwEventTypeKey  =" & EventType & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey & ")"
   If EventType2 > 0 And EventDetailKey2 > 0 Then
      sSQL = sSQL & " OR (mwEventFactLogSN.mwEventTypeKey  =" & EventType2 & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey2 & ")"
   End If
   If EventType3 > 0 And EventDetailKey3 > 0 Then
      sSQL = sSQL & " OR (mwEventFactLogSN.mwEventTypeKey  =" & EventType3 & " AND mwEventFactLogSN.mwEventDetailKey = " & EventDetailKey3 & ")"
   End If
   sSQL = sSQL & " ) AND (mwEventFactTypeSN.SN_ActivityTypeCode = 'L' OR mwEventFactTypeSN.SN_ActivityTypeCode = 'X') "
   If Trim(EventTypeSubCode) <> "" Then
      sSQL = sSQL & " AND mwEventFactTypeSN.EventTypeSubcode ='" & EventTypeSubCode & "'"
   End If
   If IsMatchFactDispSeq Then
      sSQL = sSQL & " ORDER BY mwEventFactTypeSN.DisplaySequence,mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   Else
      sSQL = sSQL & " ORDER BY mwEventFactLogSN.StartValue, mwEventFactLogSN.ID"
   End If

   loRs.Open sSQL, goCon, adOpenForwardOnly, adLockReadOnly
   Set FetchConsolidatedShipnetLaytimeRS = loRs
   'Set loRS = Nothing
   Exit Function
FunctionError:
   goSession.RaiseError "General Error in mwEventFactWorkSN.FetchConsolidatedShipnetLaytimeRS. ", Err.Number, Err.Description
   CloseRecordset loRs
End Function

